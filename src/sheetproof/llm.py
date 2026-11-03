"""M4 LLM 보완: 규칙 기반으로 변환하지 못한 수식만 LLM 에게 변환을 맡긴다.

원칙
- 사용자가 본인 API Key 를 입력한 경우에만 호출한다.
- 수식 문자열과 참조 셀의 "주소·타입"만 보낸다. 셀 값(데이터)은 보내지 않는다.
- LLM 결과는 AST 화이트리스트 검사를 통과해야 하고, 같은 검증 엔진에서
  엑셀 저장값과 일치해야만 채택한다.
"""

from __future__ import annotations

import ast
import json
import re
import types
from dataclasses import dataclass, field

from . import xl_runtime
from .analysis import Analysis
from .codegen import generate
from .verify import verify

PROVIDERS = {
    "anthropic": {"label": "Anthropic", "model": "claude-opus-5-5"},
    "openai": {"label": "OpenAI", "model": "gpt-4.1"},
    "openai_compatible": {"label": "OpenAI 호환", "model": ""},
}

ALLOWED_BUILTINS = {"abs", "min", "max", "len", "int", "float", "str", "round", "sum", "range",
                    "isinstance", "list", "bool", "True", "False", "None"}
SAFE_STR_METHODS = {"title", "capitalize", "upper", "lower", "strip", "lstrip", "rstrip", "replace",
                    "split", "join", "startswith", "endswith", "find", "count", "zfill", "ljust",
                    "rjust", "center", "isdigit", "isalpha", "swapcase"}
RUNTIME_PUBLIC = {
    n for n in dir(xl_runtime)
    if not n.startswith("_") and not isinstance(getattr(xl_runtime, n), types.ModuleType)
    and n not in {"Decimal", "InvalidOperation", "ROUND_DOWN", "ROUND_HALF_UP", "ROUND_UP",
                  "annotations", "FUNCTIONS", "PY_NAMES", "RANGE_ARG_FUNCTIONS"}
}

SYSTEM_PROMPT = """You convert one Excel formula into a single Python expression that runs on a
small Excel-semantics runtime called `xl`. Reply with JSON only: {"expr": "<python expression>",
"note": "<short Korean explanation>"}.

Rules for the expression:
- Read a single cell with g("Sheet!A1") (returns None when blank).
- Read a range with xl.rng(v, "Sheet!A1:B3") (returns a 2-D list of rows).
- Use xl functions when they exist (xl.SUM, xl.IF, xl.VLOOKUP, xl.add, xl.mul, xl.to_number,
  xl.to_text, xl.to_bool, xl.NA, xl.VALUE, xl.DIV0, ...). Excel errors are values, not exceptions.
- Plain Python is allowed: arithmetic, comparisons, conditional expressions, list/generator
  comprehensions, and builtins abs/min/max/len/int/float/str/round/sum/range/isinstance/list/bool.
- No imports, no lambdas, no dunder names, no statements. Attribute access is allowed only on `xl`
  and for plain str methods (title, upper, lower, strip, replace, split, join, ...).
- Reproduce Excel semantics exactly (blank cells, error propagation, rounding half away from zero).
"""


class LLMError(RuntimeError):
    pass


@dataclass
class LLMCandidate:
    key: str
    formula: str
    expr: str | None = None
    note: str = ""
    accepted: bool = False
    reason: str = ""


@dataclass
class LLMResult:
    candidates: list[LLMCandidate] = field(default_factory=list)
    overrides: dict[str, str] = field(default_factory=dict)
    message: str = ""


def validate_expr(expr: str) -> str | None:
    """안전하지 않은 식이면 사유를, 괜찮으면 None 을 돌려준다."""
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        return f"구문 오류: {e.msg}"
    allowed_nodes = (
        ast.Expression, ast.Call, ast.Name, ast.Load, ast.Store, ast.Constant, ast.Attribute,
        ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare, ast.IfExp, ast.Tuple, ast.List,
        ast.Subscript, ast.Slice, ast.ListComp, ast.GeneratorExp, ast.comprehension,
        ast.operator, ast.unaryop, ast.boolop, ast.cmpop, ast.keyword, ast.Starred,
    )
    comp_targets = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.comprehension):
            for t in ast.walk(node.target):
                if isinstance(t, ast.Name):
                    comp_targets.add(t.id)
    for node in ast.walk(tree):
        if not isinstance(node, allowed_nodes):
            return f"허용되지 않는 구문: {type(node).__name__}"
        if isinstance(node, ast.Attribute):
            if node.attr.startswith("_"):
                return "밑줄로 시작하는 속성은 허용되지 않습니다"
            if not (isinstance(node.value, ast.Name) and node.value.id == "xl"):
                if node.attr in SAFE_STR_METHODS:
                    continue
                return f"허용되지 않는 속성 접근: .{node.attr}"
            if node.attr not in RUNTIME_PUBLIC:
                return f"런타임에 없는 이름: xl.{node.attr}"
        if isinstance(node, ast.Name):
            if node.id.startswith("__"):
                return "dunder 이름은 허용되지 않습니다"
            if node.id not in ALLOWED_BUILTINS | {"g", "v", "xl"} | comp_targets:
                return f"허용되지 않는 이름: {node.id}"
    return None


def _describe_refs(analysis: Analysis, key: str) -> list[str]:
    fi = analysis.formulas[key]
    lines = []
    for r in fi.refs[:40]:
        if r.is_single:
            c = analysis.model.cells.get(r.key)
            if c is None:
                t = "blank"
            elif c.is_formula:
                t = "formula"
            elif c.is_date:
                t = "date serial"
            else:
                t = type(c.value).__name__
            lines.append(f"- {r.key}: {t}")
        else:
            lines.append(f"- {r.key}: range {r.max_row - r.min_row + 1}x{r.max_col - r.min_col + 1}")
    return lines


def build_prompt(analysis: Analysis, key: str) -> str:
    fi = analysis.formulas[key]
    refs = "\n".join(_describe_refs(analysis, key)) or "- (none)"
    return (
        f"Cell: {key} (current sheet: {fi.cell.sheet})\n"
        f"Formula: {fi.cell.formula}\n"
        f"Unsupported functions to translate: {', '.join(sorted(fi.unsupported)) or '-'}\n"
        f"Referenced cells (address: type, values are not shared):\n{refs}\n"
        "Always use fully qualified addresses like g(\"Sheet!A1\")."
    )


def _parse_reply(text: str) -> tuple[str, str]:
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        raise LLMError("JSON 응답을 찾을 수 없습니다")
    data = json.loads(m.group(0))
    expr = str(data.get("expr", "")).strip()
    if not expr:
        raise LLMError("expr 가 비어 있습니다")
    return expr, str(data.get("note", ""))


def make_caller(provider: str, api_key: str, model: str | None = None, base_url: str | None = None,
                client_factory=None):
    """(system, user) → 응답 텍스트 함수를 만든다. ``client_factory`` 는 테스트용 가짜 SDK 주입."""
    model = model or PROVIDERS.get(provider, {}).get("model")
    if not model:
        raise LLMError("모델 이름을 입력해 주세요")
    if provider == "anthropic":
        if client_factory is None:
            import anthropic

            client_factory = anthropic.Anthropic
        client = client_factory(api_key=api_key)

        def call(system: str, user: str) -> str:
            resp = client.messages.create(
                model=model,
                max_tokens=16000,
                system=system,
                messages=[{"role": "user", "content": user}],
            )
            if getattr(resp, "stop_reason", None) == "refusal":
                raise LLMError("모델이 요청을 거절했습니다")
            return "".join(getattr(b, "text", "") for b in resp.content
                           if getattr(b, "type", "") == "text")

        return call
    if provider in ("openai", "openai_compatible"):
        if client_factory is None:
            import openai

            client_factory = openai.OpenAI
        kwargs = {"api_key": api_key}
        if provider == "openai_compatible":
            if not base_url:
                raise LLMError("OpenAI 호환 공급자는 Base URL 이 필요합니다")
            kwargs["base_url"] = base_url
        client = client_factory(**kwargs)

        def call(system: str, user: str) -> str:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            )
            return resp.choices[0].message.content or ""

        return call
    raise LLMError(f"알 수 없는 공급자: {provider}")


def llm_candidates(analysis: Analysis) -> list[str]:
    """LLM 보완 대상: 미지원 함수 때문에 고정된 셀 (휘발성/순환/외부 참조는 제외)."""
    return [k for k, f in analysis.formulas.items()
            if f.unsupported and not f.volatile and not f.issues and f.ast is not None]


def complete_with_llm(analysis: Analysis, call, targets=None, rel_tol=1e-9, abs_tol=1e-9,
                      max_cells: int = 30, existing: dict[str, str] | None = None) -> LLMResult:
    """미지원 수식을 LLM 으로 변환하고, 검증을 통과한 것만 채택한다."""
    result = LLMResult()
    keys = llm_candidates(analysis)
    if targets:
        scope = analysis.ancestors(targets)
        keys = [k for k in keys if k in scope]
    if not keys:
        result.message = "LLM 보완 대상(미지원 함수 수식)이 없습니다."
        return result
    overrides = dict(existing or {})
    for k in keys[:max_cells]:
        cand = LLMCandidate(k, analysis.formulas[k].cell.formula)
        result.candidates.append(cand)
        try:
            cand.expr, cand.note = _parse_reply(call(SYSTEM_PROMPT, build_prompt(analysis, k)))
        except Exception as e:  # noqa: BLE001 - 공급자 오류는 사용자에게 그대로 보여 준다
            cand.reason = f"LLM 호출 실패: {e}"
            continue
        problem = validate_expr(cand.expr)
        if problem:
            cand.reason = f"안전 검사 실패: {problem}"
            continue
        trial = dict(overrides)
        trial[k] = cand.expr
        gen = generate(analysis, targets, trial)
        rep = verify(analysis, gen, rel_tol, abs_tol)
        r = rep.results.get(k)
        if r is not None and r.status == "match":
            cand.accepted = True
            cand.reason = "검증 통과 — 채택"
            overrides[k] = cand.expr
        else:
            cand.reason = "검증 실패 — 미채택" + (f" ({r.cause or '값 불일치'})" if r else "")
    if len(keys) > max_cells:
        result.message = f"대상 {len(keys)}개 중 {max_cells}개만 처리했습니다."
    result.overrides = overrides
    accepted = sum(1 for c in result.candidates if c.accepted)
    result.message = (result.message + " " if result.message else "") + \
        f"LLM 변환 {len(result.candidates)}개 중 {accepted}개가 검증을 통과해 채택되었습니다."
    return result
