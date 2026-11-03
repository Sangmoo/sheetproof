"""M2 변환: AST → 파이썬 코드.

생성되는 모듈 구조::

    import xl_runtime as xl

    INPUTS = {...}          # 입력 셀 (엑셀 상수) 기본값
    FIXED = {...}           # 변환하지 않고 저장값으로 고정한 수식 셀
    FORMULAS = {...}        # 셀 → 원본 수식 (문서용)

    def compute(inputs=None):
        v = {**INPUTS, **FIXED}
        ...
        v["Sheet1!C2"] = xl.normalize(xl.mul(g("Sheet1!A2"), g("Sheet1!B2")))
        return v

검증 엔진은 이 모듈 소스를 그대로 실행하므로, 내보낸 코드 = 검증된 코드이다.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field

from openpyxl.utils.cell import get_column_letter

from . import parser as P
from .analysis import Analysis, ResolvedRef
from .workbook import resolve_bounds, resolve_name
from .xl_runtime import PY_NAMES, RANGE_ARG_FUNCTIONS, XLError

BINARY_FUNCS = {
    "+": "add", "-": "sub", "*": "mul", "/": "div", "^": "pow", "&": "concat",
    "=": "eq", "<>": "ne", "<": "lt", "<=": "le", ">": "gt", ">=": "ge",
}
UNARY_FUNCS = {"-": "neg", "+": "pos"}
ERROR_CONSTS = {
    "#NULL!": "xl.NULL", "#DIV/0!": "xl.DIV0", "#VALUE!": "xl.VALUE", "#REF!": "xl.REF",
    "#NAME?": "xl.NAME", "#NUM!": "xl.NUM", "#N/A": "xl.NA",
}


class Unsupported(Exception):
    pass


def py_value(v) -> str:
    """런타임 값 → 파이썬 소스 리터럴."""
    if isinstance(v, XLError):
        return ERROR_CONSTS.get(v.code, f"xl.err({v.code!r})")
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return "xl.NUM"
        return repr(v)
    if isinstance(v, (bool, int, str)) or v is None:
        return repr(v)
    if isinstance(v, (_dt.date, _dt.datetime)):
        return repr(str(v))
    return repr(v)


class _ExprGen:
    def __init__(self, analysis: Analysis, sheet: str):
        self.a = analysis
        self.model = analysis.model
        self.sheet = sheet
        self._names: list[str] = []

    def ref_expr(self, r: ResolvedRef, as_range: bool) -> str:
        if r.is_single and not as_range:
            return f"g({r.key!r})"
        return f"xl.rng(v, {r.key!r})"

    def resolve_ref(self, node: P.Ref, sheet: str) -> ResolvedRef:
        if node.external:
            raise Unsupported("외부 참조")
        target = node.sheet if node.sheet is not None else sheet
        if target not in self.model.sheet_dims:
            return None
        b = resolve_bounds(self.model, target, node.address)
        return ResolvedRef(target, *b)

    def gen(self, node, sheet: str | None = None, as_range: bool = False) -> str:
        sheet = sheet or self.sheet
        if isinstance(node, P.Number):
            return repr(node.value)
        if isinstance(node, P.String):
            return repr(node.value)
        if isinstance(node, P.Bool):
            return "True" if node.value else "False"
        if isinstance(node, P.Error):
            return ERROR_CONSTS.get(node.code, f"xl.err({node.code!r})")
        if isinstance(node, P.Missing):
            return "xl.MISSING"
        if isinstance(node, P.Ref):
            r = self.resolve_ref(node, sheet)
            if r is None:
                return "xl.REF"
            return self.ref_expr(r, as_range)
        if isinstance(node, P.Name):
            dn = resolve_name(self.model, node.name, sheet)
            if dn is None or dn.ast is None:
                return "xl.NAME"
            if node.name.upper() in self._names:
                raise Unsupported("이름 순환 정의")
            self._names.append(node.name.upper())
            try:
                return self.gen(dn.ast, dn.scope or sheet, as_range)
            finally:
                self._names.pop()
        if isinstance(node, P.Unary):
            return f"xl.{UNARY_FUNCS[node.op]}({self.gen(node.operand, sheet)})"
        if isinstance(node, P.Postfix):
            if node.op != "%":
                raise Unsupported(f"후위 연산자 {node.op}")
            return f"xl.percent({self.gen(node.operand, sheet)})"
        if isinstance(node, P.Binary):
            fn = BINARY_FUNCS.get(node.op)
            if fn is None:
                raise Unsupported(f"연산자 {node.op}")
            return f"xl.{fn}({self.gen(node.left, sheet)}, {self.gen(node.right, sheet)})"
        if isinstance(node, P.Call):
            from .xl_runtime import FUNCTIONS

            if node.name not in FUNCTIONS:
                raise Unsupported(f"함수 {node.name}")
            wrap = node.name in RANGE_ARG_FUNCTIONS
            args = ", ".join(self.gen(x, sheet, as_range=wrap) for x in node.args)
            return f"xl.{PY_NAMES.get(node.name, node.name)}({args})"
        if isinstance(node, P.Array):
            rows = ", ".join("[" + ", ".join(self.gen(x, sheet) for x in r) + "]" for r in node.rows)
            return f"[{rows}]"
        raise Unsupported(type(node).__name__)


@dataclass
class CellCode:
    key: str
    formula: str
    expr: str | None  # None 이면 고정값
    status: str  # converted / fixed / llm
    reason: str = ""


@dataclass
class GeneratedModule:
    source: str
    module_name: str
    cells: dict[str, CellCode]
    inputs: dict[str, object]
    fixed: dict[str, object]
    order: list[str]
    targets: list[str] = field(default_factory=list)

    @property
    def converted(self) -> list[str]:
        return [k for k, c in self.cells.items() if c.status in ("converted", "llm")]


def _ident(s: str) -> str:
    s = re.sub(r"\W+", "_", s.lower()).strip("_")
    if not s or s[0].isdigit():
        s = "m_" + s
    return s


def module_name_for(filename: str) -> str:
    stem = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1].rsplit(".", 1)[0]
    return _ident(stem) + "_model"


def parse_targets(analysis: Analysis, text: str) -> list[str]:
    """``"Sheet1!D2:D10, Summary!B3"`` → 수식 셀 키 목록."""
    keys: list[str] = []
    if not text or not text.strip():
        return keys
    default_sheet = analysis.model.sheets[0]
    for part in re.split(r"[,;\n]+", text):
        part = part.strip()
        if not part:
            continue
        sheet, address, _ = P.split_reference(part)
        sheet = sheet or default_sheet
        if sheet not in analysis.model.sheet_dims:
            raise ValueError(f"시트 '{sheet}' 가 없습니다")
        c1, r1, c2, r2 = resolve_bounds(analysis.model, sheet, address.upper())
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                k = f"{sheet}!{get_column_letter(c)}{r}"
                if k in analysis.formulas:
                    keys.append(k)
    return keys


def generate(analysis: Analysis, targets: list[str] | None = None,
             overrides: dict[str, str] | None = None) -> GeneratedModule:
    """분석 결과로 파이썬 모듈 소스를 만든다.

    targets: 생성할 출력 셀 (None 이면 전체). 지정하면 그 셀들의 상위 셀만 포함한다.
    overrides: 셀 → 파이썬 식 (LLM 변환 결과 등). ``xl.guard`` 로 감싸 실행한다.
    """
    overrides = overrides or {}
    model = analysis.model
    scope = analysis.ancestors(targets) if targets else None

    cells: dict[str, CellCode] = {}
    fixed: dict[str, object] = {}
    order = [k for k in analysis.order if scope is None or k in scope]
    cyclic = [k for k in analysis.formulas if k not in set(analysis.order)
              and (scope is None or k in scope)]

    for k in order + cyclic:
        fi = analysis.formulas[k]
        formula = fi.cell.formula
        if k in overrides:
            cells[k] = CellCode(k, formula, f"xl.guard(lambda: {overrides[k]})", "llm", "LLM 변환")
            continue
        reason = ""
        if fi.issues:
            reason = ", ".join(dict.fromkeys(fi.issues))
        elif fi.volatile:
            reason = "휘발성 함수 " + ", ".join(sorted(fi.volatile))
        elif fi.unsupported:
            reason = "미지원 함수 " + ", ".join(sorted(fi.unsupported))
        expr = None
        if not reason:
            try:
                expr = _ExprGen(analysis, fi.cell.sheet).gen(fi.ast)
            except Unsupported as e:
                reason = f"변환 불가 ({e})"
            except (ValueError, TypeError) as e:
                reason = f"변환 불가 ({e})"
        if expr is None:
            cells[k] = CellCode(k, formula, None, "fixed", reason)
            fixed[k] = fi.cell.cached
        else:
            cells[k] = CellCode(k, formula, expr, "converted")

    needed = scope if scope is not None else None
    inputs = {}
    for k, cell in model.cells.items():
        if cell.is_formula:
            continue
        if analysis.kinds.get(k) != "input":
            continue
        if needed is not None and k not in needed:
            continue
        inputs[k] = cell.value

    src = _render(model.filename, inputs, fixed, cells, order, cyclic, analysis)
    return GeneratedModule(src, module_name_for(model.filename), cells, inputs, fixed,
                           order, list(targets or []))


def _dict_literal(name: str, d: dict, comment: str) -> list[str]:
    lines = [f"# {comment}", f"{name} = {{"]
    for k, val in d.items():
        lines.append(f"    {k!r}: {py_value(val)},")
    lines.append("}")
    return lines


def _render(filename, inputs, fixed, cells, order, cyclic, analysis) -> str:
    n_conv = sum(1 for c in cells.values() if c.status != "fixed")
    lines = [
        f'"""sheetproof 가 {filename!s} 에서 생성한 계산 모델.',
        "",
        f"변환 수식 {n_conv}개 / 고정값 {len(fixed)}개.",
        "compute(inputs) 에 입력 셀 값을 넘기면 모든 수식 셀 값을 계산해 dict 로 돌려준다.",
        "키 형식은 '시트!A1' 이다. 런타임 헬퍼 xl_runtime.py 가 같은 폴더에 있어야 한다.",
        '"""',
        "",
        "import xl_runtime as xl",
        "",
    ]
    lines += _dict_literal("INPUTS", inputs, "입력 셀 (엑셀 상수) 기본값")
    lines.append("")
    fixed_comments = {k: cells[k].reason for k in fixed}
    lines.append("# 변환하지 않고 엑셀 저장값으로 고정한 수식 셀 (사유는 주석 참고)")
    lines.append("FIXED = {")
    for k, val in fixed.items():
        lines.append(f"    {k!r}: {py_value(val)},  # {fixed_comments[k]}")
    lines.append("}")
    lines.append("")
    lines.append("# 셀 → 원본 엑셀 수식")
    lines.append("FORMULAS = {")
    for k, c in cells.items():
        lines.append(f"    {k!r}: {c.formula!r},")
    lines.append("}")
    lines.append("")
    outputs = [k for k in order if analysis.kinds.get(k) == "output"]
    lines.append(f"OUTPUTS = {outputs!r}")
    lines.append("")
    lines.append("")
    lines.append("def compute(inputs=None):")
    lines.append('    """입력 셀 값(dict)을 받아 모든 셀 값을 계산한다."""')
    lines.append("    v = dict(INPUTS)")
    lines.append("    v.update(FIXED)")
    lines.append("    if inputs:")
    lines.append("        v.update(inputs)")
    lines.append("    g = v.get")
    for k in order:
        c = cells[k]
        if c.status == "fixed":
            continue
        lines.append(f"    # {k}: {_comment(c.formula)}")
        lines.append(f"    v[{k!r}] = xl.normalize({c.expr})")
    if not any(cells[k].status != "fixed" for k in order):
        lines.append("    pass")
    lines.append("    return v")
    lines.append("")
    return "\n".join(lines)


def _comment(formula: str) -> str:
    return formula.replace("\n", " ")[:200]
