"""M3 검증: 생성 코드를 실제로 실행해 엑셀 저장값과 셀 단위로 대조한다."""

from __future__ import annotations

import math
import sys
import types
from dataclasses import dataclass, field

from . import xl_runtime
from .analysis import Analysis
from .codegen import GeneratedModule
from .xl_runtime import XLError, values_equal

STATUS_LABELS = {
    "match": "일치",
    "mismatch": "불일치",
    "fixed": "고정값(검증 제외)",
    "no_cache": "저장값 없음",
}


@dataclass
class CellResult:
    key: str
    formula: str
    status: str
    expected: object = None
    actual: object = None
    cause: str = ""
    root: bool = False  # 불일치의 근본 원인 셀 여부


@dataclass
class VerifyReport:
    results: dict[str, CellResult]
    rel_tol: float
    abs_tol: float
    error: str | None = None  # 모듈 실행 자체가 실패한 경우
    values: dict = field(default_factory=dict)

    def count(self, status: str) -> int:
        return sum(1 for r in self.results.values() if r.status == status)

    @property
    def checked(self) -> int:
        return self.count("match") + self.count("mismatch")

    @property
    def match_rate(self) -> float:
        return self.count("match") / self.checked if self.checked else 0.0

    @property
    def mismatches(self) -> list[CellResult]:
        rs = [r for r in self.results.values() if r.status == "mismatch"]
        return sorted(rs, key=lambda r: (not r.root, r.key))

    @property
    def passed(self) -> bool:
        return self.error is None and self.checked > 0 and self.count("mismatch") == 0


def load_module(source: str, name: str = "sheetproof_generated"):
    """생성 소스를 모듈로 실행한다. ``import xl_runtime`` 은 패키지 런타임으로 연결."""
    mod = types.ModuleType(name)
    prev = sys.modules.get("xl_runtime")
    sys.modules["xl_runtime"] = xl_runtime
    try:
        exec(compile(source, f"<{name}>", "exec"), mod.__dict__)  # noqa: S102 - 직접 생성한 코드
    finally:
        if prev is None:
            sys.modules.pop("xl_runtime", None)
        else:
            sys.modules["xl_runtime"] = prev
    return mod


def fmt_value(v) -> str:
    if v is None:
        return "(빈 값)"
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return repr(v) if abs(v) < 1e16 else f"{v:.15g}"
    if isinstance(v, str):
        return repr(v)
    if isinstance(v, list):
        return f"배열 {len(v)}x{len(v[0]) if v else 0}"
    return str(v)


def _guess_cause(analysis: Analysis, key: str, expected, actual, bad_parents: list[str]) -> str:
    fi = analysis.formulas.get(key)
    if bad_parents:
        return "상위 셀 불일치 전파: " + ", ".join(bad_parents[:3]) + (" …" if len(bad_parents) > 3 else "")
    if isinstance(actual, list):
        return "결과가 배열입니다 (동적 배열/암시적 교차는 미지원)"
    if isinstance(expected, XLError) and not isinstance(actual, XLError):
        return f"엑셀은 오류값 {expected.code} — 오류 발생 조건 처리 차이"
    if isinstance(actual, XLError) and not isinstance(expected, XLError):
        return f"생성 코드가 오류값 {actual.code} 반환 — 함수 의미론/타입 변환 확인 필요"
    if isinstance(expected, XLError) and isinstance(actual, XLError):
        return "오류값 종류 차이"
    num = (int, float)
    if isinstance(expected, num) and isinstance(actual, num) and not isinstance(expected, bool) \
            and not isinstance(actual, bool):
        diff = abs(expected - actual)
        scale = max(abs(expected), abs(actual), 1e-300)
        if diff / scale < 1e-6:
            return f"미세한 수치 차이 ({diff:.3g}) — 부동소수/반올림, 허용 오차 조정 고려"
        if fi and fi.cell.is_date:
            return "날짜 계산 차이 (1900 체계/월말 처리 확인)"
        if fi and {"ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "TEXT"} & fi.functions:
            return "반올림/절사 처리 차이"
        return "계산 로직 차이 (함수 의미론 확인 필요)"
    if type(expected) is not type(actual):
        return f"값 타입 차이 (엑셀 {type(expected).__name__} / 코드 {type(actual).__name__})"
    if isinstance(expected, str):
        if expected.strip() == str(actual).strip():
            return "공백 처리 차이"
        if expected.lower() == str(actual).lower():
            return "대소문자 차이"
        return "텍스트 결과 차이 (TEXT 서식/문자열 함수 확인)"
    return "값 차이"


def verify(analysis: Analysis, gen: GeneratedModule, rel_tol: float = 1e-9,
           abs_tol: float = 1e-9, source: str | None = None) -> VerifyReport:
    """생성 모듈을 실행하고 수식 셀마다 엑셀 저장값과 비교한다.

    ``source`` 를 주면 그 소스를 실행한다 (테스트에서 일부러 틀린 코드를 넣을 때 사용).
    """
    src = source if source is not None else gen.source
    results: dict[str, CellResult] = {}
    try:
        mod = load_module(src)
        values = mod.compute()
    except Exception as e:  # noqa: BLE001
        for k, c in gen.cells.items():
            results[k] = CellResult(k, c.formula, "mismatch", analysis.formulas[k].cell.cached,
                                    None, f"생성 코드 실행 실패: {e}", True)
        return VerifyReport(results, rel_tol, abs_tol, error=f"{type(e).__name__}: {e}")

    for k, c in gen.cells.items():
        cell = analysis.formulas[k].cell
        expected = cell.cached
        if c.status == "fixed":
            results[k] = CellResult(k, c.formula, "fixed", expected, expected, c.reason)
            continue
        actual = values.get(k)
        if expected is None and not analysis.model.has_cached_values:
            results[k] = CellResult(k, c.formula, "no_cache", None, actual, "엑셀 저장값 없음")
            continue
        if isinstance(actual, float) and math.isnan(actual):
            actual = XLError("#NUM!")
        ok = values_equal(expected, actual, rel_tol, abs_tol)
        results[k] = CellResult(k, c.formula, "match" if ok else "mismatch", expected, actual)

    # 원인 추정: 위상 순서대로 보면서 상위 불일치 여부 확인
    bad: set[str] = set()
    for k in gen.order:
        r = results.get(k)
        if r is None or r.status != "mismatch":
            continue
        parents = sorted(d for d in analysis.formulas[k].deps if d in bad)
        r.root = not parents
        r.cause = _guess_cause(analysis, k, r.expected, r.actual, parents)
        bad.add(k)
    return VerifyReport(results, rel_tol, abs_tol, values=values)
