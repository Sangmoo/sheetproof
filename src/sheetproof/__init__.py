"""sheetproof — 엑셀 수식을 "검증된 코드"로 바꾸는 도구."""

from .analysis import Analysis, analyze
from .codegen import GeneratedModule, generate, parse_targets
from .export import export_bundle
from .verify import VerifyReport, verify
from .workbook import WorkbookError, load_workbook

__version__ = "0.1.0"


def run(source, filename=None, targets_text: str = "", rel_tol: float = 1e-9, abs_tol: float = 1e-9):
    """읽기 → 분석 → 변환 → 검증을 한 번에 수행한다."""
    model = load_workbook(source, filename)
    analysis = analyze(model)
    targets = parse_targets(analysis, targets_text) if targets_text else None
    gen = generate(analysis, targets)
    report = verify(analysis, gen, rel_tol, abs_tol)
    return analysis, gen, report


__all__ = [
    "Analysis", "GeneratedModule", "VerifyReport", "WorkbookError",
    "analyze", "export_bundle", "generate", "load_workbook", "parse_targets", "run", "verify",
]
