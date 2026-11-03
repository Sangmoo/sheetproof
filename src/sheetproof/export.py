"""M5 내보내기: 파이썬 모듈 + 런타임 + pytest + HTML 검증 보고서."""

from __future__ import annotations

import datetime as _dt
import html
import tempfile
import zipfile
from pathlib import Path

from . import xl_runtime
from .analysis import Analysis
from .codegen import GeneratedModule, py_value
from .verify import STATUS_LABELS, VerifyReport, fmt_value
from .viz import render_svg

RUNTIME_PATH = Path(xl_runtime.__file__)


def render_pytest(gen: GeneratedModule, report: VerifyReport) -> str:
    verified = [k for k, r in report.results.items() if r.status in ("match", "mismatch")]
    lines = [
        f'"""sheetproof 가 생성한 검증 테스트: {gen.module_name} 의 계산 결과가',
        "엑셀에 저장된 값과 셀 단위로 같은지 확인한다.",
        "",
        f"실행: pytest test_{gen.module_name}.py",
        '"""',
        "",
        "import pytest",
        "",
        "import xl_runtime as xl",
        f"import {gen.module_name} as model",
        "",
        f"REL_TOL = {report.rel_tol!r}",
        f"ABS_TOL = {report.abs_tol!r}",
        "",
        "# 셀 → 엑셀이 저장한 계산 결과",
        "EXPECTED = {",
    ]
    for k in verified:
        lines.append(f"    {k!r}: {py_value(report.results[k].expected)},")
    lines += [
        "}",
        "",
        "",
        '@pytest.fixture(scope="module")',
        "def values():",
        "    return model.compute()",
        "",
        "",
        "@pytest.mark.parametrize(\"cell\", list(EXPECTED))",
        "def test_cell_matches_excel(values, cell):",
        "    expected = EXPECTED[cell]",
        "    actual = values[cell]",
        "    assert xl.values_equal(expected, actual, REL_TOL, ABS_TOL), (",
        '        f"{cell}: expected {expected!r}, got {actual!r} ({model.FORMULAS[cell]})"',
        "    )",
        "",
    ]
    return "\n".join(lines)


_CSS = """
body{font-family:system-ui,-apple-system,'Segoe UI','Malgun Gothic',sans-serif;margin:24px;color:#0f172a;background:#fff}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:28px 0 8px;border-bottom:1px solid #e2e8f0;padding-bottom:4px}
.muted{color:#64748b;font-size:13px}.cards{display:flex;gap:12px;flex-wrap:wrap;margin:16px 0}
.card{border:1px solid #e2e8f0;border-radius:10px;padding:12px 16px;min-width:120px}
.card b{display:block;font-size:22px}.ok{color:#15803d}.bad{color:#dc2626}.warn{color:#b45309}
table{border-collapse:collapse;width:100%;font-size:12.5px}th,td{border:1px solid #e2e8f0;padding:4px 8px;text-align:left;vertical-align:top}
th{background:#f8fafc}code{font-family:ui-monospace,Consolas,monospace;font-size:12px}
tr.mismatch td{background:#fef2f2}tr.fixed td{background:#fff7ed}
"""


def _e(x) -> str:
    return html.escape(str(x))


def summary_cards(report: VerifyReport) -> str:
    rate = report.match_rate * 100
    cls = "ok" if report.passed else ("bad" if report.count("mismatch") else "warn")
    cards = [
        (f"{rate:.2f}%", "일치율", cls),
        (report.checked, "검증한 수식", ""),
        (report.count("match"), "일치", "ok"),
        (report.count("mismatch"), "불일치", "bad" if report.count("mismatch") else ""),
        (report.count("fixed"), "고정값(제외)", "warn" if report.count("fixed") else ""),
    ]
    return '<div class="cards">' + "".join(
        f'<div class="card"><b class="{c}">{_e(v)}</b><span class="muted">{_e(t)}</span></div>'
        for v, t, c in cards) + "</div>"


def render_report(analysis: Analysis, gen: GeneratedModule, report: VerifyReport) -> str:
    st = analysis.stats
    now = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    verdict = ("<p class='ok'><b>검증 통과</b> — 변환된 모든 수식이 엑셀 저장값과 일치합니다.</p>"
               if report.passed else
               "<p class='bad'><b>불일치 있음</b> — 아래 목록의 근본 원인 셀부터 확인하세요.</p>"
               if report.count("mismatch") else
               "<p class='warn'><b>검증할 수식이 없습니다.</b></p>")
    if report.error:
        verdict = f"<p class='bad'><b>생성 코드 실행 실패:</b> {_e(report.error)}</p>"
    parts = [
        "<!doctype html><html lang='ko'><head><meta charset='utf-8'>",
        "<meta name='viewport' content='width=device-width,initial-scale=1'>",
        f"<title>sheetproof 검증 보고서 - {_e(analysis.model.filename)}</title><style>{_CSS}</style></head><body>",
        f"<h1>sheetproof 검증 보고서</h1><div class='muted'>{_e(analysis.model.filename)} · {now} · "
        f"허용 오차 rel={report.rel_tol:g}, abs={report.abs_tol:g}</div>",
        summary_cards(report), verdict,
        "<h2>통합문서 요약</h2><table>",
        f"<tr><th>시트</th><td>{st['sheets']}</td><th>수식</th><td>{st['formulas']}</td>"
        f"<th>변환 가능</th><td>{st['convertible']}</td></tr>",
        f"<tr><th>입력 셀</th><td>{st['inputs']}</td><th>중간 계산</th><td>{st['intermediates']}</td>"
        f"<th>출력 셀</th><td>{st['outputs']}</td></tr>",
        f"<tr><th>최대 계산 깊이</th><td>{st['max_depth']}</td><th>이름 정의</th><td>{st['names']}</td>"
        f"<th>사용 함수</th><td>{_e(', '.join(st['functions']) or '-')}</td></tr></table>",
    ]
    if analysis.warnings:
        parts.append("<h2>경고</h2><ul>")
        for w in analysis.warnings:
            parts.append(f"<li class='{'bad' if w.level == 'error' else 'warn'}'>{_e(w.message)}</li>")
        parts.append("</ul>")
    mism = report.mismatches
    if mism:
        parts.append("<h2>불일치 셀</h2><table><tr><th>셀</th><th>수식</th><th>엑셀 저장값</th>"
                     "<th>코드 결과</th><th>원인 추정</th></tr>")
        for r in mism:
            parts.append(f"<tr class='mismatch'><td><code>{_e(r.key)}</code>{' ★' if r.root else ''}</td>"
                         f"<td><code>{_e(r.formula)}</code></td><td>{_e(fmt_value(r.expected))}</td>"
                         f"<td>{_e(fmt_value(r.actual))}</td><td>{_e(r.cause)}</td></tr>")
        parts.append("</table><p class='muted'>★ = 근본 원인 셀 (상위 셀은 모두 일치)</p>")
    parts.append("<h2>의존 그래프</h2>")
    parts.append(render_svg(analysis, gen.targets or None, {r.key for r in mism}))
    parts.append("<h2>전체 셀 결과</h2><table><tr><th>셀</th><th>상태</th><th>수식</th>"
                 "<th>엑셀 저장값</th><th>코드 결과</th><th>비고</th></tr>")
    for k, r in report.results.items():
        parts.append(f"<tr class='{r.status}'><td><code>{_e(k)}</code></td><td>{_e(STATUS_LABELS[r.status])}</td>"
                     f"<td><code>{_e(r.formula)}</code></td><td>{_e(fmt_value(r.expected))}</td>"
                     f"<td>{_e(fmt_value(r.actual))}</td><td>{_e(r.cause)}</td></tr>")
    parts.append("</table><p class='muted'>sheetproof — 엑셀 수식을 검증된 코드로.</p></body></html>")
    return "".join(parts)


def export_bundle(analysis: Analysis, gen: GeneratedModule, report: VerifyReport,
                  out_dir: str | Path | None = None) -> dict[str, Path]:
    """파일을 쓰고 경로를 돌려준다: model, runtime, test, report, zip."""
    out = Path(out_dir or tempfile.mkdtemp(prefix="sheetproof_"))
    out.mkdir(parents=True, exist_ok=True)
    name = gen.module_name
    files = {
        "model": out / f"{name}.py",
        "runtime": out / "xl_runtime.py",
        "test": out / f"test_{name}.py",
        "report": out / f"{name}_report.html",
    }
    files["model"].write_text(gen.source, encoding="utf-8")
    files["runtime"].write_text(RUNTIME_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    files["test"].write_text(render_pytest(gen, report), encoding="utf-8")
    files["report"].write_text(render_report(analysis, gen, report), encoding="utf-8")
    zpath = out / f"{name}_sheetproof.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files.values():
            z.write(p, p.name)
    files["zip"] = zpath
    return files
