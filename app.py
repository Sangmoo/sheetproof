"""sheetproof — Hugging Face Space 진입점 (Gradio)."""

from __future__ import annotations

import html
import os
import shutil
import sys
import tempfile
from pathlib import Path

import gradio as gr

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from sheetproof import WorkbookError, analyze, generate, load_workbook, parse_targets, verify  # noqa: E402
from sheetproof.export import export_bundle, summary_cards  # noqa: E402
from sheetproof.llm import PROVIDERS, LLMError, complete_with_llm, llm_candidates, make_caller  # noqa: E402
from sheetproof.verify import STATUS_LABELS, fmt_value  # noqa: E402
from sheetproof.viz import render_mermaid, render_svg  # noqa: E402

SAMPLES_DIR = Path(__file__).resolve().parent / "samples"
KIND_LABELS = {"input": "입력", "intermediate": "중간 계산", "output": "출력"}

CSS = """
.cards{display:flex;gap:12px;flex-wrap:wrap;margin:8px 0}
.card{border:1px solid var(--border-color-primary);border-radius:10px;padding:10px 14px;min-width:110px}
.card b{display:block;font-size:22px}.card span{font-size:12px;opacity:.75}
.ok{color:#16a34a}.bad{color:#dc2626}.warn{color:#d97706}
"""

INTRO = """
# sheetproof — 엑셀 수식을 "검증된 코드"로

`.xlsx` 를 올리면 **① 셀 의존 관계를 분석·시각화**하고 **② 계산 로직을 Python 코드로 변환**한 뒤
**③ 엑셀에 저장된 계산 결과와 셀 단위로 대조해 같은 값이 나오는지 증명**합니다.

- 업로드 파일은 서버에 저장하지 않으며, 처리 직후 임시 파일을 삭제합니다.
- LLM 보완은 **본인 API Key 를 입력한 경우에만** 호출되며, 수식과 셀 주소·타입만 전송합니다(셀 값은 전송하지 않음).
"""


def _state(analysis=None, gen=None, report=None, targets=None, overrides=None):
    return {"analysis": analysis, "gen": gen, "report": report, "targets": targets,
            "overrides": overrides or {}}


def _warnings_md(analysis) -> str:
    if not analysis.warnings:
        return "경고 없음 ✅"
    icon = {"error": "⛔", "warn": "⚠️", "info": "ℹ️"}
    lines = []
    for w in analysis.warnings:
        cells = f" <small>({', '.join(w.cells[:6])}{' …' if len(w.cells) > 6 else ''})</small>" if w.cells else ""
        lines.append(f"- {icon.get(w.level, '•')} {w.message}{cells}")
    return "\n".join(lines)


def _stats_md(analysis) -> str:
    s = analysis.stats
    funcs = ", ".join(f"`{k}`×{v}" for k, v in list(s["functions"].items())[:20]) or "-"
    per_sheet = ", ".join(f"{k} {v}" for k, v in s["per_sheet"].items())
    return (
        f"**{html.escape(analysis.model.filename)}** — 시트 {s['sheets']}개 · 값이 있는 셀 {s['cells']}개 · "
        f"수식 {s['formulas']}개 (변환 가능 {s['convertible']}개)\n\n"
        f"| 입력 셀 | 중간 계산 | 출력 셀 | 최대 계산 깊이 | 이름 정의 |\n|---|---|---|---|---|\n"
        f"| {s['inputs']} | {s['intermediates']} | {s['outputs']} | {s['max_depth']} | {s['names']} |\n\n"
        f"시트별 수식: {per_sheet}\n\n사용 함수: {funcs}"
    )


def _cells_table(analysis):
    rows = []
    for k, kind in analysis.kinds.items():
        cell = analysis.model.cells[k]
        if cell.is_formula:
            fi = analysis.formulas[k]
            note = "" if fi.convertible else ", ".join(
                dict.fromkeys(fi.issues + sorted(fi.unsupported) + sorted(fi.volatile)))
            rows.append([k, KIND_LABELS[kind], cell.formula, fmt_value(cell.cached), note])
        else:
            rows.append([k, KIND_LABELS[kind], "", fmt_value(cell.value), ""])
    order = {"입력": 0, "중간 계산": 1, "출력": 2}
    rows.sort(key=lambda r: (order[r[1]], r[0]))
    return rows


def _mismatch_rows(report):
    return [[r.key + (" ★" if r.root else ""), r.formula, fmt_value(r.expected), fmt_value(r.actual), r.cause]
            for r in report.mismatches]


def _all_rows(report):
    return [[k, STATUS_LABELS[r.status], r.formula, fmt_value(r.expected), fmt_value(r.actual), r.cause]
            for k, r in report.results.items()]


def _verify_html(report) -> str:
    if report.error:
        return f"<p class='bad'><b>생성 코드 실행 실패:</b> {html.escape(report.error)}</p>"
    if report.checked == 0:
        msg = "검증할 수식이 없습니다 (모두 고정값이거나 엑셀 저장값이 없음)."
        return summary_cards(report) + f"<p class='warn'>{msg}</p>"
    if report.passed:
        verdict = "<p class='ok'><b>✅ 검증 통과</b> — 변환된 모든 수식이 엑셀 저장값과 일치합니다.</p>"
    else:
        verdict = ("<p class='bad'><b>❌ 불일치 있음</b> — ★ 표시(근본 원인 셀)부터 확인하세요. "
                   "나머지는 상위 셀 불일치가 전파된 것입니다.</p>")
    return summary_cards(report) + verdict


def _export(analysis, gen, report):
    out = tempfile.mkdtemp(prefix="sheetproof_out_")
    files = export_bundle(analysis, gen, report, out)
    return [str(files[k]) for k in ("zip", "model", "runtime", "test", "report")]


def _pipeline(analysis, targets, overrides, rel_tol, abs_tol):
    gen = generate(analysis, targets, overrides)
    report = verify(analysis, gen, float(rel_tol), float(abs_tol))
    mism = {r.key for r in report.mismatches}
    graph = render_svg(analysis, targets, mism)
    return gen, report, graph


def run_all(file, targets_text, rel_tol, abs_tol):
    """업로드 → 분석 → 변환 → 검증 → 내보내기."""
    empty = ("", "", "", [], "", "", [], [], None, "", _state())
    if file is None:
        return ("파일을 올려 주세요.",) + empty[1:]
    path = file if isinstance(file, str) else getattr(file, "name", None)
    if not path or not os.path.exists(path):
        return ("업로드 파일은 처리 직후 삭제됩니다. 다시 실행하려면 파일을 다시 올려 주세요.",) + empty[1:]
    try:
        try:
            model = load_workbook(path, Path(path).name)
        finally:
            _cleanup_upload(path)
        analysis = analyze(model)
        targets = parse_targets(analysis, targets_text) if targets_text and targets_text.strip() else None
        if targets_text and targets_text.strip() and not targets:
            raise ValueError("지정한 범위에 수식 셀이 없습니다.")
    except (WorkbookError, ValueError) as e:
        return (f"⛔ {e}",) + empty[1:]
    gen, report, graph = _pipeline(analysis, targets, {}, rel_tol, abs_tol)
    files = _export(analysis, gen, report)
    n_llm = len(llm_candidates(analysis))
    llm_hint = (f"LLM 보완 대상(미지원 함수 수식): <b>{n_llm}개</b> — ④ 탭에서 변환할 수 있습니다." if n_llm
                else "LLM 보완 대상이 없습니다.")
    return (
        _stats_md(analysis), _warnings_md(analysis), graph, _cells_table(analysis),
        render_mermaid(analysis, targets), gen.source, _mismatch_rows(report), _all_rows(report),
        files, _verify_html(report) + f"<p>{llm_hint}</p>",
        _state(analysis, gen, report, targets),
    )


def run_llm(state, provider, model_name, api_key, base_url, rel_tol, abs_tol):
    """미지원 수식만 LLM 으로 변환 → 검증 통과 시 채택."""
    no_change = (gr.update(), gr.update(), gr.update(), gr.update(), gr.update(), gr.update(), state)
    if not state or state.get("analysis") is None:
        return ("먼저 파일을 올리고 실행해 주세요.",) + no_change
    if not api_key or not api_key.strip():
        return ("🔑 API Key 를 입력한 경우에만 LLM 을 호출합니다. Key 는 저장되지 않습니다.",) + no_change
    analysis = state["analysis"]
    try:
        call = make_caller(provider, api_key.strip(), model_name.strip() or None, (base_url or "").strip() or None)
    except (LLMError, ImportError) as e:
        return (f"⛔ {e}",) + no_change
    res = complete_with_llm(analysis, call, state.get("targets"), float(rel_tol), float(abs_tol),
                            existing=state.get("overrides"))
    lines = [res.message, ""]
    for c in res.candidates:
        mark = "✅" if c.accepted else "❌"
        lines.append(f"- {mark} `{c.key}` `{c.formula}` → {c.reason}")
        if c.expr:
            lines.append(f"  - 식: `{c.expr}`" + (f" — {c.note}" if c.note else ""))
    gen, report, graph = _pipeline(analysis, state.get("targets"), res.overrides, rel_tol, abs_tol)
    files = _export(analysis, gen, report)
    new_state = _state(analysis, gen, report, state.get("targets"), res.overrides)
    return ("\n".join(lines), gen.source, _mismatch_rows(report), _all_rows(report), files,
            _verify_html(report), graph, new_state)


def _cleanup_upload(path):
    """업로드 임시 파일을 지운다 (개인정보 보호)."""
    try:
        if path and os.path.exists(path):
            os.remove(path)
            parent = os.path.dirname(path)
            if parent and not os.listdir(parent):
                shutil.rmtree(parent, ignore_errors=True)
    except OSError:
        pass


def load_sample(name):
    """샘플을 임시 복사본으로 넘긴다 (원본은 지우지 않도록)."""
    if not name:
        return None
    src = SAMPLES_DIR / name
    tmp = Path(tempfile.mkdtemp(prefix="sheetproof_sample_")) / name
    shutil.copy(src, tmp)
    return str(tmp)


def build_app() -> gr.Blocks:
    samples = sorted(p.name for p in SAMPLES_DIR.glob("*.xlsx"))
    with gr.Blocks(title="sheetproof", css=CSS, theme=gr.themes.Soft(), delete_cache=(1800, 3600)) as demo:
        gr.Markdown(INTRO)
        state = gr.State(_state())
        with gr.Row():
            with gr.Column(scale=2):
                file = gr.File(label="엑셀 파일 (.xlsx / .xlsm, 최대 10MB)", file_types=[".xlsx", ".xlsm"],
                               type="filepath")
                with gr.Row():
                    sample = gr.Dropdown(samples, label="또는 샘플 선택", value=None)
            with gr.Column(scale=1):
                targets = gr.Textbox(label="대상 셀 범위 (선택)", placeholder="예: 정산!H2:H8, 주문!F9")
                with gr.Row():
                    rel_tol = gr.Number(value=1e-9, label="상대 허용 오차")
                    abs_tol = gr.Number(value=1e-9, label="절대 허용 오차")
                run_btn = gr.Button("분석 · 변환 · 검증 실행", variant="primary")
        verdict = gr.HTML()
        with gr.Tabs():
            with gr.Tab("① 분석"):
                stats = gr.Markdown()
                warns = gr.Markdown()
                graph = gr.HTML(label="의존 그래프")
                with gr.Accordion("Mermaid 코드", open=False):
                    mermaid = gr.Code(language="markdown", label="flowchart")
                cells = gr.Dataframe(headers=["셀", "분류", "수식", "값", "고정 사유"], label="셀 분류",
                                     wrap=True, max_height=420,
                                     column_widths=["16%", "10%", "40%", "16%", "18%"])
            with gr.Tab("② 변환 코드"):
                code = gr.Code(language="python", label="생성된 Python 모듈")
            with gr.Tab("③ 검증"):
                mism = gr.Dataframe(headers=["셀", "수식", "엑셀 저장값", "코드 결과", "원인 추정"],
                                    label="불일치 셀 (★ = 근본 원인)", wrap=True, max_height=420,
                                    column_widths=["16%", "34%", "15%", "15%", "20%"])
                allr = gr.Dataframe(headers=["셀", "상태", "수식", "엑셀 저장값", "코드 결과", "비고"],
                                    label="전체 결과", wrap=True, max_height=480,
                                    column_widths=["16%", "12%", "32%", "13%", "13%", "14%"])
            with gr.Tab("④ LLM 보완 (선택)"):
                gr.Markdown("규칙 기반으로 변환하지 못한 **미지원 함수 수식만** LLM 이 변환합니다. "
                            "LLM 결과도 같은 검증 엔진을 통과해야만 채택됩니다. "
                            "전송 범위: 수식 문자열, 참조 셀 주소와 타입 (셀 값 제외).")
                with gr.Row():
                    provider = gr.Dropdown([(v["label"], k) for k, v in PROVIDERS.items()],
                                           value="anthropic", label="공급자")
                    model_name = gr.Textbox(value=PROVIDERS["anthropic"]["model"], label="모델")
                with gr.Row():
                    api_key = gr.Textbox(label="API Key (저장되지 않음)", type="password")
                    base_url = gr.Textbox(label="Base URL (OpenAI 호환일 때)", placeholder="https://.../v1")
                llm_btn = gr.Button("미지원 수식 LLM 변환 + 검증")
                llm_out = gr.Markdown()
            with gr.Tab("⑤ 내보내기"):
                gr.Markdown("ZIP 에는 모델 모듈(.py), 런타임(xl_runtime.py), pytest 검증 코드, HTML 보고서가 들어 있습니다. "
                            "압축을 풀고 `pytest` 를 실행하면 같은 검증을 재현할 수 있습니다.")
                files = gr.Files(label="다운로드")

        outputs = [stats, warns, graph, cells, mermaid, code, mism, allr, files, verdict, state]
        sample.change(load_sample, sample, file, api_name=False)
        run_btn.click(run_all, [file, targets, rel_tol, abs_tol], outputs, api_name=False)
        provider.change(lambda p: PROVIDERS[p]["model"], provider, model_name, api_name=False)
        llm_btn.click(run_llm, [state, provider, model_name, api_key, base_url, rel_tol, abs_tol],
                      [llm_out, code, mism, allr, files, verdict, graph, state], api_name=False)
    return demo


demo = build_app()

if __name__ == "__main__":
    demo.launch(max_file_size="10mb", show_api=False)
