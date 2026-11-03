"""앱 핸들러 스모크 테스트 (Gradio 서버 없이 함수 직접 호출)."""

import os

import app


def test_run_all_with_sample(samples_dir):
    path = app.load_sample("settlement.xlsx")
    out = app.run_all(path, "", 1e-9, 1e-9)
    stats, warns, graph, cells, mermaid, code, mism, allr, files, verdict, state = out
    assert "settlement.xlsx" in stats
    assert "<svg" in graph and "def compute" in code
    assert mism == [] and "검증 통과" in verdict
    assert len(files) == 5 and all(os.path.exists(f) for f in files)
    assert not os.path.exists(path)  # 업로드 임시 파일 삭제
    assert (samples_dir / "settlement.xlsx").exists()  # 샘플 원본은 그대로


def test_run_all_with_targets():
    out = app.run_all(app.load_sample("settlement.xlsx"), "정산!H2", 1e-9, 1e-9)
    state = out[-1]
    assert set(state["gen"].cells) == {"정산!D2", "정산!E2", "정산!F2", "정산!G2", "정산!H2"}


def test_run_all_errors():
    assert "올려" in app.run_all(None, "", 1e-9, 1e-9)[0]
    out = app.run_all(app.load_sample("no_cache.xlsx"), "", 1e-9, 1e-9)
    assert "cached value" in out[1]


def test_llm_requires_key():
    state = app.run_all(app.load_sample("errors_text.xlsx"), "", 1e-9, 1e-9)[-1]
    msg = app.run_llm(state, "anthropic", "claude-opus-5-5", "", "", 1e-9, 1e-9)[0]
    assert "API Key" in msg


def test_blocks_hide_api():
    for dep in app.demo.config["dependencies"]:
        assert dep.get("api_name") in (None, False)
