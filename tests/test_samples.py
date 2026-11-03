"""실제 Excel 로 재계산·저장한 샘플에 대해 분석 → 변환 → 검증 전체를 돌린다."""

import pytest

from sheetproof import WorkbookError, analyze, generate, load_workbook, parse_targets, run, verify
from sheetproof.viz import render_mermaid, render_svg

FULL_MATCH = ["settlement.xlsx", "price_lookup.xlsx", "sumifs_summary.xlsx", "dates.xlsx",
              "errors_text.xlsx"]


@pytest.mark.parametrize("name", FULL_MATCH)
def test_sample_verifies_against_excel(samples_dir, name):
    analysis, gen, report = run(samples_dir / name)
    assert report.error is None
    assert report.checked > 0
    bad = [(m.key, m.formula, m.expected, m.actual) for m in report.mismatches]
    assert not bad, bad
    assert report.match_rate == 1.0


def test_settlement_analysis(samples_dir):
    model = load_workbook(samples_dir / "settlement.xlsx")
    a = analyze(model)
    assert {"세율", "할인율"} <= {n.name for n in model.names.values()}
    assert a.kinds["정산!K1"] == "input"          # 이름 정의로 참조되는 상수
    assert a.kinds["정산!D2"] == "intermediate"
    assert a.kinds["정산!H8"] == "output"
    assert a.stats["formulas"] == a.stats["convertible"]
    assert a.order.index("정산!D2") < a.order.index("정산!F2") < a.order.index("정산!H8")
    assert not a.cycles


def test_volatile_and_unsupported_are_fixed(samples_dir):
    a, gen, report = run(samples_dir / "errors_text.xlsx")
    assert report.results["Checks!E33"].status == "fixed"   # TODAY()
    assert report.results["Checks!E32"].status == "fixed"   # PROPER/REPT
    assert report.results["Checks!E34"].status == "match"   # 고정된 셀에 의존해도 검증
    codes = {w.code for w in a.warnings}
    assert {"volatile", "unsupported"} <= codes


def test_name_error_from_older_excel_is_fixed(samples_dir):
    """XLOOKUP 을 모르는 Excel 에서 저장하면 #NAME? 이 저장된다 → 경고 후 고정."""
    a, gen, report = run(samples_dir / "price_lookup.xlsx")
    fi = a.formulas["주문!I2"]
    if fi.cell.cached == fi.cell.cached.__class__("#NAME?"):
        assert report.results["주문!I2"].status == "fixed"
    else:  # Microsoft 365 등에서 다시 저장한 경우
        assert report.results["주문!I2"].status == "match"


def test_no_cached_values_warning(samples_dir):
    model = load_workbook(samples_dir / "no_cache.xlsx")
    assert not model.has_cached_values
    assert model.warnings[0].code == "no_cached_values"
    a = analyze(model)
    report = verify(a, generate(a))
    assert report.checked == 0 and report.count("no_cache") == 1


def test_targets_limit_generated_cells(samples_dir):
    model = load_workbook(samples_dir / "settlement.xlsx")
    a = analyze(model)
    targets = parse_targets(a, "정산!H2")
    gen = generate(a, targets)
    assert set(gen.cells) == {"정산!D2", "정산!E2", "정산!F2", "정산!G2", "정산!H2"}
    assert set(gen.inputs) == {"정산!B2", "정산!C2", "정산!K1", "정산!K2"}
    assert verify(a, gen).passed


def test_compute_with_new_inputs(samples_dir):
    from sheetproof.verify import load_module

    a, gen, _ = run(samples_dir / "settlement.xlsx")
    mod = load_module(gen.source)
    base = mod.compute()
    changed = mod.compute({"정산!B2": 1000})
    assert changed["정산!D2"] == 1000 * 35.5
    assert changed["정산!H8"] != base["정산!H8"]


def test_graph_rendering(samples_dir):
    a, gen, report = run(samples_dir / "sumifs_summary.xlsx")
    svg = render_svg(a)
    assert svg.count("<svg") == 1 and "Summary!B2" in svg
    mer = render_mermaid(a)
    assert mer.startswith("flowchart LR") and "-->" in mer


def test_rejects_bad_files(tmp_path):
    p = tmp_path / "x.xls"
    p.write_bytes(b"junk")
    with pytest.raises(WorkbookError):
        load_workbook(p)
    p2 = tmp_path / "x.xlsx"
    p2.write_bytes(b"not a zip")
    with pytest.raises(WorkbookError):
        load_workbook(p2)


def test_cycle_detection(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A3"] = "=A2+1", "=A1+1", "=A1*2"
    p = tmp_path / "cycle.xlsx"
    wb.save(p)
    a = analyze(load_workbook(p))
    assert a.cycles == [["Sheet!A1", "Sheet!A2"]]
    assert "순환 참조 셀에 의존" in a.formulas["Sheet!A3"].issues
    gen = generate(a)
    assert all(c.status == "fixed" for c in gen.cells.values())
