"""검증 엔진이 틀린 변환을 잡아내는지, 내보낸 파일이 독립적으로 동작하는지 확인한다."""

import subprocess
import sys
import zipfile

from sheetproof import run, verify
from sheetproof.export import export_bundle


def _mutate(source: str, key: str, old: str, new: str) -> str:
    target = f"v[{key!r}] = "
    lines = source.splitlines()
    for i, line in enumerate(lines):
        if line.strip().startswith(target):
            assert old in line, line
            lines[i] = line.replace(old, new, 1)
            return "\n".join(lines)
    raise AssertionError(f"{key} not found")


def test_injected_bug_is_detected_with_root_cause(samples_dir):
    a, gen, report = run(samples_dir / "settlement.xlsx")
    assert report.passed
    bad_src = _mutate(gen.source, "정산!D3", "xl.mul(", "xl.add(")  # 수량*단가 → 수량+단가
    bad = verify(a, gen, source=bad_src)
    assert not bad.passed
    roots = [r.key for r in bad.mismatches if r.root]
    assert roots == ["정산!D3"]
    downstream = {r.key for r in bad.mismatches if not r.root}
    assert {"정산!F3", "정산!H3", "정산!D8"} <= downstream
    assert all("상위 셀 불일치 전파" in bad.results[k].cause for k in downstream)


def test_rounding_bug_is_detected(samples_dir):
    a, gen, _ = run(samples_dir / "settlement.xlsx")
    # G3 = ROUND(3675*0.1, 0) = 368 (ROUNDDOWN 이면 367)
    bad_src = _mutate(gen.source, "정산!G3", "xl.ROUND(", "xl.ROUNDDOWN(")
    bad = verify(a, gen, source=bad_src)
    assert "정산!G3" in {r.key for r in bad.mismatches if r.root}


def test_tolerance_controls_match(samples_dir):
    a, gen, _ = run(samples_dir / "settlement.xlsx")
    bad_src = _mutate(gen.source, "정산!D2", "xl.mul(", "(lambda a, b: xl.mul(a, b) * (1 + 1e-7))(")
    assert not verify(a, gen, source=bad_src).results["정산!D2"].status == "match"
    loose = verify(a, gen, rel_tol=1e-6, source=bad_src)
    assert loose.results["정산!D2"].status == "match"


def test_broken_module_reports_error(samples_dir):
    a, gen, _ = run(samples_dir / "settlement.xlsx")
    rep = verify(a, gen, source=gen.source + "\nraise RuntimeError('boom')\n")
    assert rep.error and "boom" in rep.error and not rep.passed


def test_export_bundle_runs_standalone(samples_dir, tmp_path):
    a, gen, report = run(samples_dir / "sumifs_summary.xlsx")
    files = export_bundle(a, gen, report, tmp_path / "out")
    with zipfile.ZipFile(files["zip"]) as z:
        names = set(z.namelist())
    assert names == {f"{gen.module_name}.py", "xl_runtime.py", f"test_{gen.module_name}.py",
                     f"{gen.module_name}_report.html"}
    html = files["report"].read_text(encoding="utf-8")
    assert "100.00%" in html and "<svg" in html
    # 내보낸 pytest 를 패키지 없이 별도 프로세스에서 실행
    proc = subprocess.run([sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider",
                           files["test"].name], cwd=files["test"].parent,
                          capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert f"{report.checked} passed" in proc.stdout
