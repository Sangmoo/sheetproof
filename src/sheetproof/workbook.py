"""통합문서 읽기: 수식과 엑셀이 저장한 계산 결과(cached value)를 함께 읽는다."""

from __future__ import annotations

import datetime as _dt
import io
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from . import xl_runtime as xl
from .parser import FormulaParseError, parse_formula, split_reference

MAX_FILE_BYTES = 10 * 1024 * 1024


class WorkbookError(ValueError):
    """사용자에게 그대로 보여줄 수 있는 읽기 오류."""


@dataclass
class Cell:
    sheet: str
    coord: str
    row: int
    col: int
    formula: str | None = None  # "=..." (수식 셀만)
    value: object = None  # 상수 셀의 값
    cached: object = None  # 수식 셀의 엑셀 저장값
    is_date: bool = False
    number_format: str = "General"
    array_ref: str | None = None  # 배열 수식 범위

    @property
    def key(self) -> str:
        return f"{self.sheet}!{self.coord}"

    @property
    def is_formula(self) -> bool:
        return self.formula is not None


@dataclass
class DefinedName:
    name: str
    text: str
    scope: str | None = None  # 시트 범위 이름이면 시트 이름
    ast: object = None
    error: str | None = None


@dataclass
class Warning_:
    level: str  # "error" | "warn" | "info"
    code: str
    message: str
    cells: list[str] = field(default_factory=list)


@dataclass
class WorkbookModel:
    filename: str
    sheets: list[str]
    cells: dict[str, Cell]
    names: dict[str, DefinedName]
    sheet_dims: dict[str, tuple[int, int]]  # sheet → (max_row, max_col)
    warnings: list[Warning_] = field(default_factory=list)
    has_cached_values: bool = True

    def by_sheet(self, sheet: str):
        return (c for c in self.cells.values() if c.sheet == sheet)

    @property
    def formula_cells(self) -> list[Cell]:
        return [c for c in self.cells.values() if c.is_formula]


def _convert_value(v, data_type: str | None = None):
    """openpyxl 값 → 런타임 값 (날짜는 일련번호로)."""
    if data_type == "e" and isinstance(v, str):
        return xl.XLError(v)
    if isinstance(v, _dt.datetime):
        f = float(to_excel(v))
        return int(f) if f.is_integer() else f
    if isinstance(v, _dt.date):
        return int(to_excel(v))
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, _dt.timedelta):
        return v.total_seconds() / 86400
    if isinstance(v, str) and v in xl.XLError.CODES and data_type == "e":
        return xl.XLError(v)
    return v


def _detect_external_links(data: bytes) -> list[str]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return [n for n in z.namelist() if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
    except zipfile.BadZipFile:
        return []


def _has_macros(data: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return any(n.lower().endswith("vbaproject.bin") for n in z.namelist())
    except zipfile.BadZipFile:
        return False


def _has_pivots(data: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return any(n.startswith("xl/pivotTables/") for n in z.namelist())
    except zipfile.BadZipFile:
        return False


def load_workbook(source, filename: str | None = None) -> WorkbookModel:
    """경로 또는 bytes 에서 통합문서를 읽는다."""
    if isinstance(source, (str, Path)):
        path = Path(source)
        filename = filename or path.name
        data = path.read_bytes()
    else:
        data = bytes(source)
        filename = filename or "workbook.xlsx"

    low = filename.lower()
    if low.endswith(".xls"):
        raise WorkbookError("구형 .xls 형식은 지원하지 않습니다. Excel 에서 .xlsx 로 다시 저장해 주세요.")
    if len(data) > MAX_FILE_BYTES:
        raise WorkbookError(f"파일이 너무 큽니다 ({len(data) / 1e6:.1f}MB). 최대 10MB 까지 지원합니다.")
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise WorkbookError("올바른 .xlsx/.xlsm 파일이 아닙니다.")

    warnings: list[Warning_] = []
    keep_vba = low.endswith(".xlsm")
    try:
        wb_f = openpyxl.load_workbook(io.BytesIO(data), data_only=False, keep_vba=False)
        wb_v = openpyxl.load_workbook(io.BytesIO(data), data_only=True, keep_vba=False)
    except Exception as e:  # noqa: BLE001
        raise WorkbookError(f"통합문서를 열 수 없습니다: {e}") from e

    if keep_vba or _has_macros(data):
        warnings.append(Warning_("warn", "macro", "VBA 매크로가 포함되어 있습니다. 매크로는 분석·변환 대상이 아니며 무시합니다."))
    ext = _detect_external_links(data)
    if ext:
        warnings.append(Warning_("warn", "external_link",
                                 f"외부 통합문서 링크 {len(ext)}개가 있습니다. 외부 참조 수식은 엑셀 저장값으로 고정합니다."))
    if _has_pivots(data):
        warnings.append(Warning_("info", "pivot", "피벗 테이블이 있습니다. 피벗 테이블은 분석 대상에서 제외합니다."))

    cells: dict[str, Cell] = {}
    sheet_dims: dict[str, tuple[int, int]] = {}
    sheets = []
    array_multi: list[str] = []
    for ws in wb_f.worksheets:
        sheets.append(ws.title)
        wv = wb_v[ws.title]
        sheet_dims[ws.title] = (ws.max_row or 1, ws.max_column or 1)
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                cell = Cell(ws.title, c.coordinate, c.row, c.column,
                            is_date=bool(c.is_date), number_format=c.number_format or "General")
                formula = None
                if isinstance(v, ArrayFormula):
                    formula = v.text
                    cell.array_ref = v.ref
                    if v.ref and ":" in v.ref:
                        array_multi.append(cell.key)
                elif isinstance(v, DataTableFormula):
                    formula = None
                    cell.value = None
                elif c.data_type == "f" or (isinstance(v, str) and v.startswith("=") and len(v) > 1):
                    formula = v
                if formula is not None:
                    if not formula.startswith("="):
                        formula = "=" + formula
                    cell.formula = formula
                    cv = wv[c.coordinate]
                    cell.cached = _convert_value(cv.value, cv.data_type)
                else:
                    cell.value = _convert_value(v, c.data_type)
                cells[cell.key] = cell

    if array_multi:
        warnings.append(Warning_("warn", "array_formula",
                                 f"여러 셀에 걸친 배열 수식 {len(array_multi)}개는 지원하지 않습니다 (저장값으로 고정).",
                                 array_multi[:50]))

    # 이름 정의
    names: dict[str, DefinedName] = {}

    def add_name(name, text, scope=None):
        dn = DefinedName(name, text or "", scope)
        try:
            dn.ast = parse_formula(dn.text)
        except FormulaParseError as e:
            dn.error = str(e)
        key = (f"{scope}!" if scope else "") + name.upper()
        names[key] = dn

    try:
        for name, d in wb_f.defined_names.items():
            if name.startswith("_xlnm."):
                continue
            add_name(name, d.attr_text)
        for ws in wb_f.worksheets:
            for name, d in getattr(ws, "defined_names", {}).items():
                if name.startswith("_xlnm."):
                    continue
                add_name(name, d.attr_text, ws.title)
    except AttributeError:  # pragma: no cover - 구버전 openpyxl
        pass

    # cached value 존재 여부
    formulas = [c for c in cells.values() if c.is_formula]
    missing = [c for c in formulas if c.cached is None]
    has_cached = True
    if formulas and len(missing) == len(formulas):
        has_cached = False
        warnings.insert(0, Warning_(
            "error", "no_cached_values",
            "수식의 계산 결과(cached value)가 파일에 없습니다. openpyxl 등으로 만든 파일일 수 있습니다. "
            "Excel(또는 LibreOffice)에서 한 번 열어 저장한 뒤 다시 올려 주세요. 검증을 할 수 없습니다."))
    elif missing:
        warnings.append(Warning_("info", "some_cached_missing",
                                 f"수식 {len(missing)}개는 저장된 계산 결과가 비어 있습니다 (빈 문자열 결과일 수 있음).",
                                 [c.key for c in missing[:50]]))

    return WorkbookModel(filename, sheets, cells, names, sheet_dims, warnings, has_cached)


def resolve_bounds(model: WorkbookModel, sheet: str, address: str):
    """주소를 (min_col, min_row, max_col, max_row) 로. 전체 열/행은 사용 범위로 자른다."""
    addr = address.replace("$", "")
    min_col, min_row, max_col, max_row = range_boundaries(addr)
    max_r, max_c = model.sheet_dims.get(sheet, (1, 1))
    if min_row is None:
        min_row, max_row = 1, max_r
    if min_col is None:
        min_col, max_col = 1, max_c
    return min_col, min_row, max_col, max_row


def resolve_name(model: WorkbookModel, name: str, sheet: str):
    """이름 정의를 찾는다 (시트 범위 우선)."""
    return model.names.get(f"{sheet}!{name.upper()}") or model.names.get(name.upper())


__all__ = [
    "Cell", "DefinedName", "WorkbookModel", "WorkbookError", "Warning_",
    "load_workbook", "resolve_bounds", "resolve_name", "split_reference",
]
