"""M1 분석: 수식 파싱, 참조 추출, 의존 그래프, 셀 분류, 경고."""

from __future__ import annotations

from collections import Counter, defaultdict, deque
from dataclasses import dataclass, field

from openpyxl.utils.cell import get_column_letter

from . import parser as P
from .workbook import Cell, Warning_, WorkbookModel, resolve_bounds, resolve_name
from .xl_runtime import FUNCTIONS, XLError

VOLATILE_FUNCTIONS = frozenset({
    "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO",
})
DYNAMIC_FUNCTIONS = frozenset({
    "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "MAP", "REDUCE",
    "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSEROWS",
    "CHOOSECOLS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND", "TEXTSPLIT",
})
SUPPORTED_FUNCTIONS = frozenset(FUNCTIONS)


@dataclass
class ResolvedRef:
    sheet: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int
    via_name: str | None = None

    @property
    def is_single(self) -> bool:
        return self.min_col == self.max_col and self.min_row == self.max_row

    @property
    def address(self) -> str:
        a = f"{get_column_letter(self.min_col)}{self.min_row}"
        if self.is_single:
            return a
        return f"{a}:{get_column_letter(self.max_col)}{self.max_row}"

    @property
    def key(self) -> str:
        return f"{self.sheet}!{self.address}"

    @property
    def size(self) -> int:
        return (self.max_col - self.min_col + 1) * (self.max_row - self.min_row + 1)


@dataclass
class FormulaInfo:
    cell: Cell
    ast: object = None
    parse_error: str | None = None
    refs: list[ResolvedRef] = field(default_factory=list)
    deps: set[str] = field(default_factory=set)  # 실제 값이 있는 셀 키
    functions: set[str] = field(default_factory=set)
    unsupported: set[str] = field(default_factory=set)
    volatile: set[str] = field(default_factory=set)
    issues: list[str] = field(default_factory=list)  # 고정 처리 사유

    @property
    def key(self) -> str:
        return self.cell.key

    @property
    def convertible(self) -> bool:
        return self.ast is not None and not self.unsupported and not self.volatile and not self.issues


@dataclass
class Analysis:
    model: WorkbookModel
    formulas: dict[str, FormulaInfo]
    dependents: dict[str, set[str]]  # 셀 → 그 셀을 참조하는 수식 셀
    kinds: dict[str, str]  # 셀 → input / intermediate / output
    order: list[str]  # 수식 셀 위상 정렬 순서 (순환 셀 제외)
    cycles: list[list[str]]
    warnings: list[Warning_]
    stats: dict

    def ancestors(self, keys) -> set[str]:
        """주어진 셀들이 의존하는 모든 셀 (자기 자신 포함)."""
        seen = set()
        stack = list(keys)
        while stack:
            k = stack.pop()
            if k in seen:
                continue
            seen.add(k)
            fi = self.formulas.get(k)
            if fi:
                stack.extend(fi.deps)
        return seen


class _RefCollector:
    def __init__(self, model: WorkbookModel, info: FormulaInfo):
        self.model = model
        self.info = info
        self.sheet = info.cell.sheet
        self._name_stack: list[str] = []

    def visit(self, node, sheet: str | None = None):
        sheet = sheet or self.sheet
        for n in P.walk(node):
            if isinstance(n, P.Call):
                self.info.functions.add(n.name)
                if n.name in VOLATILE_FUNCTIONS:
                    self.info.volatile.add(n.name)
                elif n.name not in SUPPORTED_FUNCTIONS:
                    self.info.unsupported.add(n.name)
            elif isinstance(n, P.Ref):
                self.add_ref(n, sheet)
            elif isinstance(n, P.Name):
                self.add_name(n, sheet)

    def add_ref(self, ref: P.Ref, sheet: str, via: str | None = None):
        if ref.external:
            self.info.issues.append("외부 통합문서 참조")
            return
        target = ref.sheet if ref.sheet is not None else sheet
        if target not in self.model.sheet_dims:
            self.info.issues.append(f"존재하지 않는 시트 참조 '{target}'")
            return
        try:
            b = resolve_bounds(self.model, target, ref.address)
        except (ValueError, TypeError):
            self.info.issues.append(f"해석할 수 없는 참조 '{ref.address}'")
            return
        self.info.refs.append(ResolvedRef(target, *b, via_name=via))

    def add_name(self, name: P.Name, sheet: str):
        dn = resolve_name(self.model, name.name, sheet)
        if dn is None:
            self.info.issues.append(f"정의되지 않은 이름 '{name.name}'")
            return
        if dn.error or dn.ast is None:
            self.info.issues.append(f"이름 '{name.name}' 해석 실패")
            return
        if name.name.upper() in self._name_stack:
            self.info.issues.append(f"이름 '{name.name}' 순환 정의")
            return
        self._name_stack.append(name.name.upper())
        if isinstance(dn.ast, P.Ref):
            self.add_ref(dn.ast, dn.scope or sheet, via=name.name)
        else:
            self.visit(dn.ast, dn.scope or sheet)
        self._name_stack.pop()


def _expand_deps(model: WorkbookModel, refs: list[ResolvedRef], sheet_index) -> set[str]:
    deps: set[str] = set()
    for r in refs:
        if r.size <= 64:
            for row in range(r.min_row, r.max_row + 1):
                for col in range(r.min_col, r.max_col + 1):
                    k = f"{r.sheet}!{get_column_letter(col)}{row}"
                    if k in model.cells:
                        deps.add(k)
        else:
            for c in sheet_index.get(r.sheet, ()):
                if r.min_row <= c.row <= r.max_row and r.min_col <= c.col <= r.max_col:
                    deps.add(c.key)
    return deps


def _tarjan(nodes: list[str], edges: dict[str, set[str]]) -> list[list[str]]:
    """강한 연결 요소 (크기 2 이상 또는 자기 참조) = 순환."""
    index = {}
    low = {}
    on_stack = set()
    stack: list[str] = []
    result: list[list[str]] = []
    counter = 0
    for root in nodes:
        if root in index:
            continue
        work = [(root, iter(sorted(edges.get(root, ()))))]
        index[root] = low[root] = counter
        counter += 1
        stack.append(root)
        on_stack.add(root)
        while work:
            v, it = work[-1]
            advanced = False
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter
                    counter += 1
                    stack.append(w)
                    on_stack.add(w)
                    work.append((w, iter(sorted(edges.get(w, ())))))
                    advanced = True
                    break
                if w in on_stack:
                    low[v] = min(low[v], index[w])
            if advanced:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[v])
            if low[v] == index[v]:
                comp = []
                while True:
                    w = stack.pop()
                    on_stack.discard(w)
                    comp.append(w)
                    if w == v:
                        break
                if len(comp) > 1 or v in edges.get(v, ()):
                    result.append(sorted(comp))
    return result


def analyze(model: WorkbookModel) -> Analysis:
    warnings = list(model.warnings)
    sheet_index = defaultdict(list)
    for c in model.cells.values():
        sheet_index[c.sheet].append(c)

    formulas: dict[str, FormulaInfo] = {}
    for cell in model.formula_cells:
        info = FormulaInfo(cell)
        try:
            info.ast = P.parse_formula(cell.formula)
        except P.FormulaParseError as e:
            info.parse_error = str(e)
            info.issues.append("수식 해석 실패")
        if cell.array_ref and ":" in cell.array_ref:
            info.issues.append("여러 셀 배열 수식")
        if info.ast is not None:
            _RefCollector(model, info).visit(info.ast)
        if cell.cached == XLError("#NAME?") and info.functions:
            # 저장한 Excel 버전이 함수를 모르면 #NAME? 이 저장된다 (예: Excel 2019 의 XLOOKUP)
            info.issues.append("엑셀 저장값이 #NAME? (저장한 Excel 이 함수를 지원하지 않음)")
        dyn = info.functions & DYNAMIC_FUNCTIONS
        if dyn:
            info.issues.append("동적 배열/LAMBDA 함수 (" + ", ".join(sorted(dyn)) + ")")
            info.unsupported -= dyn
        info.deps = _expand_deps(model, info.refs, sheet_index)
        formulas[cell.key] = info

    # 의존 관계
    dependents: dict[str, set[str]] = defaultdict(set)
    for k, fi in formulas.items():
        for d in fi.deps:
            dependents[d].add(k)

    kinds: dict[str, str] = {}
    for k, cell in model.cells.items():
        if cell.is_formula:
            kinds[k] = "intermediate" if dependents.get(k) else "output"
        elif dependents.get(k):
            kinds[k] = "input"

    # 위상 정렬 (Kahn)
    fkeys = list(formulas)
    fdeps = {k: {d for d in formulas[k].deps if d in formulas} for k in fkeys}
    indeg = {k: len(fdeps[k]) for k in fkeys}
    rev: dict[str, list[str]] = defaultdict(list)
    for k, ds in fdeps.items():
        for d in ds:
            rev[d].append(k)
    queue = deque(sorted((k for k in fkeys if indeg[k] == 0), key=_sort_key))
    order: list[str] = []
    while queue:
        k = queue.popleft()
        order.append(k)
        for n in rev.get(k, ()):
            indeg[n] -= 1
            if indeg[n] == 0:
                queue.append(n)
    cycles: list[list[str]] = []
    if len(order) < len(fkeys):
        remaining = [k for k in fkeys if k not in set(order)]
        cycles = _tarjan(remaining, fdeps)
        in_cycle = {k for comp in cycles for k in comp}
        for comp in cycles:
            warnings.append(Warning_("error", "cycle",
                                     f"순환 참조: {' → '.join(comp[:8])}{' …' if len(comp) > 8 else ''}", comp))
        for k in remaining:
            formulas[k].issues.append("순환 참조" if k in in_cycle else "순환 참조 셀에 의존")

    # 경고 집계
    vol = [k for k, f in formulas.items() if f.volatile]
    if vol:
        names = sorted({n for k in vol for n in formulas[k].volatile})
        warnings.append(Warning_("warn", "volatile",
                                 f"휘발성 함수({', '.join(names)}) 수식 {len(vol)}개는 엑셀 저장값으로 고정하고 검증에서 제외합니다.",
                                 vol[:50]))
    uns = [k for k, f in formulas.items() if f.unsupported and not f.volatile]
    if uns:
        names = Counter(n for k in uns for n in formulas[k].unsupported)
        warnings.append(Warning_("warn", "unsupported",
                                 "지원하지 않는 함수: " + ", ".join(f"{n}({c})" for n, c in names.most_common())
                                 + f" — 수식 {len(uns)}개는 저장값으로 고정합니다 (LLM 보완 가능).", uns[:50]))
    issues = Counter(i for f in formulas.values() for i in f.issues
                     if i not in ("순환 참조", "순환 참조 셀에 의존"))
    for issue, n in issues.items():
        warnings.append(Warning_("warn", "fixed", f"{issue}: 수식 {n}개를 저장값으로 고정합니다.",
                                 [k for k, f in formulas.items() if issue in f.issues][:50]))

    func_counter = Counter(n for f in formulas.values() for n in f.functions)
    kind_counter = Counter(kinds.values())
    stats = {
        "sheets": len(model.sheets),
        "cells": len(model.cells),
        "formulas": len(formulas),
        "constants": sum(1 for c in model.cells.values() if not c.is_formula),
        "inputs": kind_counter.get("input", 0),
        "intermediates": kind_counter.get("intermediate", 0),
        "outputs": kind_counter.get("output", 0),
        "convertible": sum(1 for f in formulas.values() if f.convertible),
        "functions": dict(func_counter.most_common()),
        "per_sheet": {s: sum(1 for f in formulas.values() if f.cell.sheet == s) for s in model.sheets},
        "names": len(model.names),
        "max_depth": _max_depth(order, fdeps),
    }
    return Analysis(model, formulas, dict(dependents), kinds, order, cycles, warnings, stats)


def _max_depth(order, fdeps) -> int:
    depth: dict[str, int] = {}
    for k in order:
        depth[k] = 1 + max((depth.get(d, 0) for d in fdeps[k]), default=0)
    return max(depth.values(), default=0)


def _sort_key(key: str):
    sheet, _, coord = key.rpartition("!")
    col = "".join(ch for ch in coord if ch.isalpha())
    row = "".join(ch for ch in coord if ch.isdigit())
    return sheet, int(row or 0), len(col), col


def depths(analysis: Analysis) -> dict[str, int]:
    """수식 셀의 계산 단계 (입력 셀 = 0)."""
    d: dict[str, int] = {}
    for k in analysis.order:
        fi = analysis.formulas[k]
        d[k] = 1 + max((d.get(x, 0) for x in fi.deps), default=0)
    return d
