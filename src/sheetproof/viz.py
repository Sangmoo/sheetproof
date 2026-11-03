"""의존 그래프 시각화 (SVG / Mermaid)."""

from __future__ import annotations

import html
from collections import defaultdict

from .analysis import Analysis

KIND_STYLE = {
    "input": ("#dbeafe", "#1d4ed8", "입력"),
    "intermediate": ("#f1f5f9", "#475569", "중간 계산"),
    "output": ("#dcfce7", "#15803d", "출력"),
    "range": ("#ede9fe", "#6d28d9", "범위"),
    "fixed": ("#ffedd5", "#c2410c", "고정값"),
}
MISMATCH_STROKE = "#dc2626"

NODE_W = 168
NODE_H = 26
COL_GAP = 70
ROW_GAP = 10


def build_graph(analysis: Analysis, focus: list[str] | None = None, max_nodes: int = 150):
    """참조 단위 그래프: 수식 셀, 단일 입력 셀, 범위 노드."""
    formulas = analysis.formulas
    keys = list(analysis.order) + [k for k in formulas if k not in set(analysis.order)]
    if focus:
        scope = analysis.ancestors(focus)
        keys = [k for k in keys if k in scope]
    truncated = False
    if len(keys) > max_nodes:
        truncated = True
        # 출력 셀과 가까운 쪽(뒤쪽)을 우선 보여준다
        keys = keys[-max_nodes:]
    kset = set(keys)

    nodes: dict[str, dict] = {}
    edges: set[tuple[str, str]] = set()
    for k in keys:
        fi = formulas[k]
        fixed = not fi.convertible
        nodes[k] = {"id": k, "kind": "fixed" if fixed else analysis.kinds.get(k, "output"),
                    "label": k, "title": fi.cell.formula}
    for k in keys:
        fi = formulas[k]
        for r in fi.refs:
            if r.is_single:
                src = r.key
                if src not in nodes:
                    if src in formulas and src not in kset:
                        continue
                    cell = analysis.model.cells.get(src)
                    nodes[src] = {"id": src, "kind": "input", "label": src,
                                  "title": "" if cell is None else repr(cell.value)}
                edges.add((src, k))
            else:
                rid = "R:" + r.key
                if rid not in nodes:
                    if len(nodes) > max_nodes * 1.5:
                        truncated = True
                        continue
                    nodes[rid] = {"id": rid, "kind": "range",
                                  "label": (r.via_name + " = " if r.via_name else "") + r.key,
                                  "title": f"{r.size}개 셀"}
                    for d in fi.deps:
                        if d in kset and d in formulas:
                            c = formulas[d].cell
                            if r.sheet == c.sheet and r.min_row <= c.row <= r.max_row \
                                    and r.min_col <= c.col <= r.max_col:
                                edges.add((d, rid))
                edges.add((rid, k))
    return nodes, sorted(edges), truncated


def _levels(nodes, edges):
    preds = defaultdict(list)
    for a, b in edges:
        preds[b].append(a)
    level: dict[str, int] = {}

    def lv(n, stack=()):
        if n in level:
            return level[n]
        if n in stack:
            return 0
        ps = [p for p in preds[n] if p in nodes]
        level[n] = 0 if not ps else 1 + max(lv(p, stack + (n,)) for p in ps)
        return level[n]

    import sys

    old = sys.getrecursionlimit()
    sys.setrecursionlimit(max(old, 10000))
    try:
        for n in nodes:
            lv(n)
    finally:
        sys.setrecursionlimit(old)
    return level


def render_svg(analysis: Analysis, focus=None, mismatches: set[str] | None = None,
               max_nodes: int = 150) -> str:
    nodes, edges, truncated = build_graph(analysis, focus, max_nodes)
    mismatches = mismatches or set()
    if not nodes:
        return "<p>수식이 없습니다.</p>"
    level = _levels(nodes, edges)
    cols: dict[int, list[str]] = defaultdict(list)
    for n in nodes:
        cols[level[n]].append(n)
    pos = {}
    for lvl, members in cols.items():
        members.sort()
        for i, n in enumerate(members):
            pos[n] = (20 + lvl * (NODE_W + COL_GAP), 40 + i * (NODE_H + ROW_GAP))
    width = 40 + (max(cols) + 1) * (NODE_W + COL_GAP)
    height = 60 + max(len(m) for m in cols.values()) * (NODE_H + ROW_GAP)

    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
             f'viewBox="0 0 {width} {height}" font-family="ui-monospace,Consolas,monospace" font-size="11">']
    parts.append('<defs><marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" '
                 'markerHeight="6" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#94a3b8"/>'
                 '</marker></defs>')
    parts.append(f'<rect width="{width}" height="{height}" fill="#ffffff"/>')
    # 범례
    lx = 20
    for kind, (fill, stroke, label) in KIND_STYLE.items():
        parts.append(f'<rect x="{lx}" y="10" width="12" height="12" rx="3" fill="{fill}" stroke="{stroke}"/>')
        parts.append(f'<text x="{lx + 16}" y="20" fill="#334155">{label}</text>')
        lx += 30 + 12 * len(label)
    if mismatches:
        parts.append(f'<rect x="{lx}" y="10" width="12" height="12" rx="3" fill="#fff" '
                     f'stroke="{MISMATCH_STROKE}" stroke-width="2"/>')
        parts.append(f'<text x="{lx + 16}" y="20" fill="#334155">불일치</text>')
    for a, b in edges:
        if a not in pos or b not in pos:
            continue
        x1, y1 = pos[a][0] + NODE_W, pos[a][1] + NODE_H / 2
        x2, y2 = pos[b][0], pos[b][1] + NODE_H / 2
        mx = (x1 + x2) / 2
        parts.append(f'<path d="M{x1},{y1} C{mx},{y1} {mx},{y2} {x2 - 2},{y2}" fill="none" '
                     f'stroke="#cbd5e1" stroke-width="1" marker-end="url(#arr)"/>')
    for n, (x, y) in pos.items():
        node = nodes[n]
        fill, stroke, _ = KIND_STYLE[node["kind"]]
        sw = 1
        if n in mismatches:
            stroke, sw = MISMATCH_STROKE, 2.5
        label = node["label"]
        if len(label) > 24:
            label = label[:23] + "…"
        title = html.escape(f'{node["label"]}  {node["title"] or ""}')
        parts.append(f'<g><title>{title}</title><rect x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" '
                     f'rx="5" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'
                     f'<text x="{x + 8}" y="{y + 17}" fill="#0f172a">{html.escape(label)}</text></g>')
    parts.append("</svg>")
    note = ""
    if truncated:
        note = (f'<p style="color:#b45309;font-size:12px">노드가 많아 {max_nodes}개까지만 표시했습니다. '
                '"대상 셀 범위"를 지정하면 해당 셀의 상위 그래프만 볼 수 있습니다.</p>')
    return (note + '<div style="overflow:auto;max-height:640px;border:1px solid #e2e8f0;'
            'border-radius:8px;background:#fff">' + "".join(parts) + "</div>")


def render_mermaid(analysis: Analysis, focus=None, max_nodes: int = 150) -> str:
    nodes, edges, _ = build_graph(analysis, focus, max_nodes)
    ids = {n: f"n{i}" for i, n in enumerate(nodes)}
    lines = ["flowchart LR"]
    for n, node in nodes.items():
        label = node["label"].replace('"', "'")
        shape = f'[("{label}")]' if node["kind"] == "range" else f'["{label}"]'
        lines.append(f"    {ids[n]}{shape}:::{node['kind']}")
    for a, b in edges:
        if a in ids and b in ids:
            lines.append(f"    {ids[a]} --> {ids[b]}")
    for kind, (fill, stroke, _) in KIND_STYLE.items():
        lines.append(f"    classDef {kind} fill:{fill},stroke:{stroke}")
    return "\n".join(lines)
