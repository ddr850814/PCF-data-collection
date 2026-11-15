"""Rules-based single-line isometric PDF rendering with weld capture.

True isometric projection after north-axis/quadrant orientation, with
picture-only departures: not-to-scale warp, plan shear tilting flat
45-degree runs off vertical (RenderOptions.skew_deg / set_skew), and chain
layout in schematic mode. Welds draw as dots -- numbered balloons live in
the weld-tracking overlay -- with page positions captured in the same pass.
Callouts are one style for everything; continuation callouts mark every
declared sheet exit (see continuations.py); dimension lines and their text
lie on isometric planes (see _plan_dim_geometry()).
"""

from __future__ import annotations

import json
import math
import re
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from datetime import date as _date
from itertools import permutations
from pathlib import Path

from . import geometry
from .continuations import Continuation, reference_names
from .cutlist import CutPiece, build_cutlist
from .dims import (DEFAULT_OFFSET_MM, Dimension, flange_face,
                   infer_dimensions, load_sidecar, merge_with_sidecar)
from .model import Component, PcfDocument, Point
from .sheetpack import PagePlan, SheetInfo, plan as _plan_sheets
from .skeys import (classify, connection_from_skey, is_annotation_symbol,
                    is_slip_on_flange, is_weld_coupling, looks_like_flange,
                    reducer_is_eccentric, type_from_skey, valve_subtype)
from .splitmerge import split_pages
from .units import (coord_to_inch_factor, format_bore, format_feet_inches,
                    format_length, format_mm, is_metric_units)
from .welds import (GAP_COUPLING_MAX_INCHES, Weld, apply_number_overrides,
                    component_spools, derive_welds, weld_key)

Vec2 = tuple[float, float]
Vec3 = tuple[float, float, float]

_COS30 = math.cos(math.radians(30))
_SIN30 = math.sin(math.radians(30))

_PAPERS = {
    "tabloid": (431.8, 279.4),   # 11x17 - the industry standard iso sheet
    "a3": (420.0, 297.0),
    "letter": (279.4, 215.9),
    "a4": (297.0, 210.0),
}

_QUADRANT_INDEX = {"NE": 0, "NW": 1, "SW": 2, "SE": 3}


@dataclass
class RenderOptions:
    # 11x17 (tabloid/ledger) is the construction-industry standard iso sheet.
    paper: str = "tabloid"
    north_axis: str = "Y"            # Y | X | Y-decreasing | X-decreasing
    # ISOGEN 默认视角：北向标指向图面左上（西偏北 30°）——对照实测
    # 2105 参考图（Y/SE 组合：+Y→210°左上、+X→330°右上、EL→竖直）
    quadrant: str = "SE"             # NE | NW | SW | SE
    logo_path: str | None = None
    company_lines: list[str] = field(default_factory=list)
    project: str = ""
    job: str = ""
    drawing_no: str = ""
    revision: str = "0"
    date: str = ""
    requirements: list[str] = field(default_factory=list)
    # Numbered NOTES block on the drawing face, e.g. ["PWHT REQUIRED",
    # "CRITICAL CHECK VALVE INSPECTION REQUIRED"].
    notes: list[str] = field(default_factory=list)
    # DRAWING PROPERTIES title-block column, e.g. {"LINE SPEC": "CS150",
    # "MATERIAL": "A106-B"}. Insertion order is print order.
    properties: dict[str, str] = field(default_factory=dict)
    show_dimensions: bool = True
    # Page length (mm) of the bar drawn across a flange's gasket face; 按比例
    # 图按孔径拟合，此值仅在缺孔径时兜底。
    flange_bar_mm: float = 3.8
    show_bom: bool = True
    show_cut_table: bool = True
    detect_ffw: bool = False
    dims_sidecar: str | None = None  # read overrides from here when set
    # Callout crowding: 1.0 default, 1.35 spread, 0.74 pack; clamped [0.4, 3.0]. Only
    # leadered material callouts move; weld dots and pipe labels are fixed to geometry.
    balloon_spread: float = 1.0
    # Balloon every instance of a material item so a fitter can identify each
    # fitting; False = one balloon per BOM row (quieter, fittings unidentified).
    balloon_every_instance: bool = True
    # Schematic drawing: runs compress/stretch so spans land in [nts_min_mm, nts_max_mm],
    # sheet stamped NOT TO SCALE. Warps ONLY the picture; lengths/BOM/dimension text stay true.
    not_to_scale: bool = True
    nts_min_mm: float = 8.0
    nts_max_mm: float = 70.0
    # Visual skew (degrees): tilts flat 45/135/225/315-degree runs off vertical (they would
    # project dead vertical like a riser). Only when such a run exists; picture-only. 0 disables.
    skew_deg: float = 10.0
    # Hybrid mode: grouping attribute naming the spools (e.g. SPOOL-IDENTIFIER). Spool regions
    # get labels; welds joining two spools become field welds ("F.W." tag, "Field Weld" CSV note).
    spool_key: str | None = None
    # 管线清单实时预览：只画图面内容（管线/符号/焊点/尺寸/注释/北向标/
    # NTS 戳），不出图框、标题栏、BOM/切割表等"图纸家具"。
    preview: bool = False
    # Sheet exits: one entry per END-CONNECTION-PIPELINE (point = CO-ORDS, name = reference),
    # built by continuations.build_continuation_map() over the rendered document set.
    continuations: list[Continuation] = field(default_factory=list)
    # 图名 → PcfDocument（批量渲染传入）：分图幻影件需邻图文档画断点对侧虚线幻影；
    # 单图渲染为空时幻影退化为只画接续标记。
    neighbour_docs: dict[str, PcfDocument] | None = None
    # 预览点击改号的编号覆盖表：{weld_key: 编号}，允许重复号，只动覆盖点；
    # 编号不落库（无 sidecar），GUI 经 render_preview_scene 传入。
    weld_number_overrides: dict[str, int] | None = None


@dataclass
class WeldPosition:
    number: int
    label: str
    joint: str | None
    nps: str
    x_mm: float          # weld dot, page mm (top-left origin) -- the anchor
    y_mm: float          # the weld-tracking overlay places its own balloon
    world_x: float
    world_y: float
    world_z: float
    pipeline: str
    description: str | None
    neighbors: str
    uci: str = ""        # UNIQUE-COMPONENT-IDENTIFIER of the weld record
    spool: str = ""      # spool value ("SP-1"; "SP-1 / SP-2" at a boundary)
    notes: str = ""      # "Field Weld" for inter-spool joints (hybrid mode)
    sheet: str = ""      # 分页页号（"p1"；单页输出为空）


@dataclass
class BomLine:
    item: str
    description: str
    nps: str
    cls: str             # component class (PIPE, VALVE, FLANGE, ...)
    qty: int
    length_in: float     # total length for PIPE rows (inches); 0 otherwise
    weight: float
    ucis: str = ""       # ";"-joined UNIQUE-COMPONENT-IDENTIFIERs in the row
    code: str = ""       # original ITEM-CODE (balloon uses sequential item no.)
    sheet: str = ""      # 分页页号（"p1"；单页输出为空）


@dataclass
class Callout:
    item: str            # item-code (material) or cut label (pipe)
    kind: str            # "material" | "pipe"
    x_mm: float          # label centre, page mm
    y_mm: float
    anchor_x_mm: float   # component the callout points at
    anchor_y_mm: float
    leadered: bool


@dataclass
class RenderResult:
    pdf_path: Path
    page_width_mm: float
    page_height_mm: float
    welds: list[WeldPosition]
    dimensions: list[Dimension]
    cut_pieces: list[CutPiece]
    bom: list[BomLine] = field(default_factory=list)
    callouts: list[Callout] = field(default_factory=list)
    continuations: list[Continuation] = field(default_factory=list)
    bom_truncated: bool = False
    cut_truncated: bool = False
    warnings: list[str] = field(default_factory=list)
    pages: int = 1       # NTS 1:1 自动分页后的页数（单页恒 1）


# ---------------------------------------------------------------------------
# Projection


def _orient_plan(x: float, y: float, north_axis: str) -> tuple[float, float]:
    """Map PCF plan coordinates to (east, north)."""
    axis = north_axis.upper()
    if axis == "Y":
        return x, y
    if axis == "X":
        return -y, x
    if axis == "Y-DECREASING":
        return -x, -y
    if axis == "X-DECREASING":
        return y, -x
    return x, y


def _rotate_quadrant(e: float, n: float, quadrant: str) -> tuple[float, float]:
    k = _QUADRANT_INDEX.get(quadrant.upper(), 0)
    for _ in range(k):
        e, n = n, -e
    return e, n


class Projector:
    def __init__(self, options: RenderOptions):
        self.north_axis = options.north_axis
        self.quadrant = options.quadrant
        # Optional not-to-scale warp: monotonic per-axis world remap before projection;
        # directions/topology survive, only drawn span lengths change.
        self.warp = None
        # Plan shear that tilts the flat 45-degree run off vertical; see
        # set_skew(). 0.0 = pure isometric.
        self.skew = 0.0
        # Page transform, set by fit():
        self.scale = 1.0
        self.offset = (0.0, 0.0)
        # 偏心变径正算出口带动小端侧管线的图面平移（世界坐标键）。
        self.point_offsets: dict[Vec3, Vec2] = {}
        # 出图模式：True=按比例（符号铺满实际端距），False=不按比例示意
        # （符号恒等比 sx=sy，端点/尺寸锚随符号端面走）。
        self.proportional = not options.not_to_scale
        # 链式布局（示意）专用：整图示意缩放（管件同乘、搭积木不叠件）与最近点平移
        # 回退（未映射派生点——尺寸盒角点/设备管口——跟随最近已布几何）。
        self.schematic_scale = 1.0
        self.offset_nearest = False
        # NTS 1:1 分页落位（chain_preset）：raw 是真实比例，派生点直接加最近点
        # 平移会把 raw 差量按真实长度画出页外（尺寸偏移盒 sprawled）；置 True 时
        # 最近点回退改用局部压缩比 σ 缩放 raw 差量（链式方向=raw 轴测方向）。
        self.nearest_local_scale = False

    def set_skew(self, degrees: float) -> None:
        """Tilt the plan diagonal that projects vertical by ``degrees``.

        The shear is ``e, n += k * (e - n) / 2``, which leaves ``e - n``
        (and so the page y of a plan point) untouched while sliding ``e + n``
        (page x) by an amount proportional to it. Elevation is not involved
        at all, so risers stay exactly vertical; the diagonal run picks up a
        horizontal component and tilts by ``atan(k * sqrt(3))``.
        """
        self.skew = math.tan(math.radians(degrees)) / math.sqrt(3.0)
        self._axis_dirs_cache = None   # the projected axes just moved

    def _project(self, e: float, n: float, z: float) -> Vec2:
        if self.skew:
            shift = self.skew * (e - n) / 2
            e, n = e + shift, n + shift
        return ((e + n) * _COS30, (e - n) * _SIN30 - z)

    def raw(self, p: Vec3) -> Vec2:
        if self.warp is not None:
            p = self.warp(p)
        e, n = _orient_plan(p[0], p[1], self.north_axis)
        e, n = _rotate_quadrant(e, n, self.quadrant)
        return self._project(e, n, p[2])

    def fit(self, points: list[Vec3], region: tuple[float, float, float, float]) -> None:
        """Scale/centre the projection into a page-mm region (x, y, w, h)."""
        projected = [self.raw(p) for p in points]
        xs = [p[0] for p in projected]
        ys = [p[1] for p in projected]
        span_x = max(xs) - min(xs) if xs else 1.0
        span_y = max(ys) - min(ys) if ys else 1.0
        span_x = max(span_x, 1e-6)
        span_y = max(span_y, 1e-6)
        rx, ry, rw, rh = region
        margin = 0.82  # leave room for balloons and dimensions
        self.scale = min(rw * margin / span_x, rh * margin / span_y)
        cx = (max(xs) + min(xs)) / 2 if xs else 0.0
        cy = (max(ys) + min(ys)) / 2 if ys else 0.0
        self.offset = (rx + rw / 2 - cx * self.scale, ry + rh / 2 - cy * self.scale)

    def page(self, p: Vec3, *, apply_offset: bool = True) -> Vec2:
        sx, sy = self.raw(p)
        page = (sx * self.scale + self.offset[0], sy * self.scale + self.offset[1])
        if apply_offset:
            off = self.point_offsets.get(p)
            if off is None and self.offset_nearest and self.point_offsets:
                # 链式布局：未映射派生点（尺寸盒角点/设备管口/流向箭头/管架等）跟随邻近
                # 几何；最近两点沿连线按投影比例插值平移，避免管线两端平移量不同时掉线。
                best: tuple[float, Vec3, Vec2] | None = None
                second: tuple[float, Vec3, Vec2] | None = None
                for k, d in self.point_offsets.items():
                    dd = ((k[0] - p[0]) ** 2 + (k[1] - p[1]) ** 2
                          + (k[2] - p[2]) ** 2)
                    if best is None or dd < best[0]:
                        second = best
                        best = (dd, k, d)
                    elif second is None or dd < second[0]:
                        second = (dd, k, d)
                if best is not None and second is not None and \
                        second[0] <= best[0] * 25.0 + 1e-9:
                    _d1, k1, t1 = best
                    _d2, k2, t2 = second
                    seg = (k2[0] - k1[0], k2[1] - k1[1], k2[2] - k1[2])
                    denom = seg[0] ** 2 + seg[1] ** 2 + seg[2] ** 2
                    if denom > 1e-12:
                        rel = ((p[0] - k1[0]) * seg[0]
                               + (p[1] - k1[1]) * seg[1]
                               + (p[2] - k1[2]) * seg[2]) / denom
                        rel = max(0.0, min(1.0, rel))
                        off = (t1[0] + (t2[0] - t1[0]) * rel,
                               t1[1] + (t2[1] - t1[1]) * rel)
                if off is None:
                    off = best[2] if best is not None else (0.0, 0.0)
                if best is not None and self.nearest_local_scale:
                    # 分页 1:1：以最近两点的 打包距/raw距 估局部压缩比 σ，把
                    # 派生点的 raw 差量缩到打包尺度（方向沿用 raw 轴测方向）。
                    _dd1, k1, d1 = best
                    r1 = self.raw(k1)
                    sigma = 1.0
                    if second is not None:
                        _dd2, k2, d2 = second
                        r2 = self.raw(k2)
                        num = math.hypot((r2[0] - r1[0]) * self.scale + d2[0] - d1[0],
                                         (r2[1] - r1[1]) * self.scale + d2[1] - d1[1])
                        den = math.hypot((r2[0] - r1[0]) * self.scale,
                                         (r2[1] - r1[1]) * self.scale)
                        if den > 1e-9:
                            sigma = num / den
                    base1 = (r1[0] * self.scale + self.offset[0] + d1[0],
                             r1[1] * self.scale + self.offset[1] + d1[1])
                    return (base1[0] + (page[0] - self.offset[0]
                                       - r1[0] * self.scale) * sigma,
                            base1[1] + (page[1] - self.offset[1]
                                       - r1[1] * self.scale) * sigma)
            dx, dy = off if off is not None else (0.0, 0.0)
            page = (page[0] + dx, page[1] + dy)
        return page

    def north_arrow_angle(self) -> float:
        """Screen angle (degrees, CCW from +x) of the north direction."""
        e, n = _rotate_quadrant(0.0, 1.0, self.quadrant)
        sx, sy = self._project(e, n, 0.0)
        return math.degrees(math.atan2(-sy, sx))


# ---------------------------------------------------------------------------
# Not-to-scale warp
#
# Monotonic per-axis world remap so drawn spans land in [nts_min_mm, nts_max_mm];
# topology preserved exactly (touching stays touching, order kept) and sub-gap
# features (modelled 1/16" weld root gaps) stay untouched. RENDERING ONLY:
# lengths/BOM/dimension text use the true model coordinates.

_NTS_ITERATIONS = 3
# World gaps at or below this (inches) are never stretched: weld root gaps
# and other modelling artifacts must stay visually negligible.
_NTS_TINY_INCHES = 0.30


def _build_nts_warp(points: list[Vec3], scale: float, min_mm: float,
                    max_mm: float, tiny_world: float,
                    keep: list[set[tuple[float, float]]] | None = None,
                    rigid: list[list[tuple[float, float, float]]] | None = None,
                    targets: list[list[tuple[float, float, float]]] | None = None):
    """Monotonic per-axis remap; distances are clamped via the current page
    scale (mm targets / scale = world-unit targets).
    
    ``rigid``: per-axis (lo, hi, s) intervals remapped with one shared slope,
    keeping non-axis-parallel runs' exact direction (independent per-gap
    clamps would swing collinear elbow legs to different angles).
    ``targets``: per-axis (lo, hi, target_mm) intervals mapped to a fixed
    page length -- inline fittings print at symbol size, only straight pipe
    absorbs stretching. ``rigid`` wins over ``keep`` where they overlap;
    ``keep`` intervals (e.g. an eccentric reducer's weld-end offset) stay at
    true proportion, and any gap wholly inside a protected interval is left
    alone.
    """
    from bisect import bisect_right

    min_world = min_mm / scale
    max_world = max(max_mm, min_mm) / scale
    axes: list[tuple[list[float], list[float]]] = []
    for ax in range(3):
        protected = sorted(keep[ax]) if keep else []
        locked = sorted(rigid[ax]) if rigid else []
        forced = sorted(targets[ax]) if targets else []
        vals = sorted({round(p[ax], 6) for p in points})
        src = [vals[0]]
        dst = [vals[0]]
        # ``locked``/``forced`` are sorted and non-overlapping, so only the last interval
        # starting at or before a gap can contain it; binary search keeps this linear.
        lows = [lo for lo, _hi, _s in locked]
        flows = [lo for lo, _hi, _t in forced]
        for v in vals[1:]:
            gap = v - src[-1]
            shared = None
            if locked:
                j = bisect_right(lows, src[-1]) - 1
                if j >= 0 and v <= locked[j][1]:
                    shared = locked[j][2]
            hit = None
            if shared is None and forced:
                # 零件示意尺寸目标：仅作用于不在 rigid 斜线段区间内的
                # 间隙（rigid 区间必须整段共享斜率，否则斜管被剪弯）
                j = bisect_right(flows, src[-1]) - 1
                if j >= 0 and v <= forced[j][1]:
                    lo, hi, tgt = forced[j]
                    # 区间整体映射到 target_mm：内部多段按世界长度等比分配
                    hit = (tgt / scale) * (gap / (hi - lo)) if hi > lo else None
            if shared is not None:
                gap *= shared
            elif hit is not None:
                gap = hit
            else:
                inside = any(lo <= src[-1] and v <= hi
                             for lo, hi in protected)
                if gap > tiny_world and not inside:
                    gap = min(max(gap, min_world), max_world)
            src.append(v)
            dst.append(dst[-1] + gap)
        axes.append((src, dst))

    def _remap(src: list[float], dst: list[float], v: float) -> float:
        if v <= src[0]:
            return dst[0] + (v - src[0])
        if v >= src[-1]:
            return dst[-1] + (v - src[-1])
        i = bisect_right(src, v)
        span = src[i] - src[i - 1]
        t = (v - src[i - 1]) / span if span > 0 else 0.0
        return dst[i - 1] + t * (dst[i] - dst[i - 1])

    def warp(p: Vec3) -> Vec3:
        return (_remap(axes[0][0], axes[0][1], p[0]),
                _remap(axes[1][0], axes[1][1], p[1]),
                _remap(axes[2][0], axes[2][1], p[2]))

    return warp


# A direction component below this (inches) is modelling noise, not a genuinely
# second axis; far under anything a fabricator lays out or the connection tolerance.
_SKEW_AXIS_FLOOR_INCHES = 1.0 / 64.0


def _skew_segments(runs: list[tuple[Vec3, Vec3]], points: list[Vec3],
                   tol: float, inch_factor: float) -> list[tuple[Vec3, Vec3]]:
    """Skewed runs, cut at every model point that sits on them.
    
    Runs parallel to a world axis are excluded, judged absolutely against the
    fabrication floor (_SKEW_AXIS_FLOOR_INCHES), never as a fraction of run
    length: a fraction swings near-axis runs by degrees and splits collinear
    runs (long pipe vs. its short elbow leg) into different treatments.
    Stations are the model's own points, never reconstructed from t.
    """
    from bisect import bisect_left, bisect_right

    floor = _SKEW_AXIS_FLOOR_INCHES / inch_factor
    # One sorted index per axis, so a run only tests points inside its own coordinate
    # slab -- the naive product is runs x points, thousands on a large sheet.
    indexed = [sorted(points, key=lambda p, ax=ax: p[ax]) for ax in range(3)]
    keys = [[p[ax] for p in indexed[ax]] for ax in range(3)]

    segments: list[tuple[Vec3, Vec3]] = []
    for a, b in runs:
        d = geometry.sub(b, a)
        length = geometry.length(d)
        if length < 1e-9:
            continue
        if sum(1 for k in range(3) if abs(d[k]) > floor) < 2:
            continue
        u = tuple(c / length for c in d)
        # Take the slab that excludes the most (not the narrowest axis): a flat offset
        # has zero Z extent, and every point at that elevation is in the Z slab.
        slabs = []
        for ax in range(3):
            lo, hi = sorted((a[ax], b[ax]))
            slabs.append((bisect_left(keys[ax], lo - tol),
                          bisect_right(keys[ax], hi + tol), ax))
        start, stop, pick = min(slabs, key=lambda s: s[1] - s[0])
        box = [sorted((a[ax], b[ax])) for ax in range(3)]
        candidates = [p for p in indexed[pick][start:stop]
                      if all(box[ax][0] - tol <= p[ax] <= box[ax][1] + tol
                             for ax in range(3))]
        # Stations are the model's own points, never reconstructed from t: the warp's
        # breakpoints are model coordinates rounded to six places; an edge off by a float
        # ulp stops covering the gap it was built for.
        stations = {0.0: a, length: b}
        for p in candidates:
            off = geometry.sub(p, a)
            t = sum(off[k] * u[k] for k in range(3))
            if t <= tol or t >= length - tol:
                continue
            if geometry.length(tuple(off[k] - t * u[k] for k in range(3))) <= tol:
                stations[t] = p
        ordered = [stations[t] for t in sorted(stations)]
        segments.extend(zip(ordered, ordered[1:]))
    return segments


def _skew_rigid_spans(proj: Projector, segments: list[tuple[Vec3, Vec3]],
                      tiny_world: float, min_mm: float, max_mm: float,
                      ) -> list[list[tuple[float, float, float]]] | None:
    """Coordinate intervals the warp must remap with one shared slope.
    
    Each skewed segment claims its extent on all three axes with the slope
    that lands it in [nts_min_mm, nts_max_mm]; overlapping claims merge into
    one group under one slope, iterated to a fixed point, so members map as
    p -> s*p + c and directions stay exact. The degenerate single-group
    sheet prints at true proportion. Returns None for orthogonal sheets
    (the common case).
    """
    if not segments:
        return None

    # Page length of each segment as the warp leaves it: rigid map is p -> s*p + c, so
    # its projected length is s * (unwarped projected length) * scale; measure unwarped,
    # then solve for s. Sub-tolerance gaps stay alone, exactly as the clamp does.
    warp, proj.warp = proj.warp, None
    try:
        lengths: list[float | None] = []
        for a, b in segments:
            pa, pb = proj.raw(a), proj.raw(b)
            page = math.hypot(pb[0] - pa[0], pb[1] - pa[1]) * proj.scale
            # Sub-tolerance gaps are modelled weld root gaps: they ride the group's
            # slope but must not *set* it -- asking a 1/16" gap to reach nts_min_mm
            # would blow the group up by orders of magnitude.
            lengths.append(None if page < 1e-9
                           or geometry.dist(a, b) <= tiny_world else page)
    finally:
        proj.warp = warp

    boxes = [[(round(min(a[ax], b[ax]), 6), round(max(a[ax], b[ax]), 6))
              for ax in range(3)] for a, b in segments]

    group = list(range(len(segments)))

    def root(i: int) -> int:
        while group[i] != i:
            group[i] = group[group[i]]
            i = group[i]
        return i

    # "No two groups overlap on any axis" equals "no two claims from different groups
    # overlap" (a group's claims are its members'), so one sorted sweep per axis chains
    # touching claims to the fixed point -- no iterating, no pairwise comparison.
    for ax in range(3):
        order = sorted((i for i in range(len(segments))
                        if boxes[i][ax][1] > boxes[i][ax][0]),
                       key=lambda i: boxes[i][ax])
        reach = float("-inf")
        head = None
        for i in order:
            lo, hi = boxes[i][ax]
            if head is not None and lo < reach:
                group[root(i)] = root(head)
                reach = max(reach, hi)
            else:
                head, reach = i, hi

    # One slope per group: take the smallest slope lifting the group's shortest segment
    # to nts_min_mm, capped so its longest stays under nts_max_mm; when the span ratio
    # exceeds max/min the cap wins and the group prints at true proportion (1.0 when
    # neither bound bites).
    bounds: dict[int, tuple[float, float]] = {}
    for i, page in enumerate(lengths):
        if page is None:
            continue
        r = root(i)
        lo, hi = bounds.get(r, (0.0, float("inf")))
        bounds[r] = (max(lo, min_mm / page), min(hi, max_mm / page))
    slope: dict[int, float] = {}
    for r, (lo, hi) in bounds.items():
        slope[r] = min(max(1.0, lo), hi) if lo <= hi else hi

    spans: list[list[tuple[float, float, float]]] = []
    for ax in range(3):
        claims = sorted((boxes[i][ax][0], boxes[i][ax][1], slope.get(root(i), 1.0))
                        for i in range(len(segments))
                        if boxes[i][ax][1] > boxes[i][ax][0])
        merged: list[tuple[float, float, float]] = []
        for lo, hi, s in claims:
            if merged and lo < merged[-1][1] and s == merged[-1][2]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], hi), s)
            else:
                merged.append((lo, hi, s))
        spans.append(merged)
    return spans


# ---------------------------------------------------------------------------
# Isometric skew of flat 45-degree runs
#
# A run flat to the ground on a 45/135/225/315 bearing projects dead vertical and reads
# like a riser; Projector.set_skew() applies the drafters' visual trick -- a plan shear
# moving only the colliding diagonal ~10 degrees off vertical (risers stay exactly
# vertical, plan axes shift ~2 degrees). Applied only when such a run exists;
# picture-only: dimensions/cut lengths/BOM still come from the true model.

# |dz| under this fraction of a run's length counts as flat to the ground.
_SKEW_FLAT_FRACTION = 0.05
# A flat run drawn within ~8.5 degrees of vertical is reading as a riser.
_SKEW_VERTICAL_TAN = 0.15


def _straight_runs(comps: list[tuple[Component, str]]) -> list[tuple[Vec3, Vec3]]:
    """The straight pipe centrelines the sheet draws, in world coordinates.

    Pipe bodies plus each leg of an elbow -- the elbow's end-to-end chord is
    deliberately excluded, since on a 90 it is a plan diagonal that is never
    drawn (the symbol turns at the corner).
    """
    runs: list[tuple[Vec3, Vec3]] = []
    for comp, _ref in comps:
        cls = classify(comp.type, comp.skey)
        eps = [p.xyz for p in comp.end_points]
        if len(eps) < 2:
            continue
        if cls == "PIPE":
            runs.append((eps[0], eps[1]))
        elif cls == "ELBOW" and comp.centre_points:
            corner = comp.centre_points[0].xyz
            runs.append((eps[0], corner))
            runs.append((corner, eps[1]))
    return runs


def _draws_flat_run_vertically(proj: Projector,
                               runs: list[tuple[Vec3, Vec3]]) -> bool:
    """True when some run flat to the ground projects (near) vertical.

    Judged through the projector as configured, warp included: whether the
    collision actually happens on this sheet is a question about the picture,
    not about the model.
    """
    for a, b in runs:
        delta = geometry.sub(b, a)
        world_len = geometry.length(delta)
        if world_len < 1e-9 or abs(delta[2]) > _SKEW_FLAT_FRACTION * world_len:
            continue  # a riser or a sloped run: vertical is the honest answer
        pa, pb = proj.raw(a), proj.raw(b)
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        if abs(dy) > 1e-9 and abs(dx) <= _SKEW_VERTICAL_TAN * abs(dy):
            return True
    return False


# ---------------------------------------------------------------------------
# 2D helpers


def _v2_sub(a: Vec2, b: Vec2) -> Vec2:
    return (a[0] - b[0], a[1] - b[1])


def _v2_add(a: Vec2, b: Vec2) -> Vec2:
    return (a[0] + b[0], a[1] + b[1])


def _v2_scale(a: Vec2, s: float) -> Vec2:
    return (a[0] * s, a[1] * s)


def _v2_len(a: Vec2) -> float:
    return math.hypot(a[0], a[1])


def _v2_norm(a: Vec2) -> Vec2:
    mag = _v2_len(a)
    if mag < 1e-9:
        return (0.0, 0.0)
    return (a[0] / mag, a[1] / mag)


def _v2_perp(a: Vec2) -> Vec2:
    return (-a[1], a[0])


def _v2_mid(a: Vec2, b: Vec2) -> Vec2:
    return ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)


def _cap_height_mm(font_pts: float) -> float:
    """Helvetica digit cap-height in mm (pt -> mm is 0.3528)."""
    return font_pts * 0.3528 * 0.72


# ---------------------------------------------------------------------------
# Renderer
#
# 双后端：_Canvas 与各 _draw_* 只面对统一绘图接口（fpdf2 同名方法 +
# local_context/rotation/transform 上下文），由具体后端决定落向——_FpdfBackend
# 原样转发 fpdf2 写出 PDF；_SceneBackend 收集自包含 SceneItem（颜色/线宽/字号/
# 变换全部随项快照）供 GUI 主线程用 tk.Canvas 回放矢量预览。


class _FpdfBackend:
    """fpdf2 直通后端：全部调用原样转发给 FPDF 实例。"""

    def __init__(self, pdf):
        self._pdf = pdf

    def __getattr__(self, name):        # set_line_width/line/text/...
        return getattr(self._pdf, name)

    @contextmanager
    def tags(self, *tags_):
        """语义标签只属于矢量场景；PDF 出图时原样忽略。"""
        yield


# 字宽度量单例：布局决策（尺寸文字放线内/外、标签盒宽、_fit_text）依赖
# get_string_width，预览回放必须与 PDF 同一套 Helvetica 度量才能保证同版式。
# fpdf2 导入推迟到首次使用。
_metrics_fpdf = None


def _metrics():
    global _metrics_fpdf
    if _metrics_fpdf is None:
        from fpdf import FPDF
        _metrics_fpdf = FPDF(unit="mm")
        _metrics_fpdf.add_page()
    return _metrics_fpdf


def _metric_safe_text(text: str) -> str:
    """与 _Pdf.normalize_text 同款降级：cp1252 写不出的字符替换掉。"""
    try:
        text.encode("cp1252")
        return text
    except UnicodeEncodeError:
        return text.encode("cp1252", "replace").decode("cp1252")


def _mat_mul(a, b):
    """6 元组 PDF 矩阵复合 a·b（先作用 b 再作用 a）。"""
    a0, a1, a2, a3, a4, a5 = a
    b0, b1, b2, b3, b4, b5 = b
    return (a0 * b0 + a2 * b1,
            a1 * b0 + a3 * b1,
            a0 * b2 + a2 * b3,
            a1 * b2 + a3 * b3,
            a0 * b4 + a2 * b5 + a4,
            a1 * b4 + a3 * b5 + a5)


def _mat_apply(m, p):
    """矩阵 [a b c d e f] 作用于点：x' = a.x + c.y + e, y' = b.x + d.y + f."""
    x, y = p
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


@dataclass
class SceneItem:
    """矢量场景图元：坐标为页 mm（左上原点），绘制状态全部随项自带。

    ``matrix`` 是收集时的完整 CTM（已含 transform(...about x,y) 的平移
    合成）；回放端把图元局部坐标经它映射即可，无需理解状态机。
    """

    kind: str                                  # line|rect|ellipse|polygon|text
    pts: tuple[Vec2, ...] = ()                 # line(2 点)/polygon/text 锚点
    box: tuple[float, float, float, float] | None = None  # rect/ellipse: x,y,w,h
    text: str = ""                             # 文字内容
    size: float = 0.0                          # 字号 pt
    matrix: tuple[float, float, float, float, float, float] | None = None
    color: tuple[int, int, int] = (0, 0, 0)    # 描边/文字颜色
    fill: tuple[int, int, int] | None = None   # 填充色；None=不填充
    stroke: bool = True                        # fpdf style 含 D 才描边；纯 F 填充无边
    width: float = 0.4                         # 线宽 mm
    dash: float | None = None                  # 虚线周期 mm；None=实线
    bold: bool = False                         # 文字加粗（回放端字体权重）
    widths: tuple[float, ...] = ()             # 逐字符字宽 mm（斜切文字排布用）
    tags: tuple = ()                           # 语义标签 ("weld", 3) 等


@dataclass
class SceneModel:
    """一页图纸的矢量场景：渲染后台线程产出，GUI 主线程回放。"""

    page_w_mm: float
    page_h_mm: float
    items: list[SceneItem] = field(default_factory=list)
    ink_bbox_mm: tuple[float, float, float, float] | None = None  # 墨迹范围

    def _grow_ink(self, x0: float, y0: float, x1: float, y1: float) -> None:
        if self.ink_bbox_mm is None:
            self.ink_bbox_mm = (x0, y0, x1, y1)
        else:
            bx0, by0, bx1, by1 = self.ink_bbox_mm
            self.ink_bbox_mm = (min(bx0, x0), min(by0, y0),
                                max(bx1, x1), max(by1, y1))


class _SceneBackend:
    """矢量场景后端：与 fpdf2 调用面同形的收集器。

    当前变换矩阵（CTM）统一处理 rotation/transform 上下文：线/多边形按
    点映射，矩形升格为映射后的多边形，文字则把 CTM 原样记进 SceneItem
    交给回放端（tk 对斜切文字用逐字符映射近似）。
    """

    def __init__(self, scene: SceneModel):
        self.scene = scene
        self._stroke = (0, 0, 0)
        self._fill = (0, 0, 0)
        self._text = (0, 0, 0)
        self._width = 0.4
        self._size = 8.0
        self._bold = False
        self._dash: float | None = None
        self._ctm = None
        self._tags: tuple = ()

    # ---- 状态设置（fpdf2 同名接口）-------------------------------------
    def set_draw_color(self, r, g, b):
        self._stroke = (r, g, b)

    def set_text_color(self, r, g, b):
        self._text = (r, g, b)

    def set_fill_color(self, r, g, b):
        self._fill = (r, g, b)

    def set_line_width(self, w):
        self._width = w

    def set_dash_pattern(self, dash=0, gap=0, **_):
        self._dash = None if not dash else float(dash)

    def set_font(self, family, style="", size=8.0):
        self._size = size
        self._bold = "B" in (style or "")
        _metrics().set_font(family, style=style, size=size)

    def get_string_width(self, s):
        return _metrics().get_string_width(_metric_safe_text(s))

    # ---- 上下文（fpdf2 同名接口）----------------------------------------
    @contextmanager
    def local_context(self):
        snap = (self._stroke, self._fill, self._text, self._width,
                self._size, self._bold, self._dash, self._ctm, self._tags)
        try:
            yield self
        finally:
            (self._stroke, self._fill, self._text, self._width,
             self._size, self._bold, self._dash, self._ctm, self._tags) = snap

    @contextmanager
    def tags(self, *tags_):
        """上下文内收集的图元携带语义标签（焊点/管线/元件/标注）。"""
        prev = self._tags
        self._tags = tuple(tags_)
        try:
            yield
        finally:
            self._tags = prev

    @contextmanager
    def rotation(self, angle, x=None, y=None):
        with self.transform(_rot_matrix(angle, x or 0.0, y or 0.0)):
            yield self

    @contextmanager
    def transform(self, matrix):
        """matrix 为 fpdf Transform；about(x, y) 已由调用方合成进矩阵。"""
        m = (matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f) \
            if hasattr(matrix, "a") else tuple(matrix)
        prev = self._ctm
        self._ctm = m if prev is None else _mat_mul(m, prev)
        try:
            yield self
        finally:
            self._ctm = prev

    # ---- 图元 ------------------------------------------------------------
    def line(self, x1, y1, x2, y2):
        p, q = (x1, y1), (x2, y2)
        if self._ctm is not None:
            p, q = _mat_apply(self._ctm, p), _mat_apply(self._ctm, q)
        self.scene.items.append(SceneItem(
            "line", pts=(p, q), color=self._stroke,
            width=self._width, dash=self._dash, tags=self._tags))
        self.scene._grow_ink(min(p[0], q[0]), min(p[1], q[1]),
                             max(p[0], q[0]), max(p[1], q[1]))

    def ellipse(self, x, y, w, h, style="D", **_):
        fill = self._fill if "F" in style else None
        stroke = "D" in style
        if self._ctm is not None:
            # 变换下的椭圆按外接框四角映射后的包围盒近似（当前调用面
            # 没有这种情形，仅作保险）。
            corners = [_mat_apply(self._ctm, c)
                       for c in ((x, y), (x + w, y), (x, y + h), (x + w, y + h))]
            xs = [c[0] for c in corners]
            ys = [c[1] for c in corners]
            x, y = min(xs), min(ys)
            w, h = max(xs) - x, max(ys) - y
        self.scene.items.append(SceneItem(
            "ellipse", box=(x, y, w, h), color=self._stroke,
            fill=fill, stroke=stroke, width=self._width, tags=self._tags))
        self.scene._grow_ink(x, y, x + w, y + h)

    def rect(self, x, y, w, h, style="D", **_):
        fill = self._fill if "F" in style else None
        stroke = "D" in style
        if self._ctm is not None:
            # 斜切上下文中的矩形（iso_text 白底擦除）升格为映射后的多边形
            pts = tuple(_mat_apply(self._ctm, c)
                        for c in ((x, y), (x + w, y), (x + w, y + h), (x, y + h)))
            self.scene.items.append(SceneItem(
                "polygon", pts=pts, color=self._stroke,
                fill=fill, stroke=stroke, width=self._width, tags=self._tags))
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            self.scene._grow_ink(min(xs), min(ys), max(xs), max(ys))
            return
        self.scene.items.append(SceneItem(
            "rect", box=(x, y, w, h), color=self._stroke,
            fill=fill, stroke=stroke, width=self._width, tags=self._tags))
        self.scene._grow_ink(x, y, x + w, y + h)

    def polygon(self, points, style="D", **_):
        pts = tuple((p[0], p[1]) for p in points)
        if self._ctm is not None:
            pts = tuple(_mat_apply(self._ctm, p) for p in pts)
        fill = self._fill if "F" in style else None
        stroke = "D" in style
        self.scene.items.append(SceneItem(
            "polygon", pts=pts, color=self._stroke, fill=fill,
            stroke=stroke, width=self._width, tags=self._tags))
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        self.scene._grow_ink(min(xs), min(ys), max(xs), max(ys))

    def image(self, name, x=None, y=None, w=0.0, h=0.0, **_):
        # 预览场景不渲染位图（logo 等图纸家具）；只登记范围供 fit 使用。
        if x is not None and y is not None and w and h:
            self.scene._grow_ink(x, y, x + w, y + h)

    def text(self, x, y, txt):
        txt = _metric_safe_text(txt)
        met = _metrics()
        w = met.get_string_width(txt)
        widths = tuple(met.get_string_width(ch) for ch in txt)
        cap = _cap_height_mm(self._size)
        self.scene.items.append(SceneItem(
            "text", pts=((x, y),), text=txt, size=self._size,
            color=self._text, bold=self._bold, matrix=self._ctm,
            widths=widths, tags=self._tags))
        corners = [(x, y - cap), (x + w, y - cap), (x, y), (x + w, y)]
        if self._ctm is not None:
            corners = [_mat_apply(self._ctm, c) for c in corners]
        xs = [c[0] for c in corners]
        ys = [c[1] for c in corners]
        self.scene._grow_ink(min(xs), min(ys), max(xs), max(ys))


def _rot_matrix(angle_deg: float, cx: float, cy: float):
    """绕 (cx, cy) 旋转 angle_deg 度，+角度视觉逆时针（与 fpdf2 出图一致）。

    场景/页面坐标 y 向下，视觉逆时针须取 (cos, -sin, sin, cos)；fpdf2 自己的
    (cos, sin, -sin, cos) 定义在 y 向上的内容空间，数值直接搬进 y 向下场景
    会视觉镜像——此前预览旋转文字（变径就地标签、F.W.）曾与出图反向即此因。
    """
    rad = math.radians(angle_deg)
    cos_a, sin_a = math.cos(rad), math.sin(rad)
    rot = (cos_a, -sin_a, sin_a, cos_a, 0.0, 0.0)
    t1 = (1.0, 0.0, 0.0, 1.0, cx, cy)
    t2 = (1.0, 0.0, 0.0, 1.0, -cx, -cy)
    return _mat_mul(t1, _mat_mul(rot, t2))


class _Canvas:
    """Thin drawing wrapper over fpdf2 with collision tracking for labels."""

    def __init__(self, pdf, bounds: tuple[float, float, float, float] | None = None,
                 spread: float = 1.0):
        self.pdf = pdf
        self.label_boxes: list[tuple[float, float, float, float]] = []
        # Solid geometry drawn so far; dimension placement scores candidate
        # positions against these segments to pick the least-colliding side.
        self.segments: list[tuple[Vec2, Vec2]] = []
        self.bounds = bounds  # labels are kept inside this (x, y, w, h)
        # Balloon crowding control: >1 pushes labels further apart with
        # longer leaders, <1 packs them tighter and tolerates near-touching.
        self.spread = spread
        self.pad = 0.6 * spread

    @contextmanager
    def tagged(self, *tags):
        """本上下文内画出的图元在场景后端携带语义标签；fpdf 后端忽略。"""
        with self.pdf.tags(*tags):
            yield

    def line(self, a: Vec2, b: Vec2, width: float = 0.4, record: bool = True) -> None:
        self.pdf.set_line_width(width)
        self.pdf.line(a[0], a[1], b[0], b[1])
        if record:
            self.segments.append((a, b))

    def dashed(self, a: Vec2, b: Vec2, width: float = 0.2, dash: float = 1.2,
               record: bool = False) -> None:
        self.pdf.set_line_width(width)
        with self.pdf.local_context():
            self.pdf.set_dash_pattern(dash=dash, gap=dash * 0.8)
            self.pdf.line(a[0], a[1], b[0], b[1])
        if record:
            self.segments.append((a, b))

    def polyline(self, pts: list[Vec2], width: float = 0.4, close: bool = False,
                 record: bool = True) -> None:
        self.pdf.set_line_width(width)
        for p, q in zip(pts, pts[1:]):
            self.pdf.line(p[0], p[1], q[0], q[1])
        if close and len(pts) > 2:
            self.pdf.line(pts[-1][0], pts[-1][1], pts[0][0], pts[0][1])
        if record:
            self.segments.extend(zip(pts, pts[1:]))
            if close and len(pts) > 2:
                self.segments.append((pts[-1], pts[0]))

    def circle(self, centre: Vec2, r: float, width: float = 0.3, fill: bool = False) -> None:
        self.pdf.set_line_width(width)
        style = "DF" if fill else "D"
        self.pdf.ellipse(centre[0] - r, centre[1] - r, r * 2, r * 2, style=style)

    def dot(self, centre: Vec2, r: float = 0.8) -> None:
        self.pdf.ellipse(centre[0] - r, centre[1] - r, r * 2, r * 2, style="F")

    def text(self, pos: Vec2, message: str, size: float = 7.0, angle: float = 0.0,
             centred: bool = True) -> None:
        self.pdf.set_font("Helvetica", size=size)
        width = self.pdf.get_string_width(message)
        x, y = pos
        if angle:
            with self.pdf.rotation(angle, x, y):
                self.pdf.text(x - (width / 2 if centred else 0), y, message)
        else:
            self.pdf.text(x - (width / 2 if centred else 0), y, message)

    def iso_text(self, baseline: Vec2, message: str, size: float,
                 along: Vec2, up: Vec2) -> None:
        """Draw text skewed onto an isometric plane, not merely rotated.
        
        ``along`` is the baseline direction and ``up`` the glyph-rise direction --
        two projected world axes -- so glyphs lie on the callout's plane (rotate
        + oblique, AutoCAD isometric-style). Text is centred on ``baseline``
        along it and knocked out white first in the same frame.
        """
        try:
            from fpdf.drawing_primitives import Transform
        except ImportError:  # fpdf2 < 2.8
            from fpdf.drawing import Transform

        self.pdf.set_font("Helvetica", size=size)
        width = self.pdf.get_string_width(message)
        cap = _cap_height_mm(size)
        x, y = baseline
        # PDF matrix [a b c d e f]: x' = a.x + c.y + e, y' = b.x + d.y + f. Text
        # space has x along the baseline and y *down*: (1,0) -> along, (0,1) -> -up.
        matrix = Transform(along[0], along[1], -up[0], -up[1], 0.0, 0.0)
        pad = 0.35
        with self.pdf.transform(matrix.about(x, y)):
            with self.pdf.local_context():
                self.pdf.set_fill_color(255, 255, 255)
                self.pdf.rect(x - width / 2 - pad, y - cap - pad,
                              width + 2 * pad, cap + 2 * pad, style="F")
            self.pdf.text(x - width / 2, y, message)

    def place_label(self, anchor: Vec2, w: float, h: float, away_from: Vec2,
                    min_leader: float = 5.0, avoid_geometry: bool = False,
                    clearance: float = 0.0) -> Vec2:
        """Find a label centre near ``anchor`` that avoids earlier labels.
        
        Candidates sweep rings of directions at increasing radius, natural
        placement (straight out from the drawing centre, shortest leader) first.
        With ``avoid_geometry``, drawn segments are scored too, so a label keeps
        off the centreline it names; ``clearance`` widens the box for that test
        (page mm, before spread). If nothing is clean the least-bad candidate
        wins; the anchor itself is never returned.
        """
        pos, box, _score = self._best_placement(
            anchor, w, h, away_from, min_leader, avoid_geometry, clearance)
        self.label_boxes.append(box)
        return pos

    def place_label_multi(self, anchors: list[Vec2], w: float, h: float,
                          away_from: Vec2, min_leader: float = 5.0,
                          avoid_geometry: bool = False, clearance: float = 0.0,
                          anchor_cost=None) -> tuple[Vec2, Vec2]:
        """Place one label, choosing among several places to point at.
        
        Scores each anchor's best placement (plus ``anchor_cost`` pricing a foot
        landing on a neighbour or weld dot) and takes the winning pair, so a
        callout can step along a run to somewhere clear. Anchors arrive
        best-first, so ties keep the caller's preferred anchor. Returns
        ``(anchor, position)``.
        """
        best: tuple[float, Vec2, Vec2, tuple] | None = None
        for anchor in anchors:
            pos, box, score = self._best_placement(
                anchor, w, h, away_from, min_leader, avoid_geometry, clearance)
            score += anchor_cost(anchor) if anchor_cost else 0.0
            if best is None or score < best[0]:
                best = (score, anchor, pos, box)
            if score <= 0.0:
                break       # the preferred anchor is clean; stop looking
        assert best is not None
        self.label_boxes.append(best[3])
        return best[1], best[2]

    def _best_placement(self, anchor: Vec2, w: float, h: float, away_from: Vec2,
                        min_leader: float, avoid_geometry: bool,
                        clearance: float,
                        ) -> tuple[Vec2, tuple[float, float, float, float], float]:
        """The sweep behind ``place_label``, without committing the box.

        Split out so an anchor can be *scored* before it is chosen. Returns
        ``(position, box, score)``; a score of 0.0 means nothing was hit.
        """
        outward = _v2_norm(_v2_sub(anchor, away_from))
        if outward == (0.0, 0.0):
            outward = (0.7071, -0.7071)
        leader = min_leader * self.spread
        step = 2.5 * self.spread

        # Angular offsets from `outward`, nearest first, alternating sides.
        angles = [0.0]
        for k in range(1, 7):
            angles.extend((k * math.pi / 6, -k * math.pi / 6))
        rings = [leader + r * step * 1.6 for r in range(4)]

        # Cost of *wanting* a candidate, independent of what it collides with:
        # short leaders and small angles off outward are preferred.
        cands = sorted(
            ((dist, ang) for dist in rings for ang in angles),
            key=lambda da: 0.05 * da[0] + 0.4 * abs(da[1]),
        )

        grow = clearance * self.spread
        near = (self._segments_near(anchor, rings[-1] + max(w, h) + grow)
                if avoid_geometry else ())

        best: Vec2 | None = None
        best_box: tuple[float, float, float, float] | None = None
        best_score = 0.0
        for dist, ang in cands:
            ca, sa = math.cos(ang), math.sin(ang)
            direction = (outward[0] * ca - outward[1] * sa,
                         outward[0] * sa + outward[1] * ca)
            candidate = self._clamp(_v2_add(anchor, _v2_scale(direction, dist)), w, h)
            box = (candidate[0] - w / 2, candidate[1] - h / 2, w, h)
            score = 5.0 * sum(1 for other in self.label_boxes
                              if _boxes_overlap(box, other, self.pad))
            if near:
                # Graded so that when nothing is clean the least-bad candidate is the
                # roomiest rather than merely the shortest leader: crossing a line costs
                # twice crowding it. With no clearance the two boxes coincide.
                clear_box = (box[0] - grow, box[1] - grow,
                             box[2] + 2 * grow, box[3] + 2 * grow)
                for p, q in near:
                    if not _seg_hits_box(p, q, clear_box):
                        continue
                    score += 3.0 if _seg_hits_box(p, q, box) else 1.5
            if score <= 0.0:
                return candidate, box, 0.0
            score += 0.05 * dist + 0.4 * abs(ang)
            if best is None or score < best_score:
                best, best_box, best_score = candidate, box, score

        assert best is not None and best_box is not None
        return best, best_box, best_score

    def _segments_near(self, point: Vec2, radius: float) -> list[tuple[Vec2, Vec2]]:
        """Drawn segments whose bounding box comes within `radius` of `point`.

        Placement only ever moves a label a bounded distance from its anchor,
        so scoring it against every segment on the sheet is wasted work -- on
        a busy drawing that is thousands of segments per candidate position.
        """
        px, py = point
        near = []
        for p, q in self.segments:
            if (min(p[0], q[0]) - radius <= px <= max(p[0], q[0]) + radius
                    and min(p[1], q[1]) - radius <= py <= max(p[1], q[1]) + radius):
                near.append((p, q))
        return near

    def _clamp(self, pos: Vec2, w: float, h: float) -> Vec2:
        if self.bounds is None:
            return pos
        bx, by, bw, bh = self.bounds
        x = min(max(pos[0], bx + w / 2), bx + bw - w / 2)
        y = min(max(pos[1], by + h / 2), by + bh - h / 2)
        return (x, y)


def _boxes_overlap(a: tuple[float, float, float, float],
                   b: tuple[float, float, float, float],
                   pad: float = 0.6) -> bool:
    return not (
        a[0] + a[2] + pad < b[0] or b[0] + b[2] + pad < a[0]
        or a[1] + a[3] + pad < b[1] or b[1] + b[3] + pad < a[1]
    )


def render_pdf(
    doc: PcfDocument,
    output_path: str | Path,
    options: RenderOptions | None = None,
) -> RenderResult:
    try:
        from fpdf import FPDF
        from fpdf.errors import FPDFUnicodeEncodingException
    except ImportError as exc:  # pragma: no cover
        raise ImportError(
            "fpdf2 is required for PDF rendering: pip install fpdf2"
        ) from exc

    class _Pdf(FPDF):
        """内置 helvetica 只支持 cp1252：PCF 文本经 errors=replace 读取可能
        混入 U+FFFD 等超范围字符，整张图会因此渲染失败。这里把写不出
        的字符降级为 '?'，保证出图不中断。"""

        def normalize_text(self, text: str) -> str:
            try:
                return super().normalize_text(text)
            except FPDFUnicodeEncodingException:
                return text.encode("cp1252", "replace").decode("cp1252")

    options = options or RenderOptions()
    output_path = Path(output_path)
    page_w, page_h = _PAPERS.get(options.paper.lower(), _PAPERS["tabloid"])
    pdf = _Pdf(orientation="landscape" if page_w >= page_h else "portrait",
               unit="mm", format=(page_h, page_w) if page_w >= page_h else (page_w, page_h))
    pdf.set_auto_page_break(False)
    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    backend = _FpdfBackend(pdf)
    if options.not_to_scale and not options.preview:
        # NTS 固定 1:1：超幅自动分页（测量遍装箱 → 逐页 preset 落位）
        result = _render_pdf_nts(doc, options, backend, pdf, page_w, page_h,
                                 output_path)
    else:
        pdf.add_page()
        result = _render_document(doc, options, backend, page_w, page_h,
                                  pdf_path=output_path)
    pdf.output(str(output_path))
    return result


_NTS_WINDOW_MARGIN = 0.82   # 与 fit()/链式打包同一标注带余量
_COMP_BOX_MARGIN_MM = 4.5   # 非 PIPE 件图面盒外扩（手轮/球标/法兰杠出头）


def _drawable_comps(doc: PcfDocument) -> list[tuple[Component, str]]:
    comps: list[tuple[Component, str]] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append((comp, pipeline.reference))
    return comps


def _usable_window(draw_region: tuple) -> tuple:
    rx, ry, rw, rh = draw_region
    m = _NTS_WINDOW_MARGIN
    uw, uh = rw * m, rh * m
    return (rx + (rw - uw) / 2, ry + (rh - uh) / 2, uw, uh)


def _measure_chain(doc: PcfDocument, comps: list[tuple[Component, str]],
                   draw_region: tuple, options: RenderOptions):
    """NTS 1:1 测量遍：跑一次 fit=False 链式布局产出装箱输入。

    返回 (全点平移表, SheetInfo) 或 None（链式失败 → 调用方回退 warp 整页）。
    测量投影只取方向（scale/offset 不影响链式几何），warp 必须为空。
    """
    proj = Projector(options)
    info: dict = {}
    chain = _chain_layout(proj, doc, comps, draw_region, fit=False, collect=info)
    if chain is None or not chain[0]:
        return None
    delta = chain[0]
    pos = info["pos"]
    materials = doc.materials_descriptions()
    boxes: list = []
    pts_list: list = []
    pipes: list[bool] = []
    pipe_axes: list = []
    bom_keys: list = []
    for comp, _ref in comps:
        pts = [p.xyz for p in (*comp.connection_points(), *comp.centre_points)
               if p.xyz in pos]
        pts_list.append(pts)
        cls = classify(comp.type, comp.skey)
        is_pipe = cls == "PIPE" and len(pts) >= 2
        if is_pipe:
            eps = comp.end_points
            axis = (eps[0].xyz, eps[1].xyz)
            # 零长/极短轴切不开（切点必须在管内部），不作为候选
            if geometry.dist(axis[0], axis[1]) <= 1e-3:
                is_pipe = False
        pipes.append(is_pipe)
        if is_pipe:
            pipe_axes.append(axis)
        else:
            pipe_axes.append(None)
        if pts:
            x0 = min(pos[k][0] for k in pts)
            x1 = max(pos[k][0] for k in pts)
            y0 = min(pos[k][1] for k in pts)
            y1 = max(pos[k][1] for k in pts)
            if cls not in ("PIPE", "WELD", "BOLT", "GASKET"):
                # 手轮/球标/法兰杠会伸出符号体，图面盒留出头余量
                x0 -= _COMP_BOX_MARGIN_MM
                x1 += _COMP_BOX_MARGIN_MM
                y0 -= _COMP_BOX_MARGIN_MM
                y1 += _COMP_BOX_MARGIN_MM
            boxes.append((x0, y0, x1, y1))
        else:
            boxes.append(None)
        bom_keys.append(_bom_row_key(comp, materials))
    sheet_info = SheetInfo(boxes=boxes, pts=pts_list, pipes=pipes,
                           pipe_axes=pipe_axes, bom_keys=bom_keys, pos=pos)
    return delta, sheet_info


def _page_sheet_name(options: RenderOptions, doc: PcfDocument, n: int) -> str:
    stem = options.drawing_no or (Path(doc.source_path).stem if doc.source_path
                                  else "drawing")
    return f"{stem}_SHT_{n}"


def _render_pdf_nts(doc: PcfDocument, options: RenderOptions, backend, pdf,
                    page_w: float, page_h: float, output_path: Path) -> RenderResult:
    """NTS 固定 1:1 输出编排：测量遍网格切片 →（需要时）切管分页 → 逐页落位。

    预览（render_scene）不受影响：仍走 fit 自适应缩放的整段路径。
    链式失败时回退旧整页路径（warp 兜底缩），保证出图不中断。
    """
    warnings: list[str] = []
    border = 6.0
    bom_w = page_w * 0.24 if options.show_bom else 0.0
    title_h = page_h * 0.16
    draw_region = (border + 2, border + 2,
                   page_w - 2 * border - bom_w - 4,
                   page_h - 2 * border - title_h - 4)
    comps = _drawable_comps(doc)
    if not comps:
        pdf.add_page()
        return _render_document(doc, options, backend, page_w, page_h,
                                pdf_path=output_path)
    measured = _measure_chain(doc, comps, draw_region, options)
    if measured is None:
        # 链式失败（极少数拓扑）：回退旧整页路径（含 warp 兜底缩）
        pdf.add_page()
        result = _render_document(doc, options, backend, page_w, page_h,
                                  pdf_path=output_path)
        result.warnings.append(
            "chain layout unavailable; sheet scaled to fit as before")
        return result
    delta, info = measured
    usable = _usable_window(draw_region)
    capacity = _bom_capacity(draw_region[3], options.show_cut_table)
    tol = geometry.model_tolerance(doc)
    plans, cuts, comp_cell, seg_cells, notes = _plan_sheets(
        info, usable, capacity, options.show_cut_table, conn_tol=tol)
    warnings.extend(notes)

    total = len(plans)
    page_names = [_page_sheet_name(options, doc, k + 1) for k in range(total)]
    sheet_no_of = {p.cell: k + 1 for k, p in enumerate(plans)}
    proj0 = Projector(options)   # 与测量遍同参投影（raw 基准一致）

    def pos_of(w: Vec3, pipe_idx: int) -> Vec2:
        """世界点 w（在 pipe_idx 轴上）的打包图面位：直查位表，否则轴向插值。"""
        if w in info.pos:
            return info.pos[w]
        a, b = info.pipe_axes[pipe_idx]
        pa, pb = info.pos.get(a), info.pos.get(b)
        if pa is None or pb is None:
            return proj0.raw(w)
        axis = geometry.sub(b, a)
        l2 = geometry.dot(axis, axis)
        t = 0.0 if l2 <= 1e-18 else max(0.0, min(1.0,
                geometry.dot(geometry.sub(w, a), axis) / l2))
        return (pa[0] + (pb[0] - pa[0]) * t, pa[1] + (pb[1] - pa[1]) * t)

    # （需要时）切管分页：页归属闭包按盒中心/管段轴向中点归格
    subdocs: list = []
    placed_cuts: list[Vec3] = []
    if total > 1:
        def page_of(comp: Component) -> tuple[int, int]:
            oidx = getattr(comp, "_orig_idx", -1)
            segs = seg_cells.get(oidx)
            if segs is None:
                if 0 <= oidx < len(comp_cell):
                    return comp_cell[oidx]
                return plans[0].cell
            eps = comp.end_points
            a, b = info.pipe_axes[oidx]
            axis = geometry.sub(b, a)
            l2 = geometry.dot(axis, axis)
            mid = geometry.scale(geometry.add(eps[0].xyz, eps[1].xyz), 0.5)
            t = 0.0 if l2 <= 1e-18 else geometry.dot(
                geometry.sub(mid, a), axis) / l2
            for t0, t1, cell in segs:
                if t0 - 1e-9 <= t <= t1 + 1e-9:
                    return cell
            return segs[-1][2] if t > 0.5 else segs[0][2]

        subdocs, cells, cut_notes, placed_cuts = split_pages(
            doc, [c[0] for c in cuts], page_of)
        warnings.extend(f"page split: {n}" for n in cut_notes)
        if cells != [p.cell for p in plans]:
            # 归格与切片计划不一致（不应发生）：回退整页，保证出图不中断
            pdf.add_page()
            result = _render_document(doc, options, backend, page_w, page_h,
                                      pdf_path=output_path)
            result.warnings.append(
                "page grouping mismatch; sheet rendered whole")
            return result
    else:
        subdocs = [doc]
    # 落刀点精确键（.4f 重写后的解析值）← (pipe_idx, 计划切点)
    placed_of: dict[tuple[int, Vec3], Vec3] = {}
    for (wp, pipe_idx, _t), placed in zip(cuts, placed_cuts):
        placed_of[(pipe_idx, wp)] = placed

    # 页间断点合成接续声明（声明驱动同通道：箭头 + SEE ISO / SHT n），
    # 幻影件经 neighbour_docs 指向相邻页子文档。
    neighbour_base = dict(options.neighbour_docs or {})
    for k, part in enumerate(subdocs):
        if part is not doc:
            neighbour_base[page_names[k]] = part

    results: list[RenderResult] = []
    for k in range(total):
        plan = plans[k]
        page_deltas: dict[Vec3, Vec2] = {}
        page_pts: set[Vec3] = set()
        cut_here = {pidx for pidx, _a, _b in plan.pieces}
        for cid in plan.comp_ids:
            if cid in cut_here:
                continue                # 被切管只带本页管段端点（下循环）
            for pt in info.pts[cid]:
                page_pts.add(pt)
                if pt in delta:
                    page_deltas[pt] = delta[pt]
        for pidx, wa, wb in plan.pieces:
            for w in (wa, wb):
                page_pts.add(w)
                if w in delta:
                    page_deltas[w] = delta[w]
                elif w not in info.pos:
                    # 内部切点：子文档端点是落刀精确键，显式补 delta（投影仿射）
                    wkey = (pidx, tuple(round(v, 4) for v in w))
                    placed = placed_of.get(wkey)
                    if placed is None:
                        continue
                    pp = pos_of(placed, pidx)
                    base = proj0.raw(placed)
                    page_deltas[placed] = (pp[0] - base[0], pp[1] - base[1])
                    page_pts.add(placed)

        # 页盒（未切件整盒 + 被切管本页段）中心 → 可用窗口中心 = 本页 offset；
        # 盒略超窗口时钳位贴边，防符号余量出界。
        box = None
        for cid in plan.comp_ids:
            if cid in cut_here:
                continue
            box = _union_box(box, info.boxes[cid])
        for pidx, wa, wb in plan.pieces:
            pa, pb = pos_of(wa, pidx), pos_of(wb, pidx)
            box = _union_box(box, (min(pa[0], pb[0]), min(pa[1], pb[1]),
                                   max(pa[0], pb[0]), max(pa[1], pb[1])))
        if box is None:
            box = (usable[0], usable[1], usable[0] + usable[2],
                   usable[1] + usable[3])
        ox = usable[0] + usable[2] / 2 - (box[0] + box[2]) / 2
        oy = usable[1] + usable[3] / 2 - (box[1] + box[3]) / 2
        lo, hi = usable[0] - box[0], usable[0] + usable[2] - box[2]
        if lo <= hi:
            ox = min(max(ox, lo), hi)
        lo, hi = usable[1] - box[1], usable[1] + usable[3] - box[3]
        if lo <= hi:
            oy = min(max(oy, lo), hi)
        offset = (ox, oy)

        conts = [c for c in options.continuations
                 if any(geometry.dist(c.point, pt) <= tol * 4 for pt in page_pts)]
        for wp, nb_cell, pidx, inb in plan.outs:
            nb_no = sheet_no_of.get(nb_cell)
            if nb_no is None:
                continue
            if pidx < 0:
                # 跨格组件接续（短管不切兜底）：无线号/管径引用
                conts.append(Continuation(
                    point=wp, drawing=page_names[nb_no - 1],
                    reference="", inboard=inb, bore=None,
                    sheets=[page_names[nb_no - 1]],
                    sheet_tag=f"SHT {nb_no}"))
                continue
            comp, ref = comps[pidx]
            eps = comp.end_points
            conts.append(Continuation(
                point=placed_of.get((pidx, wp), wp),
                drawing=page_names[nb_no - 1],
                reference=ref or "", inboard=inb,
                bore=eps[0].bore if eps else None,
                sheets=[page_names[nb_no - 1]], sheet_tag=f"SHT {nb_no}"))

        opt = replace(options, continuations=conts,
                      neighbour_docs=dict(neighbour_base))
        pdf.add_page()
        res = _render_document(subdocs[k], opt, backend, page_w, page_h,
                               pdf_path=output_path,
                               chain_preset={"deltas": page_deltas,
                                             "offset": offset},
                               sheet_no=k + 1, sheet_total=total)
        if total > 1:
            res.warnings = [f"[SHT {k + 1}] {w}" for w in res.warnings]
        results.append(res)
    return _merge_results(results, output_path)


def _union_box(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _merge_results(results: list[RenderResult], output_path: Path) -> RenderResult:
    base = results[0]
    return RenderResult(
        pdf_path=output_path,
        page_width_mm=base.page_width_mm,
        page_height_mm=base.page_height_mm,
        welds=[w for r in results for w in r.welds],
        dimensions=[d for r in results for d in r.dimensions],
        cut_pieces=[c for r in results for c in r.cut_pieces],
        bom=[b for r in results for b in r.bom],
        callouts=[c for r in results for c in r.callouts],
        continuations=[c for r in results for c in r.continuations],
        bom_truncated=any(r.bom_truncated for r in results),
        cut_truncated=any(r.cut_truncated for r in results),
        warnings=[w for r in results for w in r.warnings],
        pages=len(results),
    )


def render_scene(
    doc: PcfDocument,
    options: RenderOptions | None = None,
) -> tuple[SceneModel, RenderResult]:
    """矢量场景渲染：与 render_pdf 共用同一套布局/绘制管线，仅后端不同。

    返回 (SceneModel, RenderResult)：场景交给 GUI 主线程回放（不落盘、
    不经 PyMuPDF 栅格化），RenderResult 照常携带焊点页坐标等元数据
    （pdf_path 为 None）。
    """
    options = options or RenderOptions()
    page_w, page_h = _PAPERS.get(options.paper.lower(), _PAPERS["tabloid"])
    scene = SceneModel(page_w_mm=page_w, page_h_mm=page_h)
    result = _render_document(doc, options, _SceneBackend(scene), page_w, page_h)
    return scene, result


def _render_document(
    doc: PcfDocument,
    options: RenderOptions,
    backend,
    page_w: float,
    page_h: float,
    pdf_path: Path | None = None,
    chain_preset: dict | None = None,
    sheet_no: int = 0,
    sheet_total: int = 0,
) -> RenderResult:
    """chain_preset（NTS 1:1 分页路径）：{"deltas": 全点平移表, "offset": 本页
    平移}。给定时跳过 fit/链式/warp，直接按测量遍结果摆件；sheet_no/total
    进标题栏 SHEET 字段并给焊点/BOM/切割件打页号。"""
    warnings: list[str] = []

    # ---- layout regions -------------------------------------------------
    border = 6.0
    bom_w = page_w * 0.24 if (options.show_bom and not options.preview) else 0.0
    title_h = 0.0 if options.preview else page_h * 0.16
    draw_region = (
        border + 2, border + 2,
        page_w - 2 * border - bom_w - 4,
        page_h - 2 * border - title_h - 4,
    )
    spread = max(0.4, min(3.0, options.balloon_spread))
    canvas = _Canvas(backend, bounds=draw_region, spread=spread)

    comps = _drawable_comps(doc)
    if not comps:
        raise ValueError("document contains no drawable components")

    all_points = [p.xyz for comp, _ in comps for p in comp.all_points()]

    # 设备管口（END-CONNECTION-EQUIPMENT 伪组件）：管口点属于本图画面，
    # 纳入投影范围；画法见 _draw_equipment_connections。
    equip_recs = [rec for pipeline in doc.pipelines
                  for rec in pipeline.components
                  if rec.type == "END-CONNECTION-EQUIPMENT"]
    all_points += [p.xyz for rec in equip_recs for p in rec.co_ords]

    # 分图幻影件：邻图断点紧邻零件的关键点纳入投影/NTS 变形范围，幻影与本图共用
    # 同一投影（含非比例伸缩）。只影响画面，不进 BOM/切割清单/避让集合。
    for pxyz in _phantom_points(doc, options):
        all_points.append(pxyz)

    projector = Projector(options)
    reducer_runs = _reducer_runs(doc)

    # 搭积木链式布局（示意模式）：位置由拓扑走链决定；失败回退 warp。
    # 分页路径（chain_preset）已带测量遍结果，这里只落位不再布局。
    chain = None
    floored = False
    chain_on = False
    if chain_preset is not None:
        chain_on = True
        projector.scale = 1.0
        projector.offset = chain_preset["offset"]
        projector.point_offsets.update(chain_preset["deltas"])
        projector.offset_nearest = True
        projector.nearest_local_scale = True
    else:
        projector.fit(all_points, draw_region)
        if options.not_to_scale and len(all_points) > 1:
            try:
                # 超幅：直管/腿档位按缩放比收缩重摆（符号端点对刚性，打包层只压管线
                # 边），缩放比累计迭代到装得下；档位触底仍超幅时才整体兜底缩（含符号，
                # rigid off）。
                pipe_scale = 1.0
                chain = _chain_layout(projector, doc, comps, draw_region)
                for _ in range(4):
                    if chain is None or chain[2] >= 0.98:
                        break
                    pipe_scale *= chain[2]
                    if pipe_scale < 0.15:
                        pipe_scale = 0.15
                        floored = True
                    chain = _chain_layout(projector, doc, comps, draw_region,
                                          pipe_scale=pipe_scale,
                                          rigid_symbols=not floored)
                    if floored:
                        break
            except Exception as exc:
                # 失败必须可见：静默回退会让"链式没生效"伪装成布局缺陷
                chain = None
                warnings.append(f"chain layout failed ({exc}); fell back to warp")
        chain_on = chain is not None and bool(chain[0])
    if chain_on:
        if chain_preset is None:
            projector.point_offsets.update(chain[0])
            # 符号恒定：正常收敛（打包层只压管线边）时 schematic_scale=1.0，符号端点对
            # 不缩、绘制尺寸不变；仅触底兜底时整体同缩保持自洽。未映射派生点走最近点
            # 平移回退。
            projector.schematic_scale = chain[2] if floored else 1.0
            projector.offset_nearest = True
            if chain[1]:
                warnings.append(chain[1])
        # 偏心变径：小端连通侧整体平移到符号出口（按 skey 正算的小端中点，与 warp
        # 通道同一机制）。链式只把变径两端摆上主轴，偏心量由这里叠加，下游管线从
        # 符号小端中点续走。
        for k, dv in _eccentric_small_side_offsets(
                projector, doc, comps, reducer_runs,
                proportional=False).items():
            prev = projector.point_offsets.get(k)
            if prev is not None:
                projector.point_offsets[k] = _v2_add(prev, dv)
    if chain_preset is None and not chain_on and options.not_to_scale and len(all_points) > 1:
        # Fixed-count iteration: warp targets depend on page scale and vice versa,
        # so refit a few times; deterministic for same model and options.
        tiny_world = _NTS_TINY_INCHES / coord_to_inch_factor(doc.units_coords)
        # 零件示意尺寸：直管件区间固定映射为符号自然图面长度，伸缩只由直管吸收
        # （见 _fitting_symbol_targets）；偏心变径的偏心量区间仍按真实比例保护。
        keep = _eccentric_offset_pairs(doc, reducer_runs)
        sym_targets = _fitting_symbol_targets(comps, min_mm=options.nts_min_mm)
        # 只有直管/短管参与非比例伸缩；弯头腿和其他零件保持刚性。
        centrelines = [(comp.end_points[0].xyz, comp.end_points[1].xyz)
                       for comp, _ref in comps
                       if classify(comp.type, comp.skey) == "PIPE" and len(comp.end_points) >= 2]
        skew_segments = _skew_segments(
            centrelines, all_points, geometry.model_tolerance(doc),
            coord_to_inch_factor(doc.units_coords))
        for _ in range(_NTS_ITERATIONS):
            rigid = _skew_rigid_spans(projector, skew_segments, tiny_world,
                                      options.nts_min_mm, options.nts_max_mm)
            projector.warp = _build_nts_warp(
                all_points, projector.scale,
                options.nts_min_mm, options.nts_max_mm, tiny_world, keep,
                rigid, sym_targets)
            projector.fit(all_points, draw_region)

    # Flat 45-degree runs project vertical like a riser; tilt them off vertical so the
    # sheet reads (checked through the projector as configured, warp included). 链式布局
    # 不用 skew：平移表以当时投影为基准，后置 shear 会把映射整体带偏。
    skew_deg = max(0.0, min(20.0, options.skew_deg))
    if (not chain_on and skew_deg
            and _draws_flat_run_vertically(projector, _straight_runs(comps))):
        projector.set_skew(skew_deg)
        projector.fit(all_points, draw_region)

    draw_centre = (draw_region[0] + draw_region[2] / 2,
                   draw_region[1] + draw_region[3] / 2)

    flange_face_pts = _flange_faces(doc)
    if not chain_on and not projector.proportional:
        # 正算偏心变径后，小端连通侧整体接到符号 EXIT；变径自身不偏移。
        # 按比例模式不走 SKEY 符号（手绘梯形两端即真实端点），无需平移。
        projector.point_offsets = _eccentric_small_side_offsets(
            projector, doc, comps, reducer_runs,
            proportional=projector.proportional)
    if not projector.proportional:
        # 小端后直接出图时无下游管件携带平移，接续线也要从符号 EXIT 接出；只改
        # 图面锚点，标注仍打印 PCF 声明坐标（按比例下即声明坐标的真实投影）。
        projector.point_offsets.update(_continuation_eccentric_offsets(
            projector, doc, options.continuations, reducer_runs,
            proportional=projector.proportional))
    if not chain_on and not projector.proportional:
        # 等比符号端面带动连接点（仅示意模式）：管端修剪、焊点吸附、尺寸 goalpost
        # 都随符号实际端面走（端点/尺寸计算"考虑 skey"的统一通道）；按比例模式符号
        # 本就铺满端点无需平移；链式布局已产出全点位置表，二者取一。
        projector.point_offsets.update(_symbol_face_offsets(
            projector, doc, comps, reducer_runs, flange_face_pts))

    # ---- symbol pass -----------------------------------------------------
    _FLANGE_BANDS.clear()      # 每张图重置接带注册表（焊点吸附用）
    _FLANGE_WELDS.clear()
    _FLANGE_BAND_CACHE.clear()  # 接带计算缓存同图失效（proj 偏移已定型）
    # 管架定向管段缓存按图重置：它以 id(proj) 为键，批量渲染时新 Projector 可能
    # 复用已释放对象的 id，命中残留缓存会把管架对到别的图的管段上。
    _PIPE_RUNS_CACHE.clear()
    # 符号生长避让：先登记所有非管件的页坐标占位点（含尚未绘制的），
    # 供 _growth_room 判断每个零件两端还有多少可生长空间。
    _NONPIPE_PAGE_PTS.clear()
    for comp, _ref in comps:
        if classify(comp.type, comp.skey) in ("PIPE", "WELD", "BOLT", "GASKET"):
            continue
        ecc = (classify(comp.type, comp.skey) == "REDUCER"
               and reducer_is_eccentric(comp.type, comp.skey))
        for p in comp.all_points():
            _NONPIPE_PAGE_PTS.append(
                (projector.page(p.xyz,
                                apply_offset=(not ecc) or getattr(
                                    projector, "offset_nearest", False)),
                 id(comp)))
    olet_header_pts = _olet_header_points(doc)
    olet_branch_conn = _olet_branch_connected(doc)
    flange_bar_mm = max(1.0, options.flange_bar_mm)
    anchors: dict[int, Vec2] = {}
    for idx, (comp, ref) in enumerate(comps):
        # 语义标签：直管段按管线号（拆图/分图按 run 定位），其余元件按
        # UCI（缺省退回序号）。PDF 出图后端原样忽略。
        tag = (("pipe", ref) if classify(comp.type, comp.skey) == "PIPE"
               else ("component", comp.uci or str(idx)))
        try:
            with canvas.tagged(*tag):
                anchor = _draw_component(canvas, projector, comp, doc,
                                         flange_face_pts, flange_bar_mm,
                                         olet_header_pts, olet_branch_conn,
                                         reducer_runs)
        except Exception as exc:  # defensive: a bad record must not kill the sheet
            warnings.append(f"component {idx} ({comp.type}): not drawn ({exc})")
            anchor = None
        if anchor is not None:
            anchors[idx] = anchor

    if equip_recs:
        _draw_equipment_connections(canvas, projector, equip_recs, comps,
                                    draw_centre)

    # ---- welds (dots only) ------------------------------------------------
    # Numbered balloons live in the weld-tracking overlay, not on the sheet; every
    # weld's page position is still captured here so the weld-map upload has an anchor.
    welds = derive_welds(doc, detect_ffw=options.detect_ffw,
                         spool_key=options.spool_key)
    # 编号覆盖（预览点击改号/导入 JSON）：就地替换编号，允许重复；
    # 未覆盖的焊点保持派生自动号 1..N
    apply_number_overrides(welds, options.weld_number_overrides)
    weld_positions: list[WeldPosition] = []
    weld_dots: list[Vec2] = []
    for weld in welds:
        dot = projector.page(weld.xyz)
        # 焊点吸附：落在"端点重合法兰"接带簇里的焊点，吸附到法兰焊端
        # （模型上焊口与阀面挤在几毫米内，投影后落在空隙区——应落在带上）
        for top, exitp, uu in _FLANGE_BANDS:
            d = _v2_sub(dot, top)
            s = d[0] * uu[0] + d[1] * uu[1]
            lat = abs(d[0] * -uu[1] + d[1] * uu[0])
            band_len = _v2_len(_v2_sub(exitp, top))
            if -1.5 <= s <= band_len + 1.0 and lat < 2.0:
                dot = exitp
                break
        else:
            # 非重合法兰：焊点吸附到焊端锚点（与管端吸附同一焊接点）
            for w in _FLANGE_WELDS:
                if _v2_len(_v2_sub(dot, w)) < 1.5:
                    dot = w
                    break
        weld_dots.append(dot)
        # 标签附 weld_key：预览单击焊点改号按键定位，重复编号也能区分
        with canvas.tagged("weld", weld.number, weld_key(weld)):
            _draw_weld_symbol(canvas, projector, weld, dot)
        nps = format_bore(weld.bore, doc.units_bore) if weld.bore else "?"
        weld_positions.append(WeldPosition(
            number=weld.number, label=weld.label, joint=weld.joint, nps=nps,
            x_mm=round(dot[0], 2), y_mm=round(dot[1], 2),
            world_x=weld.x, world_y=weld.y, world_z=weld.z,
            pipeline=weld.pipeline, description=weld.description,
            neighbors=" | ".join(weld.neighbors),
            uci=weld.uci or "",
            spool=weld.spool, notes=weld.notes,
            sheet=f"p{sheet_no}" if sheet_no else "",
        ))

    # ---- spool labels + field-weld tags (hybrid mode) ----------------------
    if options.spool_key:
        _draw_spool_annotations(canvas, projector, comps, anchors, doc,
                                options.spool_key, welds, draw_centre)

    # ---- notes (before dims/callouts so both route around the block) -------
    if options.notes:
        _draw_notes(canvas, options.notes, draw_region)

    # ---- NOT TO SCALE stamp (registered early so callouts route around) ----
    if options.not_to_scale:
        _draw_nts_stamp(canvas, draw_region)

    # ---- continuations (before dims/callouts: they are the bigger blocks) ---
    # 分图（同管线断开）→ 邻图紧邻零件虚线幻影 + SEE ISO/SHT/坐标标注；跨管线
    # 接点（不同线号）→ 虚线示意短管 + SEE ISO 对方线号/坐标标注。
    drawn_continuations: list[Continuation] = []
    for cont in options.continuations:
        if not cont.label:
            continue
        with canvas.tagged("cont", cont.reference):
            if _junction_kind(doc, cont, options) == "line":
                _draw_line_tie_in(canvas, projector, cont, doc, options,
                                  draw_centre)
            else:
                _draw_continuation(canvas, projector, cont, doc, options,
                                   draw_centre)
                _draw_phantom_mate(canvas, projector, doc, cont, options)
        drawn_continuations.append(cont)
        if not cont.resolved:
            names = " / ".join(reference_names(cont.reference)) or cont.reference
            warnings.append(
                f"continuation at {_coordinate_text(cont.point, doc, options)} "
                f"is labelled from the PCF's own reference ({names}) "
                "-- no drawing in this render set owns the other side of that "
                "joint")

    # ---- END-POSITION 终点坐标（纯声明驱动，无几何判断、无接续线）----------
    # PCF 用 END-POSITION-OPEN/CLOSED 声明终点：按声明点直接标坐标块；与接图点
    # 重合时不重复（接图块里已带同一份坐标行）。
    tol_end = geometry.model_tolerance(doc)
    for point in _declared_end_positions(doc):
        if any(geometry.dist(point, c.point) <= tol_end
               for c in drawn_continuations):
            continue
        with canvas.tagged("endpos", _coordinate_text(point, doc, options)):
            _draw_end_position(canvas, projector, point, doc, options,
                               draw_centre)

    # ---- dimensions --------------------------------------------------------
    dimensions: list[Dimension] = []
    if options.show_dimensions:
        dimensions = infer_dimensions(doc)
        if options.dims_sidecar:
            dimensions = merge_with_sidecar(dimensions, load_sidecar(options.dims_sidecar))
        for dim in dimensions:
            if dim.suppressed:
                continue
            if not _draw_dimension(canvas, projector, dim, doc.units_coords,
                                   draw_centre):
                warnings.append(
                    f"dimension {dim.landmark_from} -> {dim.landmark_to} "
                    f"({dim.computed_text(doc.units_coords)}) too small to draw "
                    f"-- it is in the sidecar but not on the sheet")

    # ---- callouts -----------------------------------------------------------
    bom_rows = _build_bom(doc, comps)
    cut_pieces = build_cutlist(doc)
    if sheet_no:
        for piece in cut_pieces:
            piece.sheet = f"p{sheet_no}"
    pipe_labels = {c.comp_index: c.label
                   for c in cut_pieces if c.comp_index is not None}
    callouts: list[Callout] = []
    if options.show_bom and not options.preview:
        # 预览不出 BOM 表，项号球标无对照意义（管材切名/材料项号一并跳过）
        callouts = _draw_callouts(canvas, projector, comps, anchors,
                                  bom_rows, pipe_labels, draw_centre,
                                  every_instance=options.balloon_every_instance,
                                  weld_dots=weld_dots,
                                  tolerance=geometry.model_tolerance(doc) * 4)

    # ---- furniture ----------------------------------------------------------
    # 预览模式仍画北向标/注释/NTS 戳（图面内容），只跳过图框/BOM/切割表/标题栏
    # 等图纸家具。
    _draw_north_arrow(canvas, projector, (draw_region[0] + 8, draw_region[1] + 10))
    bom_truncated = cut_truncated = False
    if not options.preview:
        _draw_border(canvas, page_w, page_h, border)
        if options.show_bom:
            bom_truncated, cut_truncated = _draw_bom_panel(
                canvas, bom_rows, cut_pieces, welds, doc,
                x=page_w - border - bom_w, y=border + 2,
                w=bom_w - 2, h=page_h - 2 * border - title_h - 4,
                show_cut_table=options.show_cut_table)
        if bom_truncated or cut_truncated:
            which = " & ".join(name for name, hit in
                               (("BOM", bom_truncated), ("CUT LIST", cut_truncated)) if hit)
            warnings.append(f"{which} truncated on the sheet -- split this drawing "
                            "into smaller spools")
            _draw_overflow_warning(canvas, draw_region, which)
        _draw_title_block(canvas, doc, options, x=border, y=page_h - border - title_h,
                          w=page_w - 2 * border, h=title_h,
                          sheet=(sheet_no, sheet_total) if sheet_total > 1 else None)

    bom_lines = [
        BomLine(item=str(r["item"]), description=r["description"], nps=r["nps"],
                cls=r["class"], qty=r["qty"], length_in=round(r["length_in"], 3),
                weight=round(r["weight"], 3),
                ucis=";".join(sorted(r["_ucis"])), code=r["code"],
                sheet=f"p{sheet_no}" if sheet_no else "")
        for r in sorted(bom_rows.values(), key=lambda r: _item_sort_key(r["item"]))
    ]

    return RenderResult(
        pdf_path=pdf_path,
        page_width_mm=page_w,
        page_height_mm=page_h,
        welds=weld_positions,
        dimensions=dimensions,
        cut_pieces=cut_pieces,
        bom=bom_lines,
        callouts=callouts,
        continuations=drawn_continuations,
        bom_truncated=bom_truncated,
        cut_truncated=cut_truncated,
        warnings=warnings,
        pages=sheet_total or 1,
    )


# ---------------------------------------------------------------------------
# Spool annotations (hybrid distance-split mode)


def _box_edge_toward(pos: Vec2, target: Vec2, w: float, h: float) -> Vec2:
    """Where the ray pos->target crosses the w x h box centred on pos."""
    d = _v2_sub(target, pos)
    tx = (w / 2) / abs(d[0]) if abs(d[0]) > 1e-9 else float("inf")
    ty = (h / 2) / abs(d[1]) if abs(d[1]) > 1e-9 else float("inf")
    return _v2_add(pos, _v2_scale(d, min(tx, ty, 1.0)))


# The sheet's reading size, in points: set by the fab/spool label (the mark a fabricator
# must read first); dimension values and title-block entries held to the same size.
_SHEET_FONT = 8.0

# Whitespace a spool label must keep from any drawn line, in page mm before spread:
# not intersecting is not enough -- a box that grazes a centreline reads as crowded.
_SPOOL_LABEL_CLEARANCE_MM = 2.2

# How far off the centreline an "F.W." tag stands, in page mm before spread:
# clear of the run line and its weld dot, close enough to read as that joint's tag.
_FW_STANDOFF_MM = 3.6


def _draw_spool_annotations(canvas: _Canvas, proj: Projector,
                            comps: list[tuple[Component, str]],
                            anchors: dict[int, Vec2], doc: PcfDocument,
                            spool_key: str, welds: list[Weld],
                            draw_centre: Vec2) -> None:
    """Label each spool region and tag every inter-spool joint "F.W.".

    The dashed-outline labels distinguish spool ids from the square material
    callouts; the F.W. tags declare on the sheet face the field welds that
    the welds CSV declares in its notes column. Weld dots never move.
    """
    spool_of = component_spools(doc, spool_key)
    groups: dict[str, list[int]] = {}
    for idx, value in enumerate(spool_of):
        if value is not None:
            groups.setdefault(value, []).append(idx)

    for value, members in groups.items():
        pts = []
        for idx in members:
            point = anchors.get(idx)
            if point is None:
                comp_points = comps[idx][0].all_points()
                if not comp_points:
                    continue
                point = proj.page(comp_points[0].xyz)
            pts.append(point)
        if not pts:
            continue
        centroid = (sum(p[0] for p in pts) / len(pts),
                    sum(p[1] for p in pts) / len(pts))
        anchor = min(pts, key=lambda p: _v2_len(_v2_sub(p, centroid)))
        canvas.pdf.set_font("Helvetica", size=_SHEET_FONT)
        w = canvas.pdf.get_string_width(value) + 4.0
        h = 5.0
        # Spool labels are the first thing placed on the sheet, so without geometry
        # avoidance the very first candidate is accepted whatever it lands on --
        # struck through by the run they name.
        pos = canvas.place_label(anchor, w, h, draw_centre, min_leader=12.0,
                                 avoid_geometry=True,
                                 clearance=_SPOOL_LABEL_CLEARANCE_MM)
        corners = [(pos[0] - w / 2, pos[1] - h / 2), (pos[0] + w / 2, pos[1] - h / 2),
                   (pos[0] + w / 2, pos[1] + h / 2), (pos[0] - w / 2, pos[1] + h / 2)]
        # Knock out whatever is behind the text before drawing it: even a
        # near miss leaves a centreline running through the letters.
        with canvas.pdf.local_context():
            canvas.pdf.set_fill_color(255, 255, 255)
            canvas.pdf.rect(pos[0] - w / 2, pos[1] - h / 2, w, h, style="F")
        for a, b in zip(corners, corners[1:] + corners[:1]):
            canvas.dashed(a, b, width=0.25, dash=0.9)
        canvas.text((pos[0], pos[1] + 1.0), value, size=_SHEET_FONT)
        edge = _box_edge_toward(pos, anchor, w, h)
        if _v2_len(_v2_sub(anchor, edge)) > 1.0:
            canvas.dashed(edge, anchor, width=0.2, dash=0.9)

    for weld in welds:
        if weld.notes != "Field Weld":
            continue
        dot = proj.page(weld.xyz)
        canvas.pdf.set_font("Helvetica", size=6)
        w = canvas.pdf.get_string_width("F.W.") + 1.6
        h = 3.2
        along = None
        if weld.insert_toward is not None:
            along = _v2_norm(_v2_sub(proj.page(weld.insert_toward), dot))
        if along in (None, (0.0, 0.0)):
            # No pipe-like neighbour to take a direction from: fall back to
            # the free-placed horizontal tag with a leader.
            pos = canvas.place_label(dot, w, h, draw_centre, min_leader=3.5)
            canvas.text((pos[0], pos[1] + 0.8), "F.W.", size=6)
            edge = _box_edge_toward(pos, dot, w, h)
            if _v2_len(_v2_sub(dot, edge)) > 2.0:
                canvas.line(edge, dot, width=0.2, record=False)
            continue

        # Set parallel to the centreline and stood off it, the way a field weld is
        # tagged by hand; free-placed, four letters lying over a centreline are
        # not readable.
        angle = math.degrees(math.atan2(-along[1], along[0]))
        if angle > 90 or angle < -90:
            angle += 180          # never upside down
        off = _v2_perp(along)
        away = _v2_sub(dot, draw_centre)
        if off[0] * away[0] + off[1] * away[1] < 0:
            off = _v2_scale(off, -1.0)      # stand it off outwards
        pos = _v2_add(dot, _v2_scale(off, _FW_STANDOFF_MM * canvas.spread))
        # Knock the background out before drawing, rotated *with* the text: squared
        # to the page the knockout is much taller than the glyphs at 30 degrees and
        # takes a bite out of the run line and the weld dot next to it.
        with canvas.pdf.local_context():
            canvas.pdf.set_fill_color(255, 255, 255)
            with canvas.pdf.rotation(angle, pos[0], pos[1]):
                canvas.pdf.rect(pos[0] - w / 2, pos[1] - h / 2, w, h, style="F")
        # Reserved against the upright AABB, which is the conservative box.
        box = _text_aabb(pos, w, h, angle)
        canvas.text((pos[0], pos[1] + 0.8), "F.W.", size=6, angle=angle)
        canvas.label_boxes.append(box)


# ---------------------------------------------------------------------------
# Continuation callouts
#
# Every sheet exit: run carried a little past the hand-off with an arrowhead, plus a
# block naming the NEXT RENDERED DRAWING it continues onto plus the point's plant
# coordinates. Coordinates are printed whether or not the name resolved -- a fitter
# ties the two sheets together in the field by them, not by name. The PCF knows only
# which pipeline/spool carries on, never sheets; see continuations.py for the
# cross-matching contract.

_CONT_FONT = 5.2
_CONT_CAP = _cap_height_mm(_CONT_FONT)
_CONT_STUB_MM = 5.5          # how far the run is carried past the hand-off
_CONT_PAD_MM = 1.1           # padding inside the block
_CONT_LINE_MM = 2.6          # line pitch inside the block
_CONT_CLEARANCE_MM = 1.6     # whitespace the block keeps from drawn geometry


def _coordinate_lines(point: Vec3, doc: PcfDocument,
                      options: RenderOptions) -> list[str]:
    """The hand-off point as plant coordinates, in two printable lines.

    Written as North / East / Elevation because that is what the model's
    coordinates *are* once ``north_axis`` has said which way the plan is
    turned -- the same orientation the north arrow on this sheet shows. The
    viewing quadrant is deliberately not applied: it picks the corner the
    drawing is seen from and moves nothing in the plant.
    """
    east, north = _orient_plan(point[0], point[1], options.north_axis)

    def fmt(value: float) -> str:
        return format_length(value, doc.units_coords, force_feet=True)

    return [f"N {fmt(north)}   E {fmt(east)}", f"EL {fmt(point[2])}"]


def _coordinate_text(point: Vec3, doc: PcfDocument,
                     options: RenderOptions) -> str:
    return "  ".join(_coordinate_lines(point, doc, options))


def _draw_continuation(canvas: _Canvas, proj: Projector, cont: Continuation,
                       doc: PcfDocument, options: RenderOptions,
                       draw_centre: Vec2) -> None:
    """分图接续标记：SEE ISO 管线号/SHT/坐标文字块（仅分图型）。

    断口处不再画箭头/短管——虚线幻影接续线已表达方向与对侧内容。
    """
    at = proj.page(cont.point)
    lines = _cont_callout_lines(doc, cont, options)
    canvas.pdf.set_font("Helvetica", size=_CONT_FONT)
    text_w = max(canvas.pdf.get_string_width(line) for line in lines)
    w = text_w + 2 * _CONT_PAD_MM
    h = len(lines) * _CONT_LINE_MM + 2 * _CONT_PAD_MM - (_CONT_LINE_MM - _CONT_CAP)
    pos = canvas.place_label(at, w, h, draw_centre, min_leader=7.0,
                             avoid_geometry=True, clearance=_CONT_CLEARANCE_MM)

    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.3)
        canvas.pdf.rect(pos[0] - w / 2, pos[1] - h / 2, w, h, style="DF")
    top = pos[1] - h / 2 + _CONT_PAD_MM + _CONT_CAP
    for n, line in enumerate(lines):
        canvas.text((pos[0], top + n * _CONT_LINE_MM), line, size=_CONT_FONT)

    edge = _box_edge_toward(pos, at, w, h)
    if _v2_len(_v2_sub(at, edge)) > 0.8:
        canvas.line(edge, at, width=0.15)


def _declared_end_positions(doc: PcfDocument) -> list[Vec3]:
    """END-POSITION-OPEN/CLOSED 声明的管线终点坐标。

    纯声明驱动：PCF 在哪声明终点，图面就在哪标坐标，不做任何几何
    判断（不查连通、不查是否管段中部）——用户定稿要求。
    """
    return [pt.xyz for comp in doc.all_components()
            if comp.type.upper().startswith("END-POSITION")
            for pt in comp.co_ords]


def _draw_end_position(canvas: _Canvas, proj: Projector, point: Vec3,
                       doc: PcfDocument, options: RenderOptions,
                       draw_centre: Vec2) -> None:
    """终点坐标标注：白底坐标块 + 引线钉在声明点上，不画接续线。"""
    at = proj.page(point)
    lines = _coordinate_lines(point, doc, options)
    canvas.pdf.set_font("Helvetica", size=_CONT_FONT)
    text_w = max(canvas.pdf.get_string_width(line) for line in lines)
    w = text_w + 2 * _CONT_PAD_MM
    h = len(lines) * _CONT_LINE_MM + 2 * _CONT_PAD_MM - (_CONT_LINE_MM - _CONT_CAP)
    pos = canvas.place_label(at, w, h, draw_centre, min_leader=7.0,
                             avoid_geometry=True, clearance=_CONT_CLEARANCE_MM)

    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.3)
        canvas.pdf.rect(pos[0] - w / 2, pos[1] - h / 2, w, h, style="DF")
    top = pos[1] - h / 2 + _CONT_PAD_MM + _CONT_CAP
    for n, line in enumerate(lines):
        canvas.text((pos[0], top + n * _CONT_LINE_MM), line, size=_CONT_FONT)

    edge = _box_edge_toward(pos, at, w, h)
    if _v2_len(_v2_sub(at, edge)) > 0.8:
        canvas.line(edge, at, width=0.15)


_PHANTOM_STUB_MM = 10.0   # 幻影直管虚线延伸段的图面长度上限
_LINE_TIE_STUB_MM = 6.0   # 跨管线虚线示意短管的图面长度


def _sheet_line_refs(doc: PcfDocument) -> set[str]:
    """本图所有管线号（PIPELINE-REFERENCE 属性）。"""
    return {p.reference.strip() for p in doc.pipelines
            if p.reference and p.reference.strip()}


def _junction_kind(doc: PcfDocument, cont: Continuation,
                   options: RenderOptions) -> str:
    """接点类型："split"=同管线分图断点（画幻影件），"line"=跨管线接点。

    两侧图纸共享任一管线号 → 分图幻影；否则是两条独立管线的对接 →
    画虚线短管 + SEE ISO 对方线号文字。未解析（对侧图不在渲染集）时
    用 PCF 自己声明的引用判断：引用里的线号不等于本图管线号 → 跨管线。
    """
    own = _sheet_line_refs(doc)
    if cont.sheet_tag is not None:
        return "split"    # 人工合成的页间断点，无需查邻图
    if cont.resolved:
        for name in cont.sheets:
            ndoc = (options.neighbour_docs or {}).get(name)
            if ndoc is None:
                continue
            if own & _sheet_line_refs(ndoc):
                return "split"
        return "line"
    nos = _cont_ref_nos(cont.reference)
    if nos and not (own & set(nos)):
        return "line"
    if not nos:
        return "line"    # 什么都不接 → 标注退化为只有坐标
    return "split"


def _sheet_tag(name: str) -> str:
    """邻图图名里的页号（SHT n）；取不到就退回图名本身。

    本工具的分图文件命名是 <母图>_SHT_<n> / <母图>_SPL_<值>
    （见 project.SHEET_RE / KEY_SHEET_RE），页号直接从图名提取。
    """
    m = re.search(r"[_\s-]SHT[_\s-]?(\d+)$", name, re.IGNORECASE)
    if m:
        return f"SHT {int(m.group(1))}"
    m = re.search(r"_SPL_(.+)$", name, re.IGNORECASE)
    if m:
        return f"SHT {m.group(1)}"
    return name


def _declared_ref_lines(ref: str) -> list[str]:
    """声明的 CONNECTION-REFERENCE 逐字成行：$ = 注释框内换行。

    ISOGEN 里 ``$`` 是打印字段内的换行符，所以 ``2105\\21053000$
    2105-BR3S-31051`` 在注释框里就是两行（每行按惯例去掉路径前缀）。
    同点重复声明以 ``" / "`` 合并时逐条展开，各占自己的行。
    """
    out: list[str] = []
    for part in ref.split(" / "):
        out.extend(reference_names(part))
    return out


def _cont_callout_lines(doc: PcfDocument, cont: Continuation,
                        options: RenderOptions) -> list[str]:
    """接点标注文字（图面规范）：

    SEE ISO 块逐字来自 PCF 声明（首行冠 SEE ISO，$ 起的段各占一行）；
    分图型再补 "SHT n" 页号行（页号按渲染实际情况取自图名）；末尾是
    声明 CO-ORDS 的坐标行。声明为空时退化为只有坐标。
    """
    coords = _coordinate_lines(cont.point, doc, options)
    names = _declared_ref_lines(cont.reference)
    lines = [f"SEE ISO {names[0]}", *names[1:]] if names else []
    if _junction_kind(doc, cont, options) == "split":
        if cont.sheet_tag is not None:
            tags = [cont.sheet_tag]
        else:
            tags = [_sheet_tag(n) for n in cont.sheets]
        if tags:
            tag_line = " / ".join(tags)
            # 无线号引用（跨格组件接续）时页号行冠 SEE ISO，与声明型同风格
            lines.append(tag_line) if lines else lines.append(
                f"SEE ISO {tag_line}")
    if not lines:
        return list(coords)    # 不接任何东西 → 只写坐标
    return [*lines, *coords]


def _neighbour_mates(doc: PcfDocument, cont: Continuation,
                     options: RenderOptions) -> list[Component]:
    """邻图在接点上的紧邻零件（分图幻影件的素材，每图至多一两件）。"""
    tol = geometry.model_tolerance(doc) * 4
    # 本图 CONTINUATION 副本（导出器双写的对侧边界件）：已按虚线画出，
    # 同 UCI 的邻图零件不再重复画幻影
    local_copies = {c.uci for c in doc.real_components()
                    if c.is_continuation_copy() and c.uci}
    out: list[Component] = []
    seen: set[int] = set()
    for name in cont.sheets:
        ndoc = (options.neighbour_docs or {}).get(name)
        if ndoc is None:
            continue
        ncomps = ndoc.real_components()
        for idx in geometry.components_at_point(ncomps, cont.point, tol):
            comp = ncomps[idx]
            if comp.uci and comp.uci in local_copies:
                continue
            if id(ncomps[idx]) not in seen:
                seen.add(id(ncomps[idx]))
                out.append(ncomps[idx])
    return out


def _phantom_points(doc: PcfDocument, options: RenderOptions) -> list[Vec3]:
    """需要纳入本图投影范围的幻影件关键点（只统计分图型接点）。

    跨管线接点的对方几何一律不画，其坐标也不进投影范围。"""
    pts: list[Vec3] = []
    for cont in options.continuations:
        if not cont.label or not cont.resolved:
            continue
        if _junction_kind(doc, cont, options) != "split":
            continue
        for comp in _neighbour_mates(doc, cont, options):
            pts.extend(p.xyz for p in comp.all_points())
    return pts


def _draw_phantom_mate(canvas: _Canvas, proj: Projector, doc: PcfDocument,
                       cont: Continuation, options: RenderOptions) -> None:
    """分图幻影件：把邻图断点对侧的紧邻零件按虚线画进本图。

    直管 → 虚线延伸段；管件 → SKEY 符号虚线（幻影）；焊点 → 虚线小圆。
    纯画面元素：不进 BOM/切割清单、不注册避让集合、不产生气球/焊号。
    """
    at3 = cont.point
    at = proj.page(at3)
    base_raw = proj.raw(at3)
    # 幻影映射锚定断点：远端点不走 offset_nearest 回退（可能锚到世界
    # 最近但方向错误的点，把 raw 差量画歪——BR3R-30001 立管切点案例）；
    # σ 取断点→inboard 的页距/raw距局部压缩比，幻影按打包尺度贴断点展开。
    sigma = proj.scale
    if cont.inboard is not None:
        pg = proj.page(cont.inboard)
        pg_l = _v2_len(_v2_sub(pg, at))
        raw_l = _v2_len(_v2_sub(proj.raw(cont.inboard), base_raw))
        if pg_l > 1e-6 and raw_l > 1e-12:
            sigma = pg_l / raw_l

    def pmap(w: Vec3) -> Vec2:
        if w in proj.point_offsets:
            return proj.page(w)
        r = proj.raw(w)
        return (at[0] + (r[0] - base_raw[0]) * sigma,
                at[1] + (r[1] - base_raw[1]) * sigma)

    for comp in _neighbour_mates(doc, cont, options):
        cls = classify(comp.type, comp.skey)
        if cls in ("BOLT", "GASKET", "SUPPORT", "MESSAGE"):
            continue
        eps = comp.end_points
        if cls == "PIPE" and len(eps) >= 2:
            # 直管虚线延伸段：沿对方直管走向画示意（图面长度封顶），
            # 与本图断点处的实线端头相接
            near = 0 if (geometry.dist(eps[0].xyz, at3)
                         <= geometry.dist(eps[-1].xyz, at3)) else -1
            far = eps[1] if near == 0 else eps[-2]
            pa = pmap(eps[near].xyz)
            pb = pmap(far.xyz)
            d = _v2_sub(pb, pa)
            ln = _v2_len(d)
            if ln > 1e-6:
                end = _v2_add(at, _v2_scale(d, min(ln, _PHANTOM_STUB_MM) / ln))
                canvas.dashed(at, end, width=0.35)
            continue
        if cls == "WELD" or comp.type.upper().startswith("WELD"):
            # 幻影焊点：虚线小圆（ISOGEN 幻影焊缝画法）
            canvas.pdf.set_line_width(0.25)
            with canvas.pdf.local_context():
                canvas.pdf.set_dash_pattern(dash=0.9, gap=0.7)
                canvas.pdf.ellipse(at[0] - 0.7, at[1] - 0.7, 1.4, 1.4, style="D")
            continue
        if len(eps) >= 2:
            # 幻影管件：SKEY 符号虚线，符号首点对齐断点（与设备管口
            # 虚线画法同模式，法兰等符号按真实朝向落在幻影位置）
            e0 = pmap(eps[0].xyz)
            e1 = pmap(eps[-1].xyz)
            u = _v2_norm(_v2_sub(e1, e0))
            if u == (0.0, 0.0):
                continue
            across = _iso_cross_dir(proj, u)
            run3 = tuple(eps[-1].xyz[i] - eps[0].xyz[i] for i in range(3))
            if _draw_symbol_2pt(canvas, comp.skey, e0, e1, proj=proj,
                                run3=run3, wp=at3, across=across,
                                proportional=proj.proportional,
                                width=0.3, dashed=True):
                continue
            # ASC 库缺符号时的手写虚线回退：弯头画两腿折线，三通/支管台
            # 画主管加支管腿，其余退化为两端虚线
            if cls == "ELBOW" and comp.centre_points:
                corner = pmap(comp.centre_points[0].xyz)
                canvas.dashed(e0, corner, width=0.3)
                canvas.dashed(corner, e1, width=0.3)
            elif cls in ("TEE", "OLET") and comp.centre_points:
                centre = pmap(comp.centre_points[0].xyz)
                canvas.dashed(e0, centre, width=0.3)
                canvas.dashed(centre, e1, width=0.3)
                for bp in (*comp.branch1_points, *comp.branch2_points):
                    canvas.dashed(centre, pmap(bp.xyz), width=0.3)
            else:
                canvas.dashed(e0, e1, width=0.3)


def _junction_out_dir(doc: PcfDocument, cont: Continuation,
                      proj: Projector, draw_centre: Vec2) -> Vec2:
    """跨管线虚线短管的延伸方向（图面向量，从接点指向外侧）。

    优先 inboard→接点；两者重合时（接点常落在场焊的两重合端点上）
    改用本图紧邻零件的本体点→接点方向（OLET 取支管腿、直管/弯头取
    远端→接点）；再退化才退回图心方向。
    """
    at = proj.page(cont.point)
    if cont.inboard is not None and geometry.dist(cont.inboard,
                                                  cont.point) > 1e-9:
        v = _v2_norm(_v2_sub(at, proj.page(cont.inboard)))
        if v != (0.0, 0.0):
            return v
    tol = geometry.model_tolerance(doc) * 4
    lcomps = doc.real_components()
    best: tuple[float, Vec3] | None = None
    for idx in geometry.components_at_point(lcomps, cont.point, tol):
        comp = lcomps[idx]
        body: Vec3 | None = None
        if comp.centre_points:
            body = comp.centre_points[0].xyz
        elif comp.branch1_points:
            body = comp.branch1_points[0].xyz
        elif len(comp.end_points) >= 2:
            near = min(comp.end_points,
                       key=lambda p: geometry.dist(p.xyz, cont.point))
            if geometry.dist(near.xyz, cont.point) > tol:
                body = near.xyz
        if body is None or geometry.dist(body, cont.point) <= tol:
            continue
        d = geometry.dist(body, cont.point)
        if best is None or d < best[0]:
            best = (d, body)
    if best is not None:
        v = _v2_norm(_v2_sub(at, proj.page(best[1])))
        if v != (0.0, 0.0):
            return v
    return _v2_norm(_v2_sub(at, draw_centre)) or (0.7071, -0.7071)


def _cont_ref_nos(ref: str) -> list[str]:
    """接点声明引用里的对方线号（引用串可能以 ' / ' 连接多条）。"""
    out: list[str] = []
    for part in ref.split(" / "):
        names = reference_names(part)
        if names:
            out.append(names[-1])
    return out


def _tie_in_geometry_present(doc: PcfDocument, cont: Continuation) -> bool:
    """跨管线接点上是否已有"延续示意"几何，接续短管可省。

    双写边界段（CONTINUATION 副本，已按虚线画出）或本图支管台/三通
    本身就说明"线从这里接出去"，再画示意短管就重复了。管端对管端的
    真对接（接点上只有管端/焊缝）没有这种示意，短管照画。匹配用
    all_points：支管台的接图点常是其 CENTRE-POINT（主管中心线），
    不在 connection_points 里。
    """
    tol = geometry.model_tolerance(doc) * 4
    for comp in doc.real_components():
        cls = classify(comp.type, comp.skey)
        if cls in ("WELD", "BOLT", "GASKET", "SUPPORT", "MESSAGE"):
            continue
        if not (comp.is_continuation_copy() or cls in ("OLET", "TEE", "CROSS")):
            continue
        if any(geometry.dist(p.xyz, cont.point) <= tol
               for p in comp.all_points()):
            return True
    return False


def _draw_line_tie_in(canvas: _Canvas, proj: Projector, cont: Continuation,
                      doc: PcfDocument, options: RenderOptions,
                      draw_centre: Vec2) -> None:
    """跨管线延续：接点画虚线示意短管 + SEE ISO 对方线号文字。

    两条管线数据独立，对方管件一律不画——虚线沿本图管线在接点的
    逼近方向（inboard→接点，重合时用紧邻零件本体方向）延长，
    示意"线在那边继续"。接点上已有延续示意（双写副本/支管台/三通）
    时省略短管，文字块直接锚在接点，避免与边界几何重复。"""
    at = proj.page(cont.point)
    if _tie_in_geometry_present(doc, cont):
        tip = at
    else:
        out = _junction_out_dir(doc, cont, proj, draw_centre)
        tip = _v2_add(at, _v2_scale(out, _LINE_TIE_STUB_MM))
        canvas.dashed(at, tip, width=0.35)

    lines = _cont_callout_lines(doc, cont, options)
    canvas.pdf.set_font("Helvetica", size=_CONT_FONT)
    text_w = max(canvas.pdf.get_string_width(line) for line in lines)
    w = text_w + 2 * _CONT_PAD_MM
    h = len(lines) * _CONT_LINE_MM + 2 * _CONT_PAD_MM - (_CONT_LINE_MM - _CONT_CAP)
    pos = canvas.place_label(tip, w, h, draw_centre, min_leader=7.0,
                             avoid_geometry=True, clearance=_CONT_CLEARANCE_MM)

    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.3)
        canvas.pdf.rect(pos[0] - w / 2, pos[1] - h / 2, w, h, style="DF")
    top = pos[1] - h / 2 + _CONT_PAD_MM + _CONT_CAP
    for n, line in enumerate(lines):
        canvas.text((pos[0], top + n * _CONT_LINE_MM), line, size=_CONT_FONT)

    edge = _box_edge_toward(pos, tip, w, h)
    if _v2_len(_v2_sub(tip, edge)) > 0.8:
        canvas.line(edge, tip, width=0.15)


# ---------------------------------------------------------------------------
# Component symbols


def _olet_header_points(doc: PcfDocument) -> dict[int, tuple[Vec3, Vec3]]:
    """For each olet, the two end points of the header run it sits on.
    
    A set-on olet carries only its CENTRE-POINT and branch point -- no end
    points of its own, and connectivity does not join it to the header; the
    run behind its centre is found here so the fitting draws tangent to the
    pipeline. Run end points are returned rather than a direction: the
    header is drawn as the straight segment between those projections, so a
    direction from them is exactly parallel to the drawn line (a model-space
    sample is not, post-warp).
    """
    comps = doc.real_components()
    tol = geometry.model_tolerance(doc)
    out: dict[int, tuple[Vec3, Vec3, float]] = {}
    for comp in comps:
        if classify(comp.type, comp.skey) != "OLET" or not comp.centre_points:
            continue
        centre = comp.centre_points[0].xyz
        for other in comps:
            if other is comp or len(other.end_points) < 2:
                continue
            a, b = other.end_points[0].xyz, other.end_points[-1].xyz
            axis = tuple(q - p for p, q in zip(a, b))
            length = geometry.length(axis)
            if length <= tol:
                continue
            # Distance from the olet centre to this component's axis, and how
            # far along it the foot of that perpendicular falls.
            t = sum((centre[i] - a[i]) * axis[i] for i in range(3)) / (length ** 2)
            if not (-0.01 <= t <= 1.01):
                continue
            foot = tuple(a[i] + t * axis[i] for i in range(3))
            if geometry.dist(centre, foot) > tol * 4:
                continue
            # 第三个元素带上宿主管口径：变径就地标签（'大X小'）需要母管
            # 口径——OLET 自身记录只写支管 bore（CENTRE/BRANCH 两点）
            out[id(comp)] = (a, b, other.bore or 0.0)
            break
    return out


_RUN_AXIS_CLASSES = {"PIPE", "FLANGE", "VALVE"}


def _run_axis_at(comp: Component, at: Vec3, tol: float
                 ) -> tuple[Vec3, Vec3] | None:
    """Two points on ``comp``'s centreline through the port ``at``.
    
    Straight-bodied components hand back their own ends; a fitting must say
    which leg the caller is on -- an elbow's end-to-end chord is the bend
    diagonal and a tee's chord is its run whichever port you arrived at.
    Returns None when ``at`` is not on this component or the leg is too
    short.
    """
    cls = classify(comp.type, comp.skey)
    eps = comp.end_points
    ends = (eps[0].xyz, eps[-1].xyz) if len(eps) >= 2 else None

    def on_run() -> tuple[Vec3, Vec3] | None:
        if ends is None or geometry.dist(*ends) <= tol:
            return None
        if min(geometry.dist(at, ends[0]), geometry.dist(at, ends[1])) > tol:
            return None
        return ends

    if cls in _RUN_AXIS_CLASSES:
        return on_run()

    if cls in ("ELBOW", "BEND"):
        # The leg the caller is on: its weld end out to the corner.
        if ends is None or not comp.centre_points:
            return None
        corner = comp.centre_points[0].xyz
        for end in ends:
            if (geometry.dist(at, end) <= tol
                    and geometry.dist(end, corner) > tol):
                return (end, corner)
        return None

    if cls in ("TEE", "CROSS"):
        run = on_run()
        if run is not None:
            return run
        # Off the branch instead: centre out to the branch port.
        if not comp.centre_points:
            return None
        centre = comp.centre_points[0].xyz
        for br in (*comp.branch1_points, *comp.branch2_points):
            if (geometry.dist(at, br.xyz) <= tol
                    and geometry.dist(centre, br.xyz) > tol):
                return (centre, br.xyz)
    return None


def _reducer_runs(doc: PcfDocument) -> dict[int, tuple[Vec3, Vec3]]:
    """For each reducer, the end points of the run it is welded into.
    
    An eccentric reducer's two weld ends sit on different centrelines, so
    the chord between them is tilted off the pipe axis and a symbol built on
    it has no edge parallel to the pipeline -- the one thing an eccentric
    reducer must show. The axis comes from a neighbour via ``_run_axis_at``
    (elbows/tees count, on the leg the reducer is actually on), preferring
    the large end; neighbour end points are returned so projections give an
    exactly parallel direction (cf. ``_olet_header_points``).
    """
    comps = doc.real_components()
    tol = geometry.model_tolerance(doc)
    out: dict[int, tuple[Vec3, Vec3]] = {}
    for comp in comps:
        if classify(comp.type, comp.skey) != "REDUCER" or len(comp.end_points) < 2:
            continue
        first, last = comp.end_points[0], comp.end_points[-1]
        large_first = (first.bore or 0.0) >= (last.bore or 0.0)
        ordered = (first, last) if large_first else (last, first)
        for want in ordered:            # large end first
            hit = next((run for other in comps
                        if other is not comp
                        and (run := _run_axis_at(other, want.xyz, tol))), None)
            if hit is not None:
                out[id(comp)] = hit
                break
    return out


def _eccentric_small_end_exit(
        proj: Projector, comp: Component, run: tuple[Vec3, Vec3],
        proportional: bool = False) -> Vec2 | None:
    """偏心变径 SKEY 符号的小端 EXIT 图面坐标。"""
    from .asc_symbols import find_symbol_full

    if not reducer_is_eccentric(comp.type, comp.skey) or len(comp.end_points) < 2:
        return None
    first, last = comp.end_points[0], comp.end_points[1]
    if not first.bore or not last.bore or first.bore <= last.bore:
        return None
    found = find_symbol_full(comp.skey or "")
    if not found:
        return None
    sym, _orient, stop = found
    if not stop or abs(stop[1]) <= 1e-9:
        return None
    ys = [c for seg in sym for c in (seg[1], seg[3])]
    big_half = (max(ys) - min(ys)) / 2.0
    small_half = big_half * last.bore / first.bore
    centre = math.copysign(big_half - small_half, stop[1])
    # 取含平移表的图面位置：warp 通道调用时表为空（行为不变），
    # 链式通道必须基于摆放后的位置正算出口
    p0 = proj.page(first.xyz)
    p1 = proj.page(last.xyz)
    a = proj.page(run[0])
    b = proj.page(run[1])
    axis = _v2_sub(b, a)
    ln = _v2_len(axis)
    if ln <= 1e-9:
        return None
    u = (axis[0] / ln, axis[1] / ln)
    diff = _v2_sub(p1, p0)
    if u[0] * diff[0] + u[1] * diff[1] < 0:
        u = (-u[0], -u[1])
    run3 = tuple(last.xyz[i] - first.xyz[i] for i in range(3))
    v = _across_v(proj, run3, first.xyz, u) or _v2_perp(u)
    # 与符号分支相同的定向镜像：小端偏侧由模型实测偏心方向决定，
    # 否则管线平移与符号出口各偏一侧（30001 曾因此顶平/底平画反）
    flat3 = _reducer_flat_offset(comp, first, last, run)
    ea = proj.page(first.xyz)
    eb = proj.page(tuple(first.xyz[i] + flat3[i] for i in range(3)))
    e2 = _v2_sub(eb, ea)
    if _v2_len(e2) > 1e-9 and (e2[0] * v[0] + e2[1] * v[1]) * centre < 0:
        v = (-v[0], -v[1])
    # 出口跟随符号：按比例=实际端距投影；示意=等比自然长度（与 _draw_symbol_2pt
    # 同款规则），小端管线落符号真实出口；schematic_scale 时出口偏移同倍缩小。
    sy = 4.5 / (2.0 * big_half)
    sy *= getattr(proj, "schematic_scale", 1.0)
    if proportional:
        d = (diff[0] * u[0] + diff[1] * u[1]) or _v2_len(diff)
        sx = (1.0 if stop[0] > 0 else -1.0) * d / abs(stop[0])
    else:
        sx = (1.0 if stop[0] > 0 else -1.0) * sy
    return (p0[0] + stop[0] * sx * u[0] + centre * sy * v[0],
            p0[1] + stop[0] * sx * u[1] + centre * sy * v[1])


def _eccentric_small_side_offsets(
        proj: Projector, doc: PcfDocument,
        comps: list[tuple[Component, str | None]],
        runs: dict[int, tuple[Vec3, Vec3]],
        proportional: bool = False) -> dict[Vec3, Vec2]:
    """正算偏心变径出口带动小端连通侧的图面平移。

    符号保持 4.5 mm 大端和口径比小端；小端侧后续管线整体平移到正算
    EXIT，不画折线连接。变径自身不偏移，尺寸/BOM 仍取真实模型值。
    """
    real = [comp for comp, _ref in comps]
    tol = geometry.model_tolerance(doc)
    graph = geometry.ConnectivityGraph(real, tol)
    offsets: dict[Vec3, Vec2] = {}
    for idx, comp in enumerate(real):
        if id(comp) not in runs or len(comp.end_points) < 2:
            continue
        last = comp.end_points[1]
        exit_pt = _eccentric_small_end_exit(
            proj, comp, runs[id(comp)], proportional=proportional)
        if exit_pt is None:
            continue
        delta = _v2_sub(exit_pt, proj.page(last.xyz))
        if _v2_len(delta) <= 0.001:
            continue
        starts = [j for j in graph.adjacency.get(idx, ())
                  if any(geometry.dist(pt.xyz, last.xyz) <= tol * 4
                         for pt in real[j].connection_points())]
        seen = set(starts)
        stack = list(starts)
        while stack:
            current = stack.pop()
            for nxt in graph.adjacency.get(current, ()):
                if nxt != idx and nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        for j in seen:
            for pt in real[j].all_points():
                offsets[pt.xyz] = delta
    return offsets


def _continuation_eccentric_offsets(
        proj: Projector, doc: PcfDocument, continuations: list[Continuation],
        runs: dict[int, tuple[Vec3, Vec3]],
        proportional: bool = False) -> dict[Vec3, Vec2]:
    """偏心变径小端直接接图时，把声明点投到 SKEY 小端 EXIT。"""
    points = [cont.point for cont in continuations if cont.label]
    if not points:
        return {}
    tol = geometry.model_tolerance(doc) * 4
    offsets: dict[Vec3, Vec2] = {}
    for comp in doc.real_components():
        if id(comp) not in runs or len(comp.end_points) < 2:
            continue
        small = comp.end_points[1].xyz
        matched = [point for point in points
                   if geometry.dist(point, small) <= tol]
        if not matched:
            continue
        exit_pt = _eccentric_small_end_exit(
            proj, comp, runs[id(comp)], proportional=proportional)
        if exit_pt is None:
            continue
        for point in matched:
            base = proj.page(point, apply_offset=False)
            offsets[point] = _v2_sub(exit_pt, base)
    return offsets


def _natural_run_length(skey: str | None) -> float | None:
    """符号等比自然图面长度（|收笔点X| × 4.5 / 符号Y跨度）。

    与 _draw_symbol_2pt 的等比制同源：符号沿管轴不变形，长度恒为
    自然长度。无符号或几何退化（无收笔点/零高度）返回 None。
    """
    try:
        from .asc_symbols import find_symbol_full
        found = find_symbol_full(skey or "")
    except Exception:
        return None
    if not found:
        return None
    sym, _orient, stop = found
    if not stop or abs(stop[0]) <= 1e-9:
        return None
    ys = [c for seg in sym for c in (seg[1], seg[3])]
    perp = (max(ys) - min(ys)) if ys else 0.0
    if perp <= 1e-9:
        return None
    return abs(stop[0]) * 4.5 / perp


def _chain_pipe_len(world_len: float, scale: float = 1.0) -> float:
    """链式布局的直管图面长度档位（真实长度单调分档，保长短关系）。
    
    基准档位 ×1.5（管线拉长腾出注释空间，固定倍数不做自适应填满）。
    ``scale`` 是管线自适应因子：首轮 1.0，超幅时按缩放比重摆。
    """
    if world_len < 500.0:
        b = 12.0
    elif world_len < 1500.0:
        b = 22.5
    elif world_len < 4000.0:
        b = 36.0
    else:
        b = 51.0
    return max(6.0, b * scale)


_chain_debug: list[str] = []   # 链式布局调试日志（临时，验证后删）


def _chain_layout(proj: "Projector", doc: PcfDocument,
                  comps: list[tuple[Component, str | None]],
                  draw_region: tuple,
                  pipe_scale: float = 1.0,
                  rigid_symbols: bool = True,
                  fit: bool = True,
                  collect: dict | None = None,
                  ) -> tuple[dict[Vec3, Vec2], str | None, float] | None:
    """搭积木链式布局（示意模式专用）：位置由拓扑走链决定。

    管件固定尺寸（符号自然长度；弯头/三通腿 _SCHEMATIC_LEG_MM）、
    直管按真实长度单调分档，逐件拼接：上一件出口面 = 下一件入口面，
    永不断开；每段方向 = 该管段世界轴向的无 warp 投影（轴测 30° 族），
    弯头按模型转角转。坐标不参与图面定位，只提供方向；尺寸/坐标文字
    仍取模型真值。多连通组竖排打包；整体超幅时等比缩放并返回警告。

    fit=False（NTS 固定 1:1 测量遍）：不缩放、不居中、不警告——打包后
    内容平移到原点，缩放由分页装箱接管。collect 给定时回填打包测量
    数据：pos（打包后绝对图面位表，不含 free_drop）。

    返回 (全点图面平移表, 警告|None)；与 point_offsets 通道一致。
    失败返回 None，调用方回退 warp 布局。
    """
    from collections import deque

    _chain_debug.clear()     # 批量渲染多图时不累积（防内存无限增长）
    real = [c for c, _r in comps]
    if len(real) < 2:
        return None
    tol = geometry.model_tolerance(doc)
    graph = geometry.ConnectivityGraph(real, tol)

    def _page(p: Vec3) -> Vec2:
        return proj.page(p, apply_offset=False)

    def _pdir(a: Vec3, b: Vec3) -> Vec2 | None:
        d = _v2_sub(_page(b), _page(a))
        n = _v2_len(d)
        return (d[0] / n, d[1] / n) if n > 1e-9 else None

    def _sym_dir(a: Vec3, b: Vec3) -> Vec2 | None:
        # 直通符号件（阀/法兰/变径…）走向取主轴投影：偏心变径两端点带垂直偏心量，
        # 直接投影弦线会画歪符号；主轴占比 <90% 是真斜管（45° 平面走向），保留弦线。
        w = (b[0] - a[0], b[1] - a[1], b[2] - a[2])
        ln = math.sqrt(w[0] * w[0] + w[1] * w[1] + w[2] * w[2])
        if ln <= 1e-9:
            return None
        k = max(range(3), key=lambda i: abs(w[i]))
        if abs(w[k]) / ln < 0.9:
            return _pdir(a, b)
        b2 = list(a)
        b2[k] += w[k]
        return _pdir(a, (b2[0], b2[1], b2[2]))

    def _cps(c: Component) -> list[Vec3]:
        pts = [p.xyz for p in c.connection_points()]
        # OLET/TEE-STUB 的主管贴附点存在 CENTRE-POINT（connection_points 不收）：
        # 不收进来走链走不过主管——支管台两侧管子各成一组，后到直管横跨间隙。
        if classify(c.type, c.skey) == "OLET":
            pts += [p.xyz for p in c.centre_points]
        return pts

    # 垫片两端点隔开的"共点"是假共点：垫片厚小于整图连接容差 tol，BFS 会跨过垫片
    # 致符号绘制与布局错位、垫片缝被吃掉——真实走链必须穿过垫片，不能跳过。
    gasket_ends: list[tuple[Vec3, Vec3]] = []
    for c in real:
        if classify(c.type, c.skey) == "GASKET":
            eps = [p.xyz for p in c.end_points]
            if len(eps) >= 2:
                gasket_ends.append((eps[0], eps[1]))

    def _across_gasket(a: Vec3, b: Vec3) -> bool:
        for ga, gb in gasket_ends:
            if (geometry.dist(a, ga) <= 1e-6
                    and geometry.dist(b, gb) <= 1e-6) or (
                    geometry.dist(a, gb) <= 1e-6
                    and geometry.dist(b, ga) <= 1e-6):
                return True
        return False

    # 邻接边 + 共享点（双向都要建：BFS 从任一侧都能走过去，否则图被
    # 砍成单向、走不到的组件会被当成独立组另摆——接头脱开的根因）
    edges: dict[int, list[tuple[int, Vec3]]] = {}
    for i, ci in enumerate(real):
        pts_i = _cps(ci)
        for j in graph.adjacency.get(i, ()):
            if any(k == j for k, _s in edges.get(i, ())):
                continue
            for cp in pts_i:
                if any(geometry.dist(cp, q) <= tol
                       and not _across_gasket(cp, q) for q in _cps(real[j])):
                    edges.setdefault(i, []).append((j, cp))
                    break

    pos: dict[Vec3, Vec2] = {}
    placed: set[int] = set()
    # 摆放谱系：parent[点] = (摆放时依托的上游点, 摆件索引)；children 反查。
    # 环路闭合修正要靠它从闭合点回溯到分叉点、并级联搬移侧枝。
    parent: dict[Vec3, tuple[Vec3, int]] = {}
    children: dict[Vec3, list[Vec3]] = {}
    # 待修闭合：(已占用点, 锚点, 锚点→应有位置的偏移)。直管闭环不修——
    # 直管本就是弹性件，横跨现状间隙即设计行为。
    closures: list[tuple[Vec3, Vec3, Vec2]] = []

    def _write(idx: int, pt: Vec3, where: Vec2, from_pt: Vec3) -> bool:
        """写入摆放点并登记谱系；已占用则不动（闭环守卫）。"""
        if pt in pos:
            return False
        pos[pt] = where
        parent[pt] = (from_pt, idx)
        children.setdefault(from_pt, []).append(pt)
        return True

    def _place(idx: int, comp: Component, s: Vec3, p_in: Vec2, din: Vec2,
               frontier: list[tuple[Vec3, Vec2]]) -> None:
        """从入口点 s（图面位置 p_in、来向 din）摆放一个管件。

        把其余连接点/中心点写入 pos，并把 (出口点, 出口方向) 追加进
        frontier 供 BFS 继续走邻件。目标点已被摆放（环路闭合）时：
        刚性件（弯头/三通/符号/点件）登记闭合修正，直管直接横跨。
        """
        cls = classify(comp.type, comp.skey)
        # 垫片厚小于整图连接容差：不能用 tol 把另一端滤成“同点”，否则垫片图面空隙
        # 被吞。1e-6 只给垫片——通用化会让 BFS 穿过短管/点件产生闭合差级联。
        point_tol = 1e-6 if cls == "GASKET" else tol
        others = [c for c in _cps(comp) if geometry.dist(c, s) > point_tol]
        if cls == "PIPE" and not others:
            # 短管拯救：两端距在 tol 内的短管也要占住图面档位长度，否则远端被吸回
            # 近端画成零长度；只写占位不延伸 frontier——短管不做走链通道（防闭合差连锁）。
            exact = [c for c in _cps(comp) if geometry.dist(c, s) > 1e-6]
            if exact:
                far = max(exact, key=lambda c: geometry.dist(s, c))
                d = _pdir(s, far) or din
                _write(idx, far, _v2_add(p_in, _v2_scale(
                    d, _chain_pipe_len(geometry.dist(s, far), pipe_scale))), s)
            return
        if cls == "PIPE" and others:
            far = max(others, key=lambda c: geometry.dist(s, c))
            d = _pdir(s, far) or din
            req = _v2_add(p_in, _v2_scale(
                d, _chain_pipe_len(geometry.dist(s, far), pipe_scale)))
            if _write(idx, far, req, s):
                frontier.append((far, d))
            else:
                _chain_debug.append(
                    f"PIPE chord: idx={idx} req={req} occ={pos[far]} "
                    f"owner={parent.get(far)}")
        elif cls == "ELBOW" and others:
            corner = _v2_add(p_in, _v2_scale(din, _SCHEMATIC_LEG_MM))
            far = others[0]
            # 出口腿方向 = 中心→出口的世界轴投影；用 s→far 弦线会把 90°
            # 弯头画成 30°/72° 等错误转角（弦线是两腿对角线）
            c3 = comp.centre_points[0].xyz if comp.centre_points else s
            d = _pdir(c3, far) or _pdir(s, far) or din
            anchor = c3
            if comp.centre_points:
                if not _write(idx, c3, corner, s):
                    _chain_debug.append(
                        f"ELBOW centre occupied: idx={idx} "
                        f"existing={pos[c3]} corner={corner}")
            req = _v2_add(corner, _v2_scale(d, _SCHEMATIC_LEG_MM))
            if _write(idx, far, req, anchor):
                frontier.append((far, d))
            else:
                closures.append((far, anchor, _v2_sub(req, pos[anchor])))
                _chain_debug.append(
                    f"ELBOW closure: idx={idx} err={_v2_sub(req, pos[far])}")
            _chain_debug.append(
                f"ELBOW placed: idx={idx} "
                f"leg0={_v2_len(_v2_sub(pos[c3], pos[s])):.2f} "
                f"leg1={_v2_len(_v2_sub(pos[far], pos[c3])):.2f} "
                f"far_owner={parent.get(far)}")
        elif cls == "TEE":
            centre = _v2_add(p_in, _v2_scale(din, _SCHEMATIC_LEG_MM))
            c3 = (comp.centre_points[0].xyz if comp.centre_points else s)
            if comp.centre_points:
                _write(idx, c3, centre, s)
            for c in others:
                d = _pdir(c3, c) or _pdir(s, c) or din
                req = _v2_add(centre, _v2_scale(d, _SCHEMATIC_LEG_MM))
                if _write(idx, c, req, c3):
                    frontier.append((c, d))
                else:
                    closures.append((c, c3, _v2_sub(req, pos[c3])))
        elif cls in ("WELD", "BOLT", "MESSAGE", "SUPPORT"):
            # 点件：直通，其余接点吸附在同一位置
            for c in others:
                if not _write(idx, c, p_in, s) and \
                        _v2_len(_v2_sub(pos[c], p_in)) > 0.4:
                    closures.append((c, s, (0.0, 0.0)))
            if others:
                frontier.append((others[0], din))
        elif cls == "CAP":
            for c in others:
                _write(idx, c, p_in, s)
        else:
            # 双端符号件（阀/法兰/变径/短接/仪表…）：固定自然长度直通，多余接点按
            # 腿长展开；垫片不画符号但占住一小段图面长度，否则法兰面与阀面贴死。
            natural = _natural_run_length(comp.skey)
            if natural is None:
                natural = 1.0 if cls == "GASKET" else 6.0
            if not others:
                return
            far = max(others, key=lambda c: geometry.dist(s, c))
            if cls in ("FLANGE", "GASKET"):
                # 法兰/垫片自身不带偏心偏移：方向随邻管弦线（真斜管主轴占比
                # 可 >0.9，主轴吸附会把符号拧到 30° 族轴测轴上、与邻管折线脱节）
                d = _pdir(s, far) or din
            else:
                d = _sym_dir(s, far) or din
            req = _v2_add(p_in, _v2_scale(d, natural))
            if _write(idx, far, req, s):
                if comp.centre_points:
                    _write(idx, comp.centre_points[0].xyz,
                           _v2_mid(p_in, pos[far]), s)
                frontier.append((far, d))
            else:
                closures.append((far, s, _v2_sub(req, p_in)))
            for c in others:
                if c is far:
                    continue
                dc = _pdir(s, c) or din
                req = _v2_add(p_in, _v2_scale(dc, _SCHEMATIC_LEG_MM))
                if not _write(idx, c, req, s):
                    closures.append((c, s, _v2_sub(req, p_in)))

    def _anc(p: Vec3) -> list[Vec3]:
        """p 的祖先链（含 p），沿 parent 回到组根。"""
        chain = [p]
        while p in parent:
            p = parent[p][0]
            chain.append(p)
        return chain

    def _shift_subtree(root: Vec3, delta: Vec2, skip: set[Vec3]) -> None:
        """root 子树整体平移 delta（刚性跟随，保持内部几何）。"""
        stack = [root]
        while stack:
            q = stack.pop()
            if q in skip:
                continue
            skip.add(q)
            pos[q] = _v2_add(pos[q], delta)
            stack.extend(children.get(q, ()))

    def _absorb_closures() -> None:
        """把环路闭合差摊进闭合侧路径上最顺向的直管。

        闭合件（弯头/三通/符号/点件）保持固定几何不动；从闭合点回溯到
        与入口侧的公共祖先（分叉点），该侧路径上的点整体平移 err——
        平移量在选定直管处从 err 跳变到 0，于是只有那根管子被拉长
        （沿自身轴最好时零侧倾），其余管件刚性跟随，处处不脱开。
        """
        for q, anchor, off in closures:
            req = _v2_add(pos[anchor], off)
            err = _v2_sub(req, pos[q])
            if _v2_len(err) < 0.4:
                continue
            chain_q = _anc(q)
            fork = next((p for p in chain_q if p in anchor_set(anchor)),
                        None)
            if fork is not None:
                path = chain_q[:chain_q.index(fork)]  # q .. fork 之前
            else:
                # 跨树闭合（q 由别的组摆放）：无公共祖先，沿 q 自身
                # 谱系向上找最近直管吸收（管子拉长/侧倾兜底）
                path = chain_q[:-1]                   # q .. 根之前
            # 选吸收管：路径上被 PIPE 摆放、方向与 err 最顺向的一步
            best_i, best_dot = -1, 0.0
            for i in range(len(path)):
                pq, pi = path[i], parent[path[i]][0]
                if pi not in pos:
                    continue
                _pp, idx = parent[path[i]]
                if classify(real[idx].type, real[idx].skey) != "PIPE":
                    continue
                seg = _v2_sub(pos[pq], pos[pi])
                n = _v2_len(seg)
                if n < 2.0:
                    continue
                dot = abs((seg[0] * err[0] + seg[1] * err[1]) / n)
                if dot > best_dot:
                    best_i, best_dot = i, dot
            if best_i < 0:
                # 闭合侧无直管可吸收（如点件跨树闭合）：刚性平移较小
                # 的那一侧整组，再把闭合两侧合并；否则后续竖排打包会把
                # 已经对齐的两组重新分开，将跨组符号拉成长线。
                bq, ba = key_box.get(q), key_box.get(anchor)
                tgt, delta = None, err
                if bq is not None and ba is not None:
                    if len(group_boxes[bq][4]) <= len(group_boxes[ba][4]):
                        tgt = bq
                    else:
                        tgt, delta = ba, (-err[0], -err[1])
                elif bq is not None:
                    tgt = bq
                elif ba is not None:
                    tgt, delta = ba, (-err[0], -err[1])
                if tgt is not None:
                    box = group_boxes[tgt]
                    for kk in box[4]:
                        pos[kk] = _v2_add(pos[kk], delta)
                    if bq is not None and ba is not None and bq != ba:
                        other = ba if tgt == bq else bq
                        merged = group_boxes[other][4]
                        box[4].extend(merged)
                        for kk in merged:
                            key_box[kk] = tgt
                        group_boxes[other][4] = []
                    elif (bq is None) != (ba is None):
                        # 无包盒的一侧属于正在摆放的当前组，待当前组登记后再合并。
                        pending_merges.add(bq if bq is not None else ba)
                    xs = [pos[k][0] for k in box[4]]
                    ys = [pos[k][1] for k in box[4]]
                    box[0], box[1] = min(xs), min(ys)
                    box[2], box[3] = max(xs), max(ys)
                _chain_debug.append(
                    f"closure box-shift: err={_v2_len(err):.1f} tgt={tgt}")
                continue
            _chain_debug.append(
                f"closure absorbed: err={_v2_len(err):.1f}mm "
                f"path={len(path)} best_i={best_i}")
            moved = set(path[:best_i + 1])
            for p in path[:best_i + 1]:
                pos[p] = _v2_add(pos[p], err)
                # 侧枝（不在路径上的子树）随挂点平移
                for c in children.get(p, ()):
                    if c not in moved:
                        _shift_subtree(c, err, moved)
        closures.clear()

    def anchor_set(anchor: Vec3) -> set[Vec3]:
        return set(_anc(anchor))

    # BFS 逐组布局：组内从该件最长的跨起步，方向取其投影
    group_boxes: list[list] = []
    key_box: dict[Vec3, int] = {}   # 摆放点 → 所属包盒（锚定组合并用）
    for start in range(len(real)):
        if start in placed:
            continue
        comp = real[start]
        cps = _cps(comp)
        placed.add(start)
        _chain_debug.append(
            f"GROUP start: idx={start} {comp.type} anchor="
            f"{next((c for c in cps if c in pos), None)}")
        if not cps:
            continue
        # 跨组共点：接点已在已摆结构上（如支座坐在弯头端点）——锚在原位置接着摆，
        # 不另起原点、不参与竖排打包，否则共点被撕成两个图面位置。
        anchor0 = next((c for c in cps if c in pos), None)
        if anchor0 is not None:
            s = anchor0
            others0 = [c for c in cps if geometry.dist(c, anchor0) > tol]
            din = (_pdir(anchor0, max(
                others0, key=lambda c: geometry.dist(anchor0, c)))
                if others0 else (1.0, 0.0)) or (1.0, 0.0)
        else:
            if len(cps) >= 2:
                pairs = [(a, b) for a in cps for b in cps
                         if a != b and geometry.dist(a, b) > tol]
                if pairs:
                    s, f = max(pairs, key=lambda ab: geometry.dist(*ab))
                else:                 # 端点全部重合（平焊法兰等）：按点件摆
                    s, f = cps[0], None
            else:
                s, f = cps[0], None
            din = _pdir(s, f) if f is not None else (1.0, 0.0)
        # 弯头/三通起步：行进方向取 s→中心（s→f 是弦线，方向错）
        if comp.centre_points and classify(comp.type, comp.skey) in (
                "ELBOW", "TEE"):
            din = _pdir(s, comp.centre_points[0].xyz) or din
        if din is None:
            din = (1.0, 0.0)
        before = set(pos)
        pending_merges: set[int] = set()
        if anchor0 is None:
            pos[s] = (0.0, 0.0)
        frontier: list[tuple[Vec3, Vec2]] = []
        _place(start, comp, s, pos[s], din, frontier)
        # 起步件近端（s 侧）的邻件也要走链：frontier 只带远端出口，近端邻件落不到链
        # 上会各自另起组（盲法兰被甩出阀组）。行进方向取反（从 s 向组外走）。
        frontier.append((s, _v2_scale(din, -1.0)))
        # 非管优先、直管靠后（deferred）：环路尽量由直管闭合——直管是弹性件，横跨
        # 现状间隙即可；无管可闭合时由 _absorb_closures 摊进闭合侧最顺向的直管。
        pending = deque((start, pt, d) for pt, d in frontier)
        deferred: deque[tuple[int, Vec3, Vec2, int, Vec3]] = deque()
        while pending or deferred:
            if pending:
                owner, pt, d = pending.popleft()
                for j, shared in edges.get(owner, ()):
                    if j in placed or geometry.dist(shared, pt) > tol:
                        continue
                    if classify(real[j].type, real[j].skey) == "PIPE":
                        deferred.append((owner, pt, d, j, shared))
                        _chain_debug.append(
                            f"defer PIPE: j={j} owner={owner} shared={shared}")
                        continue
                    placed.add(j)
                    if not _write(j, shared, pos[pt], pt) and \
                            _v2_len(_v2_sub(pos[shared], pos[pt])) > 0.4:
                        # 共点已被他组占用且位置不一致：登记闭合修正
                        closures.append((shared, pt, (0.0, 0.0)))
                    nb_front: list[tuple[Vec3, Vec2]] = []
                    _place(j, real[j], shared, pos[shared], d, nb_front)
                    pending.extend((j, p2, d2) for p2, d2 in nb_front)
            else:
                _o, pt, d, j, shared = deferred.popleft()
                if j in placed:
                    continue
                placed.add(j)
                if not _write(j, shared, pos[pt], pt) and \
                        _v2_len(_v2_sub(pos[shared], pos[pt])) > 0.4:
                    closures.append((shared, pt, (0.0, 0.0)))
                nb_front = []
                _place(j, real[j], shared, pos[shared], d, nb_front)
                pending.extend((j, p2, d2) for p2, d2 in nb_front)
        _absorb_closures()
        keys = [k for k in pos if k not in before]
        if anchor0 is not None:
            # 锚定组：并入锚点所在的包盒，打包时随主结构一起平移，
            # 否则主结构平移后锚定件被撕在原地（腿被拽长的根因）
            bi = key_box.get(anchor0)
            if bi is not None and (keys or pending_merges):
                box = group_boxes[bi]
                box[4].extend(keys)
                for k in keys:
                    key_box[k] = bi
                for other in pending_merges:
                    if other == bi:
                        continue
                    merged = group_boxes[other][4]
                    box[4].extend(merged)
                    for k in merged:
                        key_box[k] = bi
                    group_boxes[other][4] = []
                xs = [pos[k][0] for k in box[4]]
                ys = [pos[k][1] for k in box[4]]
                box[0], box[1] = min(xs), min(ys)
                box[2], box[3] = max(xs), max(ys)
            continue
        if keys:
            bi = len(group_boxes)
            key_box.update({k: bi for k in keys})
            group_boxes.append([0.0, 0.0, 0.0, 0.0, keys])
            box = group_boxes[bi]
            for other in pending_merges:
                if other == bi:
                    continue
                merged = group_boxes[other][4]
                box[4].extend(merged)
                for k in merged:
                    key_box[k] = bi
                group_boxes[other][4] = []
            xs = [pos[k][0] for k in box[4]]
            ys = [pos[k][1] for k in box[4]]
            box[0], box[1] = min(xs), min(ys)
            box[2], box[3] = max(xs), max(ys)

    if not pos:
        return None

    # 中段贴附件（支座/仪表挂在管壁中段、不与任何接点重合）自成松散组：排除出竖排
    # 打包，打包后再按所挂直管的最终图面位置整组平移吸附（打包前吸附会被撕掉）。
    pipe_segs: list[tuple[Vec3, Vec3]] = []
    for c in real:
        if classify(c.type, c.skey) != "PIPE":
            continue
        ep = [p.xyz for p in c.connection_points()]
        if len(ep) >= 2 and ep[0] in pos and ep[1] in pos:
            pipe_segs.append((ep[0], ep[1]))
    to_snap: list[tuple[list[Vec3], Vec3, Vec3, Vec3, float]] = []
    kept_boxes = []
    free_drop: set[Vec3] = set()
    for box in group_boxes:
        keys = box[4]
        if not keys:             # 跨组闭合时已并入其他包盒
            continue
        hit = None
        for k in keys:
            for a, b in pipe_segs:
                lab = geometry.dist(a, b)
                if lab < 1e-6 or min(geometry.dist(k, a),
                                     geometry.dist(k, b)) <= tol:
                    continue
                if geometry.point_segment_distance(k, a, b) <= tol:
                    ab = (b[0] - a[0], b[1] - a[1], b[2] - a[2])
                    # 投影分数并钳位：3D 距离/段长只对共线点成立，端外
                    # 离轴点会算出 >1 的"比例"，插值冲出段外（掉线）
                    t = ((k[0] - a[0]) * ab[0] + (k[1] - a[1]) * ab[1]
                         + (k[2] - a[2]) * ab[2]) / (lab * lab)
                    hit = (k, a, b, min(1.0, max(0.0, t)))
                    break
            if hit:
                break
        if hit:
            to_snap.append((keys, *hit))
        elif len(keys) == 1:
            # 离轴单点松散件（支座/仪表挂在管端外、不在任何管线轴上）：不进打包、
            # 不写 point_offsets——绘制时 page() 最近段插值钳到所挂管线最终图面位置。
            free_drop.add(keys[0])
        else:
            kept_boxes.append(box)
    group_boxes = kept_boxes
    # 贴附组的键集合：bbox 统计排除（贴附单点组的 pos 还停在打包前的
    # 原点附近，计入会把包围盒撑爆、错误压小适配比例）
    snap_keys = {k for keys, _k, _a, _b, _t in to_snap for k in keys}

    # 组竖排打包（y 向下为正，组间留 14mm），再整体居中；
    # 超幅只缩不放（管件固定尺寸是布局语义，放大失真）
    cursor_y = 0.0
    for gx0, gy0, gx1, gy1, keys in group_boxes:
        dx, dy = -gx0, cursor_y - gy0
        for k in keys:
            pos[k] = (pos[k][0] + dx, pos[k][1] + dy)
        cursor_y += (gy1 - gy0) + 14.0
    # 贴附件吸附：改在最终 rebuild 之后按 new_pos 线段插值执行——打包后
    # 此刻吸附会被随后的弹性拉伸撕掉；未覆盖点（孤立接点/中心点）也必须
    # 在打包之后吸附到世界最近的已布点，否则留在打包前的旧位置上（飞出）。
    mapped = list(pos)
    unmap_twin: dict[Vec3, Vec3] = {}
    for comp in real:
        for p in (*comp.connection_points(), *comp.centre_points):
            if p.xyz in pos:
                continue
            near = min(mapped, key=lambda k: geometry.dist(k, p.xyz))
            pos[p.xyz] = pos[near]
            unmap_twin[p.xyz] = near
    xs = [p[0] for k, p in pos.items()
          if k not in free_drop and k not in snap_keys]
    ys = [p[1] for k, p in pos.items()
          if k not in free_drop and k not in snap_keys]
    bx0, by0, bx1, by1 = min(xs), min(ys), max(xs), max(ys)
    bw = max(bx1 - bx0, 1e-6)
    bh = max(by1 - by0, 1e-6)
    rx, ry, rw, rh = draw_region
    m = 0.82
    if fit:
        s = min(1.0, rw * m / bw, rh * m / bh)
        cx_t, cy_t = rx + rw / 2.0, ry + rh / 2.0
        bcx, bcy = (bx0 + bx1) / 2.0, (by0 + by1) / 2.0
        warn = None
        if s < 0.999:
            warn = f"chain layout scaled to {s:.2f} to fit the sheet"
    else:
        # 1:1 测量遍：只平移到原点，不缩放不居中（分页装箱接管适配）
        s = 1.0
        warn = None
        cx_t, cy_t = -bx0, -by0
        bcx, bcy = 0.0, 0.0
    # 按边弹性缩放（符号恒定的关键）：超幅只压管线边（直管/弯头/三通摆放
    # 的边 ×s），符号/垫片/点件边刚性跟随长度不变——全局绕中心乘 s 会把
    # 符号端点对一起压近、符号随之画小。rigid_symbols=False 极端兜底时
    # 全体弹性退回整体缩小。沿摆放谱系按写入序重放：parent 先于 child。
    flex_cls = ("PIPE", "ELBOW", "TEE") if rigid_symbols else ()
    new_pos: dict[Vec3, Vec2] = {}
    for k in pos:
        if k in free_drop:
            continue
        info = parent.get(k)
        twin = unmap_twin.get(k)
        if twin is not None and twin in new_pos:
            new_pos[k] = new_pos[twin]
        elif info is None or info[0] not in new_pos:
            p = pos[k]
            new_pos[k] = ((p[0] - bcx) * s + cx_t, (p[1] - bcy) * s + cy_t)
        else:
            idx = info[1]
            cls = classify(real[idx].type, real[idx].skey)
            f = s if cls in flex_cls else 1.0
            base = new_pos[info[0]]
            src = pos[info[0]]
            p = pos[k]
            new_pos[k] = (base[0] + (p[0] - src[0]) * f,
                          base[1] + (p[1] - src[1]) * f)
    # 贴附件二次吸附（最终空间）：按世界比例 t 插值到管线段的最终图面
    # 位置——贴附点沿用打包后旧位置会被 rebuild 弹性拉伸撕离管线，整组平移归位。
    for keys, k, a, b, t in to_snap:
        if a not in new_pos or b not in new_pos or k not in new_pos:
            continue
        fa, fb = new_pos[a], new_pos[b]
        tgt = (fa[0] + (fb[0] - fa[0]) * t, fa[1] + (fb[1] - fa[1]) * t)
        d = (tgt[0] - new_pos[k][0], tgt[1] - new_pos[k][1])
        for kk in keys:
            if kk in new_pos:
                new_pos[kk] = (new_pos[kk][0] + d[0], new_pos[kk][1] + d[1])
    # 派生接点不在包盒 keys 中，贴附平移后必须再次跟随其映射接点，
    # 否则同一支架的近重合点会一个吸附到管线、另一个留在原点附近。
    for k, twin in unmap_twin.items():
        if k in new_pos and twin in new_pos:
            new_pos[k] = new_pos[twin]
    out: dict[Vec3, Vec2] = {}
    for k, tgt in new_pos.items():
        base = _page(k)
        d = (tgt[0] - base[0], tgt[1] - base[1])
        if _v2_len(d) > 1e-6:
            out[k] = d
    if collect is not None:
        collect["pos"] = dict(new_pos)
    return out, warn, s


def _symbol_face_offsets(proj: "Projector", doc: PcfDocument,
                         comps: list[tuple[Component, str | None]],
                         reducer_runs: dict[int, tuple[Vec3, Vec3]],
                         flange_face_pts: dict[int, Vec3] | None,
                         ) -> dict[Vec3, Vec2]:
    """等比符号端面带动连接点的图面平移（端点/尺寸锚随符号走）。

    双端符号件恒等比后 EXIT 落在自然长度处而非模型端点：把两侧模型连接
    点的投影平移到符号实际端面（经 point_offsets 点对点匹配，同一模型点
    被多个管件注册取均值保持相连）。弯头/三通不注册——腿画到模型端点、
    角度随管线，短腿由 _grown_leg_end 补长。偏心变径由
    _eccentric_small_side_offsets 单独接管以免双算。与绘制端
    （_draw_component 符号调度）同一套规则的预计算镜像，两侧必须同步。
    """
    real = [c for c, _r in comps]
    desired: dict[Vec3, list[Vec2]] = {}

    def _want(pt: Vec3, page_pt: Vec2) -> None:
        desired.setdefault(pt, []).append(page_pt)

    for comp in real:
        cls = classify(comp.type, comp.skey)
        if cls not in ("VALVE", "FLANGE", "REDUCER", "COUPLING", "FILTER",
                       "INSTRUMENT", "MULTIPORT"):
            continue
        if len(comp.end_points) < 2:
            continue
        if reducer_is_eccentric(comp.type, comp.skey):
            continue
        natural = _natural_run_length(comp.skey)
        if natural is None:
            continue
        ep0, ep1 = comp.end_points[0], comp.end_points[1]
        e0 = proj.page(ep0.xyz, apply_offset=False)
        e1 = proj.page(ep1.xyz, apply_offset=False)
        sp0, sp1 = e0, e1
        if cls == "FLANGE":
            fpt = (flange_face_pts or {}).get(id(comp))
            if _v2_len(_v2_sub(e1, e0)) < 0.5:
                band = _flange_band(proj, doc, comp, flange_face_pts, pre=True)
                if band is not None:
                    sp0 = band[0]
            elif fpt is not None:
                face = proj.page(fpt, apply_offset=False)
                if _v2_len(_v2_sub(face, e1)) < _v2_len(_v2_sub(face, e0)):
                    sp0, sp1 = e1, e0
        du = _v2_sub(sp1, sp0)
        ln = _v2_len(du)
        if ln <= 1e-9:
            continue    # 方向无法确定（重合且无接带基准），交给绘制端重合分支
        u = (du[0] / ln, du[1] / ln)
        exit_page = _v2_add(sp0, _v2_scale(u, natural))
        flipped = sp0 is e1          # START 对准的是 end_points[1]（面在 e1 侧）
        near_pt, far_pt = ((ep1.xyz, ep0.xyz) if flipped
                           else (ep0.xyz, ep1.xyz))
        _want(near_pt, sp0)
        _want(far_pt, exit_page)

    out: dict[Vec3, Vec2] = {}
    for pt, pages in desired.items():
        base = proj.page(pt, apply_offset=False)
        mean = (sum(p[0] for p in pages) / len(pages),
                sum(p[1] for p in pages) / len(pages))
        delta = (mean[0] - base[0], mean[1] - base[1])
        if _v2_len(delta) > 0.05:
            out[pt] = delta
    return out


def _eccentric_offset_pairs(doc: PcfDocument,
                            runs: dict[int, tuple[Vec3, Vec3]],
                            ) -> list[set[tuple[float, float]]]:
    """Coordinate-value intervals the not-to-scale warp must leave alone.

    An eccentric reducer's two weld ends sit on centrelines offset across the
    run by half the outside-diameter difference -- fitting geometry, not a
    pipe leg: stretching it to the minimum visible span bends the pipeline and
    hides the flat side the offset *is*. Returns one set per axis of (lo, hi)
    intervals rounded to the warp's quantisation; every gap inside an interval
    keeps true proportion.
    """
    keep: list[set[tuple[float, float]]] = [set(), set(), set()]
    for comp in doc.real_components():
        if classify(comp.type, comp.skey) != "REDUCER" or len(comp.end_points) < 2:
            continue
        if not reducer_is_eccentric(comp.type, comp.skey):
            continue
        run = runs.get(id(comp))
        if run is None:
            continue
        first, last = comp.end_points[0], comp.end_points[-1]
        d = tuple(b - a for a, b in zip(first.xyz, last.xyz))
        u = geometry.normalize(tuple(q - p for p, q in zip(run[0], run[1])))
        axial = geometry.dot(d, u)
        across = tuple(d[i] - axial * u[i] for i in range(3))
        for ax in range(3):
            if abs(across[ax]) <= 1e-9:
                continue
            lo, hi = sorted((round(first.xyz[ax], 6), round(last.xyz[ax], 6)))
            if lo != hi:
                keep[ax].add((lo, hi))
    return keep


def _fitting_symbol_targets(comps: list[tuple[Component, str]],
                            min_mm: float = 8.0,
                            height_mm: float = 4.5
                            ) -> list[list[tuple[float, float, float]]]:
    """直管件（阀门/法兰/同心变径/短接等单轴双端件）的示意尺寸目标。

    每个件的坐标区间在 warp 中固定映射为 max(符号自然图面长度, min_mm)，
    上限封顶示意尺寸，非比例伸缩只由直管吸收。跨轴件（弯头/三通）不走此
    机制：腿即管段空间，由间隙 clamp 与腿外延处理；偏心变径由 keep 保护
    真实偏心量。返回 [axis] -> [(lo, hi, target_mm)]。
    """
    out: list[list[tuple[float, float, float]]] = [[], [], []]
    try:
        from .asc_symbols import find_symbol_full
    except Exception:
        return out
    for comp, _ref in comps:
        c = classify(comp.type, comp.skey)
        if c in ("PIPE", "WELD", "BOLT", "ELBOW", "TEE", "SUPPORT"):
            continue
        if reducer_is_eccentric(comp.type, comp.skey):
            continue
        eps = comp.end_points
        if len(eps) < 2:
            continue
        a, b = eps[0].xyz, eps[-1].xyz
        axes = [ax for ax in range(3)
                if abs(round(b[ax], 6) - round(a[ax], 6)) > 1e-9]
        if len(axes) != 1:
            continue
        ax = axes[0]
        lo, hi = sorted((round(a[ax], 6), round(b[ax], 6)))
        if c == "GASKET":
            # 垫片模型厚度（~1.5mm）示意比例下不可见：warp 固定为 0.6mm 图面
            # 间隙露出垫片位，法兰 EXIT 仍精确落焊端，无需平移符号
            out[ax].append((lo, hi, _GASKET_GAP_MM))
            continue
        found = find_symbol_full(comp.skey or "")
        if not found:
            continue
        sym = found[0]
        xs = [c for seg in sym for c in (seg[0], seg[2])]
        ys = [c for seg in sym for c in (seg[1], seg[3])]
        if not xs or not ys:
            continue
        perp = max(ys) - min(ys)
        if perp <= 1e-9:
            continue
        # 原始比例：区间 = 符号等比锚距（START→EXIT × 高度/横向跨度），
        # sx=sy 恒成立，符号不拉伸不压缩
        anchor = abs(found[2][0]) if found[2] and found[2][0] else 1.0
        natural = anchor * height_mm / perp
        out[ax].append((lo, hi, natural))
    return out


def _olet_branch_connected(doc: PcfDocument) -> set[int]:
    """Ids of olets whose outlet actually has something connected to it.

    BRANCH1-POINT says where the outlet *faces*, not that anything is on it;
    running the branch centreline out unconditionally litters the sheet with
    bare stubs off spare or capped fittings. "Connected" = any connection
    point of another real component landing on the branch point at
    connectivity-graph tolerance; components on other sheets do not count --
    a branch piped elsewhere is not this drawing's to show.
    """
    comps = doc.real_components()
    olets = [c for c in comps
             if classify(c.type, c.skey) == "OLET" and c.branch1_points]
    if not olets:
        return set()
    tol = geometry.model_tolerance(doc)
    out: set[int] = set()
    for comp in olets:
        branch = comp.branch1_points[0].xyz
        for other in comps:
            if other is comp:
                continue
            if any(geometry.dist(branch, pt.xyz) <= tol
                   for pt in other.connection_points()):
                out.add(id(comp))
                break
    return out


def _flange_faces(doc: PcfDocument) -> dict[int, Vec3]:
    """Gasket-face point of every flange, keyed by component identity.

    Resolved once for the whole document because it needs the connectivity
    graph, which the per-component symbol pass has no business rebuilding.
    """
    comps = doc.real_components()
    tol = geometry.model_tolerance(doc)
    graph = geometry.ConnectivityGraph(comps, tol)
    faces: dict[int, Vec3] = {}
    for idx, comp in enumerate(comps):
        if not looks_like_flange(comp.type, comp.description, comp.skey):
            continue
        face = flange_face(comp, comps, graph, idx, tol)
        if face is not None:
            faces[id(comp)] = face
    return faces


def _draw_nozzle_tag(canvas: _Canvas, anchor: Vec2, ref: str,
                     draw_centre: Vec2) -> None:
    """设备管口信息块：CONNECTION-REFERENCE 拆成"管口号/连接信息"两行。

    PCF 的 CONNECTION-REFERENCE 形如 ``P39004/Suction$80 mm RFFE 150``：
    ``$`` 前是 设备号/管口号，后是 尺寸 连接型式 压力等级。画成带边框的
    两行块并引线指向管口，先于尺寸/球标放置并注册进避让集合。
    """
    tag, _, info = ref.partition("$")
    lines = [tag.strip()]
    if info.strip():
        lines.append(info.strip())
    if not lines[0]:
        return
    size = 4.5
    canvas.pdf.set_font("Helvetica", size=size)
    cap = _cap_height_mm(size)
    pad = 0.6
    line_h = cap * 1.45
    box_w = max(canvas.pdf.get_string_width(t) for t in lines) + 2 * pad
    box_h = len(lines) * line_h + 2 * pad
    pos = canvas.place_label(anchor, box_w, box_h, draw_centre,
                             min_leader=4.0, avoid_geometry=True,
                             clearance=0.6)
    edge = _box_edge_toward(pos, anchor, box_w, box_h)
    canvas.line(anchor, edge, width=0.15)
    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.25)
        canvas.pdf.rect(pos[0] - box_w / 2, pos[1] - box_h / 2,
                        box_w, box_h, style="DF")
    top = pos[1] - box_h / 2 + pad + cap
    for i, text in enumerate(lines):
        canvas.text((pos[0], top + i * line_h), text, size=size)


def _draw_equipment_connections(canvas: _Canvas, proj: Projector,
                                equip_recs: list, comps: list,
                                draw_centre: Vec2) -> None:
    """设备管口：官方 NZ 符号以虚线画在设备侧，法兰面落在连接点。"""
    ends: list[tuple[Vec2, Vec2, Vec3, Vec3]] = []   # (端点, 来向, 3D 端, 3D 来向)
    for comp, _ref in comps:
        ep3 = comp.end_points
        if len(ep3) < 2:
            continue
        e0, e1 = proj.page(ep3[0].xyz), proj.page(ep3[1].xyz)
        ends.append((e0, e1, ep3[0].xyz, ep3[1].xyz))
        ends.append((e1, e0, ep3[1].xyz, ep3[0].xyz))
    for rec in equip_recs:
        pts = rec.co_ords
        if not pts:
            continue
        nz = proj.page(pts[0].xyz)
        best = None
        for pe, ps, we, ws in ends:
            d2 = (pe[0] - nz[0]) ** 2 + (pe[1] - nz[1]) ** 2
            if best is None or d2 < best[0]:
                best = (d2, pe, ps, we, ws)
        if best is None:
            continue
        _, pe, ps, we, ws = best
        dx, dy = pe[0] - nz[0], pe[1] - nz[1]
        gap = (dx * dx + dy * dy) ** 0.5
        if gap > 1e-6:
            u = (dx / gap, dy / gap)     # 管口 → 管线
        else:
            sx, sy = ps[0] - pe[0], ps[1] - pe[1]
            sl = (sx * sx + sy * sy) ** 0.5
            u = (sx / sl, sy / sl) if sl > 1e-6 else (1.0, 0.0)
        ref = rec.get("CONNECTION-REFERENCE") or ""
        tail = ref.split("$")[-1].split()
        face = next((t.upper() for t in tail
                     if t.isalpha() and t.upper() != "MM"), "")
        skey = "NZWS" if ("W" in face or face.startswith("B")) else "NZFE"
        try:
            from .asc_symbols import find_symbol_full
            found = find_symbol_full(skey)
        except Exception:
            found = None
        run3 = tuple(ws[i] - we[i] for i in range(3))
        # 与普通 SKEY 零件一致：端面沿垂直于管轴的世界横轴投影，
        # 不强制采用法兰标注的竖直面杠，避免管口与相邻零件不在同一平面。
        across = _across_v(proj, run3, we, u) or _v2_perp(u)
        h = 4.5                           # 与普通 SKEY 符号一致的图面高度
        if found is None:
            canvas.dashed(_v2_add(nz, _v2_scale(across, h / 2)),
                          _v2_sub(nz, _v2_scale(across, h / 2)), width=0.3)
            continue
        segs, _orient, stop = found
        span = abs(stop[0]) if stop else 0.0
        ys = [c for seg in segs for c in (seg[1], seg[3])]
        perp = (max(ys) - min(ys)) if ys else 1.0
        if span <= 1e-6 or perp <= 1e-6:
            continue
        d = span * (h / perp)                # 等比缩放后的 START→EXIT 距离
        # NZ 的 EXIT 是法兰面，START 是设备侧颈端；颈端必须远离管线，
        # 不能叠画到法兰、垫片和管道内部。
        p0 = _v2_sub(nz, _v2_scale(u, d))
        if gap > 1.0:
            canvas.dashed(pe, nz, width=0.3)
        _draw_symbol_2pt(canvas, skey, p0, nz, proj=proj,
                         run3=run3, wp=we, across=across, height_mm=h,
                         proportional=proj.proportional,
                         width=0.3, dashed=True)
        # 管口信息块：锚在设备侧颈端，引线指向管口本体
        if ref:
            _draw_nozzle_tag(canvas, p0, ref, draw_centre)


def _draw_component(canvas: _Canvas, proj: Projector, comp: Component,
                    doc: PcfDocument,
                    flange_face_pts: dict[int, Vec3] | None = None,
                    flange_bar_mm: float = 2 * DEFAULT_OFFSET_MM,
                    olet_header_pts: dict[int, tuple[Vec3, Vec3, float]] | None = None,
                    olet_branch_connected: set[int] | None = None,
                    reducer_runs: dict[int, tuple[Vec3, Vec3]] | None = None,
                    ) -> Vec2 | None:
    cls = classify(comp.type, comp.skey)
    # 宽度体系按出图模式分开：按比例=孔径正比宽度 2·_bore_half_w(D)，法兰
    # 杠与法兰端口的阀/方块再乘 B16.5 突面比例 _flange_face_ratio(D)（1/2"
    # ≈1.64→24"≈1.14），整组同宽读作一个法兰连接；不按比例（示意）退回
    # 固定尺寸（ISOGEN 惯例），缺孔径时按比例模式回退 flange_bar_mm 固定杠。
    prop_mode = bool(getattr(proj, "proportional", False))
    bore_in = _bore_to_in(comp.bore, doc.units_bore) if prop_mode else None
    if bore_in is not None:
        flange_bar = 2 * _bore_half_w(bore_in) * _flange_face_ratio(bore_in)
        flange_widen = (_flange_face_ratio(bore_in)
                        if _flanged_ports(comp) else 1.0)
    else:
        flange_bar = flange_bar_mm
        flange_widen = 1.0
    # 对侧边界件副本（CONTINUATION）：整件画虚线示意（ISOGEN 客侧边界画法）
    cont_copy = comp.is_continuation_copy()
    # 链式布局整图缩放：示意腿长同乘 schematic_scale（位置表已缩放）
    leg_mm = _SCHEMATIC_LEG_MM * getattr(proj, "schematic_scale", 1.0)
    eccentric = cls == "REDUCER" and reducer_is_eccentric(comp.type, comp.skey)
    # 链式布局下所有点都有平移表项，偏心变径也必须带偏移绘制；
    # apply_offset=False 只在 warp 通道（变径自身不偏移）有意义
    off_ok = (not eccentric) or getattr(proj, "offset_nearest", False)
    eps = [proj.page(p.xyz, apply_offset=off_ok) for p in comp.end_points]
    cps = [proj.page(p.xyz) for p in comp.centre_points]
    b1s = [proj.page(p.xyz) for p in comp.branch1_points]
    b2s = [proj.page(p.xyz) for p in comp.branch2_points]
    cos = [proj.page(p.xyz) for p in comp.co_ords]
    pps = [proj.page(p.xyz) for p in comp.port_points]

    # 法兰常被误导出为 COUPLING 记录：符号派发兼看 description，BOM 仍只用
    # classify() 不受影响。符号库优先（统一入口）：SKEY 查到几何的双端组件
    # 先画官方符号（START→eps[0]、EXIT→eps[1] 双点仿射），手写分支仅作回退；
    # 必须先于下方 flange_face 特殊分支。SAFETY-DISC 在 classify 归 GASKET
    # 但官方符号必须画——只让普通垫片走"不画、项号并父标"分支。
    sym_first = cls == "GASKET" and (
        (comp.type or "").upper().replace("_", "-").startswith("SAFETY-DISC")
        or type_from_skey(comp.skey) == "SAFETY-DISC")
    # 焊接联管节（gap 小）不画符号——焊点语义由焊点呈现，中心线保持
    # 连续（与下方 COUPLING 原生分支同一判定）；gap 大的仍走符号。
    skip_sym = False
    if cls == "COUPLING" and len(eps) >= 2:
        gap_in = (comp.end_points[0].distance_to(comp.end_points[1])
                  * coord_to_inch_factor(doc.units_coords))
        skip_sym = (gap_in <= GAP_COUPLING_MAX_INCHES
                    and is_weld_coupling(comp.type, comp.description,
                                         comp.skey))
    # 按比例模式整体走上游手绘路径（轴向=真实端点、横向=固定可读页宽），
    # 不走 SKEY 符号——符号 X 轴仿射对齐在偏心件上靠下游平移补偿，与
    # "忠实坐标"语义冲突；SAFETY-DISC（sym_first）无手绘等价物仍走符号库。
    if (cls in ("VALVE", "FLANGE", "ELBOW", "REDUCER", "COUPLING", "FILTER",
                "INSTRUMENT", "MULTIPORT", "OLET") or sym_first) \
            and not skip_sym and len(eps) >= 2 \
            and (sym_first or not proj.proportional):
        ep3 = comp.end_points
        run3 = (tuple(ep3[1].xyz[i] - ep3[0].xyz[i] for i in range(3))
                if len(ep3) >= 2 else None)
        sp0, sp1 = eps[0], eps[1]
        best = None            # (距离, 焊点=最近管端, 管段远端)
        if _v2_len(_v2_sub(eps[1], eps[0])) < 0.5:
            # 两端点几乎重合（平焊法兰 face-to-face≈0）：轴向取最近管段，
            # 方向必须指向该管段"远端"（沿管道离开接头），否则符号翻进阀门里。
            try:
                for c in doc.real_components():
                    if classify(c.type, c.skey) != "PIPE" or len(c.end_points) < 2:
                        continue
                    e0 = proj.page(c.end_points[0].xyz)
                    e1 = proj.page(c.end_points[1].xyz)
                    near, far = ((e0, e1) if _v2_len(_v2_sub(e0, sp0))
                                 <= _v2_len(_v2_sub(e1, sp0)) else (e1, e0))
                    dn = _v2_len(_v2_sub(near, sp0))
                    if dn < 2.0 and (best is None or dn < best[0]):
                        best = (dn, near, far)
            except Exception:
                best = None
            if best is not None and _v2_len(_v2_sub(best[2], sp0)) > 1.0:
                uu = _v2_norm(_v2_sub(best[2], sp0))
                sp1 = (sp0[0] + uu[0] * 6.0, sp0[1] + uu[1] * 6.0)
        if cls == "FLANGE":
            # 法兰符号：START(面线侧)锚定在垫片配合面端，并向外空开垫片
            # 厚度——法兰与阀门/对方法兰之间的垫片在图上留出间隙
            fpt = (flange_face_pts or {}).get(id(comp))
            band = _flange_band(proj, doc, comp, flange_face_pts)
            if band is not None:
                # 端点重合的法兰（平焊等）：光面端距阀面 _GASKET_GAP_MM，
                # 焊端越过焊点；主管端点与焊点随后修剪/吸附到焊端
                sp0, sp1 = band[0], band[1]
                _FLANGE_BANDS.append(band)
            elif fpt is not None:
                face = proj.page(fpt)
                if _v2_len(_v2_sub(face, eps[1])) < _v2_len(_v2_sub(face, eps[0])):
                    sp0, sp1 = eps[1], eps[0]      # 配合面在 eps[1]：翻转使 START 对准面
                # 登记焊端：管端与焊点都吸到 EXIT 锚定的这一点；不做垫片
                # 间隙平移——平移会把 EXIT 推过焊端、符号与管端/焊点错开
                _FLANGE_WELDS.append(sp1)
        axis_override = None
        if cls == "REDUCER":
            # 偏心大小头：管轴取主管走向（reducer_runs，两端同轴导出场合），
            # 符号 X→管轴、偏心量按符号 Y 保留——沿 p0→p1 对角线对齐会扭斜端面
            run = (reducer_runs or {}).get(id(comp))
            if run is not None:
                a3 = proj.page(run[0])
                b3 = proj.page(run[1])
                ln = _v2_len(_v2_sub(b3, a3))
                if ln > 1e-9:
                    ax = ((b3[0] - a3[0]) / ln, (b3[1] - a3[1]) / ln)
                    diff0 = _v2_sub(eps[1], eps[0])
                    if ax[0] * diff0[0] + ax[1] * diff0[1] < 0:
                        ax = (-ax[0], -ax[1])
                    axis_override = ax
        bore_ratio = None
        ecc_dir = None
        if cls == "REDUCER" and reducer_is_eccentric(comp.type, comp.skey):
            bore0, bore1 = ep3[0].bore, ep3[1].bore
            if bore0 and bore1 and bore0 > bore1:
                bore_ratio = bore1 / bore0
                # 模型偏心方向（顶平/底平）投影到图面，供符号定向镜像
                run0 = (reducer_runs or {}).get(id(comp))
                flat3 = _reducer_flat_offset(comp, ep3[0], ep3[-1], run0)
                c3 = ep3[0].xyz
                d3 = tuple(c3[i] + flat3[i] for i in range(3))
                pa = proj.page(c3, apply_offset=False)
                pb = proj.page(d3, apply_offset=False)
                if _v2_len(_v2_sub(pb, pa)) > 1e-9:
                    ecc_dir = _v2_sub(pb, pa)
        op_kw = {}
        if cls == "VALVE":
            # 阀柄：映射表查柄 skey，SPINDLE-DIRECTION 投影成页面方向
            op_skey = _valve_operator_skey(comp.skey)
            sd3 = _FLAT_DIRECTIONS.get(
                (comp.get("SPINDLE-DIRECTION") or "").strip().upper())
            if op_skey and sd3 and proj is not None and ep3:
                c3 = ep3[0].xyz
                pa = proj.page(c3, apply_offset=False)
                pb = proj.page(tuple(c3[i] + sd3[i] for i in range(3)),
                               apply_offset=False)
                dv = _v2_sub(pb, pa)
                ln3 = _v2_len(dv)
                if ln3 > 1e-9:
                    op_kw["operator"] = (op_skey,
                                         (dv[0] / ln3, dv[1] / ln3))
        prop_kw = {}
        if proj.proportional:
            # 按比例模式才需要生长空间（示意模式符号恒等比不生长）
            du = _v2_sub(sp1, sp0)
            dln = _v2_len(du)
            if dln > 1e-9:
                ax_u = (du[0] / dln, du[1] / dln)
                prop_kw = {
                    "room0": _growth_room(sp0, (-ax_u[0], -ax_u[1]),
                                          id(comp)),
                    "room1": _growth_room(sp1, ax_u, id(comp))}
            if cls == "FLANGE":
                prop_kw["no_grow"] = True
        # 符号路径宽度体系同样按模式分开：按比例=孔径正比宽度（法兰杠/法兰
        # 端口件=flange_bar，其余=管径宽 2·_bore_half_w）；示意=恒定 4.5 不传 width_mm
        sym_w = None
        if prop_mode and bore_in is not None:
            sym_w = (flange_bar if (cls == "FLANGE" or flange_widen != 1.0)
                     else 2 * _bore_half_w(bore_in))
        if _draw_symbol_2pt(canvas, comp.skey, sp0, sp1,
                            proj=proj, run3=run3, axis=axis_override,
                            wp=ep3[0].xyz if ep3 else None,
                            drop_isolated=(cls == "REDUCER"
                                           and not (b1s or b2s)),
                            bore_ratio=bore_ratio,
                            proportional=proj.proportional,
                            dashed=cont_copy,
                            width=0.3 if cont_copy else 0.4,
                            width_mm=sym_w,
                            ecc_dir=ecc_dir, **op_kw, **prop_kw):
            if cls == "REDUCER" and len(eps) >= 2 and not cont_copy \
                    and not proj.proportional:
                tag = _bore_tag_text((comp.end_points[0].bore,
                                      comp.end_points[-1].bore))
                if tag:
                    _draw_bore_tag(canvas, proj, _v2_mid(eps[0], eps[1]),
                                   axis_override or _v2_sub(eps[1], eps[0]),
                                   tag)
            return cps[0] if cps else _v2_mid(eps[0], eps[1])

    face_pt = (flange_face_pts or {}).get(id(comp))
    if face_pt is not None and len(eps) >= 2:
        face = proj.page(face_pt)
        face_first = _v2_len(_v2_sub(face, eps[0])) <= _v2_len(_v2_sub(face, eps[1]))
        _draw_flange(canvas, proj, eps[0], eps[1], comp, face_first, flange_bar)
        return _v2_mid(eps[0], eps[1])

    # 流向箭头（SKEY=FLOW）：优先用符号库官方符号，无符号时回退实心小三角；
    # 沿管道流向定向（FLOW=2 反向），图面符号不进 BOM/气球
    if comp.type.upper().replace("_", "-").startswith("FLOW-ARROW"):
        at = cos[0] if cos else (eps[0] if eps else None)
        if at is not None and proj is not None and doc is not None:
            run = _nearest_pipe_run_dir(proj, doc, at)
            if run is not None:
                d = run[0]
                dw = run[1] if len(run) > 1 else None
                ln = _v2_len(d)
                u = (d[0] / ln, d[1] / ln)
                if (comp.get("FLOW") or "1").strip() == "2":
                    u = (-u[0], -u[1])
                v = (-u[1], u[0])
                # 跨向取"垂直于管轴的世界轴"投影（与符号统一入口/支撑同规则）
                if dw is not None and comp.co_ords:
                    wp = comp.co_ords[0].xyz
                    axis = max(range(3), key=lambda i: abs(dw[i]))
                    across = (0.0, 1.0, 0.0) if axis == 0 else (1.0, 0.0, 0.0)
                    a2 = proj.page(wp, apply_offset=False)
                    b2 = proj.page((wp[0] + across[0], wp[1] + across[1],
                                    wp[2] + across[2]), apply_offset=False)
                    vn = (b2[0] - a2[0]) ** 2 + (b2[1] - a2[1]) ** 2
                    if vn > 1e-12:
                        vn = vn ** 0.5
                        vv = ((b2[0] - a2[0]) / vn, (b2[1] - a2[1]) / vn)
                        if abs(vv[0] * u[0] + vv[1] * u[1]) < 0.99:
                            v = vv
                try:
                    from .asc_symbols import find_symbol
                    flow_sym = find_symbol("FLOW")
                except Exception:
                    flow_sym = None
                if flow_sym:
                    sym = flow_sym[0]
                    xs = [c for seg in sym for c in (seg[0], seg[2])]
                    ys = [c for seg in sym for c in (seg[1], seg[3])]
                    extent = max(max(xs) - min(xs), max(ys) - min(ys), 1e-6)
                    scale = 4.5 / extent
                    for x1, y1, x2, y2 in sym:
                        canvas.line(
                            (at[0] + (x1 * u[0] + y1 * v[0]) * scale,
                             at[1] + (x1 * u[1] + y1 * v[1]) * scale),
                            (at[0] + (x2 * u[0] + y2 * v[0]) * scale,
                             at[1] + (x2 * u[1] + y2 * v[1]) * scale),
                            width=0.3)
                else:
                    length, width = 2.0, 1.2
                    tip = (at[0] + u[0] * length / 2, at[1] + u[1] * length / 2)
                    b1 = (at[0] - u[0] * length / 2 + v[0] * width / 2,
                          at[1] - u[1] * length / 2 + v[1] * width / 2)
                    b2 = (at[0] - u[0] * length / 2 - v[0] * width / 2,
                          at[1] - u[1] * length / 2 - v[1] * width / 2)
                    canvas.pdf.polygon([tip, b1, b2], style="F")
        return None

    if cls == "PIPE":
        if len(eps) >= 2:
            a, b = eps[0], eps[1]
            # 端点落在"端点重合法兰"接带簇里的：修剪到法兰焊端——
            # 主管接焊端，不穿越垫片空隙直抵阀面
            for fl in doc.real_components():
                if classify(fl.type, fl.skey) != "FLANGE":
                    continue
                bd = _flange_band(proj, doc, fl, flange_face_pts)
                if bd is not None:
                    top, exitp, uu = bd
                    band_len = _v2_len(_v2_sub(exitp, top))
                    for pt_name in ("a", "b"):
                        p = a if pt_name == "a" else b
                        d = _v2_sub(p, top)
                        s = d[0] * uu[0] + d[1] * uu[1]
                        lat = abs(d[0] * -uu[1] + d[1] * uu[0])
                        if -1.5 <= s <= band_len + 1.0 and lat < 2.0:
                            # 直管止于接带焊端（越过焊点 0.6mm）
                            if pt_name == "a":
                                a = exitp
                            else:
                                b = exitp
                else:
                    # 非重合法兰：管端吸附到匹配的法兰端点（焊接点）
                    fpt = (flange_face_pts or {}).get(id(fl))
                    if fpt is None or len(fl.end_points) < 2:
                        continue
                    fl_eps = [proj.page(e.xyz) for e in fl.end_points[:2]]
                    for pt_name, pend in (("a", a), ("b", b)):
                        w = min(fl_eps, key=lambda q: _v2_len(_v2_sub(pend, q)))
                        if _v2_len(_v2_sub(pend, w)) < 1.5:
                            if pt_name == "a":
                                a = w
                            else:
                                b = w
            if cont_copy:
                canvas.dashed(a, b, width=0.35)
            else:
                canvas.line(a, b, width=0.55)
            return _v2_mid(eps[0], eps[1])

    elif cls == "ELBOW":
        if len(eps) >= 2:
            corners = cps if cps else [_v2_mid(eps[0], eps[1])]
            if proj.proportional:
                # 按比例：腿=真实端点（上游画法），不做示意腿长生长
                chain = [eps[0], *corners, eps[1]]
                b_end = b1s[0] if b1s else None
            else:
                chain = [_grown_leg_end(eps[0], corners[0], natural=leg_mm,
                                        own_id=id(comp)),
                         *corners,
                         _grown_leg_end(eps[1], corners[-1], natural=leg_mm,
                                        own_id=id(comp))]
                b_end = (_grown_leg_end(b1s[0], corners[0], natural=leg_mm,
                                        own_id=id(comp))
                         if b1s else None)
            if cont_copy:
                for p, q in zip(chain, chain[1:]):
                    canvas.dashed(p, q, width=0.3)
            else:
                canvas.polyline(chain, width=0.55)
            branch = b_end
            if branch is not None and corners:
                if cont_copy:
                    canvas.dashed(corners[0], branch, width=0.3)
                else:
                    canvas.line(corners[0], branch, width=0.55)
            return corners[0]

    elif cls == "TEE":
        anchor = None
        centre = cps[0] if cps else None
        if cont_copy:
            def stroke(p, q, _w=0.3, **_kw):
                canvas.dashed(p, q, width=0.3)
        else:
            stroke = canvas.line
        if len(eps) >= 2:
            c0 = centre if centre is not None else _v2_mid(eps[0], eps[1])
            if proj.proportional:
                # 按比例：腿=真实端点（上游画法），不做示意腿长生长
                stroke(eps[0], eps[1], width=0.55)
            else:
                stroke(_grown_leg_end(eps[0], c0, natural=leg_mm,
                                      own_id=id(comp)),
                       _grown_leg_end(eps[1], c0, natural=leg_mm,
                                      own_id=id(comp)),
                       width=0.55)
            anchor = c0
        if centre is not None:
            # 变径三通就地标签：支管端旁标 '大X小'（主管口径取端点 bore）
            br_bores = ([p.bore for p in comp.branch1_points]
                        + [p.bore for p in comp.branch2_points])
            for leg_i, (branch, bb) in enumerate(zip((*b1s, *b2s), br_bores)):
                if proj.proportional:
                    stroke(centre, branch, width=0.55)
                    continue
                tip = _grown_leg_end(branch, centre, natural=leg_mm,
                                     own_id=id(comp))
                stroke(centre, tip, width=0.55)
                if not cont_copy:
                    run_bore = max((p.bore for p in comp.end_points),
                                   default=None)
                    tag = _bore_tag_text((run_bore, bb))
                    if tag:
                        # 标在三通中心点旁、文字顺主管向（run 轴）；
                        # 多支腿（CROSS）逐腿抬 gap 防叠字
                        u = _v2_sub(eps[1], eps[0]) if len(eps) >= 2 \
                            else _v2_sub(tip, centre)
                        _draw_bore_tag(canvas, proj, centre, u, tag,
                                       gap=3.0 + 2.5 * leg_i)
            anchor = centre
        return anchor

    elif cls == "OLET":
        centre = cps[0] if cps else (eps[0] if eps else None)
        branch = b1s[0] if b1s else (eps[1] if len(eps) > 1 else None)
        if centre is not None and branch is not None:
            run = (olet_header_pts or {}).get(id(comp))
            header = (_v2_sub(proj.page(run[1]), proj.page(run[0]))
                      if run else None)
            connected = (olet_branch_connected is None
                         or id(comp) in olet_branch_connected)
            # 变径就地标签：支管端旁 '大X小'。支管 bore 在自身记录，
            # 母管口径由 _olet_header_points 几何匹配宿主管时带回
            if not cont_copy and not proj.proportional:
                host_bore = run[2] if run and len(run) > 2 else None
                tag = _bore_tag_text((host_bore,
                                      comp.branch1_points[0].bore
                                      if comp.branch1_points else None))
                if tag:
                    # 文字顺着宿主管向（几何匹配的 run 向），锚仍在支管端外
                    u = header or _v2_sub(branch, centre)
                    _draw_bore_tag(canvas, proj,
                                   _v2_add(branch, _v2_scale(_v2_norm(_v2_sub(branch, centre)), 1.2)),
                                   u, tag)
            # SKEY 符号库优先（WT 支管台等）：主管中心→支管端两接点对齐，
            # 跨向取世界轴投影；支管未连接或按比例模式保持 _draw_olet 原行为
            if connected and not proj.proportional:
                c3 = (comp.centre_points[0].xyz if comp.centre_points
                      else comp.end_points[0].xyz)
                b3 = comp.branch1_points[0].xyz
                r3 = tuple(b3[i] - c3[i] for i in range(3))
                if _draw_symbol_2pt(canvas, comp.skey, centre, branch,
                                    proj=proj, run3=r3, wp=c3,
                                    proportional=proj.proportional,
                                    dashed=cont_copy,
                                    width=0.3 if cont_copy else 0.4):
                    return _v2_mid(centre, branch)
            return _draw_olet(canvas, proj, centre, branch, header, connected)
        if centre is not None:
            canvas.circle(centre, 0.8, width=0.3)
            return centre

    elif cls == "VALVE":
        if len(eps) >= 2:
            mid = cps[0] if cps else _v2_mid(eps[0], eps[1])
            _draw_valve(canvas, proj, eps, mid, comp, bore_in,
                        widen=flange_widen)
            return mid

    elif cls == "FLANGE":
        # Only reached when the face could not be resolved (a single-ended
        # or point-only flange record).
        if len(eps) >= 2:
            _draw_flange(canvas, proj, eps[0], eps[1], comp, True, flange_bar)
            return _v2_mid(eps[0], eps[1])
        if len(eps) == 1:
            _draw_bar(canvas, eps[0], (1.0, 0.0), flange_bar / 2, width=0.6)
            return eps[0]

    elif cls == "REDUCER":
        if len(eps) >= 2:
            _draw_reducer(canvas, proj, comp, eps,
                          (reducer_runs or {}).get(id(comp)), doc.units_bore)
            if b1s:
                canvas.line(_v2_mid(eps[0], eps[1]), b1s[0], width=0.55)
            if not cont_copy and not proj.proportional:
                tag = _bore_tag_text((comp.end_points[0].bore,
                                      comp.end_points[-1].bore))
                if tag:
                    _draw_bore_tag(canvas, proj, _v2_mid(eps[0], eps[1]),
                                   _v2_sub(eps[1], eps[0]), tag)
            return _v2_mid(eps[0], eps[1])

    elif cls == "GASKET":
        # 垫片不另画符号或球标，其项号合并到父件球标。
        if len(eps) >= 2:
            return None

    elif cls == "CAP":
        if eps:
            end = eps[0]
            axis = _v2_norm(_v2_sub(eps[0], eps[1])) if len(eps) > 1 else (1.0, 0.0)
            _draw_cap(canvas, proj, end, axis)
            return end

    elif cls == "COUPLING":
        if len(eps) >= 2:
            # 联管节 gap 在模型空间测量（页空间会被 NTS warp 拉大误判）：
            # 超过阈值才画本体线+横杠，小 gap 焊耦合并入焊点语义不画符号
            gap_in = (comp.end_points[0].distance_to(comp.end_points[1])
                      * coord_to_inch_factor(doc.units_coords))
            if gap_in <= GAP_COUPLING_MAX_INCHES and is_weld_coupling(
                    comp.type, comp.description, comp.skey):
                # 焊接联管节经焊点派生呈现、不画符号，但中心线仍须连续：
                # warp 可把 1/16" gap 拉到近 1cm，不描线会在每个焊点处断线
                canvas.line(eps[0], eps[1], width=0.55)
                return None
            axis = _v2_norm(_v2_sub(eps[1], eps[0]))
            mid = _v2_mid(eps[0], eps[1])
            canvas.line(eps[0], eps[1], width=0.9)
            _draw_bar(canvas, mid, axis, 1.4)
            return mid

    elif cls == "SUPPORT":
        at = cos[0] if cos else (eps[0] if eps else None)
        if at is not None:
            _draw_support_symbol(canvas, at, comp, proj, doc)
            return at

    elif cls in ("INSTRUMENT", "FILTER", "MULTIPORT"):
        if len(eps) >= 2:
            _draw_box_on_axis(canvas, proj, eps[0], eps[1], bore_in,
                              widen=flange_widen)
            return _v2_mid(eps[0], eps[1])
        if pps and cps:
            for pp in pps:
                canvas.line(cps[0], pp, width=0.4)
            canvas.circle(cps[0], 1.0, width=0.3)
            return cps[0]
        if eps and cos:
            canvas.dashed(eps[0], cos[0])
            canvas.circle(cos[0], 1.6, width=0.3)
            return cos[0]

    elif cls == "BOLT":
        return None

    elif cls == "GAP":
        if len(eps) >= 2:
            canvas.dashed(eps[0], eps[1], width=0.3)
            return _v2_mid(eps[0], eps[1])

    elif cls == "CONNECTION":
        # Nothing: 接图点由 continuation 阶段负责画——图上该出现的是下一张
        # 图名，本记录只知道管线接续；render_pdf 的 real_components() 到不了这里
        return None

    elif cls == "MESSAGE":
        if cos:
            message = comp.get("TEXT") or comp.get("MESSAGE") or ""
            if message:
                canvas.text(cos[0], message, size=5.5)
        return None

    elif cls == "WELD":
        # 焊点符号由焊点阶段画；此处再描本体线会在每个焊点下多一条线，
        # socket weld 变成"点上有线"
        if len(eps) >= 2:
            return _v2_mid(eps[0], eps[1])
        return eps[0] if eps else (cos[0] if cos else None)

    else:
        if len(eps) >= 3 and _v2_len(_v2_sub(eps[1], eps[0])) > 1e-6:
            # 未知多口件（≥3 端点）：画 NC 式箱体（Isogen NC 符号——矩形箱
            # 体，两端接点+上下 tap）；第 3 个起的其余端点从箱体中心各拉一条
            # tap 线到真实坐标。
            a, b = eps[0], eps[1]
            u = _v2_norm(_v2_sub(b, a))
            ep3 = comp.end_points
            run3 = (tuple(ep3[1].xyz[i] - ep3[0].xyz[i] for i in range(3))
                    if len(ep3) >= 2 else None)
            v = (_across_v(proj, run3, ep3[0].xyz, u) or _v2_perp(u))
            ln = max(_v2_len(_v2_sub(b, a)), 4.0)
            mid = _v2_mid(a, b)
            h = 4.5
            corners = [
                _v2_add(mid, _v2_add(_v2_scale(u, ln / 2), _v2_scale(v, h / 2))),
                _v2_add(mid, _v2_add(_v2_scale(u, ln / 2), _v2_scale(v, -h / 2))),
                _v2_add(mid, _v2_add(_v2_scale(u, -ln / 2), _v2_scale(v, -h / 2))),
                _v2_add(mid, _v2_add(_v2_scale(u, -ln / 2), _v2_scale(v, h / 2))),
            ]
            canvas.polyline([*corners, corners[0]], width=0.45)
            for tap in eps[2:]:
                canvas.line(mid, tap, width=0.45)
            return mid
        if len(eps) >= 2:
            canvas.line(eps[0], eps[1], width=0.45)
            return _v2_mid(eps[0], eps[1])
        if eps and cps:
            canvas.line(eps[0], cps[0], width=0.45)
            return cps[0]
        if cos:
            canvas.dot(cos[0], r=0.7)
            return cos[0]

    return None


def _draw_bar(canvas: _Canvas, centre: Vec2, axis: Vec2, half_len: float,
              width: float = 0.45) -> None:
    perp = _v2_perp(axis)
    a = _v2_add(centre, _v2_scale(perp, half_len))
    b = _v2_add(centre, _v2_scale(perp, -half_len))
    canvas.line(a, b, width=width)


# How much better one candidate axis must score to beat the fixed preference
# order. See _iso_cross_dir -- this is a tie band, not a float epsilon.
_ISO_CROSS_TIE_BAND = 1e-3


def _iso_cross_dir(proj: Projector, run: Vec2) -> Vec2:
    """A projected world axis to draw across ``run``, on an isometric plane.

    Screen-perpendicular bars read as flat marks pasted onto an isometric
    picture; the world axis most nearly perpendicular puts the bar on an iso
    plane at 30 degrees. Elevation first so a flange on any horizontal run
    draws upright; the tie band must stay wide -- runs along an iso axis leave
    the other two axes scoring identically, so ties decide by fixed order.
    """
    dirs = _axis_screen_dirs(proj)
    best, best_score = None, 0.0
    for d in (dirs[2], dirs[0], dirs[1]):
        score = abs(d[0] * run[1] - d[1] * run[0])   # |cross| = perpendicularity
        if score > best_score + _ISO_CROSS_TIE_BAND:
            best, best_score = d, score
    if best is None:
        return _v2_perp(run)
    return best


# How far back from the gasket face the black face bar stands, in page mm --
# opens a black|white|black gap on a bolted pair that reads as the gasket.
_FLANGE_GASKET_GAP_MM = 0.55


def _draw_flange(canvas: _Canvas, proj: Projector, e0: Vec2, e1: Vec2,
                 comp: Component, face_first: bool, bar_mm: float) -> None:
    """Flange body plus one long iso-plane bar just behind the gasket face.

    Bar on the flanged end only (welded end already has its weld dot), ~2x a
    dimension offset long; body stops _FLANGE_GASKET_GAP_MM short of the face
    with an unrecorded white knockout across it -- on a bolted pair that draws
    black|white|black, reading as the gasket.
    """
    face, back = (e0, e1) if face_first else (e1, e0)
    run = _v2_norm(_v2_sub(face, back))
    if run == (0.0, 0.0):
        run = (1.0, 0.0)
    # A flange drawn shorter than the gap itself keeps its full body: better
    # a joint with no gasket gap than a flange with no body.
    gap = (_FLANGE_GASKET_GAP_MM
           if _v2_len(_v2_sub(face, back)) > 3 * _FLANGE_GASKET_GAP_MM else 0.0)
    bar_at = _v2_sub(face, _v2_scale(run, gap))
    canvas.line(back, bar_at, width=0.55)
    across = _iso_cross_dir(proj, run)
    half = bar_mm / 2
    if is_slip_on_flange(comp.description):
        half *= 1.15
    # Knock the joint clear first, so a neighbour drawn earlier through this
    # point cannot fill the gasket gap in.
    with canvas.pdf.local_context():
        canvas.pdf.set_draw_color(255, 255, 255)
        canvas.line(_v2_add(face, _v2_scale(across, half)),
                    _v2_add(face, _v2_scale(across, -half)),
                    width=_FLANGE_GASKET_GAP_MM * 0.9, record=False)
    a = _v2_add(bar_at, _v2_scale(across, half))
    b = _v2_add(bar_at, _v2_scale(across, -half))
    canvas.line(a, b, width=0.6)
    if "BLIND" in comp.type.upper() or "blind" in (comp.description or "").lower():
        # Blind: nothing passes, so double the face bar.
        offset = _v2_scale(run, -0.9)
        canvas.line(_v2_add(a, offset), _v2_add(b, offset), width=0.6)


_GASKET_GAP_MM = 0.6    # 法兰符号与配合面之间的垫片间隙（图面视觉厚度）

# 本轮渲染中已画出的"端点重合法兰"接带（top=光面端, exit=焊端, uu=轴向），
# 供焊点吸附（焊点应落在法兰带上而不是模型投影的空隙里）
_FLANGE_BANDS: list = []
# _flange_band 记忆化缓存：(id(comp), pre) → 结果。直管修剪循环每根管×
# 每个法兰都要问一次接带，不缓存时 O(管×法兰×全组件) 三重扫描。
_FLANGE_BAND_CACHE: dict[tuple[int, bool], tuple | None] = {}

# 非重合法兰的焊端锚点（页坐标）：焊点圆点吸附用，每张图重置
_FLANGE_WELDS: list[Vec2] = []

# 非管件占位点（页坐标, 组件 id）：符号生长避让用，每张图重置。
# PIPE/WELD/BOLT/GASKET 不登记——它们与零件同轴相邻，会把生长空间全锁死。
_NONPIPE_PAGE_PTS: list[tuple[Vec2, int]] = []

_SCHEMATIC_LEG_MM = 6.75  # 弯头/三通腿外延目标（page mm，×1.5 管线放大）


def _growth_room(p: Vec2, direction: Vec2, own_id: int | None = None) -> float:
    """符号端点 p 沿 direction 可生长的图面距离（mm）。

    直接焊接的邻件（占位点贴着 p）锁死返回 0；隔管邻件留 1.5mm 净距；
    无遮挡返回 1e9。组件自身端点按 own_id 排除。
    """
    best = 1e9
    for (qx, qy), cid in _NONPIPE_PAGE_PTS:
        if own_id is not None and cid == own_id:
            continue
        dx, dy = qx - p[0], qy - p[1]
        proj = dx * direction[0] + dy * direction[1]
        if proj <= -0.3 or proj >= best + 0.8:   # 只看前方（含贴身焊接点）
            continue
        lat = abs(dx * direction[1] - dy * direction[0])
        if lat > 2.5:
            continue
        if proj < 0.8:            # 贴身：焊接邻件，锁死
            return 0.0
        best = proj - 1.5
    return max(best, 0.0)


def _grown_leg_end(end: Vec2, corner: Vec2, natural: float = _SCHEMATIC_LEG_MM,
                   own_id: int | None = None) -> Vec2:
    """弯头/三通腿从 corner 向 end 外延到示意腿长，受邻件空间限制。"""
    d = _v2_sub(end, corner)
    ln = _v2_len(d)
    if ln <= 1e-9 or ln >= natural:
        return end
    u = (d[0] / ln, d[1] / ln)
    grow = min(natural - ln, _growth_room(end, u, own_id))
    return _v2_add(end, _v2_scale(u, grow)) if grow > 0 else end


def _flange_band(proj, doc, comp, flange_face_pts, pre: bool = False):
    """端点重合法兰（平焊等）的接带布局：(top, exit, uu) 或 None。

    top = 光面端（从配合面向管侧退 _GASKET_GAP_MM）；exit = 焊端（越过焊点
    0.6mm）；uu = 从接头指向管道远端的单位轴向。主管端点/焊点应修剪·吸附
    到 exit——实现"主管接焊端、阀门空开光面端"的接头布局。
    pre=True 供 _symbol_face_offsets 预计算：全部用无偏移投影
    （apply_offset=False），在连接点偏移合并之前建立基准。
    """
    eps3 = comp.end_points
    if len(eps3) < 2:
        return None
    _key = (id(comp), pre)
    if _key in _FLANGE_BAND_CACHE:
        return _FLANGE_BAND_CACHE[_key]
    pp = (lambda p: proj.page(p, apply_offset=False)) if pre else proj.page
    e0 = pp(eps3[0].xyz)
    e1 = pp(eps3[1].xyz)
    if _v2_len(_v2_sub(e1, e0)) >= 0.5:
        _FLANGE_BAND_CACHE[_key] = None
        return None
    best = None
    try:
        for c in doc.real_components():
            if classify(c.type, c.skey) != "PIPE" or len(c.end_points) < 2:
                continue
            p0 = pp(c.end_points[0].xyz)
            p1 = pp(c.end_points[1].xyz)
            near, far = ((p0, p1) if _v2_len(_v2_sub(p0, e0))
                         <= _v2_len(_v2_sub(p1, e0)) else (p1, p0))
            dn = _v2_len(_v2_sub(near, e0))
            if dn < 2.0 and (best is None or dn < best[0]):
                best = (dn, near, far)
    except Exception:
        _FLANGE_BAND_CACHE[_key] = None
        return None
    if best is None or _v2_len(_v2_sub(best[2], e0)) <= 1.0:
        _FLANGE_BAND_CACHE[_key] = None
        return None
    weld = best[1]
    uu = _v2_norm(_v2_sub(best[2], weld))
    fpt = (flange_face_pts or {}).get(id(comp))
    if fpt is not None:
        face = pp(fpt)
        top = _v2_add(face, _v2_scale(uu, _GASKET_GAP_MM))
    else:
        top = _v2_add(weld, _v2_scale((-uu[0], -uu[1]), 1.2))
    span = max(_v2_len(_v2_sub(weld, top)) + 0.6, 1.2)
    band = (top, _v2_add(top, _v2_scale(uu, span)), uu)
    _FLANGE_BAND_CACHE[_key] = band
    return band


def _across_v(proj: "Projector | None", run3: tuple | None,
              wp: tuple | None, u: Vec2) -> Vec2 | None:
    """垂直于管轴的世界轴的图面投影。

    优先取管轴分量最小的世界轴：纯 X/Y/Z 管仍沿用原来的横轴规则，
    X/Z 空间斜管则稳定跨 Y。不能用“最大分量”决定横轴——理论等长的
    两个斜管分量只要有浮点尾数差，阀与法兰就会跨到不同世界轴。
    """
    if proj is None or run3 is None or wp is None:
        return None
    for axis in sorted(range(3), key=lambda i: (abs(run3[i]), i)):
        across = tuple(1.0 if i == axis else 0.0 for i in range(3))
        # 点偏移只平移最终图面位置；方向两点必须同时取原始投影，
        # 否则单侧偏移会被误算成旋转。
        a2 = proj.page(wp, apply_offset=False)
        b2 = proj.page((wp[0] + across[0], wp[1] + across[1], wp[2] + across[2]),
                       apply_offset=False)
        dx, dy = b2[0] - a2[0], b2[1] - a2[1]
        ln = (dx * dx + dy * dy) ** 0.5
        if ln <= 1e-12:
            continue
        vv = (dx / ln, dy / ln)
        if abs(vv[0] * u[0] + vv[1] * u[1]) < 0.99:
            return vv
    return None


_VALVE_OPS: dict[str, str] | None = None


def _valve_operators() -> dict[str, str]:
    """阀体 SKEY -> 阀柄 SKEY 映射（symbols/operators.json）。

    阀体与阀柄在符号库里是两个独立符号（柄需要按 SPINDLE-DIRECTION
    单独适配朝向，不能并进阀体几何）。by_skey 键支持 "VV**" 形式
    （** = 两个端接字符），"default" 兜底所有未列出的阀门。
    """
    global _VALVE_OPS
    if _VALVE_OPS is None:
        ops: dict[str, str] = {}
        try:
            from .asc_symbols import _candidate_dirs
            for d in _candidate_dirs():
                p = d / "operators.json"
                if p.is_file():
                    data = json.loads(p.read_text(encoding="utf-8"))
                    ops.update(data.get("by_skey") or {})
                    if data.get("default"):
                        ops.setdefault("**", data["default"])
                    break
        except (OSError, ValueError):
            pass
        _VALVE_OPS = ops
    return _VALVE_OPS


def _valve_operator_skey(skey: str | None) -> str | None:
    """阀柄 SKEY：operators.json 显式映射优先（带 ** 前缀通配），其次
    ASC 符号库厂商配对（501 行第 2 名，如 `501 VB 03SP` → VB** 挂杠杆
    03SP），最后落 default。"""
    if not skey:
        return None
    ops = _valve_operators()
    s = skey.upper()
    for pat, op in ops.items():
        p = pat.upper()
        if p == "**":
            continue
        if (p.endswith("**") and s.startswith(p[:-2])) or p == s:
            return op
    from .asc_symbols import operator_skey_of
    paired = operator_skey_of(s)
    if paired:
        return paired
    return ops.get("**")


def _draw_symbol_2pt(canvas: _Canvas, skey: str | None, p0: Vec2, p1: Vec2,
                     proj: "Projector | None" = None,
                     run3: tuple | None = None, wp: tuple | None = None,
                     axis: Vec2 | None = None,
                     drop_isolated: bool = False,
                     width: float = 0.4,
                     height_mm: float = 4.5,
                     across: Vec2 | None = None,
                     dashed: bool = False,
                     bore_ratio: float | None = None,
                     proportional: bool = False,
                     room0: float | None = None,
                     room1: float | None = None,
                     no_grow: bool = False,
                     ecc_dir: Vec2 | None = None,
                     width_mm: float | None = None,
                     operator: tuple[str, Vec2] | None = None) -> bool:
    """SKEY 符号库双接点对齐绘制（ASC 优先 → skeys.json 官方库）。

    两种出图模式（proportional）：按比例=True 时符号铺满实际端距——
    sx=端距/跨度，短端距生长（room0/room1 限邻件空间）、超限钳制
    |sx|≤2·sy，同心变径带轴向残差校正；不按比例（示意，默认）时符号
    恒等比——sx 恒等于 sy，START 锚 p0、EXIT 落在等比自然长度处，不
    为凑端点变形，端点/尺寸锚由 _symbol_face_offsets 随符号端面平移。
    符号 X 轴直接映射管轴 u
    （span 取 EXIT 的 X 分量，偏心大小头离轴出口不扭斜端面），普通符号
    u=p0→p1。横向 v 默认取"垂直于管轴的世界轴"的图面投影（等轴测 30°
    斜向族，与支撑符号同规则）；调用方可用 across 直接指定横向方向
    （设备管口符号与管线法兰杠同用 _iso_cross_dir，保证法兰面对齐同一
    斜向族）。符号整体高度（Y 跨度）固定 height_mm；偏心大小头按口径比
    正算小端高度和偏心量，X 基始终沿主管轴，平齐底边保持平行。小端侧
    管线由 Projector.point_offsets 整体接到正算出口，不画折线连接。
    未查到符号或几何退化时返回 False，调用方走手写回退分支。
    drop_isolated：剔除与任何其他腿不共端点的孤立腿——大小头的 tap 标记
    线（支管接入位置指示，RE** 中部短竖线），PCF 无 BRANCH-POINT 时不画
    （ISOGEN 同约定：真有支管接上才显示）。
    """
    try:
        from .asc_symbols import find_symbol_full
        found = find_symbol_full(skey or "")
    except Exception:
        return False
    if not found:
        return False
    sym, orient, stop = found
    if (axis is not None and bore_ratio is not None and stop
            and 0.0 < bore_ratio < 1.0 and abs(stop[1]) > 1e-9):
        # 偏心变径正算：大端全高固定 height_mm，小端全高=大端×口径比，
        # 小端中心偏移=(大端高-小端高)/2。保留 RE** 轮廓与平齐边，
        # 不用页面端点残差反推高度。
        ys0 = [c for seg in sym for c in (seg[1], seg[3])]
        big_half = (max(ys0) - min(ys0)) / 2.0
        small_half = big_half * bore_ratio
        centre = math.copysign(big_half - small_half, stop[1])
        if centre < 0:
            small_bottom, small_top = -big_half, centre + small_half
        else:
            small_bottom, small_top = centre - small_half, big_half
        x1 = stop[0]
        x2 = max(c for seg in sym for c in (seg[0], seg[2]))
        sym = [
            (0.0, -big_half, 0.0, big_half),
            (0.0, big_half, x1, small_top),
            (x1, small_top, x2, small_bottom),
            (x2, small_bottom, 0.0, -big_half),
            *sym[4:],
        ]
        stop = (x1, centre)
    diff = _v2_sub(p1, p0)
    d = (diff[0] * diff[0] + diff[1] * diff[1]) ** 0.5
    # 轴向跨度取收笔点的 X 分量：偏心大小头（RE** EXIT=(548,-136)）的
    # 出口离轴，若用 START→EXIT 对角线对准管轴会把整个符号扭斜、端面
    # 不再垂直管轴——直接映射符号 X→管轴，偏心量保留为侧向偏移
    span = abs(stop[0]) if stop else 0.0
    if d <= 1e-6 or span <= 1e-6:
        return False
    # 收笔点在 -X 侧（官方 NZFE 定义 EXIT=(-800,0)，面杠在 START 的负向）
    # 时镜像符号 X 轴：START 仍落 p0，形状左右翻转
    sgn = 1.0 if (stop and stop[0] > 0) else -1.0
    # axis：外部指定的管轴（偏心大小头用主管走向，不用 p0→p1 对角线）
    if axis is not None:
        u = axis
        if proportional:
            d = (diff[0] * u[0] + diff[1] * u[1]) or d   # 轴向投影距离
    else:
        u = (diff[0] / d, diff[1] / d)
    # 横向：调用方指定（管口符号与管线法兰杠同族）优先，否则端面线沿
    # "垂直于管轴的世界轴"的图面投影（等轴测 30° 斜向族）；无投影信息时
    # 回退图面垂直
    v = across or (_across_v(proj, run3, wp, u) or _v2_perp(u))
    if ecc_dir is not None and abs(stop[1]) > 1e-9:
        # 偏心方向按模型实测定向：符号小端偏在 sign(stop[1])·v 侧，
        # 模型偏心投影（ecc_dir）与之相反时镜像符号——顶平/底平不画反
        if (ecc_dir[0] * v[0] + ecc_dir[1] * v[1]) * stop[1] < 0:
            v = (-v[0], -v[1])
    ys = [c for seg in sym for c in (seg[1], seg[3])]
    perp = (max(ys) - min(ys)) if ys else 1.0
    if perp < 0.05:
        # 全水平线段的扁平符号：Y 跨度≈0 会让 sy 爆到 1e6 量级，画出
        # 公里级线段并污染碰撞几何——弃符号走手绘回退
        return False
    sy = height_mm / max(perp, 1e-6)
    if proj is not None and getattr(proj, "schematic_scale", 1.0) != 1.0:
        # 链式布局整图缩放：符号尺寸与位置表同乘，保持搭积木间距
        sy *= proj.schematic_scale
    if proportional:
        # 按比例模式：符号长度严格等于实际端距（sx=端距/跨度）——不生长、
        # 不钳制，符号两端与真实坐标投影精确对齐；只有高度保持可读
        # （符号是示意表达，压缩到端距宽的阀/法兰在图上会缩成窄条，
        # 这是"真实比例"语义下的预期可读性代价）。
        sx = sgn * d / span
    else:
        # 不按比例（示意）：符号恒等比，sx 恒等于 sy（镜像号除外）——
        # 轴向不单独拉伸，也不对个别符号加大/缩小。START 锚 p0 沿管轴
        # 伸出自然长度；与端点的差值由 _symbol_face_offsets 带动连接点
        # 消化（管端/焊点/尺寸锚随符号端面走，warp 目标同源等比锚距）。
        sx = sgn * sy
    if width_mm is not None:
        # 横向宽度按孔径独立覆盖（与手绘回退/法兰杠同宽）：恒等比只在缺省时
        # 成立——height_mm 的 4.5 同时校准轴向自然长度（VB** X 跨 800→9.0mm、
        # FLWN 500→4.09mm 均等于链式占位），改高度会连带缩放长度、链上留死区，
        # 所以宽度走单独通道。
        sy = width_mm / max(perp, 1e-6)
        if proj is not None and getattr(proj, "schematic_scale", 1.0) != 1.0:
            sy *= proj.schematic_scale
    # 符号基向量：P = p0 + x·bx + y·by（普通符号 bx=sx·u, by=sy·v）。
    # 偏心变径不做任何 X 剪切：bx=sx·u、by=sy·v，平齐边自然与主管平行。
    bx = (sx * u[0], sx * u[1])
    by = (sy * v[0], sy * v[1])
    if proportional and axis is not None and abs(stop[1]) <= 1e-9:
        # 按比例模式的同心符号：校正轴向数值残差，EXIT 精确落 p1。
        rx = diff[0] - d * u[0]
        ry = diff[1] - d * u[1]
        bx = (bx[0] + rx / span, bx[1] + ry / span)
    x0, y0 = p0
    if drop_isolated and len(sym) > 1:
        # 孤立腿 = 与其他任何腿都不共端点的腿：轮廓腿围成闭环共端点，tap 标记线
        # 浮在体内（端点落在轮廓边线内部）——按 ISOGEN 约定组件无 BRANCH-POINT 时不画
        def _near(p, q, tol=1.0):
            return abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol
        kept = []
        for i, s1 in enumerate(sym):
            a1, b1 = (s1[0], s1[1]), (s1[2], s1[3])
            touches = any(
                i != j and (_near(a1, (s2[0], s2[1]))
                            or _near(a1, (s2[2], s2[3]))
                            or _near(b1, (s2[0], s2[1]))
                            or _near(b1, (s2[2], s2[3])))
                for j, s2 in enumerate(sym))
            if touches:
                kept.append(s1)
        if kept:
            sym = kept
    draw = canvas.dashed if dashed else canvas.line
    pts: list[Vec2] = []
    for x1, y1, x2, y2 in sym:
        # 虚线符号（设备管口）是实体几何：记入 segments，尺寸放置与
        # 球标布线才会避让——否则 goalpost 直接压在管口符号上。
        a = (x0 + x1 * bx[0] + y1 * by[0],
             y0 + x1 * bx[1] + y1 * by[1])
        b = (x0 + x2 * bx[0] + y2 * by[0],
             y0 + x2 * bx[1] + y2 * by[1])
        pts.append(a)
        pts.append(b)
        draw(a, b, width=width, record=True)
    if operator and pts:
        # 阀柄叠加：柄是独立 skey 的符号（01SP=T 形柄、03SP=杠杆、04SP=手轮、
        # 10SP=疏水器浮子…，独立块见 ALIAS.asc 的 N**SP），厂商坐标系以阀中心
        # 连接点为 START、阀杆沿符号 +X 伸出；已归一到 START=(0,0)，锚到阀中心
        # 连接点、+X 旋到 SPINDLE-DIRECTION 图面投影，缩放沿用阀体高度标尺 sy。
        op_skey, sp2 = operator
        from .asc_symbols import find_symbol_full
        op_sym = find_symbol_full(op_skey)[0] if op_skey else None
        dl = (sp2[0] * sp2[0] + sp2[1] * sp2[1]) ** 0.5
        if op_sym and dl > 1e-9:
            ex = (sp2[0] / dl, sp2[1] / dl)          # 符号 +X=阀杆伸出方向
            ey = (ex[1], -ex[0])
            anchor = (x0 + stop[0] / 2.0 * bx[0],
                      y0 + stop[0] / 2.0 * bx[1])    # 阀中心连接点
            for x1, y1, x2, y2 in op_sym:
                draw((anchor[0] + sy * (x1 * ex[0] + y1 * ey[0]),
                      anchor[1] + sy * (x1 * ex[1] + y1 * ey[1])),
                     (anchor[0] + sy * (x2 * ex[0] + y2 * ey[0]),
                      anchor[1] + sy * (x2 * ex[1] + y2 * ey[1])),
                     width=width, record=True)
    return True


def _bore_to_in(bore: float | None, units_bore: str) -> float | None:
    """关键点孔径换算成英寸（UNITS-BORE 感知）；缺孔径返回 None。"""
    if bore is None:
        return None
    return bore / 25.4 if units_bore.startswith("MM") else bore


def _bore_half_w(bore_in: float | None) -> float:
    """按比例图的符号横半宽（页 mm）：与孔径严格成正比（过原点直线）。

    1.2mm/英寸孔径：1" 全宽 2.4、2" 4.8、6" 14.4——宽度随管径翻倍。
    仅按比例（真实比例）模式使用；不按比例图用固定尺寸（阀 1.9，
    见 _draw_component）。缺孔径回退 1.9（固定宽度，两模式兜底一致）；
    1/8" 下限只挡脏数据（负/零孔径）。
    """
    if bore_in is None:
        return 1.9
    return 1.2 * max(bore_in, 0.125)


# 法兰突面(垫片面)直径 ÷ 管外径，随标称管径 D(英寸)变化的平滑函数。法兰垫
# 片面不随管径等比放大：小管径相对更"肥"（1/2" 突面是管外径的 1.64 倍），
# 越大越瘦（24" 只剩 1.14 倍）。下式对 ASME B16.5 Class 150 全表（1/2"~24"
# 21 档，各压力等级同尺寸突面直径相同）最小二乘拟合，最大偏差 0.05（≤0.15mm 不可辨），
# 免查表、任意孔径连续取值。数据锚点：1/2"→1.64、2"→1.53、6"→1.28、12"→1.18、24"→1.14。
def _flange_face_ratio(bore_in: float) -> float:
    """突面直径/管外径（B16.5 表拟合曲线）；bore_in 为英寸孔径。"""
    # 钳到最小标称管径 1/8"：导出脏数据（负/零/近 -6.3 孔径）会让分母
    # 归零或比例爆炸，炸掉整张图；钳后恒取曲线单调段，图面不受扰
    return 0.962 + 4.30 / (max(bore_in, 0.125) + 6.30)


def _flanged_ports(comp: Component) -> bool:
    """任一端口为法兰连接：END-POINT 行尾连接码 FL 优先，缺码按 SKEY
    尾两字符兜底（FL/SO，见 skeys.connection_from_skey）。"""
    if any(pt.connection_type == "FL" for pt in comp.end_points):
        return True
    return connection_from_skey(comp.skey) == "FL"


def _draw_valve(canvas: _Canvas, proj: Projector, eps: list[Vec2],
                mid: Vec2, comp: Component,
                bore_in: float | None = None,
                widen: float = 1.0) -> None:
    # 回退：程序化领结 + 子类型装饰（SKEY 符号库命中时不会走到这里，
    # 统一入口见 _draw_component 顶部的 _draw_symbol_2pt）
    half_w = _bore_half_w(bore_in) * widen
    for end in eps[:4]:
        axis = _v2_norm(_v2_sub(mid, end))
        # 底边沿轴测平面横轴（与法兰面杠/变径端面同族），不用页面垂直
        # ——页面垂直不在任何轴测平面上，按比例长阀读起来像贴上去的贴纸
        perp = _iso_cross_dir(proj, axis) if proj is not None else _v2_perp(axis)
        a = _v2_add(end, _v2_scale(perp, half_w))
        b = _v2_add(end, _v2_scale(perp, -half_w))
        canvas.polyline([a, mid, b, a], width=0.4)
    subtype = valve_subtype(comp.description, comp.skey)
    if subtype == "BALL":
        with canvas.pdf.local_context():
            canvas.pdf.set_fill_color(255, 255, 255)
            canvas.circle(mid, 1.1, width=0.35, fill=True)
    elif subtype == "GLOBE":
        canvas.dot(mid, r=1.0)
    elif subtype == "CHECK":
        axis = _v2_norm(_v2_sub(eps[1], eps[0])) if len(eps) > 1 else (1.0, 0.0)
        perp = _v2_perp(axis)
        tip = _v2_add(mid, _v2_scale(perp, -2.2))
        canvas.line(mid, tip, width=0.35)
        canvas.dot(tip, r=0.5)
    elif subtype == "BUTTERFLY":
        axis = _v2_norm(_v2_sub(eps[1], eps[0])) if len(eps) > 1 else (1.0, 0.0)
        perp = _v2_perp(axis)
        canvas.line(_v2_add(mid, _v2_scale(perp, 1.8)),
                    _v2_add(mid, _v2_scale(perp, -1.8)), width=0.7)
    elif subtype == "PLUG":
        canvas.polyline([
            (mid[0] - 1.1, mid[1] - 1.1), (mid[0] + 1.1, mid[1] - 1.1),
            (mid[0] + 1.1, mid[1] + 1.1), (mid[0] - 1.1, mid[1] + 1.1),
        ], width=0.35, close=True)
    elif subtype == "RELIEF":
        canvas.text((mid[0], mid[1] - 2.4), "PSV", size=5.0)


_FLAT_DIRECTIONS: dict[str, Vec3] = {
    "UP": (0, 0, 1), "DOWN": (0, 0, -1),
    "NORTH": (0, 1, 0), "SOUTH": (0, -1, 0),
    "EAST": (1, 0, 0), "WEST": (-1, 0, 0),
}


def _reducer_flat_offset(comp: Component, big: Point, small: Point,
                         run: tuple[Vec3, Vec3] | None) -> Vec3:
    """Which way the flat of an eccentric reducer faces, in model space.

    Measured, not taken on trust: the small end's centreline is offset across
    the run *towards the flat*, by half the difference in outside diameters.
    The two sizes and the two centrelines are the whole rule, so it is read
    off the geometry rather than from FLAT-DIRECTION -- which only ever names
    a world axis, so it cannot describe a rolled run, and is missing from
    plenty of exports.

    FLAT-DIRECTION remains the fallback, for exports that put both weld ends
    on one centreline and leave the eccentricity implied.
    """
    if run is not None:
        d = tuple(s - b for b, s in zip(big.xyz, small.xyz))
        u = geometry.normalize(tuple(q - p for p, q in zip(run[0], run[1])))
        across = tuple(d[i] - geometry.dot(d, u) * u[i] for i in range(3))
        # 1% of the chord: a real eccentric offset is inches against a body of
        # inches, so anything smaller means the export did not model one.
        if geometry.length(across) > 0.01 * max(geometry.length(d), 1e-9):
            return across
    return _FLAT_DIRECTIONS.get((comp.get("FLAT-DIRECTION") or "").upper(),
                                (0, 0, -1))


def _draw_reducer(canvas: _Canvas, proj: Projector, comp: Component,
                  eps: list[Vec2], run: tuple[Vec3, Vec3] | None = None,
                  units_bore: str = "INCH") -> None:
    """Trapezoid spanning the reducer's own two weld ends, on an iso plane.

    Both body ends sit on the reducer's own projected weld points (never built
    out along a neighbour's run axis). ``run`` only squares the body -- the
    chord leans across an eccentric's run and faces squared to it lean with it.
    Cross direction is a projected world axis (`_iso_cross_dir`); the flat side
    is schematic, its side from the model (_reducer_flat_offset); the centreline
    offset stays true proportion (_eccentric_offset_pairs guards the warp).
    """
    pts = comp.end_points
    large_first = (pts[0].bore or 0.0) >= (pts[-1].bore or 0.0)
    big_pt, small_pt = (pts[0], pts[-1]) if large_first else (pts[-1], pts[0])
    big, small = (eps[0], eps[1]) if large_first else (eps[1], eps[0])
    # 两端半高按出图模式分开：按比例图各随自己的孔径（UNITS-BORE
    # 感知）；不按比例图用固定 1.9/0.9（原始示意尺寸），缺数据不画矩形
    prop_mode = bool(getattr(proj, "proportional", False))
    if prop_mode:
        big_in = _bore_to_in(big_pt.bore, units_bore)
        small_in = _bore_to_in(small_pt.bore, units_bore)
        half_big = _bore_half_w(big_in) if big_in is not None else 1.9
        half_small = _bore_half_w(small_in) if small_in is not None else 0.9
    else:
        half_big, half_small = 1.9, 0.9

    axis = (0.0, 0.0)
    if run is not None:
        axis = _v2_norm(_v2_sub(proj.page(run[1]), proj.page(run[0])))
        if axis != (0.0, 0.0):
            towards = _v2_sub(small, big)
            if axis[0] * towards[0] + axis[1] * towards[1] < 0:
                axis = _v2_scale(axis, -1.0)
    if axis == (0.0, 0.0):
        axis = _v2_norm(_v2_sub(small, big))
    if axis == (0.0, 0.0):
        return

    if reducer_is_eccentric(comp.type, comp.skey):
        offset = _reducer_flat_offset(comp, big_pt, small_pt, run)
        across = _v2_norm(_v2_sub(proj.raw(offset), proj.raw((0, 0, 0))))
        step = _step_across(_v2_sub(small, big), axis, across)
        if across != (0.0, 0.0) and step is not None:
            # How far the centreline moves across the run *on the page*: the body
            # must be at least that deep or the small end falls off the bottom and
            # the two long edges cross in a bowtie; growing the big end by exactly
            # the step also keeps the flat parallel to the run (flat is always +1).
            half_big = max(half_big, half_small + max(step, 0.0))
    else:
        across = _iso_cross_dir(proj, axis)

    if across == (0.0, 0.0):
        across = _iso_cross_dir(proj, axis)

    canvas.polyline([
        _v2_add(big, _v2_scale(across, half_big)),
        _v2_add(small, _v2_scale(across, half_small)),
        _v2_add(small, _v2_scale(across, -half_small)),
        _v2_add(big, _v2_scale(across, -half_big)),
    ], width=0.4, close=True)


def _step_across(chord: Vec2, axis: Vec2, across: Vec2) -> float | None:
    """The ``across`` component of ``chord``, in page mm.

    Solves ``chord = a*axis + b*across`` for ``b``. ``axis`` and ``across``
    are both projected *world* axes and so are not perpendicular on the page
    -- a dot product would give the wrong answer.
    """
    det = axis[0] * across[1] - axis[1] * across[0]
    if abs(det) < 1e-6:
        return None
    return (axis[0] * chord[1] - axis[1] * chord[0]) / det


def _bore_tag_text(bores) -> str | None:
    """变径就地标签文本：'大X小'（降序）；等径、缺径或不足两口征返回 None。

    口径取 PCF 记录的原始数值（本语料为 mm/DN），整数省略小数位。
    """
    vals = [float(b) for b in bores if b]
    if len(vals) < 2:
        return None
    big, small = max(vals), min(vals)
    if big - small < 1e-9:
        return None

    def fmt(b: float) -> str:
        return str(int(round(b))) if abs(b - round(b)) < 1e-9 else f"{b:.1f}"

    return f"{fmt(big)}X{fmt(small)}"


def _draw_bore_tag(canvas: _Canvas, proj: "Projector", anchor: Vec2,
                   along: Vec2, text: str, gap: float = 3.0) -> None:
    """变径件旁的 '大X小' 就地标签：沿管向倾斜、跨向抬 gap 避开本体线。

    anchor 为标签参考点（变径本体中点、三通中心点或支管端），along 为主管向。
    页面 y 向下，text 旋转角取 atan2(-u.y, u.x)（同 F.W. 标签约定）。
    """
    u = _v2_norm(along)
    if u == (0.0, 0.0):
        return
    if u[0] < 0 or (u[0] == 0.0 and u[1] < 0):
        u = (-u[0], -u[1])   # 翻转方向保持文字正立可读
    across = _iso_cross_dir(proj, u)
    at = _v2_add(anchor, _v2_scale(across, gap))
    canvas.text(at, text, size=5.0,
                angle=math.degrees(math.atan2(-u[1], u[0])))


def _draw_cap(canvas: _Canvas, proj: "Projector", end: Vec2,
              axis_into_pipe: Vec2) -> None:
    # 弧面沿轴测平面横轴（与阀领结/法兰杠/方块端面同族）
    perp = _iso_cross_dir(proj, axis_into_pipe) if proj is not None \
        else _v2_perp(axis_into_pipe)
    r = 1.5
    pts = []
    for step in range(9):
        theta = math.pi * step / 8
        along = -math.sin(theta) * r * 0.9
        across = math.cos(theta) * r
        pts.append(_v2_add(end, _v2_add(_v2_scale(perp, across),
                                        _v2_scale(axis_into_pipe, along))))
    canvas.polyline(pts, width=0.4)


def _draw_box_on_axis(canvas: _Canvas, proj: "Projector", e0: Vec2,
                      e1: Vec2, bore_in: float | None = None,
                      widen: float = 1.0) -> None:
    axis = _v2_norm(_v2_sub(e1, e0))
    # 端面沿轴测平面横轴（与阀领结/法兰杠同族），不用页面垂直
    perp = _iso_cross_dir(proj, axis) if proj is not None else _v2_perp(axis)
    # 按比例图横宽随孔径正比；不按比例图回原始 0.35×自身长度（长度
    # 已是坐标端距，方块读作一段紧凑的盒）
    if bore_in is not None:
        half = _bore_half_w(bore_in) * widen
    else:
        half = max(_v2_len(_v2_sub(e1, e0)) * 0.35, 1.2)
    corners = [
        _v2_add(e0, _v2_scale(perp, half)),
        _v2_add(e1, _v2_scale(perp, half)),
        _v2_add(e1, _v2_scale(perp, -half)),
        _v2_add(e0, _v2_scale(perp, -half)),
    ]
    canvas.polyline(corners, width=0.35, close=True)


# ---------------------------------------------------------------------------
# Dimensions
#
# Dimension lines and extension bars run along projected world-axis directions
# (isometric skew), so callouts lie on an iso plane instead of screen-
# perpendicular. Point pairs that are not axis-aligned get an *offset box*: a
# dashed rectangle (two axes differ) or cuboid (all three) whose axis-aligned
# legs are dimensioned individually, each on the side that collides least.

_AXIS_VECS: tuple[Vec3, ...] = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
_DIM_FONT = _SHEET_FONT
_DIM_CAP = _cap_height_mm(_DIM_FONT)

# Arrowhead: length tip-to-base and half the width across the barbs.
_ARROW_MM = 1.4
_ARROW_HALF_MM = 0.45
# A dimension line must be at least four arrowheads long before both arrows fit
# between its goalposts; under that the arrows go outside, opposing each other.
_MIN_ARROWS_INSIDE_MM = 4 * _ARROW_MM
# Clear space each side of a value that breaks the dimension line.
_DIM_TEXT_GAP_MM = 0.9
# How far a leadered value stands off the thing it points at.
_DIM_LEADER_MM = 4.0
# How far the arrow tails reach past the goalposts in the "outside" style.
_OUTSIDE_TAIL_MM = 3.6

# Whether an offset leg counts as "spanned" is decided on the page, not in the
# model: the leg must draw at least this many mm AND exceed a modelling artifact
# in the world. Model-diagonal thresholds made offset boxes vanish on large
# spools; page-space also accounts for the warp, which stretches those legs.
_MIN_LEG_MM = 1.5
_MIN_LEG_INCH = 0.25


def _axis_screen_dirs(proj: Projector) -> tuple[Vec2, Vec2, Vec2]:
    """Screen-space unit direction of each world axis (+X, +Y, +Z).

    Constant for a given projector once oriented, so the result is cached on
    it -- this sits inside the per-candidate planning loop. set_skew() drops
    the cache, being the one thing that moves the axes after init.
    """
    dirs = getattr(proj, "_axis_dirs_cache", None)
    if dirs is None:
        origin = proj.raw((0.0, 0.0, 0.0))
        dirs = tuple(_v2_norm(_v2_sub(proj.raw(v), origin)) for v in _AXIS_VECS)
        proj._axis_dirs_cache = dirs
    return dirs


def _segs_cross(a: Vec2, b: Vec2, c: Vec2, d: Vec2) -> bool:
    """Intersection test; shared endpoints do not count, but near-collinear
    overlap does (a dimension line lying along a parallel pipe run is a
    collision even though the segments never properly cross)."""
    def shrink(p: Vec2, q: Vec2) -> tuple[Vec2, Vec2]:
        return (_v2_add(p, _v2_scale(_v2_sub(q, p), 0.03)),
                _v2_add(q, _v2_scale(_v2_sub(p, q), 0.03)))

    def orient(p: Vec2, q: Vec2, r: Vec2) -> float:
        return (q[0] - p[0]) * (r[1] - p[1]) - (q[1] - p[1]) * (r[0] - p[0])

    a, b = shrink(a, b)
    c, d = shrink(c, d)
    o1, o2 = orient(a, b, c), orient(a, b, d)
    o3, o4 = orient(c, d, a), orient(c, d, b)
    if o1 * o2 < 0 and o3 * o4 < 0:
        return True
    ab = _v2_sub(b, a)
    ab_len = _v2_len(ab)
    if ab_len < 1e-9:
        return False
    # |orient| / ab_len is the point-to-line distance: both cd endpoints within
    # 0.3mm of ab's line = near-collinear; overlap iff cd's projection hits (0,1).
    if abs(o1) < 0.3 * ab_len and abs(o2) < 0.3 * ab_len:
        t1 = ((c[0] - a[0]) * ab[0] + (c[1] - a[1]) * ab[1]) / (ab_len * ab_len)
        t2 = ((d[0] - a[0]) * ab[0] + (d[1] - a[1]) * ab[1]) / (ab_len * ab_len)
        lo, hi = min(t1, t2), max(t1, t2)
        return hi > 0.0 and lo < 1.0
    return False


def _segs_close(a: Vec2, b: Vec2, c: Vec2, d: Vec2, tol: float = 0.35) -> bool:
    """True when the segments pass within ``tol`` mm of each other.

    ``_segs_cross`` misses the skew-past-an-endpoint case -- a goalpost
    whose tip grazes a nozzle stroke at 0.05mm without either segment
    crossing the other's line. This tests true minimum distance with a
    bounding-box prefilter, so the per-candidate cost stays bounded.
    """
    if (max(a[0], b[0]) + tol < min(c[0], d[0])
            or max(c[0], d[0]) + tol < min(a[0], b[0])
            or max(a[1], b[1]) + tol < min(c[1], d[1])
            or max(c[1], d[1]) + tol < min(a[1], b[1])):
        return False

    def pt_seg(p: Vec2, s: Vec2, e: Vec2) -> float:
        se = _v2_sub(e, s)
        length = _v2_len(se) or 1e-9
        v = _v2_sub(p, s)
        t = max(0.0, min(1.0, (v[0] * se[0] + v[1] * se[1]) / (length * length)))
        return _v2_len(_v2_sub(p, _v2_add(s, _v2_scale(se, t))))

    return (pt_seg(c, a, b) <= tol or pt_seg(d, a, b) <= tol
            or pt_seg(a, c, d) <= tol or pt_seg(b, c, d) <= tol)


def _seg_hits_box(p: Vec2, q: Vec2,
                  box: tuple[float, float, float, float]) -> bool:
    bx, by, bw, bh = box
    if (bx <= p[0] <= bx + bw and by <= p[1] <= by + bh) or \
       (bx <= q[0] <= bx + bw and by <= q[1] <= by + bh):
        return True
    corners = [(bx, by), (bx + bw, by), (bx + bw, by + bh), (bx, by + bh)]
    return any(_segs_cross(p, q, c1, c2)
               for c1, c2 in zip(corners, corners[1:] + corners[:1]))


def _text_aabb(centre: Vec2, w: float, h: float,
               angle_deg: float) -> tuple[float, float, float, float]:
    theta = math.radians(angle_deg)
    ww = abs(w * math.cos(theta)) + abs(h * math.sin(theta))
    hh = abs(w * math.sin(theta)) + abs(h * math.cos(theta))
    return (centre[0] - ww / 2, centre[1] - hh / 2, ww, hh)


def _iso_text_frame(run: Vec2, offset: Vec2) -> tuple[Vec2, Vec2]:
    """Baseline and glyph-rise directions for text on a callout's own plane.

    A dimension's plane is spanned by the run it measures and the direction
    its goalposts are offset along -- both projected world axes -- so text
    belongs on that plane too, sheared onto it rather than merely rotated to
    the dimension line's angle. Rotated text keeps its glyphs square to the
    baseline and reads as a flat label pasted onto an isometric picture.

    Two orientations lie on any plane and only they: ``(along, up)`` and its
    negation -- anything else mirrors the letters. The one whose baseline
    reads left to right wins, and a dimension drawn dead vertical reads
    bottom to top, which is the same rule the trade applies to a plain
    orthographic vertical dimension.
    """
    along = run
    if run[0] < -1e-9 or (abs(run[0]) <= 1e-9 and run[1] > 0):
        along = _v2_scale(run, -1.0)
    up = offset
    # Positive determinant, i.e. cross(up, along) > 0, is what keeps the
    # glyphs the right way round rather than mirrored.
    if up[0] * along[1] - up[1] * along[0] <= 0:
        up = _v2_scale(offset, -1.0)
    return along, up


def _skew_text_box(baseline: Vec2, width: float, cap: float, along: Vec2,
                   up: Vec2) -> tuple[float, float, float, float]:
    """Upright bounding box of text drawn on an isometric plane."""
    corners = [
        _v2_add(baseline, _v2_add(_v2_scale(along, du), _v2_scale(up, dv)))
        for du in (-width / 2, width / 2) for dv in (0.0, cap)
    ]
    xs = [p[0] for p in corners]
    ys = [p[1] for p in corners]
    return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))


def _plan_dim_geometry(canvas: _Canvas, a: Vec2, b: Vec2, direction: Vec2,
                       offset_mm: float, text: str) -> dict:
    """Geometry for one dimension pushed off the measured segment along
    ``direction`` (a projected world-axis unit vector).

    The value is placed in one of the three ways the trade draws it, chosen
    by how much room the dimension has between its goalposts:

    ``inline``
        The text **breaks the dimension line** and sits in the gap, arrows
        pointing outwards at each end. The default whenever the span can
        carry the text and still leave an arrow's worth of line each side.

    ``leader``
        The line is drawn full length with both arrows, and the text stands
        off it with a **leader back to the middle of the double arrow**.
        Used when breaking the line would leave stubs too short to read as
        a dimension.

    ``outside``
        Below four arrowheads of span there is no room for arrows between
        the goalposts at all, so the two arrows go **outside them, opposing
        each other**, and the text is leadered off the base of one. This is
        the standard treatment of a short dimension and it is common on a
        fabrication iso: nipples, weld gaps and reducer offsets all land
        here.

    In every case the text lies on the isometric plane the goalposts define
    (see ``_iso_text_frame``), not flat on the page.
    """
    off = _v2_scale(direction, offset_mm)
    da, db = _v2_add(a, off), _v2_add(b, off)
    gap, over = _v2_scale(direction, 0.8), _v2_scale(direction, 1.2)
    ext = [(_v2_add(a, gap), _v2_add(da, over)),
           (_v2_add(b, gap), _v2_add(db, over))]
    axis = _v2_norm(_v2_sub(db, da))
    span = _v2_len(_v2_sub(db, da))
    mid = _v2_mid(da, db)
    back = _v2_scale(axis, -1.0)
    along, up = _iso_text_frame(axis, direction)
    canvas.pdf.set_font("Helvetica", size=_DIM_FONT)
    tw = canvas.pdf.get_string_width(text)
    stand_off = _v2_scale(direction, _DIM_LEADER_MM + _DIM_CAP / 2)

    leader: Vec2 | None = None
    if span < _MIN_ARROWS_INSIDE_MM:
        style = "outside"
        tail_a = _v2_add(da, _v2_scale(back, _OUTSIDE_TAIL_MM))
        tail_b = _v2_add(db, _v2_scale(axis, _OUTSIDE_TAIL_MM))
        segs = [(tail_a, da), (db, tail_b)]
        if span > 0.25:
            segs.append((da, db))
        arrows = [(da, back), (db, axis)]
        leader = tail_a
        centre = _v2_add(tail_a, stand_off)
    elif span >= tw + 2 * (_ARROW_MM * 1.6 + _DIM_TEXT_GAP_MM):
        style = "inline"
        half = _v2_scale(axis, tw / 2 + _DIM_TEXT_GAP_MM)
        segs = [(da, _v2_sub(mid, half)), (_v2_add(mid, half), db)]
        arrows = [(da, axis), (db, back)]
        centre = mid
    else:
        style = "leader"
        segs = [(da, db)]
        arrows = [(da, axis), (db, back)]
        leader = mid
        centre = _v2_add(mid, stand_off)

    baseline = _v2_sub(centre, _v2_scale(up, _DIM_CAP / 2))
    text_box = _skew_text_box(baseline, tw, _DIM_CAP, along, up)
    return {"da": da, "db": db, "axis": axis, "ext": ext, "style": style,
            "segs": segs, "arrows": arrows, "leader": leader,
            "text_centre": centre, "baseline": baseline, "along": along,
            "up": up, "text": text, "text_box": text_box,
            "direction": direction}


def _dim_leader_edge(geom: dict) -> Vec2 | None:
    """Where a leadered dimension's line leaves its text box."""
    if geom["leader"] is None:
        return None
    _bx, _by, bw, bh = geom["text_box"]
    return _box_edge_toward(geom["text_centre"], geom["leader"],
                            bw + 0.8, bh + 0.8)


def _score_dim_geometry(canvas: _Canvas, geom: dict, away: Vec2) -> float:
    """Collision cost of a candidate placement (lower is better)."""
    score = 0.0
    edge = _dim_leader_edge(geom)
    for seg in canvas.segments:
        for p, q in geom["segs"]:
            if _segs_cross(p, q, seg[0], seg[1]):
                score += 3.0
        if _seg_hits_box(seg[0], seg[1], geom["text_box"]):
            score += 2.0
        for e in geom["ext"]:
            if (_segs_cross(e[0], e[1], seg[0], seg[1])
                    or _segs_close(e[0], e[1], seg[0], seg[1])):
                score += 1.0
        if edge is not None and _segs_cross(edge, geom["leader"], seg[0], seg[1]):
            score += 1.0
    for box in canvas.label_boxes:
        if _boxes_overlap(geom["text_box"], box, 0.3):
            score += 5.0
    if canvas.bounds is not None:
        bx, by, bw, bh = canvas.bounds
        tx, ty, tw, th = geom["text_box"]
        if tx < bx or ty < by or tx + tw > bx + bw or ty + th > by + bh:
            score += 10.0
    # Everything else equal, push towards the open side of the sheet.
    score -= 0.5 * (geom["direction"][0] * away[0] + geom["direction"][1] * away[1])
    return score


def _plan_best_dim(canvas: _Canvas, proj: Projector, p3a: Vec3, p3b: Vec3,
                   text: str, offset_mm: float, span_axes: list[int],
                   draw_centre: Vec2, *, coaxial: bool = False) -> tuple[dict | None, float]:
    """Best-scoring placement for a dimension between two 3D points.

    Offset candidates are the world axes NOT spanned by the measured
    segment (both signs), so the leader bars and dimension line stay on an
    isometric plane. Falls back to the screen perpendicular only for fully
    skew segments (all three axes spanned, e.g. an overridden diagonal).
    """
    a, b = proj.page(p3a), proj.page(p3b)
    if coaxial:
        # 偏心变径的大小端在平行中心线上，符号布局给小端侧额外横向平移；
        # 尺寸不能直接连两个实际图面端点（会沿偏心斜弦画）——以大端为锚，
        # 小端投影回原始管轴：去横向位移、保留轴向位移。
        raw_a = proj.page(p3a, apply_offset=False)
        raw_b = proj.page(p3b, apply_offset=False)
        run = _v2_sub(raw_b, raw_a)
        run_len = _v2_len(run)
        if run_len > 1e-9:
            unit = _v2_scale(run, 1.0 / run_len)
            ab = _v2_sub(b, a)
            axial = ab[0] * unit[0] + ab[1] * unit[1]
            b = _v2_add(a, _v2_scale(unit, axial))
    screen_axis = _v2_norm(_v2_sub(b, a))
    if screen_axis == (0.0, 0.0):
        return None, 0.0
    dirs = _axis_screen_dirs(proj)
    cands = [_v2_scale(dirs[i], sign)
             for i in range(3) if i not in span_axes
             for sign in (1.0, -1.0)]
    cands = [d for d in cands
             if abs(d[0] * screen_axis[0] + d[1] * screen_axis[1]) < 0.95]
    if not cands:
        perp = _v2_perp(screen_axis)
        cands = [perp, _v2_scale(perp, -1.0)]
    away = _v2_norm(_v2_sub(_v2_mid(a, b), draw_centre))
    best, best_score = None, 0.0
    for direction in cands:
        geom = _plan_dim_geometry(canvas, a, b, direction, offset_mm, text)
        score = _score_dim_geometry(canvas, geom, away)
        if best is None or score < best_score:
            best, best_score = geom, score
    return best, best_score


def _goalpost_pieces(a: Vec2, b: Vec2,
                     segments: list[tuple[Vec2, Vec2]]) -> list[tuple[Vec2, Vec2]]:
    """Trim a goalpost against near-parallel strokes it would duplicate.

    An equipment nozzle's face bar is centred on the gasket face -- exactly
    where a "to equipment face" dimension's endpoint sits -- and often runs
    the same direction as that end's goalpost. Drawing both puts two lines
    0.05mm apart on the sheet. The bar already serves as the extension
    line, so the goalpost draws only the part the bar does not cover.
    Perpendicular structures (pipes, elbow legs) are never parallel enough
    to trim, so ordinary goalposts pass through untouched.
    """
    ab = _v2_sub(b, a)
    length = _v2_len(ab)
    if length <= 1e-6:
        return [(a, b)]
    u = (ab[0] / length, ab[1] / length)
    cuts: list[tuple[float, float]] = []
    for s, e in segments:
        se = _v2_sub(e, s)
        sl = _v2_len(se)
        if sl <= 1e-6:
            continue
        v = (se[0] / sl, se[1] / sl)
        if abs(u[0] * v[0] + u[1] * v[1]) < 0.98:
            continue                      # not near-parallel
        if not _segs_close(a, b, s, e, 0.3):
            continue                      # not alongside it
        t1 = (s[0] - a[0]) * u[0] + (s[1] - a[1]) * u[1]
        t2 = (e[0] - a[0]) * u[0] + (e[1] - a[1]) * u[1]
        lo, hi = min(t1, t2), max(t1, t2)
        if hi > 0.1 and lo < length - 0.1:
            cuts.append((max(lo, 0.0), min(hi, length)))
    if not cuts:
        return [(a, b)]
    cuts.sort()
    pieces: list[tuple[Vec2, Vec2]] = []
    start = 0.0
    for lo, hi in cuts:
        if lo > start + 0.1:
            pieces.append((_v2_add(a, _v2_scale(u, start)),
                           _v2_add(a, _v2_scale(u, lo))))
        start = max(start, hi)
    if start < length - 0.1:
        pieces.append((_v2_add(a, _v2_scale(u, start)), b))
    return pieces


def _render_dim(canvas: _Canvas, geom: dict) -> None:
    for e in geom["ext"]:
        for p, q in _goalpost_pieces(e[0], e[1], canvas.segments):
            canvas.line(p, q, width=0.15, record=False)
    for p, q in geom["segs"]:
        canvas.line(p, q, width=0.2)  # recorded: later dims avoid it
    for tip, direction in geom["arrows"]:
        _draw_arrowhead(canvas, tip, direction)
    edge = _dim_leader_edge(geom)
    if edge is not None and _v2_len(_v2_sub(geom["leader"], edge)) > 0.6:
        canvas.line(edge, geom["leader"], width=0.15, record=False)
    canvas.iso_text(geom["baseline"], geom["text"], _DIM_FONT,
                    geom["along"], geom["up"])
    canvas.label_boxes.append(geom["text_box"])


def _span_axes(proj: Projector, p_from: Vec3, delta: Vec3,
               inch_factor: float) -> list[int]:
    """World axes this dimension genuinely travels along.

    An axis counts when its leg is both a real distance in the model (past
    the modelled-root-gap floor) and long enough on the page to be drawn and
    labelled. The page test uses the projector as-is, so the not-to-scale
    warp is included -- it stretches short legs to at least nts_min_mm, which
    is exactly when an offset box becomes worth drawing.
    """
    base = proj.page(p_from)
    axes: list[int] = []
    for i in range(3):
        if abs(delta[i]) * inch_factor < _MIN_LEG_INCH:
            continue
        leg_end = list(p_from)
        leg_end[i] += delta[i]
        if _v2_len(_v2_sub(proj.page(tuple(leg_end)), base)) >= _MIN_LEG_MM:
            axes.append(i)
    return axes


_OLET_HALF_MM = 1.9     # along the header: half the fitting's footprint
_OLET_RISE_MM = 1.7     # along the branch: how far it stands off the header
_OLET_BASE_MM = 0.9     # clearance from the header centreline (the weld dot)


def _draw_olet(canvas: _Canvas, proj: Projector, centre: Vec2, branch: Vec2,
               header: Vec2 | None, branch_connected: bool = True) -> Vec2:
    """A weldolet/sockolet/thredolet: a boss sitting on the header.

    Drawn as a rectangle **tangent to the pipeline** -- long edges parallel
    to the header run, standing off it towards the branch -- which on an
    isometric comes out as a skewed parallelogram. The branch centreline
    starts at the fitting's outer face, not at the header centreline, so the
    olet reads as a fitting welded onto the pipe rather than as a bare stub.

    ``branch_connected`` says whether anything is actually piped off the
    outlet (see ``_olet_branch_connected``). When nothing is, the boss is
    drawn on its own and the centreline is suppressed: the branch point still
    orients the fitting, but a centreline reaching out to it would promise a
    run that this spool does not have.

    Previously this was a 0.55 mm stub line plus a 0.8 mm open circle at the
    root -- and the weld pass stamps a 0.9 mm *filled* dot at that same
    point, so the circle was completely buried and the sheet showed nothing
    but a centreline leaving a dot.

    ``header`` is the page-space vector of the run the olet sits on. Without
    it (nothing found under the centre) we fall back to the iso axis most
    square to the branch, which is the header direction whenever the branch
    is perpendicular to the run -- the normal case.
    """
    out = _v2_norm(_v2_sub(branch, centre))
    if out == (0.0, 0.0):
        canvas.circle(centre, 0.8, width=0.3)
        return centre
    along = _v2_norm(header) if header else (0.0, 0.0)
    if along == (0.0, 0.0):
        along = _iso_cross_dir(proj, out)

    # Stand the boss clear of the weld dot at its root: the weld pass stamps a
    # 0.9 mm filled dot there, and a box sharing that edge is half swallowed.
    base = _v2_add(centre, _v2_scale(out, _OLET_BASE_MM))
    half = _v2_scale(along, _OLET_HALF_MM)
    rise = _v2_scale(out, _OLET_RISE_MM)
    root_a, root_b = _v2_add(base, half), _v2_sub(base, half)
    top_a, top_b = _v2_add(root_a, rise), _v2_add(root_b, rise)
    canvas.polyline([root_a, top_a, top_b, root_b], width=0.45, close=True)
    # Branch leaves the outer face of the fitting -- but only if there is a
    # branch to leave for.
    face = _v2_add(base, rise)
    if branch_connected:
        canvas.line(face, branch, width=0.55)
    return face


_SW_SPINE_MM = 2.6      # across the run: the socket shoulder
_SW_ARM_MM = 1.3        # along the run: the socket depth


def _draw_weld_symbol(canvas: _Canvas, proj: Projector, weld: Weld,
                      dot: Vec2) -> None:
    """The symbol for one weld, chosen by joint type.

    A butt weld is a filled dot (the whole joint is the circumference); a
    **socket** weld is a square bracket opening towards the inserted member.
    Drawn with ``record=False``: weld symbols sit on the run and must not
    perturb dimension placement.
    """
    if weld.joint != "SW" or weld.insert_toward is None:
        canvas.dot(dot, r=0.9)
        return

    # Along-run direction, measured on the page so the NTS warp and skew are
    # already baked in, pointing at the member that goes into the socket.
    along = _v2_norm(_v2_sub(proj.page(weld.insert_toward), dot))
    if along == (0.0, 0.0):
        canvas.dot(dot, r=0.9)
        return
    across = _iso_cross_dir(proj, along)
    half = _SW_SPINE_MM / 2
    spine_a = _v2_add(dot, _v2_scale(across, half))
    spine_b = _v2_add(dot, _v2_scale(across, -half))
    arm = _v2_scale(along, _SW_ARM_MM)
    canvas.polyline([_v2_add(spine_a, arm), spine_a, spine_b,
                     _v2_add(spine_b, arm)], width=0.5, record=False)


# --- 传统 Isogen 管架符号 ---------------------------------------------------
# SKEY 语义来自 Smart Isogen SystemMaterialData.xml 的 Supports 段：
#   01HG/空=承重支撑, ANCH=固定架, GUID=导向架, SPRG=弹簧支吊架, HANG=吊架
#   （DUCK/SKID/SLVE 等暂按承重画法）。.ASC 方向模式（501 尾参）：
#   0=顺管道方向, 2=固定方向（SUPPORT-DIRECTION UP/DOWN 翻转）,
#   3=三轴六向（ANCH 双箭头，方向由 SUPPORT-DIRECTION 的世界轴决定）。
_SUPPORT_WORLD_DIRS = {
    "UP": (0, 0, 1), "DOWN": (0, 0, -1),
    "NORTH": (0, 1, 0), "N": (0, 1, 0), "Y+": (0, 1, 0), "Y": (0, 1, 0),
    "SOUTH": (0, -1, 0), "S": (0, -1, 0), "Y-": (0, -1, 0),
    "EAST": (1, 0, 0), "E": (1, 0, 0), "X+": (1, 0, 0), "X": (1, 0, 0),
    "WEST": (-1, 0, 0), "W": (-1, 0, 0), "X-": (-1, 0, 0),
    "Z+": (0, 0, 1), "Z": (0, 0, 1), "Z-": (0, 0, -1),
}
_PIPE_RUNS_CACHE: dict[int, list] = {}


def _nearest_pipe_run_dir(proj: "Projector", doc, at: Vec2,
                          world_at: Vec3 | None = None):
    """支架点附近管道走向；返回 (页空间方向, 世界空间方向)，找不到返回 None。

    世界方向用于推导跨管的世界轴（等轴测里跨管偏移沿世界轴的 30° 投影）。
    world_at 给定时，世界距离≈0（支架真立在管线上）的管段优先，其间再按
    页面距离定夺——支架常立在弯头端口上，页面最近的可能只是邻跨接续短管。
    候选集含弯头/三通的腿（centre→端点），端口上的支架才能命中真正管轴。
    """
    runs = _PIPE_RUNS_CACHE.get(id(proj))
    if runs is None:
        runs = []
        for comp in doc.real_components():
            cls = classify(comp.type, comp.skey)
            pairs: list[tuple[Vec3, Vec3, bool]] = []
            if cls == "PIPE" and len(comp.end_points) >= 2:
                pairs.append((comp.end_points[0].xyz,
                              comp.end_points[1].xyz, False))
            elif (cls in ("ELBOW", "TEE") and comp.centre_points
                    and comp.end_points):
                cw = comp.centre_points[0].xyz
                pairs = [(cw, ep.xyz, True) for ep in comp.end_points]
            for w0, w1, leg in pairs:
                a = proj.page(w0)
                b = proj.page(w1)
                d = _v2_sub(b, a)
                if _v2_len(d) > 1e-6:
                    runs.append((a, d,
                                 (w1[0]-w0[0], w1[1]-w0[1], w1[2]-w0[2]),
                                 w0, w1, leg))
        _PIPE_RUNS_CACHE.clear()
        _PIPE_RUNS_CACHE[id(proj)] = runs
    best, best_key = None, (1e18, 1e18)
    px, py = at
    for a, d, dw, w0, w1, leg in runs:
        if world_at is None and leg:
            continue           # 页面最近回退保持旧候选集（仅 PIPE）
        ax, ay = a
        dx, dy = d
        t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        dist = math.hypot(px - ax - t * dx, py - ay - t * dy)
        if world_at is not None:
            sd = (w1[0]-w0[0], w1[1]-w0[1], w1[2]-w0[2])
            denom = sd[0]*sd[0] + sd[1]*sd[1] + sd[2]*sd[2]
            if denom > 1e-12:
                tw = ((world_at[0]-w0[0]) * sd[0]
                      + (world_at[1]-w0[1]) * sd[1]
                      + (world_at[2]-w0[2]) * sd[2]) / denom
                tw = max(0.0, min(1.0, tw))
                wdist = math.sqrt(
                    (world_at[0]-w0[0]-tw*sd[0]) ** 2
                    + (world_at[1]-w0[1]-tw*sd[1]) ** 2
                    + (world_at[2]-w0[2]-tw*sd[2]) ** 2)
            else:
                wdist = 1e18
            if wdist > 1.0:      # 1mm 内才算"立在管上"，否则维持页面最近
                continue
            key = (wdist, dist)
        else:
            key = (dist, 0.0)
        if key < best_key:
            best_key, best = key, (d, dw)
    if best is None and world_at is not None:
        return _nearest_pipe_run_dir(proj, doc, at)
    return best


def _draw_support_symbol(canvas: _Canvas, at: Vec2, comp,
                         proj=None, doc=None) -> None:
    skey = (comp.skey or "").upper()
    text = " ".join([comp.description or "",
                     comp.get("SUPPORT-TYPE") or ""]).upper()
    direction = (comp.get("SUPPORT-DIRECTION") or "DOWN").upper()
    s = -1.0 if direction == "UP" else 1.0     # 图面 y 向下；UP 时符号画到上方
    x, y = at
    w = 0.3

    # 优先使用 Isogen Symbol Editor 导出的 .ASC 符号库（symbols 目录）
    try:
        from .asc_symbols import find_symbol
        found = find_symbol(skey)
    except Exception:
        found = None
    if found:
        sym, orient = found
        xs = [c for seg in sym for c in (seg[0], seg[2])]
        ys = [c for seg in sym for c in (seg[1], seg[3])]
        extent = max(max(xs) - min(xs), max(ys) - min(ys), 1e-6)
        scale = 4.5 / extent                    # 归一化到约 4.5mm

        def draw_frame(u, v, sf):
            """按局部坐标系 (u=符号X轴, sf*v=符号Y轴) 画解码出的 ASC 几何"""
            for x1, y1, x2, y2 in sym:
                canvas.line(
                    (x + (x1 * u[0] + sf * y1 * v[0]) * scale,
                     y + (x1 * u[1] + sf * y1 * v[1]) * scale),
                    (x + (x2 * u[0] + sf * y2 * v[0]) * scale,
                     y + (x2 * u[1] + sf * y2 * v[1]) * scale),
                    width=w)

        if orient == 3 and proj is not None and comp.co_ords:
            # 三轴同画（ANCH）：X/Y/Z 三个世界轴向各画一次 ASC 几何
            wp = comp.co_ords[0].xyz
            for d3 in ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0)):
                a = proj.page(wp)
                b = proj.page((wp[0] + d3[0], wp[1] + d3[1], wp[2] + d3[2]))
                ln = math.hypot(b[0] - a[0], b[1] - a[1])
                if ln <= 1e-9:
                    continue
                u = ((b[0] - a[0]) / ln, (b[1] - a[1]) / ln)
                draw_frame(u, (-u[1], u[0]), 1.0)
            return

        # 符号局部坐标系：X 轴 = u（符号杆/线方向），Y 轴 = s*v（侧向）
        u = (1.0, 0.0)
        v = (0.0, 1.0)
        if orient == 0 and proj is not None and doc is not None:
            # 顺管道：符号 X 轴对齐管段；跨管方向取"垂直于管轴的世界轴"的投影
            # （等轴测中为 30° 斜向，立管上两条线因此有高度差）
            world_at = comp.co_ords[0].xyz if comp.co_ords else None
            run = _nearest_pipe_run_dir(proj, doc, at, world_at)
            if run is not None:
                d, dw = run
                ln = _v2_len(d)
                u = (d[0] / ln, d[1] / ln)
                across = (1.0, 0.0, 0.0)
                if dw is not None:
                    axis = max(range(3), key=lambda i: abs(dw[i]))
                    # X 向管跨 Y；Y/Z 向管跨 X（跨管轴一律取水平轴）
                    across = (0.0, 1.0, 0.0) if axis == 0 else (1.0, 0.0, 0.0)
                if comp.co_ords:
                    wp = comp.co_ords[0].xyz
                    a2 = proj.page(wp, apply_offset=False)
                    b2 = proj.page((wp[0] + across[0], wp[1] + across[1],
                                    wp[2] + across[2]), apply_offset=False)
                    vn = math.hypot(b2[0] - a2[0], b2[1] - a2[1])
                    if vn > 1e-9:
                        v = ((b2[0] - a2[0]) / vn, (b2[1] - a2[1]) / vn)
        draw_frame(u, v, s)
        return

    stem_y = 2.2

    def stem_and_base():
        canvas.line((x, y + s * 0.4), (x, y + s * stem_y), width=w)
        canvas.line((x - 1.8, y + s * stem_y), (x + 1.8, y + s * stem_y), width=w)

    if "SPRG" in skey or "SPRING" in text or "弹簧" in text:
        # 弹簧支吊架：锯齿弹簧 + 端板
        pts = [(x, y + s * 0.4)]
        for i in range(4):
            pts.append((x + (0.9 if i % 2 == 0 else -0.9), y + s * (0.9 + 0.5 * i)))
        pts.append((x, y + s * 3.1))
        canvas.polyline(pts, width=w)
        canvas.line((x - 1.6, y + s * 3.5), (x + 1.6, y + s * 3.5), width=w)
    elif "GUID" in skey or "GUIDE" in text or "STEADY" in text or "导向" in text:
        # 导向架：立柱 + 底座 + 两侧导向板
        stem_and_base()
        canvas.line((x - 1.8, y + s * stem_y), (x - 1.8, y + s * 0.8), width=w)
        canvas.line((x + 1.8, y + s * stem_y), (x + 1.8, y + s * 0.8), width=w)
    elif "ANCH" in skey or "ANCHOR" in text or "固定" in text:
        # 固定架：立柱 + 底座 + 交叉斜撑
        stem_and_base()
        canvas.line((x - 1.0, y + s * 0.8), (x + 1.0, y + s * 2.1), width=w)
        canvas.line((x + 1.0, y + s * 0.8), (x - 1.0, y + s * 2.1), width=w)
    elif "HANG" in skey or "HANGER" in text or "吊" in text:
        # 吊架：吊杆 + 顶板
        canvas.line((x, y + s * 0.4), (x, y + s * 2.6), width=w)
        canvas.line((x - 1.6, y + s * 2.6), (x + 1.6, y + s * 2.6), width=w)
    else:
        # 默认承重支撑（01HG）：立柱 + 底座 + 斜线阴影
        stem_and_base()
        for dx in (-1.3, 0.0, 1.3):
            canvas.line((x + dx, y + s * stem_y),
                        (x + dx - 0.7, y + s * (stem_y + 0.8)), width=w)


def _draw_dimension(canvas: _Canvas, proj: Projector, dim: Dimension,
                    units_coords: str, draw_centre: Vec2) -> bool:
    """Draw one dimension. False when it was inferred but could not be drawn.

    A dropped dimension used to be a bare `return`. It still went into
    `RenderResult.dimensions` and still got written to the sidecar with a
    real value, so the only evidence that a number was missing from the sheet
    was noticing the gap by eye. The caller now turns a False into a warning.
    """
    delta = tuple(t - f for f, t in zip(dim.p_from, dim.p_to))
    span_axes = _span_axes(proj, dim.p_from, delta,
                           coord_to_inch_factor(units_coords))
    if not span_axes:
        return False  # nothing measurable at this scale
    # Magnitude only: the side is chosen by collision scoring, and an
    # explicit sidecar offset_mm of 0 (flush on the run) is honoured.
    offset_mm = abs(dim.offset_mm)
    if len(span_axes) >= 2 and not dim.text:
        _draw_offset_box(canvas, proj, dim, delta, span_axes,
                         units_coords, draw_centre, offset_mm)
        return True
    # Axis-aligned (or overridden-text) segment: one dimension.
    geom, _ = _plan_best_dim(canvas, proj, dim.p_from, dim.p_to,
                             dim.computed_text(units_coords),
                             offset_mm, span_axes, draw_centre,
                             coaxial=dim.coaxial)
    if geom is None:
        return False   # both ends project to the same page point
    _render_dim(canvas, geom)
    return True


def _offset_box_corners(p_from: Vec3, delta: Vec3,
                        axes: list[int]) -> dict[frozenset, Vec3]:
    """Corners of the offset rectangle/cuboid, keyed by the set of axes
    already applied: 4 corners for two axes, 8 for three."""
    corners: dict[frozenset, Vec3] = {}
    for bits in range(1 << len(axes)):
        applied = frozenset(ax for n, ax in enumerate(axes) if bits >> n & 1)
        pt = list(p_from)
        for ax in applied:
            pt[ax] += delta[ax]
        corners[applied] = tuple(pt)
    return corners


def _offset_box_edges(corners: dict[frozenset, Vec3]) -> list[tuple[Vec3, Vec3]]:
    """Box edges connect corners whose applied-axis sets differ by one axis."""
    keys = list(corners)
    return [(corners[ka], corners[kb])
            for i, ka in enumerate(keys) for kb in keys[i + 1:]
            if len(ka ^ kb) == 1]


def _gauss_solve(m: list[list[float]], v: list[float]) -> list[float]:
    n = len(v)
    for col in range(n):
        piv = max(range(col, n), key=lambda r: abs(m[r][col]))
        m[col], m[piv] = m[piv], m[col]
        v[col], v[piv] = v[piv], v[col]
        if abs(m[col][col]) < 1e-12:
            continue
        for r in range(col + 1, n):
            f = m[r][col] / m[col][col]
            for c in range(col, n):
                m[r][c] -= f * m[col][c]
            v[r] -= f * v[col]
    x = [0.0] * n
    for r in range(n - 1, -1, -1):
        s = v[r] - sum(m[r][c] * x[c] for c in range(r + 1, n))
        x[r] = s / m[r][r] if abs(m[r][r]) > 1e-12 else 0.0
    return x


def _register_offset_box_pages(proj: Projector, p_from: Vec3, delta: Vec3,
                               axes: list[int], corners: dict) -> dict:
    """offset 盒统一角点映射，注册进 point_offsets 并返回旧值供恢复。

    派生角点各自走 offset_nearest 回退时最近点/σ 基准不一，虚线矩形会
    变梯形（一条边不沿轴测方向）。改为以实际两端页面位置 A/B 为锚，
    把 B-A 按各轴 raw 投影向量最小二乘分解（2 轴不共线即精确解），
    角点 = A + 已施轴分量之和——盒边严格沿轴测方向且两端贴实际点。
    按比例/无变形时分解精确（系数全 1），注册 offset 为 0，行为不变。
    """
    p_to = tuple(p_from[k] + delta[k] for k in range(3))
    a, b = proj.page(p_from), proj.page(p_to)
    r0 = proj.raw(p_from)
    basis = []
    for ax in axes:
        p1 = list(p_from)
        p1[ax] += delta[ax]
        r1 = proj.raw(tuple(p1))
        basis.append(((r1[0] - r0[0]) * proj.scale,
                      (r1[1] - r0[1]) * proj.scale))
    d = _v2_sub(b, a)
    n = len(axes)
    ata = [[sum(basis[i][k] * basis[j][k] for k in range(2))
            for j in range(n)] for i in range(n)]
    atd = [sum(basis[i][k] * d[k] for k in range(2)) for i in range(n)]
    for i in range(n):
        ata[i][i] += 1e-9
    coeff = _gauss_solve(ata, atd)
    saved = {}
    for key, w in corners.items():
        target = a
        for i, ax in enumerate(axes):
            if ax in key:
                target = _v2_add(target, _v2_scale(basis[i], coeff[i]))
        rp = proj.raw(w)
        raw_page = (rp[0] * proj.scale + proj.offset[0],
                    rp[1] * proj.scale + proj.offset[1])
        saved[w] = proj.point_offsets.get(w)
        proj.point_offsets[w] = (target[0] - raw_page[0],
                                 target[1] - raw_page[1])
    return saved


def _draw_offset_box(canvas: _Canvas, proj: Projector, dim: Dimension,
                     delta: Vec3, axes: list[int], units_coords: str,
                     draw_centre: Vec2, offset_mm: float) -> None:
    """Skew dimension: dashed offset rectangle (2 axes) or cuboid (3 axes),
    dimensioning each axis-aligned leg of the least-collision corner path.

    The box is the whole point of the callout -- it is what tells a fitter
    the run rolls -- so it is drawn before anything can fail, and every leg
    of the winning corner path is dimensioned even if one of them ends up
    crowded. ``axes`` has already been vetted for page-visible length, so
    no edge here is a hairline.
    """
    corners = _offset_box_corners(dim.p_from, delta, axes)
    saved = _register_offset_box_pages(proj, dim.p_from, delta, axes, corners)
    try:
        for e3a, e3b in _offset_box_edges(corners):
            a, b = proj.page(e3a), proj.page(e3b)
            if _v2_len(_v2_sub(b, a)) > 0.4:
                canvas.dashed(a, b, width=0.22, dash=1.4, record=True)

        def leg_text(ax: int) -> str:
            return format_length(abs(delta[ax]), units_coords)

        best_order: tuple[int, ...] | None = None
        best_cost = 0.0
        for order in permutations(axes):
            cost, applied, prev = 0.0, frozenset(), dim.p_from
            for ax in order:
                nxt = corners[applied | {ax}]
                _, leg_score = _plan_best_dim(canvas, proj, prev, nxt, leg_text(ax),
                                              offset_mm, [ax], draw_centre)
                cost += leg_score
                applied, prev = applied | {ax}, nxt
            if best_order is None or cost < best_cost:
                best_order, best_cost = order, cost

        # Draw the winning path leg by leg, re-picking each leg's side so a leg
        # also avoids the legs already drawn before it.
        applied, prev = frozenset(), dim.p_from
        for ax in best_order or ():
            nxt = corners[applied | {ax}]
            geom, _ = _plan_best_dim(canvas, proj, prev, nxt, leg_text(ax),
                                     offset_mm, [ax], draw_centre)
            if geom is not None:
                _render_dim(canvas, geom)
            applied, prev = applied | {ax}, nxt
    finally:
        for w, old in saved.items():
            if old is None:
                proj.point_offsets.pop(w, None)
            else:
                proj.point_offsets[w] = old


def _draw_arrowhead(canvas: _Canvas, tip: Vec2, direction: Vec2) -> None:
    """An arrowhead at ``tip`` whose barbs run back along ``direction``."""
    perp = _v2_perp(direction)
    base = _v2_add(tip, _v2_scale(direction, _ARROW_MM))
    canvas.polyline([
        _v2_add(base, _v2_scale(perp, _ARROW_HALF_MM)), tip,
        _v2_add(base, _v2_scale(perp, -_ARROW_HALF_MM)),
    ], width=0.2, record=False)


# ---------------------------------------------------------------------------
# BOM / tables / furniture


def _bom_class(comp: Component) -> str:
    """Material class for accounting: the record type, except that a flange
    shipped as a COUPLING/OTHER record is counted as the flange it is."""
    if looks_like_flange(comp.type, comp.description, comp.skey):
        return "FLANGE"
    return classify(comp.type, comp.skey)


def _bom_key(comp: Component) -> tuple:
    return (_bom_class(comp), comp.item_code or "?",
            (comp.description or "").strip().lower())


def _bom_row_key(comp: Component, materials: dict[str, str]) -> tuple | None:
    """该组件是否进 BOM；进则返回行 key，不进返回 None。

    与 _build_bom 同一过滤（单一来源），分页装箱也用它估每页行数。
    """
    cls = _bom_class(comp)
    if cls in ("CONNECTION", "MESSAGE", "GAP"):
        return None
    if comp.material_excluded():
        return None  # 双写边界段的对侧副本（CONTINUATION 件及其副料）：画但不计料
    if is_weld_coupling(comp.type, comp.description, comp.skey):
        return None  # welds are counted in the weld table, not the BOM
    if cls == "WELD" and not comp.item_code:
        return None  # 焊缝记录无料号时不进 BOM（避免图面/材料表出现 "?" 气球）
    _cu = comp.type.upper().replace("_", "-")
    if _cu.startswith("FLOW-ARROW") or (
            (comp.skey or "").upper() == "FLOW" and cls == "OTHER"):
        return None  # 流向箭头是图面符号（SKEY FLOW 曾被 FL 前缀误判为法兰），不进 BOM
    # Resolved the same way the row's own description is, below: plenty of
    # exports name the part only in the MATERIALS table.
    desc = comp.description or materials.get(comp.item_code or "", "")
    if is_annotation_symbol(comp.type, desc, comp.skey):
        return None  # a slope arrow is a drawing mark, not material
    return _bom_key(comp)


def _build_bom(doc: PcfDocument, comps: list[tuple[Component, str]]) -> dict:
    inch_factor = coord_to_inch_factor(doc.units_coords)
    materials = doc.materials_descriptions()
    rows: dict[tuple, dict] = {}
    order = 0
    for idx, (comp, _ref) in enumerate(comps):
        cls = _bom_class(comp)
        if _bom_row_key(comp, materials) is None:
            continue
        key = _bom_key(comp)
        row = rows.get(key)
        if row is None:
            bore = comp.bore
            nps = format_bore(bore, doc.units_bore) if bore else ""
            description = comp.description or materials.get(comp.item_code or "", "") \
                or comp.type
            row = rows[key] = {
                # 气球/BOM 项号固定用行顺序号（Isogen 惯例），料号存入 code 列
                "order": order, "item": str(order + 1), "code": comp.item_code or "",
                "description": description, "nps": nps, "class": cls,
                "qty": 0, "length_in": 0.0, "weight": 0.0,
                "_callout_ids": set(), "_ucis": set(),
            }
            order += 1
        row["qty"] += 1
        row["_callout_ids"].add(idx)
        if comp.uci:
            row["_ucis"].add(comp.uci)
        if comp.weight:
            row["weight"] += comp.weight
        if cls == "PIPE":
            eps = comp.end_points
            if len(eps) >= 2:
                row["length_in"] += eps[0].distance_to(eps[1]) * inch_factor
    return rows


_CALLOUT_FONT = 6.0
_CALLOUT_CAP = _cap_height_mm(_CALLOUT_FONT)
_CALLOUT_PAD_MM = 0.6        # very tight padding around the number
_CALLOUT_LEADER_MM = 6.0     # shortest leader the placer will accept
_CALLOUT_CLEARANCE_MM = 1.0  # whitespace a box keeps from drawn geometry


# Whitespace a callout's *leader foot* keeps from a neighbouring fitting's anchor
# or a weld dot, in page mm -- inside it the leader reads as pointing at the neighbour.
_ANCHOR_CLEARANCE_MM = 3.0

# Where along a pipe its cut-piece callout may point, as fractions of the run:
# midpoint first, quarter points as the escape route when a fitting is welded
# on at the middle.
_PIPE_ANCHOR_FRACTIONS = (0.5, 0.3, 0.7)


def _pipe_anchor_options(comp: Component, proj: Projector) -> list[Vec2]:
    """Points along a pipe its callout leader may land on, best first.

    A pipe's anchor was hard-wired to the midpoint of its projected body, and
    a midpoint is exactly where a branch tends to be: an olet welded into the
    middle of a run put the cut-piece leader on the olet symbol and its weld
    dot every time, with the balloon then squeezed between the olet's own
    number and the spool label. The box could move; the foot could not.
    """
    eps = comp.end_points
    if len(eps) < 2:
        return []
    a, b = proj.page(eps[0].xyz), proj.page(eps[-1].xyz)
    span = _v2_sub(b, a)
    return [_v2_add(a, _v2_scale(span, t)) for t in _PIPE_ANCHOR_FRACTIONS]


def _crowding_cost(point: Vec2, occupied: list[Vec2]) -> float:
    """What it costs to land a leader foot at ``point``.

    Priced per neighbour rather than as a yes/no, so a foot with one fitting
    nearby beats one with three, and a foot that merely comes close beats one
    that lands on top.
    """
    cost = 0.0
    for other in occupied:
        gap = _v2_len(_v2_sub(point, other))
        if gap < _ANCHOR_CLEARANCE_MM:
            cost += 4.0 * (1.0 - gap / _ANCHOR_CLEARANCE_MM)
    return cost


def _draw_callout_box(canvas: _Canvas, anchor: Vec2, label: str,
                      draw_centre: Vec2,
                      options: list[Vec2] | None = None,
                      occupied: list[Vec2] | None = None) -> tuple[Vec2, Vec2]:
    """One square or rectangular callout box with a leader to the run.

    The single style for **everything the shop has to identify** -- a fitting's
    item number and a cut piece's label alike: same kind of mark, one convention.
    Every callout stands clear of the centreline and points at its run -- bare
    labels force proximity-inference, which fails on busy sheets or wherever
    the warp brings two runs close.
    """
    canvas.pdf.set_font("Helvetica", size=_CALLOUT_FONT)
    tw = canvas.pdf.get_string_width(label)
    box_h = _CALLOUT_CAP + 2 * _CALLOUT_PAD_MM
    box_w = max(tw, _CALLOUT_CAP) + 2 * _CALLOUT_PAD_MM
    if options:
        anchor, pos = canvas.place_label_multi(
            options, box_w, box_h, draw_centre,
            min_leader=_CALLOUT_LEADER_MM, avoid_geometry=True,
            clearance=_CALLOUT_CLEARANCE_MM,
            anchor_cost=lambda p: _crowding_cost(p, occupied or []))
    else:
        pos = canvas.place_label(anchor, box_w, box_h, draw_centre,
                                 min_leader=_CALLOUT_LEADER_MM,
                                 avoid_geometry=True,
                                 clearance=_CALLOUT_CLEARANCE_MM)
    # Leader to the box *edge*, so the run of it is visible rather than
    # buried under the knockout, then the box over the top of it.
    edge = _box_edge_toward(pos, anchor, box_w, box_h)
    canvas.line(anchor, edge, width=0.15)
    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.3)
        canvas.pdf.rect(pos[0] - box_w / 2, pos[1] - box_h / 2, box_w, box_h,
                        style="DF")
    canvas.text((pos[0], pos[1] + _CALLOUT_CAP / 2), label, size=_CALLOUT_FONT)
    return anchor, pos


def _draw_callouts(canvas: _Canvas, proj: Projector, comps, anchors, bom_rows,
                   pipe_labels: dict[int, str], draw_centre,
                   every_instance: bool = True,
                   weld_dots: list[Vec2] | None = None,
                   tolerance: float = 1e-3) -> list[Callout]:
    """Draw component callouts -- one style for materials and cut pieces.

    Pipe pieces carry their cut-list name (e.g. "2A") and everything else its
    BOM item number, but both are drawn as the same square box offset from
    the centreline with a leader to the component. Pipe labels are placed
    first, so where the sheet is tight they get the shorter leaders: a cut
    piece is bound to one specific run, while an item number identifies a
    fitting that the box itself sits beside.

    With ``every_instance`` (the default) each component on the sheet gets its
    own box, so a BOM row of qty 6 balloons six times -- a fitter has to be
    able to identify the fitting in front of them, and a single balloon on one
    of six identical elbows does not do that. With it False only the first
    instance of each BOM row is ballooned.

    A **pipe** callout gets three places along its run to point at rather than
    one; a **fitting**'s stays on the fitting, because a balloon that does not
    point at the part it numbers is worse than a crowded one.
    """
    callouts: list[Callout] = []
    dots = list(weld_dots or [])

    # -- pass 1: pipe item-number callouts（Isogen 惯例：直管也标 BOM 项号） --
    for idx, (comp, _ref) in enumerate(comps):
        if idx not in anchors or classify(comp.type, comp.skey) != "PIPE":
            continue
        row = bom_rows.get(_bom_key(comp))
        label = str(row["item"]) if row else pipe_labels.get(idx)
        if not label:
            continue
        anchor = anchors[idx]
        # What the foot must keep clear of: every *other* component's anchor and
        # every weld dot -- the pipe's own anchor is one of the candidates, not
        # an obstacle.
        occupied = [p for i, p in anchors.items() if i != idx] + dots
        with canvas.tagged("callout", label):
            anchor, pos = _draw_callout_box(
                canvas, anchor, label, draw_centre,
                options=_pipe_anchor_options(comp, proj) or [anchor],
                occupied=occupied)
        callouts.append(Callout(
            item=label, kind="pipe", x_mm=round(pos[0], 2), y_mm=round(pos[1], 2),
            anchor_x_mm=round(anchor[0], 2), anchor_y_mm=round(anchor[1], 2),
            leadered=True))

    # -- pass 2: material item-number callouts ------------------------------
    # 垫片/螺栓优先并入任一端实际相邻的法兰（顺序“法兰/螺栓/垫片”）；两端都
    # 没有法兰时不并入阀门、三通等父件，单独出“螺栓/垫片”球标。
    flanges = [(i, comp) for i, (comp, _ref) in enumerate(comps)
               if looks_like_flange(comp.type, comp.description, comp.skey)]
    flange_children: dict[int, dict[str, str]] = {}
    loose_children: dict[tuple[str, str], dict[str, str]] = {}
    loose_anchors: dict[tuple[str, str], Vec2] = {}
    for idx, (child, child_ref) in enumerate(comps):
        child_type = child.type.upper()
        if child_type not in ("GASKET", "BOLT"):
            continue
        child_row = bom_rows.get(_bom_key(child))
        child_points = child.all_points()
        if child_row is None or not child_points:
            continue
        near_flange = next((
            i for i, flange in flanges
            if any(geometry.dist(p.xyz, q.xyz) <= tolerance
                   for p in child_points for q in flange.all_points())
        ), None)
        if near_flange is not None:
            group = flange_children.setdefault(near_flange, {})
        else:
            master = (child.get("MASTER-COMPONENT-IDENTIFIER") or "").strip()
            group_key = (child_ref, master or f"#{idx}")
            group = loose_children.setdefault(group_key, {})
            if group_key not in loose_anchors:
                projected = [proj.page(p.xyz) for p in child_points]
                loose_anchors[group_key] = (
                    sum(p[0] for p in projected) / len(projected),
                    sum(p[1] for p in projected) / len(projected))
        group[child_type] = str(child_row["item"])

    done: set[tuple] = set()
    for idx, (comp, _ref) in enumerate(comps):
        if idx not in anchors or classify(comp.type, comp.skey) == "PIPE":
            continue
        if comp.type.upper() in ("GASKET", "BOLT"):
            continue
        key = _bom_key(comp)
        row = bom_rows.get(key)
        if row is None:
            continue
        if not every_instance:
            if key in done:
                continue
            done.add(key)
        label = str(row["item"])
        children = flange_children.get(idx, {})
        label = "/".join([label] + [children[k] for k in ("BOLT", "GASKET")
                                    if k in children])
        anchor = anchors[idx]
        with canvas.tagged("callout", label):
            _anchor, pos = _draw_callout_box(canvas, anchor, label, draw_centre)
        callouts.append(Callout(
            item=label, kind="material",
            x_mm=round(pos[0], 2), y_mm=round(pos[1], 2),
            anchor_x_mm=round(anchor[0], 2), anchor_y_mm=round(anchor[1], 2),
            leadered=True))

    for group_key, children in loose_children.items():
        anchor = loose_anchors.get(group_key)
        if anchor is None:
            continue
        label = "/".join(children[k] for k in ("BOLT", "GASKET") if k in children)
        with canvas.tagged("callout", label):
            _anchor, pos = _draw_callout_box(canvas, anchor, label, draw_centre)
        callouts.append(Callout(
            item=label, kind="material",
            x_mm=round(pos[0], 2), y_mm=round(pos[1], 2),
            anchor_x_mm=round(anchor[0], 2), anchor_y_mm=round(anchor[1], 2),
            leadered=True))

    return callouts


def _draw_notes(canvas: _Canvas, notes: list[str], region: tuple) -> None:
    """Numbered NOTES block, top-left of the drawing region (under the north
    arrow). Its box is registered so dimensions and callouts route around it."""
    rx, ry, _rw, _rh = region
    x, y = rx + 2.0, ry + 22.0
    pdf = canvas.pdf
    pdf.set_font("Helvetica", style="B", size=6.5)
    pdf.text(x, y, "NOTES:")
    width = pdf.get_string_width("NOTES:")
    line_h = 3.0
    pdf.set_font("Helvetica", size=6.0)
    for n, note in enumerate(notes, start=1):
        text = f"{n}. {note}"[:70]
        pdf.text(x, y + n * line_h, text)
        width = max(width, pdf.get_string_width(text))
    canvas.label_boxes.append(
        (x - 1.0, y - 2.6, width + 2.0, line_h * len(notes) + 3.6))


def _draw_nts_stamp(canvas: _Canvas, region: tuple) -> None:
    """Bold NOT TO SCALE stamp, bottom-left of the drawing area. Its box is
    registered so callouts and dimensions route around it."""
    rx, ry, _rw, rh = region
    pdf = canvas.pdf
    text = "NOT TO SCALE"
    pdf.set_font("Helvetica", style="B", size=8.5)
    width = pdf.get_string_width(text)
    x, y = rx + 2.0, ry + rh - 2.5
    pdf.text(x, y, text)
    canvas.label_boxes.append((x - 1.0, y - 4.0, width + 2.0, 5.5))


def _draw_overflow_warning(canvas: _Canvas, draw_region: tuple, which: str) -> None:
    """Loud banner: a table was cut off, so the sheet must be split."""
    pdf = canvas.pdf
    rx, ry, rw, rh = draw_region
    text = f"{which} CUT OFF - SPLIT THIS SHEET"
    box_w, box_h = 100.0, 8.5
    bx = rx + (rw - box_w) / 2
    by = ry + 1.5
    with pdf.local_context():
        pdf.set_fill_color(214, 40, 40)
        pdf.set_draw_color(150, 0, 0)
        pdf.set_line_width(0.7)
        pdf.rect(bx, by, box_w, box_h, style="DF")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", style="B", size=10.0)
        tw = pdf.get_string_width(text)
        pdf.text(bx + (box_w - tw) / 2, by + box_h / 2 + 1.6, text)


def _draw_border(canvas: _Canvas, page_w: float, page_h: float, border: float) -> None:
    pdf = canvas.pdf
    pdf.set_line_width(0.7)
    pdf.rect(border, border, page_w - 2 * border, page_h - 2 * border)
    pdf.set_line_width(0.25)
    pdf.rect(border - 1.6, border - 1.6, page_w - 2 * border + 3.2,
             page_h - 2 * border + 3.2)


def _draw_north_arrow(canvas: _Canvas, proj: Projector, at: Vec2) -> None:
    angle = math.radians(proj.north_arrow_angle())
    tip = (at[0] + 6.0 * math.cos(angle), at[1] - 6.0 * math.sin(angle))
    canvas.line(at, tip, width=0.4)
    direction = _v2_norm(_v2_sub(tip, at))
    _draw_arrowhead(canvas, tip, _v2_scale(direction, -1.0))
    canvas.text((tip[0] + direction[0] * 3, tip[1] + direction[1] * 3 + 1.0), "N", size=8.0)


# BOM 面板几何（与 _draw_bom_panel 内的魔数一一对应，装箱容量估算共用）
_BOM_ROW_MM = 3.0        # BOM/切割表行高
_BOM_HEAD_MM = 9.4       # 面板头(cursor=y+4,+2.0)+表头(+3.4)
_BOM_TAIL_MM = 8.0       # 无切割表时行区底部保留
_CUT_HEAD_MM = 10.2      # CUT LIST 标题(+3.0+4.0)+表头(+3.2)
_CUT_TAIL_MM = 12.0      # 切割表行区底部保留（含 WELDS 摘要）


def _bom_capacity(region_h: float, show_cut_table: bool) -> int:
    """单页 BOM 行容量估算（分页装箱约束，取整保守）。"""
    head = _BOM_HEAD_MM + (_CUT_HEAD_MM if show_cut_table else 0.0)
    tail = _CUT_TAIL_MM if show_cut_table else _BOM_TAIL_MM
    return max(4, int((region_h - head - tail) / _BOM_ROW_MM))


def _draw_bom_panel(canvas: _Canvas, bom_rows: dict, cut_pieces: list[CutPiece],
                    welds: list[Weld], doc: PcfDocument,
                    x: float, y: float, w: float, h: float,
                    show_cut_table: bool) -> tuple[bool, bool]:
    """Draw the BOM/cut-list panel; return (bom_truncated, cut_truncated)."""
    bom_truncated = False
    cut_truncated = False
    pdf = canvas.pdf
    pdf.set_line_width(0.3)
    pdf.line(x, y, x, y + h)

    cursor = y + 4
    metric_qty = is_metric_units(doc.units_coords)
    canvas.text((x + w / 2, cursor), "BILL OF MATERIALS", size=7.5)
    cursor += 2.0

    col_item = x + 1.5
    col_desc = x + 9
    col_nps = x + w - 22
    col_qty = x + w - 9

    pdf.set_font("Helvetica", size=5.5)
    pdf.text(col_item, cursor + 2.4, "PC")
    pdf.text(col_desc, cursor + 2.4, "DESCRIPTION")
    pdf.text(col_nps, cursor + 2.4, "NPS")
    pdf.text(col_qty, cursor + 2.4, "QTY")
    cursor += 3.4
    pdf.line(x, cursor, x + w, cursor)

    rows = sorted(bom_rows.values(), key=lambda r: _item_sort_key(r["item"]))
    max_desc_chars = max(int((col_nps - col_desc) / 1.55), 8)
    for row in rows:
        if cursor > y + h - (30 if show_cut_table else 8):
            pdf.text(col_item, cursor + 2.6, "...")
            cursor += 3.0
            bom_truncated = True
            break
        if row["class"] == "PIPE":
            qty = (format_mm(row["length_in"] * 25.4) if metric_qty
                   else format_feet_inches(row["length_in"]))
        else:
            qty = str(row["qty"])
        # 图面 BOM：描述前带料号（项号列已是顺序号）
        desc = row["description"]
        if row["code"]:
            desc = f"{row['code']}  {desc}"
        desc = desc[:max_desc_chars]
        pdf.set_font("Helvetica", size=5.5)
        pdf.text(col_item, cursor + 2.6, str(row["item"]))
        pdf.text(col_desc, cursor + 2.6, desc)
        pdf.text(col_nps, cursor + 2.6, row["nps"])
        pdf.text(col_qty, cursor + 2.6, qty)
        cursor += 3.0

    if show_cut_table and cut_pieces:
        cursor += 3.0
        canvas.text((x + w / 2, cursor + 2), "CUT LIST", size=7.0)
        cursor += 4.0
        pdf.set_font("Helvetica", size=5.5)
        pdf.text(col_item, cursor + 2.4, "CUT")
        pdf.text(col_desc, cursor + 2.4, "NPS")
        pdf.text(col_nps, cursor + 2.4, "LENGTH")
        cursor += 3.2
        pdf.line(x, cursor, x + w, cursor)
        for piece in sorted(cut_pieces, key=lambda p: -p.cut_length):
            if cursor > y + h - 12:
                pdf.text(col_item, cursor + 2.6, "...")
                cursor += 3.0
                cut_truncated = True
                break
            pdf.text(col_item, cursor + 2.6, piece.label)
            pdf.text(col_desc, cursor + 2.6, piece.nps)
            pdf.text(col_nps, cursor + 2.6, piece.cut_length_text(doc.units_coords))
            cursor += 3.0

    # Weld summary at panel bottom.
    counts: dict[str, int] = {}
    for weld in welds:
        counts[weld.label] = counts.get(weld.label, 0) + 1
    if counts:
        summary = "  ".join(f"{label}: {n}" for label, n in sorted(counts.items()))
        canvas.text((x + w / 2, y + h - 2.0), f"WELDS - {summary}", size=6.0)

    return bom_truncated, cut_truncated


def _item_sort_key(item: str):
    try:
        return (0, int(item), item)
    except (TypeError, ValueError):
        return (1, 0, str(item))


# Field *labels* in the title block stay a step below their values -- the
# hierarchy the block already read with; only values are held to _SHEET_FONT.
_TITLE_LABEL_FONT = 6.5


def _fit_text(pdf, text: str, width_mm: float) -> str:
    """``text`` trimmed to fit ``width_mm`` at the current font, with an ellipsis.

    A fixed character count cannot do this job: it was 52 characters, chosen
    against a 6.5 pt font, and it cut every real drawing number mid-word
    ("...COOLING TOWER-01 02-16-2") while leaving most of the column empty.
    The column is wide; measure it.
    """
    if pdf.get_string_width(text) <= width_mm:
        return text
    ellipsis = "..."
    room = max(0.0, width_mm - pdf.get_string_width(ellipsis))
    cut = text
    while cut and pdf.get_string_width(cut) > room:
        cut = cut[:-1]
    return cut + ellipsis if cut else ""


def _draw_title_block(canvas: _Canvas, doc: PcfDocument, options: RenderOptions,
                      x: float, y: float, w: float, h: float,
                      sheet: tuple[int, int] | None = None) -> None:
    """sheet=(n, m) 时追加 SHEET n/m 字段（NTS 1:1 多页输出），字段网格
    按实际字段数自动分行。"""
    pdf = canvas.pdf
    pdf.set_line_width(0.4)
    pdf.line(x, y, x + w, y)

    logo_w = w * 0.18
    fields_x = x + logo_w + 2

    if options.logo_path and Path(options.logo_path).is_file():
        try:
            pdf.image(options.logo_path, x + 2, y + 2, w=logo_w - 4, h=h - 4,
                      keep_aspect_ratio=True)
        except Exception:
            canvas.text((x + logo_w / 2, y + h / 2), "LOGO", size=8)
    elif options.company_lines:
        for n, line in enumerate(options.company_lines[:4]):
            canvas.text((x + logo_w / 2, y + 5 + n * 4), line, size=6.5)
    pdf.line(x + logo_w, y, x + logo_w, y + h)

    pipeline_refs = ", ".join(p.reference for p in doc.pipelines[:3])
    if len(doc.pipelines) > 3:
        pipeline_refs += f" (+{len(doc.pipelines) - 3})"

    printed_date = options.date or _date.today().isoformat()
    fields = [
        ("PROJECT", options.project), ("JOB", options.job),
        ("LINE", pipeline_refs),
        ("DRAWING NO.", options.drawing_no or (Path(doc.source_path).stem
                                               if doc.source_path else "")),
        ("DATE", printed_date), ("REV", options.revision),
    ]
    if sheet is not None:
        fields.append(("SHEET", f"{sheet[0]}/{sheet[1]}"))
    n_field_rows = max(1, math.ceil(len(fields) / 2))
    # Split the width right of the logo between the two field columns and the
    # optional DRAWING PROPERTIES / REQUIREMENTS list columns.
    avail_w = w - logo_w - 4
    n_lists = (1 if options.properties else 0) + (1 if options.requirements else 0)
    fields_frac = {0: 1.0, 1: 0.64, 2: 0.5}[n_lists]
    col_w = avail_w * fields_frac / 2
    list_w = (avail_w * (1 - fields_frac) / n_lists) if n_lists else 0.0
    for n, (label, value) in enumerate(fields):
        col = n % 2
        row = n // 2
        fx = fields_x + col * col_w
        fy = y + 4 + row * (h - 6) / n_field_rows
        pdf.set_font("Helvetica", size=_TITLE_LABEL_FONT)
        pdf.text(fx, fy, label)
        pdf.set_font("Helvetica", size=_SHEET_FONT)
        pdf.text(fx, fy + 3.4, _fit_text(pdf, str(value), col_w - 4))

    # Drawing properties and requirements / quality notes on the right end.
    list_x = fields_x + 2 * col_w + 4
    if options.properties:
        _draw_title_list(pdf, list_x, y, h, "DRAWING PROPERTIES",
                         [f"{k}: {v}" for k, v in options.properties.items()])
        list_x += list_w
    if options.requirements:
        _draw_title_list(pdf, list_x, y, h, "REQUIREMENTS", options.requirements)

    if options.not_to_scale:
        canvas.text((x + w - 18, y + h - 2.0), "NOT TO SCALE", size=5.5)


def _draw_title_list(pdf, x: float, y: float, h: float, heading: str,
                     lines: list[str]) -> None:
    """One headed list column in the title block (properties, requirements)."""
    pdf.line(x - 2, y, x - 2, y + h)
    pdf.set_font("Helvetica", size=5.0)
    pdf.text(x, y + 4, heading)
    pdf.set_font("Helvetica", size=5.2)
    for n, line in enumerate(lines[:6]):
        pdf.text(x, y + 7.5 + n * 3.0, str(line)[:60])
