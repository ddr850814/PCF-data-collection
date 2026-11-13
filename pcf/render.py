"""Rules-based single-line isometric PDF rendering with weld capture.

Projection: true isometric. After the plan is oriented (north axis + view
quadrant), East contributes (+cos30, +sin30), North (+cos30, -sin30), and
Up (0, -1) in page space (y grows downward, matching PDF coordinates), so
elevation always draws straight up and both plan axes draw at 30 degrees.
A sheet carrying piping that is flat to the ground on a 45/135/225/315
bearing gets one departure from that: such a run projects dead vertical,
the way a riser reads, so a plan shear tilts it ~10 degrees off vertical
(RenderOptions.skew_deg). Risers stay vertical; see set_skew().

Symbols are dispatched per component class using the conventions proven in
the wmdb renderers: pipes are single lines, elbows turn at their
centre-point corner, valves are bow-ties with subtype decorations, reducers
are trapezoids (eccentric ones keep their FLAT-DIRECTION side flat), olets
branch off the header centreline. Flanges draw one long bar at the gasket
face -- the flanged end, not the welded one -- skewed onto an isometric
plane like the dimension callouts and about twice a dimension offset long,
so a flanged connection is unmistakable on a full sheet.

Welds are drawn as small filled dots only -- their numbered balloons live
in the weld-tracking system as a separate overlay layer, not on the sheet.
Every weld's page x/y (mm, top-left origin) is still captured during the
same layout pass that draws the dot, so the weld-map upload keeps working;
positions are never re-measured from the PDF.

Component callouts are one style for everything the shop has to identify:
a square box carrying the item number (materials) or the cut-list name
(pipe pieces), standing off the centreline with a leader to the component.
If the BOM or cut-list table overflows the sheet a loud "SPLIT SHEET"
warning is drawn and reported.

Continuation callouts mark every point the PCF declares as leaving the
sheet (END-CONNECTION-PIPELINE): an arrow along the run plus a block
printing the declared reference -- "SEE ISO <line number>" -- and the
declared point's plant coordinates. The render set only maps the
declared neighbour to the sheets actually carrying it (SHT tags,
split-junction phantoms) and warns when nothing carries it; the hand-off
points themselves are never re-derived from geometry (see
continuations.py).

Dimensional callouts have isometric skew: leader bars, dimension lines and
the *text itself* run along projected world axes, so the whole callout --
glyphs included -- lies on the isometric plane its goalposts define. Point
pairs that are not axis-aligned are decomposed into a dashed offset
rectangle (one plane) or cuboid (three axes) whose legs are dimensioned
individually on the least-colliding side. Whether a leg counts is decided
in page millimetres rather than as a fraction of the model, so a small roll
in a long line still gets its box. Three standard text placements are used
according to how much room the dimension has: text breaking the dimension
line, text leadered off a full-length line, or -- when the span is too
short even for the arrows -- arrows outside the goalposts with leadered
text. See _plan_dim_geometry().
"""

from __future__ import annotations

import json
import math
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date as _date
from itertools import permutations
from pathlib import Path

from . import geometry
from .continuations import Continuation, reference_names
from .cutlist import CutPiece, build_cutlist
from .dims import (DEFAULT_OFFSET_MM, Dimension, flange_face,
                   infer_dimensions, load_sidecar, merge_with_sidecar)
from .model import Component, PcfDocument, Point
from .skeys import (classify, is_annotation_symbol, is_slip_on_flange,
                    is_weld_coupling, looks_like_flange,
                    reducer_is_eccentric, type_from_skey, valve_subtype)
from .units import (coord_to_inch_factor, format_bore, format_feet_inches,
                    format_length, format_mm, is_metric_units)
from .welds import (GAP_COUPLING_MAX_INCHES, Weld, component_spools,
                    derive_welds)

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
    # Page length (mm) of the bar drawn across the gasket face of a flange.
    # Twice the standard dimension offset by default, which is what makes a
    # flanged connection read at a glance on a full sheet.
    flange_bar_mm: float = 2 * DEFAULT_OFFSET_MM
    show_bom: bool = True
    show_cut_table: bool = True
    detect_ffw: bool = False
    dims_sidecar: str | None = None  # read overrides from here when set
    # Callout crowding: 1.0 = default spacing, 1.35 = one "spread out" step,
    # 0.74 = one "pack closer" step. Clamped to [0.4, 3.0] at render time.
    # Controls the leadered material (item-number) callouts; weld dots and
    # in-place pipe labels are fixed to geometry and never move.
    balloon_spread: float = 1.0
    # Balloon every instance of a material item, not just the first one on the
    # sheet. A fabrication isometric has to let a fitter identify each fitting
    # in front of them, so six elbows off one BOM row need six item-number
    # boxes. Set False for the older one-balloon-per-BOM-row behaviour, which
    # gives a much quieter page at the cost of leaving fittings unidentified.
    balloon_every_instance: bool = True
    # Schematic (not-to-scale) drawing: pipe/nipple runs are compressed or
    # stretched so every visible span lands between nts_min_mm and nts_max_mm
    # on the page, and the sheet is stamped NOT TO SCALE. This warps ONLY the
    # picture -- cut lengths, BOM totals, and dimension text are computed from
    # the true 3D model and never change.
    not_to_scale: bool = True
    nts_min_mm: float = 8.0
    nts_max_mm: float = 70.0
    # Visual skew (degrees) applied to piping that is flat to the ground but
    # runs on a 45/135/225/315-degree bearing: in a true isometric such a run
    # draws as a dead vertical line, on top of how risers and downcomers
    # read. The skew tilts it off vertical so the two cannot be confused. It
    # is applied only to sheets that contain such a run, and like the
    # not-to-scale warp it changes the picture only. 0 disables it.
    skew_deg: float = 10.0
    # Hybrid (distance-split) mode: the grouping attribute whose values are
    # the spools on this sheet (e.g. SPOOL-IDENTIFIER). When set, each spool
    # region is labelled on the drawing and every weld joining two different
    # spools is declared a field weld: "F.W." tag at the dot, label FW, and
    # "Field Weld" in the welds CSV notes column.
    spool_key: str | None = None
    # 管线清单实时预览：只画图面内容（管线/符号/焊点/尺寸/注释/北向标/
    # NTS 戳），不出图框、标题栏、BOM/切割表等"图纸家具"。
    preview: bool = False
    # Where the piping leaves this sheet: one entry per END-CONNECTION-
    # PIPELINE declaration (point = its CO-ORDS, name = its reference),
    # built by continuations.build_continuation_map() over the documents
    # being rendered together -- the set only resolves which rendered
    # sheets carry the declared neighbour (SHT tags, phantom mates).
    continuations: list[Continuation] = field(default_factory=list)
    # 整套图纸的文档表（图名 → PcfDocument），由 render_pcf_files 在批量
    # 渲染时传入。分图幻影件需要邻图文档才能把断点对侧的紧邻零件画成
    # 虚线幻影；单图渲染（预览）为空时幻影件退化为只画接续标记。
    neighbour_docs: dict[str, PcfDocument] | None = None


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
        # Optional not-to-scale warp: a monotonic per-axis world-coordinate
        # remap applied before projection. Directions and topology survive;
        # only drawn span lengths change.
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
        # 链式布局（示意）专用：整图示意缩放（管件尺寸与位置同乘，保持
        # 搭积木不叠件）与最近点平移回退（未映射的派生世界点——尺寸盒
        # 角点/设备管口——跟随最近的已布几何）。
        self.schematic_scale = 1.0
        self.offset_nearest = False

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
                # 链式布局：未映射的派生世界点（尺寸盒角点/设备管口等）
                # 取世界距离最近已映射点的平移，跟随邻近几何
                best: tuple[float, Vec2] | None = None
                for k, d in self.point_offsets.items():
                    dd = ((k[0] - p[0]) ** 2 + (k[1] - p[1]) ** 2
                          + (k[2] - p[2]) ** 2)
                    if best is None or dd < best[0]:
                        best = (dd, d)
                off = best[1] if best else (0.0, 0.0)
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
# Schematic isometrics compress long runs and stretch short nipples so the
# drawing stays readable. The warp is a monotonic piecewise-linear remap of
# each world axis: the gap between adjacent coordinate planes is clamped so
# its projected length lands in [nts_min_mm, nts_max_mm]. Monotonic per-axis
# mapping preserves topology exactly -- everything that touched still
# touches, parallel runs stay parallel and ordered -- while sub-gap features
# (modelled 1/16" weld root gaps) are left untouched so they don't balloon
# into visible openings. RENDERING ONLY: cut lengths, BOM totals, and
# dimension text are always computed from the true model coordinates.

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

    ``rigid`` is a per-axis list of ``(lo, hi, s)`` intervals to remap with one
    shared slope ``s`` instead of clamping gap by gap. It is what keeps runs
    that are **not parallel to a world axis** from being deformed. A run along
    a world axis draws at a page direction that depends only on which axis it
    is on, so any monotonic per-axis remap leaves it alone; a skewed run's page
    direction is ``(sx*dx, sy*dy, sz*dz)``, and the three local slopes have
    nothing to do with each other. Clamping them independently swings the run
    to an arbitrary angle -- and swings an elbow leg and the pipe welded to it,
    collinear in the model but sitting in different cells of the breakpoint
    grid, to *different* arbitrary angles. Measured on real sheets that came
    back from review, up to 47 degrees apart: the fitting visibly not on the
    pipe. Inside a rigid interval every axis uses the same ``s``, so a skewed
    run maps as ``p -> s*p + c``: direction exactly preserved, collinear stays
    collinear, parallel stays parallel. See ``_skew_rigid_spans``.

    ``targets`` is a per-axis list of ``(lo, hi, target_mm)`` intervals to map
    to a FIXED page length: inline fittings (valves, flanges, concentric
    reducers...) whose coordinate span is forced to their symbol's natural
    page size. The schematic principle: parts print at symbol size whatever
    the scale, and only straight pipe absorbs the non-to-scale stretching --
    gaps inside a target interval are scaled proportionally so the whole
    interval lands on ``target_mm``; the space that true-proportion fittings
    would have eaten (or the [min,max] clamp would have inflated them to)
    stays with the pipe runs.

    ``rigid`` wins over ``keep`` where they overlap: a gap being scaled by a
    shared slope is not being ballooned to ``min_mm``, which is the only thing
    ``keep`` exists to prevent.

    ``keep`` is a per-axis list of coordinate *intervals* to leave at true
    proportion. The warp stretches short spans so they can be seen, but not
    every small gap is a span worth seeing: an eccentric reducer's two weld
    ends sit on centrelines half an inch to an inch apart, and blowing that
    up to ``min_mm`` puts a dog-leg in the pipeline seven times the size of
    the real offset -- and swamps the reducer's own flat side, which is the
    eccentricity drawn at proportion. See ``_eccentric_offset_pairs``.

    Any gap falling **wholly inside** a protected interval is left alone,
    rather than only a gap whose two ends *are* the interval. The stricter
    rule looked conservative and was simply broken: the moment anything else
    on the sheet sat at a coordinate between a reducer's two ends -- the rest
    of the spool, routinely, since the reducer straddles the run centreline
    -- the interval was subdivided, neither sub-gap matched the pair, and the
    offset was inflated after all. A protected interval is an inch wide at
    most, so no leg of pipe can hide inside one.
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
        # ``locked``/``forced`` are sorted and non-overlapping, so the only
        # interval that can contain a gap is the last one starting at or
        # before it. Binary search rather than scan, to keep this linear in
        # the number of coordinates on sheets that carry thousands of both.
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


# A direction component below this (inches) is modelling noise, not a second
# axis the run genuinely travels on. Far under anything a fabricator lays out,
# and far under the connection tolerance, so it never reclassifies a real run.
_SKEW_AXIS_FLOOR_INCHES = 1.0 / 64.0


def _skew_segments(runs: list[tuple[Vec3, Vec3]], points: list[Vec3],
                   tol: float, inch_factor: float) -> list[tuple[Vec3, Vec3]]:
    """Skewed runs, cut at every model point that sits on them.

    A run parallel to a world axis needs none of this: its page direction is
    the projected axis whatever the warp does to lengths. Everything else is
    cut into the pieces the sheet actually has landmarks on, so a piece can be
    compressed or stretched on its own without dragging its neighbours -- a
    45-degree elbow leg still gets stretched to a visible length beside the
    twenty-foot diagonal it welds onto.

    "Parallel to an axis" is judged **absolutely**, against a fabrication
    floor, not as a fraction of the run's own length. Two reasons it cannot
    be a fraction. A twenty-foot run that drifts two inches sideways is 1%
    off axis and still a skewed run -- left to the clamp it swings by
    degrees. And a fraction splits runs that are *collinear*: the same
    direction reads as skewed on the long pipe and as axis-parallel on the
    six-inch elbow leg welded to it, one gets a shared slope and the other
    gets clamped, and the kink this exists to remove comes back at a fifth
    of its size. The floor is well under anything a fabricator lays out, so
    only genuine modelling noise falls through it.
    """
    from bisect import bisect_left, bisect_right

    floor = _SKEW_AXIS_FLOOR_INCHES / inch_factor
    # One sorted index per axis, so a run only tests the points inside its own
    # coordinate slab rather than every point on the sheet -- the naive
    # product is runs x points, and both run to thousands on a large sheet.
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
        # Take the slab that actually excludes the most: the run's narrowest
        # axis is not it. A flat 2D offset has zero extent in Z, and in a
        # mechanical room every point at that elevation is in the Z slab.
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
        # Stations are the model's own points, never a point reconstructed
        # from t: the warp's breakpoints are the model coordinates rounded to
        # six places, and a claim whose edge misses one of them by a float ulp
        # stops covering the gap it was built for.
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

    Each skewed segment asks for the slope that lands it in
    ``[min_mm, max_mm]`` on the page -- the same target the gap-by-gap clamp
    aims at -- and claims its own extent on each of the three axes. Because a
    claim covers all three axes at once, the segment maps as ``p -> s*p + c``:
    its direction is untouched, and two segments joined end to end stay in
    line however differently they are scaled.

    Two segments whose claims overlap on any axis cannot both have their way,
    since the axis remap is a function. Those are merged into one group under
    one slope, iterated to a fixed point so that in the end no two groups
    overlap anywhere. The degenerate outcome -- one group across the whole
    sheet -- is a sheet drawn at true proportion, which is the honest answer
    for geometry that cannot be axis-warped without bending.

    Takes the segments from ``_skew_segments``, which depend only on the
    model and are cut once for the whole warp loop. Returns None when the
    sheet is orthogonal, which is the common case: the warp is then built
    exactly as it was before this existed.
    """
    if not segments:
        return None

    # Page length of each segment as the warp is about to leave it: the rigid
    # map is p -> s*p + c, so its projected length is s * (unwarped projected
    # length) * scale. Measure unwarped, then solve for s. Sub-tolerance gaps
    # (modelled weld root gaps) are left alone, exactly as the clamp does.
    warp, proj.warp = proj.warp, None
    try:
        lengths: list[float | None] = []
        for a, b in segments:
            pa, pb = proj.raw(a), proj.raw(b)
            page = math.hypot(pb[0] - pa[0], pb[1] - pa[1]) * proj.scale
            # Sub-tolerance gaps are modelled weld root gaps. They ride the
            # group's slope like everything else, but they must not be the
            # thing that *sets* it: asking a 1/16" gap to reach nts_min_mm
            # would blow the group up by two orders of magnitude.
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

    # "No two groups overlap on any axis" is the same statement as "no two
    # *claims* from different groups overlap", because a group's claims are
    # exactly its members'. So one sweep per axis reaches the fixed point --
    # sort by lower edge, and chain each claim onto the run of claims it
    # touches. No iterating to convergence, and no pairwise comparison.
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

    # One slope per group, so relative proportions *inside* a group are fixed
    # whatever it is; all it decides is how big the group is beside the rest
    # of the sheet. Take the smallest slope that lifts the group's shortest
    # segment to nts_min_mm, then cap it so its longest stays under
    # nts_max_mm. When the group's own span ratio exceeds max/min the two
    # bounds cross and no slope satisfies both: the cap wins, so the sheet is
    # never blown off the page, and the group prints at true proportion --
    # the honest answer for geometry that cannot be axis-warped without
    # bending. Staying at 1.0 when neither bound bites keeps a group that is
    # already readable exactly where the model put it.
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
# In a true isometric both plan axes project at 30 degrees, so a run that is
# flat to the ground but bearing 45/135/225/315 degrees -- half of it along
# each plan axis -- projects to a dead vertical line, indistinguishable from
# a riser or a downcomer. Drafters break the tie by hand with a visual trick:
# the diagonal run is drawn about 10 degrees off vertical. Nothing about the
# pipe changes, only the picture; the dimensions, cut lengths and BOM are all
# still taken from the true model.
#
# Projector.set_skew() is that trick as a plan shear. It moves ONLY the
# diagonal that would collide: risers stay exactly vertical, and the plan
# axes shift by about 2 degrees off 30 (a 45-degree line cannot be tilted by
# a projection that leaves both axes it is made of untouched -- something has
# to give, and 2 degrees on the axes buys 10 on the diagonal).
#
# It is applied only when the sheet actually has such a run, so an orthogonal
# spool -- the common case -- still draws as a textbook isometric.

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
# local_context/rotation/transform 上下文），由具体后端决定落向——
# _FpdfBackend 原样转发 fpdf2 真正写出 PDF；_SceneBackend 收集自包含的
# SceneItem（颜色/线宽/字号/变换全部随项快照），供 GUI 主线程用 tk.Canvas
# 回放矢量预览，不再经 PyMuPDF 栅格化。


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
# get_string_width，预览回放必须与 PDF 用同一套 Helvetica 度量才能保证
# 同版式。fpdf2 导入推迟到首次使用。
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
    """绕 (cx, cy) 旋转 angle_deg 度（逆时针为正，同 fpdf.rotation）。"""
    rad = math.radians(angle_deg)
    cos_a, sin_a = math.cos(rad), math.sin(rad)
    rot = (cos_a, sin_a, -sin_a, cos_a, 0.0, 0.0)
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
        """Draw text *skewed onto an isometric plane*, not merely rotated.

        ``along`` is the direction the baseline runs and ``up`` the direction
        the glyph strokes rise -- two projected world axes, so the letters lie
        on the same plane as the callout that carries them. Rotation alone
        cannot do this: it keeps the glyphs square to their own baseline,
        which reads as flat text pasted onto an isometric picture. What a
        drafter does by hand (and what AutoCAD's isometric text styles do) is
        rotate *and* oblique, which is exactly the linear map that sends the
        text-space axes onto the plane's two axes.

        The text is centred on ``baseline`` along the baseline direction: the
        half-width shift is applied before the transform, so it is carried
        onto the plane with everything else. The background is knocked out
        white first, in the same frame, so a value landing on a centreline
        stays readable -- and so that the knockout is the shape of the text
        rather than an upright box several times its size.
        """
        try:
            from fpdf.drawing_primitives import Transform
        except ImportError:  # fpdf2 < 2.8
            from fpdf.drawing import Transform

        self.pdf.set_font("Helvetica", size=size)
        width = self.pdf.get_string_width(message)
        cap = _cap_height_mm(size)
        x, y = baseline
        # PDF matrix [a b c d e f]: x' = a.x + c.y + e, y' = b.x + d.y + f.
        # Text space has x along the baseline and y *down*, so (1,0) -> along
        # and (0,1) -> -up.
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
        """Find a label centre near anchor that avoids earlier labels.

        Candidates are swept around a *ring* of directions at increasing
        radius, ordered so the natural placement -- straight out from the
        drawing centre, shortest leader -- is tried first and still wins
        whenever it is free. The search used to run along that one ray only,
        which meant that once it walked past the edge of the drawing region
        ``_clamp`` pinned every remaining candidate to the same point and the
        rest of the budget was spent re-testing it.

        With ``avoid_geometry`` the drawn segments are scored as well as the
        earlier label boxes. That is what keeps a spool label off the pipeline
        it names: without it the placer knows only about other labels and will
        happily drop a box straight onto a centreline.

        ``clearance`` (page mm, before spread) widens the box for that
        geometry test only, so the label has to keep real whitespace rather
        than merely not intersect. Not intersecting is a low bar: a box whose
        edge grazes a centreline scores zero and is accepted at the shortest
        leader, which is how these came out legibly placed but visually
        crowded against the run.

        If nothing is clean, the least-bad candidate is returned rather than
        the anchor itself -- a label sitting exactly on its own component was
        the old fallback and it is the worst position available.
        """
        pos, box, _score = self._best_placement(
            anchor, w, h, away_from, min_leader, avoid_geometry, clearance)
        self.label_boxes.append(box)
        return pos

    def place_label_multi(self, anchors: list[Vec2], w: float, h: float,
                          away_from: Vec2, min_leader: float = 5.0,
                          avoid_geometry: bool = False, clearance: float = 0.0,
                          anchor_cost=None) -> tuple[Vec2, Vec2]:
        """Place one label, choosing among several places to *point at*.

        ``place_label`` is good at moving the box and cannot move the leader's
        foot, so a component that happens to have a fitting sitting on its
        anchor gets a leader landing on that fitting every time -- which is
        what a pipe with an olet welded on at its midpoint looks like. Given a
        list of anchors it scores each one's best placement and takes the pair
        that wins, so the callout can step along the run to somewhere clear.

        ``anchor_cost(anchor) -> float`` prices the anchor itself: a foot
        landing on a neighbouring fitting or a weld dot is bad however good
        the box above it is. Anchors are offered best-first, so a tie leaves
        the caller's preferred anchor (the midpoint) in place.

        Returns ``(anchor, position)``.
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
                # Graded, so that when nothing is clean the least-bad
                # candidate is the roomiest one rather than merely the one
                # with the shortest leader: crossing a line costs twice what
                # crowding it does. With no clearance asked for, the two
                # boxes are the same and this is the old flat count.
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
    pdf.add_page()
    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    result = _render_document(doc, options, _FpdfBackend(pdf),
                              page_w, page_h, pdf_path=output_path)
    pdf.output(str(output_path))
    return result


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
) -> RenderResult:
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

    comps: list[tuple[Component, str]] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append((comp, pipeline.reference))
    if not comps:
        raise ValueError("document contains no drawable components")

    all_points = [p.xyz for comp, _ in comps for p in comp.all_points()]

    # 设备管口（END-CONNECTION-EQUIPMENT 伪组件）：管口点属于本图画面，
    # 纳入投影范围；画法见 _draw_equipment_connections。
    equip_recs = [rec for pipeline in doc.pipelines
                  for rec in pipeline.components
                  if rec.type == "END-CONNECTION-EQUIPMENT"]
    all_points += [p.xyz for rec in equip_recs for p in rec.co_ords]

    # 分图幻影件：邻图断点紧邻零件的关键点纳入投影/NTS 变形范围，
    # 保证幻影几何与本图共用同一投影（含非比例伸缩）。只影响画面，
    # 不进 BOM/切割清单/避让集合（那些只吃 doc 本身）。
    for pxyz in _phantom_points(doc, options):
        all_points.append(pxyz)

    projector = Projector(options)
    projector.fit(all_points, draw_region)

    reducer_runs = _reducer_runs(doc)

    # 搭积木链式布局（示意模式）：位置由拓扑走链决定；失败回退 warp
    chain = None
    if options.not_to_scale and len(all_points) > 1:
        try:
            chain = _chain_layout(projector, doc, comps, draw_region)
            for _ in range(4):
                # 超幅：直管档位按缩放比收缩重摆（管件固定尺寸不动），
                # 迭代到装得下或封顶（固定件占比越大收敛越慢）
                if chain is None or chain[2] >= 0.98:
                    break
                chain = _chain_layout(projector, doc, comps, draw_region,
                                      pipe_scale=chain[2])
        except Exception as exc:
            # 失败必须可见：静默回退会让"链式没生效"伪装成布局缺陷
            chain = None
            warnings.append(f"chain layout failed ({exc}); fell back to warp")
    chain_on = chain is not None and bool(chain[0])
    if chain_on:
        projector.point_offsets.update(chain[0])
        # 位置表已按 s 缩放：示意符号尺寸与弯头/三通腿同乘 schematic_scale
        # （否则 0.6x 位置配 1.0x 符号，管件互相压叠）；未映射派生点走
        # 最近点平移回退，跟随邻近几何
        projector.schematic_scale = min(chain[2], 1.0)
        projector.offset_nearest = True
        if chain[1]:
            warnings.append(chain[1])
        # 偏心变径：小端连通侧整体平移到符号出口（按 skey 正算的小端
        # 中点，与 warp 通道同一机制）。链式只把变径两端摆上主轴
        # （_sym_dir），偏心量由这里叠加，下游管线从符号小端中点续走
        for k, dv in _eccentric_small_side_offsets(
                projector, doc, comps, reducer_runs,
                proportional=False).items():
            prev = projector.point_offsets.get(k)
            if prev is not None:
                projector.point_offsets[k] = _v2_add(prev, dv)
    if not chain_on and options.not_to_scale and len(all_points) > 1:
        # Fixed-count iteration: the warp targets depend on the page scale
        # and the scale depends on the warped extents, so refit a few times.
        # Deterministic -- same model and options always converge identically.
        tiny_world = _NTS_TINY_INCHES / coord_to_inch_factor(doc.units_coords)
        # 零件示意尺寸：直管件区间固定映射为符号自然图面长度，伸缩只由
        # 直管吸收（见 _fitting_symbol_targets）。偏心变径的偏心量区间
        # 仍按真实比例保护。
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

    # Flat 45-degree runs project vertical, exactly like a riser; tilt them
    # off vertical so the sheet can be read. Checked against the projector as
    # already configured, so the warp is accounted for, and skipped entirely
    # when nothing on this sheet collides. 链式布局不用 skew：平移表以
    # 当时的投影为基准，后置 shear 会把映射整体带偏。
    skew_deg = max(0.0, min(20.0, options.skew_deg))
    if (not chain_on and skew_deg
            and _draws_flat_run_vertically(projector, _straight_runs(comps))):
        projector.set_skew(skew_deg)
        projector.fit(all_points, draw_region)

    draw_centre = (draw_region[0] + draw_region[2] / 2,
                   draw_region[1] + draw_region[3] / 2)

    flange_face_pts = _flange_faces(doc)
    if not chain_on:
        # 正算偏心变径后，小端连通侧整体接到符号 EXIT；变径自身不偏移。
        projector.point_offsets = _eccentric_small_side_offsets(
            projector, doc, comps, reducer_runs,
            proportional=projector.proportional)
    # 小端后直接出图时没有下游管件可携带平移，接续线也要从符号 EXIT
    # 接出；只改图面锚点，标注仍打印 PCF 声明坐标。
    projector.point_offsets.update(_continuation_eccentric_offsets(
        projector, doc, options.continuations, reducer_runs,
        proportional=projector.proportional))
    if not chain_on and not projector.proportional:
        # 等比符号端面带动连接点（仅示意模式）：管端修剪、焊点吸附、尺寸
        # goalpost 都随符号实际端面走——端点/尺寸计算"考虑 skey"的统一
        # 通道；按比例模式符号本就铺满端点，无需平移。
        # 链式布局已产出全点位置表，本通道与其二选一。
        projector.point_offsets.update(_symbol_face_offsets(
            projector, doc, comps, reducer_runs, flange_face_pts))

    # ---- symbol pass -----------------------------------------------------
    _FLANGE_BANDS.clear()      # 每张图重置接带注册表（焊点吸附用）
    _FLANGE_WELDS.clear()
    # 管架定向用的管段缓存同样按图重置：它以 id(proj) 为键，批量渲染时
    # 新 Projector 可能复用已释放对象的 id，命中上一张图的残留缓存会把
    # 管架对到别的图的管段上（31006 管架画成竖直的根因）。
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
    # The numbered balloons are drawn as an overlay layer by the weld-tracking
    # system, not on the sheet -- but every weld's page position is still
    # captured here so the weld-map upload has something to anchor to.
    welds = derive_welds(doc, detect_ffw=options.detect_ffw,
                         spool_key=options.spool_key)
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
        with canvas.tagged("weld", weld.number):
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
    # 分图（同管线断开）→ 邻图紧邻零件画虚线幻影 + SEE ISO/SHT/坐标标注；
    # 跨管线接点（不同线号）→ 虚线示意短管 + SEE ISO 对方线号/坐标标注。
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
    pipe_labels = {c.comp_index: c.label
                   for c in cut_pieces if c.comp_index is not None}
    callouts: list[Callout] = []
    if options.show_bom:
        callouts = _draw_callouts(canvas, projector, comps, anchors,
                                  bom_rows, pipe_labels, draw_centre,
                                  every_instance=options.balloon_every_instance,
                                  weld_dots=weld_dots,
                                  tolerance=geometry.model_tolerance(doc) * 4)

    # ---- furniture ----------------------------------------------------------
    # 预览模式仍画北向标/注释/NTS 戳（属于图面内容），只跳过
    # 图框/BOM/切割表/标题栏等图纸家具。
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
                          w=page_w - 2 * border, h=title_h)

    bom_lines = [
        BomLine(item=str(r["item"]), description=r["description"], nps=r["nps"],
                cls=r["class"], qty=r["qty"], length_in=round(r["length_in"], 3),
                weight=round(r["weight"], 3),
                ucis=";".join(sorted(r["_ucis"])), code=r["code"])
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
    )


# ---------------------------------------------------------------------------
# Spool annotations (hybrid distance-split mode)


def _box_edge_toward(pos: Vec2, target: Vec2, w: float, h: float) -> Vec2:
    """Where the ray pos->target crosses the w x h box centred on pos."""
    d = _v2_sub(target, pos)
    tx = (w / 2) / abs(d[0]) if abs(d[0]) > 1e-9 else float("inf")
    ty = (h / 2) / abs(d[1]) if abs(d[1]) > 1e-9 else float("inf")
    return _v2_add(pos, _v2_scale(d, min(tx, ty, 1.0)))


# The sheet's reading size, in points. The fab/spool label set it -- it is
# the mark a fabricator on a noisy shop floor has to read first -- and the
# dimension values and title-block entries are held to the same size, because
# they are read in the same conditions and at the same arm's length. They used
# to be 6.0 and 6.5/5.0 against the label's 8, which read as small print.
_SHEET_FONT = 8.0

# Whitespace a spool label must keep from any drawn line, in page mm before
# spread. Merely not intersecting is not enough: a box that grazes a
# centreline is legible but reads as crowded onto the run it names.
_SPOOL_LABEL_CLEARANCE_MM = 2.2

# How far off the centreline an "F.W." tag stands, in page mm before spread.
# Clear of a 0.55 mm run line and its weld dot, close enough to still read as
# belonging to that joint.
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
        # Spool labels are the first thing placed on the sheet, so without
        # geometry avoidance the very first candidate is accepted whatever it
        # lands on -- which is how these ended up struck through by the run
        # they name.
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

        # Set parallel to the centreline and stood off it, the way a field
        # weld is tagged by hand. Free-placed horizontally the tag came to
        # rest across the pipe at whatever angle the run happened to be, and
        # four letters lying over a centreline are not readable.
        angle = math.degrees(math.atan2(-along[1], along[0]))
        if angle > 90 or angle < -90:
            angle += 180          # never upside down
        off = _v2_perp(along)
        away = _v2_sub(dot, draw_centre)
        if off[0] * away[0] + off[1] * away[1] < 0:
            off = _v2_scale(off, -1.0)      # stand it off outwards
        pos = _v2_add(dot, _v2_scale(off, _FW_STANDOFF_MM * canvas.spread))
        # Knock the background out before drawing: a near miss still leaves a
        # centreline or an extension line running through the letters. The
        # knockout is rotated *with* the text -- squared to the page it is
        # much taller than the glyphs at 30 degrees and takes a bite out of
        # the run line and the weld dot it is standing next to.
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
# Every point where the piping runs off the sheet is marked with the run
# carried a little past the hand-off, an arrowhead on the end of it, and a
# block naming the drawing it continues onto plus the plant coordinates of
# that point. The coordinates are what lets a fitter or a surveyor tie the
# two sheets together in the field without hunting for the matching end, so
# they are printed whether or not the drawing name resolved.
#
# The name is the NEXT RENDERED DRAWING. It cannot come from the PCF: the
# file knows only which *pipeline or spool* carries on, and a sheet is not a
# spool -- a distance split puts several spools on one drawing. See
# continuations.py for how the set of sheets is cross-matched. The file's own
# reference is used only as a last resort, and says so in the warnings.

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
    """分图接续标记：沿线箭头 + SEE ISO 管线号/SHT/坐标文字块（仅分图型）。"""
    at = proj.page(cont.point)
    out: Vec2 | None = None
    if cont.inboard is not None:
        out = _v2_norm(_v2_sub(at, proj.page(cont.inboard)))
    if not out or out == (0.0, 0.0):
        out = _v2_norm(_v2_sub(at, draw_centre)) or (0.7071, -0.7071)

    tip = _v2_add(at, _v2_scale(out, _CONT_STUB_MM))
    canvas.line(at, tip, width=0.35)
    _draw_arrowhead(canvas, tip, _v2_scale(out, -1.0))

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
# 分图幻影件 / 跨管线虚线延续


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
        tags = [_sheet_tag(n) for n in cont.sheets]
        if tags:
            lines.append(" / ".join(tags))
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
            pa = proj.page(eps[near].xyz)
            pb = proj.page(far.xyz)
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
            e0 = proj.page(eps[0].xyz)
            e1 = proj.page(eps[-1].xyz)
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
                corner = proj.page(comp.centre_points[0].xyz)
                canvas.dashed(e0, corner, width=0.3)
                canvas.dashed(corner, e1, width=0.3)
            elif cls in ("TEE", "OLET") and comp.centre_points:
                centre = proj.page(comp.centre_points[0].xyz)
                canvas.dashed(e0, centre, width=0.3)
                canvas.dashed(centre, e1, width=0.3)
                for bp in (*comp.branch1_points, *comp.branch2_points):
                    canvas.dashed(centre, proj.page(bp.xyz), width=0.3)
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

    A set-on olet record carries only its CENTRE-POINT (on the header
    centreline) and its branch point -- it has no end points of its own, and
    it lands mid-pipe so the connectivity graph does not join it to the
    header either. To draw the fitting *tangent to the pipeline*, as a
    fabricator expects, we have to go and find the run its centre lies on.

    Returns the run's own end points rather than a direction, so the caller
    can project both. The header is *drawn* as the straight page segment
    between those two projections, so a direction taken from them is exactly
    parallel to the line on the sheet -- which a direction sampled a short
    way along the axis is not, once the not-to-scale warp has had its way.
    """
    comps = doc.real_components()
    tol = geometry.model_tolerance(doc)
    out: dict[int, tuple[Vec3, Vec3]] = {}
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
            out[id(comp)] = (a, b)
            break
    return out


_RUN_AXIS_CLASSES = {"PIPE", "FLANGE", "VALVE"}


def _run_axis_at(comp: Component, at: Vec3, tol: float
                 ) -> tuple[Vec3, Vec3] | None:
    """Two points on ``comp``'s centreline through the port ``at``.

    Straight-bodied components hand back their own two ends. A fitting has to
    say which of its legs the caller is standing on: an elbow's end-to-end
    chord is the diagonal across the bend, and a tee's is its run whichever
    port you arrived at, so taking either blindly gives a direction the pipe
    never goes. Returns None when ``at`` is not on this component, or when
    the leg through it is too short to give a direction.
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

    An eccentric reducer's two weld ends sit on two *different* centrelines,
    offset across the run by half the difference in outside diameters. The
    chord between them is therefore tilted off the pipe axis -- and a symbol
    built on that chord has no edge parallel to the pipeline, which is the
    one thing an eccentric reducer has to show. So the axis is taken from a
    neighbour instead.

    Returns the neighbour's own end points, like ``_olet_header_points``, so
    the caller can project both: the direction between two projections is
    exactly parallel to the line drawn on the sheet, which a direction
    sampled in model space is not once the not-to-scale warp has had its way.

    Reducers weld straight onto elbows and tees often enough that those count
    too, via ``_run_axis_at`` -- but on the leg the reducer is actually on,
    never the fitting's end-to-end chord. The neighbour on the *large* end is
    preferred, since the body is drawn from that end, but either does: both
    runs are parallel.
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
    # 出口跟随符号：按比例=实际端距投影（旧行为）；示意=等比自然
    # 长度（与 _draw_symbol_2pt 同款规则），小端管线落符号真实出口
    # schematic_scale：链式整图缩放时符号同倍缩小，出口偏移也同倍
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

    scale 是整图自适应因子：首轮 1.0，超幅时按缩放比重摆一轮，
    管件尺寸不动、只有直管档位收缩，尽量保住固定尺寸语义。
    """
    if world_len < 500.0:
        b = 8.0
    elif world_len < 1500.0:
        b = 15.0
    elif world_len < 4000.0:
        b = 24.0
    else:
        b = 34.0
    return max(4.0, b * scale)


_chain_debug: list[str] = []   # 链式布局调试日志（临时，验证后删）


def _chain_layout(proj: "Projector", doc: PcfDocument,
                  comps: list[tuple[Component, str | None]],
                  draw_region: tuple,
                  pipe_scale: float = 1.0,
                  ) -> tuple[dict[Vec3, Vec2], str | None, float] | None:
    """搭积木链式布局（示意模式专用）：位置由拓扑走链决定。

    管件固定尺寸（符号自然长度；弯头/三通腿 _SCHEMATIC_LEG_MM）、
    直管按真实长度单调分档，逐件拼接：上一件出口面 = 下一件入口面，
    永不断开；每段方向 = 该管段世界轴向的无 warp 投影（轴测 30° 族），
    弯头按模型转角转。坐标不参与图面定位，只提供方向；尺寸/坐标文字
    仍取模型真值。多连通组竖排打包；整体超幅时等比缩放并返回警告。

    返回 (全点图面平移表, 警告|None)；与 point_offsets 通道一致。
    失败返回 None，调用方回退 warp 布局。
    """
    from collections import deque

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
        # 直通符号件（阀/法兰/变径…）的走向取主轴投影：偏心变径两端点
        # 带垂直于管轴的偏心量，直接投影弦线会把符号画歪。主轴占比 <90%
        # 的是真斜管（如 45° 平面走向），保留弦线投影
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
        # OLET/TEE-STUB 的主管贴附点存在 CENTRE-POINT（connection_points
        # 不收）：不收进来走链就走不过主管——支管台两侧的管子会各成一组，
        # 后到的直管只能横跨两组之间的间隙（斜弦断桥的根因）
        if classify(c.type, c.skey) == "OLET":
            pts += [p.xyz for p in c.centre_points]
        return pts

    # 邻接边 + 共享点（双向都要建：BFS 从任一侧都能走过去，否则图被
    # 砍成单向、走不到的组件会被当成独立组另摆——接头脱开的根因）
    edges: dict[int, list[tuple[int, Vec3]]] = {}
    for i, ci in enumerate(real):
        pts_i = _cps(ci)
        for j in graph.adjacency.get(i, ()):
            if any(k == j for k, _s in edges.get(i, ())):
                continue
            for cp in pts_i:
                if any(geometry.dist(cp, q) <= tol for q in _cps(real[j])):
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
        others = [c for c in _cps(comp) if geometry.dist(c, s) > tol]
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
            # 双端符号件（阀/法兰/变径/短接/仪表…）：固定自然长度直通；
            # 多余接点按腿长展开（多口件兜底）。垫片不画符号，但要占住
            # 一小段图面长度——否则法兰面与阀面/盲板贴死，垫片空隙丢失
            natural = _natural_run_length(comp.skey)
            if natural is None:
                natural = 1.0 if cls == "GASKET" else 6.0
            if not others:
                return
            far = max(others, key=lambda c: geometry.dist(s, c))
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
                # 的那一侧整组，闭合差整体搬走，组内几何不动
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
        # 跨组共点：接点已在已摆结构上（如支座坐在弯头端点）——锚在
        # 原位置接着摆，不另起原点、不参与竖排打包，否则共点被撕成
        # 两个图面位置（弯头腿被拽长的根因之一）
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
        if anchor0 is None:
            pos[s] = (0.0, 0.0)
        frontier: list[tuple[Vec3, Vec2]] = []
        _place(start, comp, s, pos[s], din, frontier)
        # 起步件近端（s 侧）的邻件也要走链：frontier 只带远端出口，近端
        # 邻件落不到链上会各自另起组——盲法兰被甩出阀组、垫片链断开的
        # 根因。行进方向取反（从 s 向组外走）
        frontier.append((s, _v2_scale(din, -1.0)))
        # 非管优先、直管靠后（deferred）：环路尽量由直管闭合——直管是
        # 弹性件，横跨现状间隙即可；无管可闭合时由 _absorb_closures
        # 把闭合差摊进闭合侧最顺向的那根直管。
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
            if bi is not None and keys:
                box = group_boxes[bi]
                box[4].extend(keys)
                for k in keys:
                    key_box[k] = bi
                xs = [pos[k][0] for k in box[4]]
                ys = [pos[k][1] for k in box[4]]
                box[0], box[1] = min(xs), min(ys)
                box[2], box[3] = max(xs), max(ys)
            continue
        if keys:
            xs = [pos[k][0] for k in keys]
            ys = [pos[k][1] for k in keys]
            key_box.update({k: len(group_boxes) for k in keys})
            group_boxes.append([min(xs), min(ys), max(xs), max(ys), keys])

    if not pos:
        return None

    # 中段贴附件（支座/仪表等挂在管壁中段、与任何接点都不重合的件）
    # 自成一个松散组：先标记并排除出竖排打包，打包后再按所挂直管的
    # 最终图面位置整组平移吸附（打包前吸附会被主结构的平移撕掉）
    pipe_segs: list[tuple[Vec3, Vec3]] = []
    for c in real:
        if classify(c.type, c.skey) != "PIPE":
            continue
        ep = [p.xyz for p in c.connection_points()]
        if len(ep) >= 2 and ep[0] in pos and ep[1] in pos:
            pipe_segs.append((ep[0], ep[1]))
    to_snap: list[tuple[list[Vec3], Vec3, Vec3, Vec3]] = []  # keys,k,a,b
    kept_boxes = []
    for box in group_boxes:
        keys = box[4]
        hit = None
        for k in keys:
            for a, b in pipe_segs:
                lab = geometry.dist(a, b)
                if lab < 1e-6 or min(geometry.dist(k, a),
                                     geometry.dist(k, b)) <= tol:
                    continue
                if geometry.point_segment_distance(k, a, b) <= tol:
                    hit = (k, a, b)
                    break
            if hit:
                break
        if hit:
            to_snap.append((keys, hit[0], hit[1], hit[2]))
        else:
            kept_boxes.append(box)
    group_boxes = kept_boxes

    # 组竖排打包（y 向下为正，组间留 14mm），再整体居中；
    # 超幅只缩不放（管件固定尺寸是布局语义，放大失真）
    cursor_y = 0.0
    for gx0, gy0, gx1, gy1, keys in group_boxes:
        dx, dy = -gx0, cursor_y - gy0
        for k in keys:
            pos[k] = (pos[k][0] + dx, pos[k][1] + dy)
        cursor_y += (gy1 - gy0) + 14.0
    # 贴附件吸附（打包后执行）：世界线段上的比例插值到图面线段
    for keys, k, a, b in to_snap:
        lab = geometry.dist(a, b)
        t = geometry.dist(a, k) / lab           # 贴壁点近共线，线性插值
        tgt = _v2_add(pos[a], _v2_scale(_v2_sub(pos[b], pos[a]), t))
        d = _v2_sub(tgt, pos[k])
        for kk in keys:
            pos[kk] = _v2_add(pos[kk], d)
    # 未覆盖点（孤立接点/中心点）吸附到世界最近的已布点——必须在打包
    # 之后做，否则吸附点留在锚点打包前的旧位置上（飞出件的根因）
    mapped = list(pos)
    for comp in real:
        for p in (*comp.connection_points(), *comp.centre_points):
            if p.xyz in pos:
                continue
            near = min(mapped, key=lambda k: geometry.dist(k, p.xyz))
            pos[p.xyz] = pos[near]
    xs = [p[0] for p in pos.values()]
    ys = [p[1] for p in pos.values()]
    bx0, by0, bx1, by1 = min(xs), min(ys), max(xs), max(ys)
    bw = max(bx1 - bx0, 1e-6)
    bh = max(by1 - by0, 1e-6)
    rx, ry, rw, rh = draw_region
    m = 0.82
    s = min(1.0, rw * m / bw, rh * m / bh)
    cx_t, cy_t = rx + rw / 2.0, ry + rh / 2.0
    bcx, bcy = (bx0 + bx1) / 2.0, (by0 + by1) / 2.0
    warn = None
    if s < 0.999:
        warn = f"chain layout scaled to {s:.2f} to fit the sheet"
    out: dict[Vec3, Vec2] = {}
    for k, p in pos.items():
        tgt = ((p[0] - bcx) * s + cx_t, (p[1] - bcy) * s + cy_t)
        base = _page(k)
        d = (tgt[0] - base[0], tgt[1] - base[1])
        if _v2_len(d) > 1e-6:
            out[k] = d
    return out, warn, s


def _symbol_face_offsets(proj: "Projector", doc: PcfDocument,
                         comps: list[tuple[Component, str | None]],
                         reducer_runs: dict[int, tuple[Vec3, Vec3]],
                         flange_face_pts: dict[int, Vec3] | None,
                         ) -> dict[Vec3, Vec2]:
    """等比符号端面带动连接点的图面平移（端点/尺寸锚随符号走）。

    双端符号（VALVE/FLANGE/…）：符号恒等比后 EXIT 落在自然长度处而非
    模型端点——把两侧模型连接点的投影平移到符号实际端面，管端、焊点、
    尺寸 goalpost 投影时自动落在符号面上（通道与偏心变径小端一致：
    point_offsets，点对点精确匹配）。同一模型点被多个管件注册（直焊
    法兰-阀门等）时取均值——两边画到同一点保持相连。

    弯头/三通**不注册**（2026-11-13 试过固定腿长+弦线交点两版均造成
    斜向弯头与管线折角/脱开，用户定稿回退）：腿画到模型端点，角度
    随管线（正确），两侧长度不强制等长，短的由 _grown_leg_end 补长。

    法兰端点重合的按接带 top 定 START 面（含垫片空隙）；非重合法兰按
    配合面翻转让 START 对准面。偏心变径已由 _eccentric_small_side_offsets
    单独接管，此处跳过以免双算。与绘制端（_draw_component 符号调度）
    同一套规则的预计算镜像；两侧规则改动必须同步。
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
    """Coordinate-value pairs the not-to-scale warp must leave alone.

    An eccentric reducer's two weld ends are on centrelines offset across the
    run by half the difference in outside diameters -- half an inch to an
    inch on the sizes these drawings carry. That is fitting geometry, not a
    leg of pipe, and the warp has no business stretching it to its minimum
    visible span: doing so bends the pipeline by seven times the real offset
    at every reducer, and leaves the reducer symbol unable to show the flat
    side that the offset *is*.

    Returns one set per axis of ``(lo, hi)`` **intervals**, rounded to match
    the warp's own quantisation. Every gap inside such an interval is left at
    true proportion, not merely a gap that spans it exactly. The old
    exact-pair rule failed on any sheet where something else sat at a
    coordinate between the reducer's two ends -- the rest of the spool
    usually does, since the reducer straddles the run centreline -- and on
    those sheets the offset was inflated to two full ``nts_min_mm`` steps.
    That is what tore the reducer symbol off the small side of its own run.
    An interval is only ever the width of one fitting's eccentricity, so
    nothing a fabricator needs to see can hide inside it.
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

    每个件的坐标区间在 warp 中固定映射为 max(符号自然图面长度,
    min_mm)——与旧版"每个间隙 ≥nts_min"给零件的 8mm 跨度一致，但上限
    封顶在示意尺寸：不再被拉长（FLWN 曾被组斜率拉到 19mm）。非比例
    伸缩只由直管吸收。跨轴件（弯头/三通）不走此机制：它们的腿即管段
    空间，由间隙 clamp 与腿外延处理；偏心变径由 keep 保护真实偏心量。
    返回 [axis] -> [(lo, hi, target_mm)]。
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
            # 垫片模型厚度（~1.5mm）在示意比例下不可见：warp 固定为
            # 0.6mm 图面间隙——法兰面与配合件之间露出垫片位，法兰
            # EXIT 仍精确落焊端，无需平移符号
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

    An olet record always carries a BRANCH1-POINT, whether or not the branch
    was ever piped: it says where the outlet *faces*, not that anything is on
    it. Running the branch centreline out to that point unconditionally
    litters the sheet with long bare stubs leaving fittings that, on this
    spool, are spare or capped -- and a fabricator reads a centreline as
    "pipe continues here".

    "Connected" is any connection point of another real component landing on
    the olet's branch point, using the same tolerance as the connectivity
    graph. Components on *other* sheets do not count: this is a per-sheet
    document, and if the branch is piped elsewhere the stub is not this
    drawing's to show.
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
                    olet_header_pts: dict[int, Vec3] | None = None,
                    olet_branch_connected: set[int] | None = None,
                    reducer_runs: dict[int, tuple[Vec3, Vec3]] | None = None,
                    ) -> Vec2 | None:
    cls = classify(comp.type, comp.skey)
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

    # Flanges arrive misclassified often enough (weld-neck flanges exported
    # as COUPLING records) that the symbol dispatch trusts the description
    # too, not just the record type. Material accounting still uses
    # classify() alone, so the BOM is untouched by this.
    # 符号库优先（统一入口）：SKEY 查到几何的双端组件先画官方符号——
    # START(首点)→eps[0]、EXIT(收笔点)→eps[1] 双点仿射对齐。
    # 手写分支仅作无符号时的回退；支撑(SUPPORT)走自己的方向模式分支。
    # 注意：必须先于下方 flange_face 特殊分支，法兰符号才优先生效。
    # 爆破片（SAFETY-DISC，SKEY=RD）在 classify 里归 GASKET，但官方
    # 符号必须画——只让普通垫片走"不画、项号并父标"分支。
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
    if (cls in ("VALVE", "FLANGE", "ELBOW", "REDUCER", "COUPLING", "FILTER",
                "INSTRUMENT", "MULTIPORT", "OLET") or sym_first) \
            and not skip_sym and len(eps) >= 2:
        ep3 = comp.end_points
        run3 = (tuple(ep3[1].xyz[i] - ep3[0].xyz[i] for i in range(3))
                if len(ep3) >= 2 else None)
        sp0, sp1 = eps[0], eps[1]
        best = None            # (距离, 焊点=最近管端, 管段远端)
        if _v2_len(_v2_sub(eps[1], eps[0])) < 0.5:
            # 两端点几乎重合（平焊法兰 face-to-face≈0）：轴向取最近管段。
            # 方向感必须指向该管段的"远端"（从接头沿管道离开）——直接用管段
            # 方向向量的正负是随机的，会把符号翻进阀门里。
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
                # 登记焊端：管端与焊点都吸到 EXIT 锚定的这一点。
                # 不做垫片间隙平移——平移会把 EXIT 推过焊端 0.6mm，
                # 符号与管端/焊点错开（实测偏差就是它）
                _FLANGE_WELDS.append(sp1)
        axis_override = None
        if cls == "REDUCER":
            # 偏心大小头：管轴取主管走向（原项目 reducer_runs，两端同轴
            # 导出的场合），符号 X→管轴，偏心量按符号 Y 保留——不能沿
            # p0→p1 对角线对齐（会扭斜端面）
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
        if _draw_symbol_2pt(canvas, comp.skey, sp0, sp1,
                            proj=proj, run3=run3, axis=axis_override,
                            wp=ep3[0].xyz if ep3 else None,
                            drop_isolated=(cls == "REDUCER"
                                           and not (b1s or b2s)),
                            bore_ratio=bore_ratio,
                            proportional=proj.proportional,
                            dashed=cont_copy,
                            width=0.3 if cont_copy else 0.4,
                            ecc_dir=ecc_dir, **op_kw, **prop_kw):
            return cps[0] if cps else _v2_mid(eps[0], eps[1])

    face_pt = (flange_face_pts or {}).get(id(comp))
    if face_pt is not None and len(eps) >= 2:
        face = proj.page(face_pt)
        face_first = _v2_len(_v2_sub(face, eps[0])) <= _v2_len(_v2_sub(face, eps[1]))
        _draw_flange(canvas, proj, eps[0], eps[1], comp, face_first, flange_bar_mm)
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
            chain = [_grown_leg_end(eps[0], corners[0], natural=leg_mm,
                                    own_id=id(comp)),
                     *corners,
                     _grown_leg_end(eps[1], corners[-1], natural=leg_mm,
                                    own_id=id(comp))]
            if cont_copy:
                for p, q in zip(chain, chain[1:]):
                    canvas.dashed(p, q, width=0.3)
            else:
                canvas.polyline(chain, width=0.55)
            branch = b1s[0] if b1s else None
            if branch is not None and corners:
                b_end = _grown_leg_end(branch, corners[0], natural=leg_mm,
                                       own_id=id(comp))
                if cont_copy:
                    canvas.dashed(corners[0], b_end, width=0.3)
                else:
                    canvas.line(corners[0], b_end, width=0.55)
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
            stroke(_grown_leg_end(eps[0], c0, natural=leg_mm,
                                  own_id=id(comp)),
                   _grown_leg_end(eps[1], c0, natural=leg_mm,
                                  own_id=id(comp)),
                   width=0.55)
            anchor = c0
        if centre is not None:
            for branch in (*b1s, *b2s):
                stroke(centre, _grown_leg_end(branch, centre,
                                              natural=leg_mm,
                                              own_id=id(comp)),
                       width=0.55)
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
            # SKEY 符号库优先（WT 支管台等）：主管中心→支管端 两接点对齐，
            # 跨向仍取世界轴投影；支管未连接时保持 _draw_olet 原行为
            if connected:
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
            _draw_valve(canvas, eps, mid, comp)
            return mid

    elif cls == "FLANGE":
        # Only reached when the face could not be resolved (a single-ended
        # or point-only flange record).
        if len(eps) >= 2:
            _draw_flange(canvas, proj, eps[0], eps[1], comp, True, flange_bar_mm)
            return _v2_mid(eps[0], eps[1])
        if len(eps) == 1:
            _draw_bar(canvas, eps[0], (1.0, 0.0), flange_bar_mm / 2, width=0.6)
            return eps[0]

    elif cls == "REDUCER":
        if len(eps) >= 2:
            _draw_reducer(canvas, proj, comp, eps,
                          (reducer_runs or {}).get(id(comp)))
            if b1s:
                canvas.line(_v2_mid(eps[0], eps[1]), b1s[0], width=0.55)
            return _v2_mid(eps[0], eps[1])

    elif cls == "GASKET":
        # 垫片不另画符号或球标，其项号合并到父件球标。
        if len(eps) >= 2:
            return None

    elif cls == "CAP":
        if eps:
            end = eps[0]
            axis = _v2_norm(_v2_sub(eps[0], eps[1])) if len(eps) > 1 else (1.0, 0.0)
            _draw_cap(canvas, end, axis)
            return end

    elif cls == "COUPLING":
        if len(eps) >= 2:
            # Measure the gap in the MODEL, not on the page. This used to test
            # the projected page span, which the not-to-scale warp exists to
            # inflate: a 1/16" shop-weld gap got stretched past the 3 mm
            # threshold, so the coupling escaped suppression and stamped its
            # body line and bar across the weld dot -- the "line on top of a
            # weld dot" that made every weld read the same.
            gap_in = (comp.end_points[0].distance_to(comp.end_points[1])
                      * coord_to_inch_factor(doc.units_coords))
            if gap_in <= GAP_COUPLING_MAX_INCHES and is_weld_coupling(
                    comp.type, comp.description, comp.skey):
                # Weld couplings surface through weld derivation, not as a
                # symbol -- but the run still has to read as continuous. The
                # warp can stretch a 1/16" gap to most of a centimetre, and
                # leaving it unstroked breaks the centreline at every weld.
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
            _draw_box_on_axis(canvas, eps[0], eps[1])
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
        # Nothing: the continuation pass owns these points. It has to, because
        # what belongs on the sheet is the name of the NEXT DRAWING, and this
        # record knows only which pipeline carries on -- which is a different
        # thing as soon as a sheet holds more than one spool. (Unreachable
        # from render_pdf, whose component list is real_components().)
        return None

    elif cls == "MESSAGE":
        if cos:
            message = comp.get("TEXT") or comp.get("MESSAGE") or ""
            if message:
                canvas.text(cos[0], message, size=5.5)
        return None

    elif cls == "WELD":
        # The weld pass draws the joint symbol. Stroking the record's own
        # body here as well put a stray line under every weld dot, which is
        # what made socket welds read as "a line on top of a dot".
        if len(eps) >= 2:
            return _v2_mid(eps[0], eps[1])
        return eps[0] if eps else (cos[0] if cos else None)

    else:
        if len(eps) >= 3 and _v2_len(_v2_sub(eps[1], eps[0])) > 1e-6:
            # 未知多口件（≥3 端点）：画 NC 式箱体（Isogen NC 符号——矩形
            # 箱体，两端接点 + 上下 tap）。箱体沿前两个端点，其余端点
            # （第 3 个起，不限个数）从箱体中心各拉一条 tap 线到真实坐标。
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
    picture. Choosing the world axis most nearly perpendicular to the run
    instead puts the bar on an iso plane at 30 degrees, the way a flange is
    drawn by hand. Elevation is tried first so a flange on any horizontal
    run draws as an upright bar, which is the conventional reading; ties
    after that resolve in a fixed order, so the sheet is deterministic.

    The tie band has to be **wide**. Whenever a run lies along one iso axis
    -- the common case, not a corner case -- the other two are both exactly
    30 degrees off perpendicular and score identically, so the winner is
    decided entirely by the tie-break. At the old 1e-6 the two halves of a
    bolted flange pair, whose runs point at each other and differ only by
    accumulated projection noise, scored 0.866010 against 0.866041 and chose
    *different* axes: the two face bars came out 60 degrees apart and drew a
    great X across the joint. Real distinctions here are 0.866 against 0.5 or
    0, so a band of 1e-3 cannot merge two axes that are genuinely different.
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


# How far back from the gasket face the black face bar stands, in page mm.
# Two flanges bolted together are coincident in the model, so their bars
# landed on the same point and the joint drew as a *single* bar -- one
# flange, not two faces with a gasket between them. Setting each bar back by
# this much opens a black | white | black gap that reads as the gasket.
_FLANGE_GASKET_GAP_MM = 0.55


def _draw_flange(canvas: _Canvas, proj: Projector, e0: Vec2, e1: Vec2,
                 comp: Component, face_first: bool, bar_mm: float) -> None:
    """Flange body plus one long iso-plane bar just behind the gasket face.

    The bar goes on the *flanged* end only -- the welded end is just a butt
    weld like any other and is already marked by its weld dot. It is drawn
    about twice the length of a standard dimension offset so a flanged
    connection is unmistakable at a glance on a full sheet.

    The body stops ``_FLANGE_GASKET_GAP_MM`` short of the face and the bar
    stands there, with a white knockout painted across the face itself. On a
    bolted pair -- where both faces are the same point -- that draws two black
    bars a millimetre apart with clear white between them, which is how a
    fabricator reads a gasketed joint. The knockout is unrecorded, so it never
    enters the collision geometry that label placement scores against, and it
    is narrower than the setback so it cannot eat the other half's bar.
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

# 非重合法兰的焊端锚点（页坐标）：焊点圆点吸附用，每张图重置
_FLANGE_WELDS: list[Vec2] = []

# 非管件占位点（页坐标, 组件 id）：符号生长避让用，每张图重置。
# PIPE/WELD/BOLT/GASKET 不登记——它们与零件同轴相邻，会把生长空间全锁死。
_NONPIPE_PAGE_PTS: list[tuple[Vec2, int]] = []

_SCHEMATIC_LEG_MM = 4.5   # 弯头/三通腿外延目标（page mm）


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
    pp = (lambda p: proj.page(p, apply_offset=False)) if pre else proj.page
    e0 = pp(eps3[0].xyz)
    e1 = pp(eps3[1].xyz)
    if _v2_len(_v2_sub(e1, e0)) >= 0.5:
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
        return None
    if best is None or _v2_len(_v2_sub(best[2], e0)) <= 1.0:
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
    return top, _v2_add(top, _v2_scale(uu, span)), uu


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
    sy = height_mm / max(perp, 1e-6)
    if proj is not None and getattr(proj, "schematic_scale", 1.0) != 1.0:
        # 链式布局整图缩放：符号尺寸与位置表同乘，保持搭积木间距
        sy *= proj.schematic_scale
    if proportional:
        # 按比例模式：符号铺满实际端距（sx=端距/跨度），短端距生长、
        # 超限钳制——真实比例出图时符号长度跟端点走（旧行为）。
        sx = sgn * d / span
        if axis is None:
            # 短端距向两端生长到自然长度（sx=sy 等比），生长量受邻件
            # 空间约束（room0/room1）：长不出来就保持原样。
            xs_ = [c for seg in sym for c in (seg[0], seg[2])]
            natural = (max(xs_) - min(xs_)) * sy if xs_ else span * sy
            if not no_grow and d < natural:
                extra = natural - d
                r0 = 1e9 if room0 is None else room0
                r1 = 1e9 if room1 is None else room1
                total = min(extra, r0 + r1)
                g0 = min(r0, total / 2.0)
                g1 = min(r1, total - g0)
                g0 += min(r0 - g0, total - g0 - g1)
                if g0 > 0:
                    p0 = _v2_sub(p0, _v2_scale(u, g0))
                if g1 > 0:
                    p1 = _v2_add(p1, _v2_scale(u, g1))
                d += g0 + g1
                sx = sgn * d / span
            ax_ = abs(sx)
            if ax_ > sy * 2.0:
                sx = math.copysign(sy * 2.0, sx)
    else:
        # 不按比例（示意）：符号恒等比，sx 恒等于 sy（镜像号除外）——
        # 轴向不单独拉伸，也不对个别符号加大/缩小。START 锚 p0 沿管轴
        # 伸出自然长度；与端点的差值由 _symbol_face_offsets 带动连接点
        # 消化（管端/焊点/尺寸锚随符号端面走，warp 目标同源等比锚距）。
        sx = sgn * sy
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
        # 孤立腿 = 与其他任何腿都不共端点的腿：轮廓腿围成闭环共端点，
        # tap 标记线浮在体内（端点落在轮廓边线内部）——按 ISOGEN 约定，
        # 组件没有 BRANCH-POINT 时不画
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
        # 阀柄叠加：柄是独立 skey 的符号，厂商坐标系以阀中心连接点为
        # START、阀杆沿符号 +X 伸出（01SP=T 形柄、03SP=杠杆、04SP=手轮、
        # 10SP=疏水器浮子…，独立块见 symbols/ALIAS.asc 的 N**SP）。柄
        # 符号已归一到 START=(0,0)：锚到阀中心连接点（收笔点中程、两端
        # 点页面中点），符号 +X 旋到 SPINDLE-DIRECTION 图面投影方向；
        # 缩放沿用阀体高度标尺 sy——柄与阀体同一符号坐标系，等比不失真。
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


def _draw_valve(canvas: _Canvas, eps: list[Vec2], mid: Vec2, comp: Component) -> None:
    # 回退：程序化领结 + 子类型装饰（SKEY 符号库命中时不会走到这里，
    # 统一入口见 _draw_component 顶部的 _draw_symbol_2pt）
    half_w = 1.9
    for end in eps[:4]:
        axis = _v2_norm(_v2_sub(mid, end))
        perp = _v2_perp(axis)
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
                  eps: list[Vec2], run: tuple[Vec3, Vec3] | None = None) -> None:
    """Trapezoid spanning the reducer's own two weld ends, on an iso plane.

    **Both ends of the body sit on the reducer's own projected weld points.**
    The body used to be built out along a neighbour's run axis to a point
    ``big + axis * reach`` -- the small end dropped onto that axis -- so
    whenever the small end was genuinely off the axis the symbol's tip landed
    somewhere the pipeline was not, and the fitting drew as a thin wedge
    floating clear of the run it reduces. That is the bad assumption: the
    neighbour's axis says which way the pipeline *points*, and nothing about
    where this fitting *ends*.

    ``run`` still decides the direction the body is squared to, because the
    chord between the two ends will not do: on an eccentric it leans across
    the run by the centreline offset, and end faces square to that chord lean
    with it -- which is how these came out as rhomboids lying across the pipe.

    The cross direction is a **projected world axis**, the way a flange bar is
    (`_iso_cross_dir`), never a screen perpendicular. A screen perpendicular
    lies on no isometric plane, so the fitting read as a flat mark pasted onto
    the picture. On an eccentric it is the projection of the measured flat
    direction itself, so a reducer that is flat on the bottom draws with its
    flat on the elevation plane and its taper on top -- which is the whole
    thing an eccentric reducer exists to show.

    The flat side is **schematic**, like every other symbol on the sheet: one
    edge parallel to the run with both ends on it, and the small end hugging
    it so only the far side slopes. It is not the offset drawn to scale --
    at true proportion a 1" offset is a third of a millimetre on the page and
    the fitting would read as concentric. Which side it falls on *is* taken
    from the model (see ``_reducer_flat_offset``). The centreline offset
    itself stays at true proportion; ``_eccentric_offset_pairs`` keeps the
    not-to-scale warp from inflating it into a dog-leg.
    """
    pts = comp.end_points
    large_first = (pts[0].bore or 0.0) >= (pts[-1].bore or 0.0)
    big_pt, small_pt = (pts[0], pts[-1]) if large_first else (pts[-1], pts[0])
    big, small = (eps[0], eps[1]) if large_first else (eps[1], eps[0])
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
            # How far the centreline moves across the run *on the page*. The
            # body has to be at least that deep or the small end falls off
            # the bottom of it -- which drew the two long edges crossing in a
            # bowtie. Growing the big end by exactly the step is also what
            # makes the flat come out parallel to the run: on a real fitting
            # the offset *is* half the difference in diameters.
            # `across` already points at the flat, so the flat is always +1.
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


def _draw_cap(canvas: _Canvas, end: Vec2, axis_into_pipe: Vec2) -> None:
    perp = _v2_perp(axis_into_pipe)
    r = 1.5
    pts = []
    for step in range(9):
        theta = math.pi * step / 8
        along = -math.sin(theta) * r * 0.9
        across = math.cos(theta) * r
        pts.append(_v2_add(end, _v2_add(_v2_scale(perp, across),
                                        _v2_scale(axis_into_pipe, along))))
    canvas.polyline(pts, width=0.4)


def _draw_box_on_axis(canvas: _Canvas, e0: Vec2, e1: Vec2) -> None:
    axis = _v2_norm(_v2_sub(e1, e0))
    perp = _v2_perp(axis)
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
# Dimensions are drawn with *isometric skew*: the extension (leader) bars and
# the dimension line always run along projected world-axis directions, so the
# whole callout lies on an isometric plane instead of being screen-
# perpendicular. Point pairs that are not axis-aligned get an *offset box*:
# a dashed rectangle (two axes differ) or cuboid (all three differ) whose
# axis-aligned legs are dimensioned individually, each on the side that
# collides least with geometry, labels, and earlier dimensions.

_AXIS_VECS: tuple[Vec3, ...] = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
_DIM_FONT = _SHEET_FONT
_DIM_CAP = _cap_height_mm(_DIM_FONT)

# Arrowhead: length tip-to-base and half the width across the barbs.
_ARROW_MM = 1.4
_ARROW_HALF_MM = 0.45
# A dimension line has to be at least four arrowheads long before both arrows
# will fit between its goalposts and still read as a double arrow. Under that
# the arrows go outside, opposing each other -- style "outside".
_MIN_ARROWS_INSIDE_MM = 4 * _ARROW_MM
# Clear space each side of a value that breaks the dimension line.
_DIM_TEXT_GAP_MM = 0.9
# How far a leadered value stands off the thing it points at.
_DIM_LEADER_MM = 4.0
# How far the arrow tails reach past the goalposts in the "outside" style.
_OUTSIDE_TAIL_MM = 3.6

# Whether an offset leg counts as "spanned" is decided on the page, not in
# the model: a leg must draw at least this many mm AND be more than a
# modelling artifact in the world. Judging it against the model diagonal (as
# the readability threshold for whole dimensions does) made offset boxes
# vanish on large spools -- a 3" roll in a 40' line is under 1% of the
# diagonal, so the box collapsed and the callout degenerated into a single
# skewed line. Page-space is the honest test, and it accounts for the
# not-to-scale warp, which stretches exactly those short legs into view.
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
    # |orient| / ab_len is the point-to-line distance: both endpoints of cd
    # within 0.3mm of ab's line means near-collinear; overlap iff cd's
    # projection onto ab intersects the (0, 1) parameter span.
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
        # 偏心变径的大小端在平行中心线上。符号布局会给小端侧额外横向
        # 平移；尺寸不能直接连接两个实际图面端点，否则仍沿偏心斜弦画。
        # 以大端为锚，把小端图面位置投影回原始管轴：去掉横向位移，
        # 保留轴向位移（示意模式下仍与符号的实际图面长度一致）。
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

    # Stand the boss clear of the weld dot at its root: the weld pass stamps
    # a 0.9 mm filled dot on the header centreline here, and a box sharing
    # that edge is half swallowed by it.
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

    A butt weld is a filled dot -- the whole joint is the circumference, and
    that is what a fabricator expects. A **socket** weld is not: the pipe,
    nipple or swage is inserted into a socket and filleted at the shoulder,
    and the convention is a square bracket sitting at the joint, opening
    towards the inserted member. Everything used to draw as the same dot, so
    the two joint types were indistinguishable on the sheet even though the
    weld table counted them separately.

    Drawn with ``record=False``: weld symbols are small and sit on the run,
    and letting them into ``canvas.segments`` would start perturbing
    dimension placement on drawings that are otherwise unchanged.
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
#   01HG/空 = 默认承重支撑, ANCH=固定架, GUID=导向架, SPRG=弹簧支吊架,
#   HANG=吊架（DUCK/SKID/SLVE 等暂按承重画法）。
# .ASC 符号方向模式（501 尾参）：
#   0 = 顺管道方向（符号 X 轴对齐管道走向，如 01HG 两线平行于管道）
#   2 = 固定方向（SUPPORT-DIRECTION UP/DOWN 上下翻转）
#   3 = 三轴六向（如 ANCH 双箭头，方向由 SUPPORT-DIRECTION 的世界轴决定）
_SUPPORT_WORLD_DIRS = {
    "UP": (0, 0, 1), "DOWN": (0, 0, -1),
    "NORTH": (0, 1, 0), "N": (0, 1, 0), "Y+": (0, 1, 0), "Y": (0, 1, 0),
    "SOUTH": (0, -1, 0), "S": (0, -1, 0), "Y-": (0, -1, 0),
    "EAST": (1, 0, 0), "E": (1, 0, 0), "X+": (1, 0, 0), "X": (1, 0, 0),
    "WEST": (-1, 0, 0), "W": (-1, 0, 0), "X-": (-1, 0, 0),
    "Z+": (0, 0, 1), "Z": (0, 0, 1), "Z-": (0, 0, -1),
}
_PIPE_RUNS_CACHE: dict[int, list] = {}


def _nearest_pipe_run_dir(proj: "Projector", doc, at: Vec2):
    """支架点附近管道走向；返回 (页空间方向, 世界空间方向)，找不到返回 None。

    世界方向用于推导跨管的世界轴（等轴测里跨管偏移沿世界轴的 30° 投影，
    而不是页面上管轴的垂线）。
    """
    runs = _PIPE_RUNS_CACHE.get(id(proj))
    if runs is None:
        runs = []
        for comp in doc.real_components():
            if classify(comp.type, comp.skey) == "PIPE" and len(comp.end_points) >= 2:
                w0 = comp.end_points[0].xyz
                w1 = comp.end_points[1].xyz
                a = proj.page(w0)
                b = proj.page(w1)
                d = _v2_sub(b, a)
                if _v2_len(d) > 1e-6:
                    runs.append((a, d, (w1[0]-w0[0], w1[1]-w0[1], w1[2]-w0[2])))
        _PIPE_RUNS_CACHE.clear()
        _PIPE_RUNS_CACHE[id(proj)] = runs
    best, best_dist = None, 1e18
    px, py = at
    for a, d, dw in runs:
        ax, ay = a
        dx, dy = d
        t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        dist = math.hypot(px - ax - t * dx, py - ay - t * dy)
        if dist < best_dist:
            best_dist, best = dist, (d, dw)
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
            run = _nearest_pipe_run_dir(proj, doc, at)
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


def _build_bom(doc: PcfDocument, comps: list[tuple[Component, str]]) -> dict:
    inch_factor = coord_to_inch_factor(doc.units_coords)
    materials = doc.materials_descriptions()
    rows: dict[tuple, dict] = {}
    order = 0
    for idx, (comp, _ref) in enumerate(comps):
        cls = _bom_class(comp)
        if cls in ("CONNECTION", "MESSAGE", "GAP"):
            continue
        if comp.material_excluded():
            continue  # 双写边界段的对侧副本（CONTINUATION 件及其副料）：画但不计料
        if is_weld_coupling(comp.type, comp.description, comp.skey):
            continue  # welds are counted in the weld table, not the BOM
        if cls == "WELD" and not comp.item_code:
            continue  # 焊缝记录无料号时不进 BOM（避免图面/材料表出现 "?" 气球）
        _cu = comp.type.upper().replace("_", "-")
        if _cu.startswith("FLOW-ARROW") or (
                (comp.skey or "").upper() == "FLOW" and cls == "OTHER"):
            continue  # 流向箭头是图面符号（SKEY FLOW 曾被 FL 前缀误判为法兰），不进 BOM
        # Resolved the same way the row's own description is, below: plenty of
        # exports name the part only in the MATERIALS table.
        desc = comp.description or materials.get(comp.item_code or "", "")
        if is_annotation_symbol(comp.type, desc, comp.skey):
            continue  # a slope arrow is a drawing mark, not material
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


# Whitespace a callout's *leader foot* keeps from a neighbouring fitting's
# anchor or a weld dot, in page mm. Inside this the leader reads as pointing
# at the neighbour rather than at the run.
_ANCHOR_CLEARANCE_MM = 3.0

# Where along a pipe its cut-piece callout may point, as fractions of the run.
# The midpoint is first so it still wins wherever it is clear, and today's
# sheets are unchanged; the quarter points are the escape route when a fitting
# is welded on at the middle.
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

    The single style used for **everything the shop has to identify** -- a
    fitting's item number and a cut piece's label alike. They are the same
    kind of mark (this is the part in front of you, here is its number in the
    table), so drawing them differently only asks the reader to learn two
    conventions for one idea.

    Cut labels used to be written bare against the centreline with no leader.
    That saved a line and cost clarity: with no leader, which run a label
    belongs to is inferred from proximity alone, and on a busy sheet -- or
    wherever the not-to-scale warp brings two runs close -- proximity is not
    an answer. Every callout now stands clear of the centreline and points.
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
        # What the foot must keep clear of: every *other* component's anchor
        # and every weld dot. The pipe's own anchor is excluded -- it is one
        # of the candidates, and a run must not be crowded out by itself.
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
    # 垫片/螺栓优先并入任一端实际相邻的法兰，顺序为“法兰/螺栓/垫片”；
    # 两端都没有法兰时不并入阀门、三通等父件，而单独出“螺栓/垫片”球标。
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


# Field *labels* in the title block stay a step below their values, which is
# the hierarchy the block already read with. Only the values are held to
# _SHEET_FONT -- they are the part anyone actually reads off the sheet.
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
                      x: float, y: float, w: float, h: float) -> None:
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
        fy = y + 4 + row * (h - 6) / 3
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
