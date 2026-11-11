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

Continuation callouts mark every point where the piping leaves the sheet:
an arrow along the run plus a block naming the **next rendered drawing**
and the hand-off point's plant coordinates. The drawing name is resolved
across the set of sheets being rendered together (see continuations.py) --
never taken from the file's own spool/pipeline reference, which names the
neighbouring piping rather than the neighbouring sheet.

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

import math
from dataclasses import dataclass, field
from datetime import date as _date
from itertools import permutations
from pathlib import Path

from . import geometry
from .continuations import Continuation
from .cutlist import CutPiece, build_cutlist
from .dims import (DEFAULT_OFFSET_MM, Dimension, flange_face,
                   infer_dimensions, load_sidecar, merge_with_sidecar)
from .model import Component, PcfDocument, Point
from .skeys import (classify, is_annotation_symbol, is_slip_on_flange,
                    is_weld_coupling, looks_like_flange, valve_subtype)
from .units import bore_to_nps, coord_to_inch_factor, format_feet_inches, format_nps
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
    quadrant: str = "NE"             # NE | NW | SW | SE
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
    # Where the piping leaves this sheet, and for which drawing. Resolved by
    # continuations.build_continuation_map() over every document about to be
    # rendered together, because the answer is the name of the NEXT SHEET --
    # not the neighbouring spool, which is all the PCF itself knows. Empty
    # means the sheet was rendered without knowledge of its neighbours, and
    # only the file's own END-CONNECTION references (if any) are available.
    continuations: list[Continuation] = field(default_factory=list)


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

    def page(self, p: Vec3) -> Vec2:
        sx, sy = self.raw(p)
        return (sx * self.scale + self.offset[0], sy * self.scale + self.offset[1])

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
                    rigid: list[list[tuple[float, float, float]]] | None = None):
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
        vals = sorted({round(p[ax], 6) for p in points})
        src = [vals[0]]
        dst = [vals[0]]
        # ``locked`` is sorted and non-overlapping, so the only interval that
        # can contain a gap is the last one starting at or before it. Binary
        # search rather than scan, to keep this linear in the number of
        # coordinates on sheets that carry thousands of both.
        lows = [lo for lo, _hi, _s in locked]
        for v in vals[1:]:
            gap = v - src[-1]
            shared = None
            if locked:
                j = bisect_right(lows, src[-1]) - 1
                if j >= 0 and v <= locked[j][1]:
                    shared = locked[j][2]
            if shared is not None:
                gap *= shared
            else:
                inside = any(lo <= src[-1] and v <= hi for lo, hi in protected)
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
        cls = classify(comp.type)
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
    except ImportError as exc:  # pragma: no cover
        raise ImportError(
            "fpdf2 is required for PDF rendering: pip install fpdf2"
        ) from exc

    options = options or RenderOptions()
    output_path = Path(output_path)
    warnings: list[str] = []

    page_w, page_h = _PAPERS.get(options.paper.lower(), _PAPERS["tabloid"])
    pdf = FPDF(orientation="landscape" if page_w >= page_h else "portrait",
               unit="mm", format=(page_h, page_w) if page_w >= page_h else (page_w, page_h))
    pdf.set_auto_page_break(False)
    pdf.add_page()
    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)

    # ---- layout regions -------------------------------------------------
    border = 6.0
    bom_w = page_w * 0.24 if options.show_bom else 0.0
    title_h = page_h * 0.16
    draw_region = (
        border + 2, border + 2,
        page_w - 2 * border - bom_w - 4,
        page_h - 2 * border - title_h - 4,
    )
    spread = max(0.4, min(3.0, options.balloon_spread))
    canvas = _Canvas(pdf, bounds=draw_region, spread=spread)

    comps: list[tuple[Component, str]] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append((comp, pipeline.reference))
    if not comps:
        raise ValueError("document contains no drawable components")

    all_points = [p.xyz for comp, _ in comps for p in comp.all_points()]

    projector = Projector(options)
    projector.fit(all_points, draw_region)

    reducer_runs = _reducer_runs(doc)

    if options.not_to_scale and len(all_points) > 1:
        # Fixed-count iteration: the warp targets depend on the page scale
        # and the scale depends on the warped extents, so refit a few times.
        # Deterministic -- same model and options always converge identically.
        tiny_world = _NTS_TINY_INCHES / coord_to_inch_factor(doc.units_coords)
        keep = _eccentric_offset_pairs(doc, reducer_runs)
        centrelines = _straight_runs(comps)
        skew_segments = _skew_segments(
            centrelines, all_points, geometry.model_tolerance(doc),
            coord_to_inch_factor(doc.units_coords))
        for _ in range(_NTS_ITERATIONS):
            rigid = _skew_rigid_spans(projector, skew_segments, tiny_world,
                                      options.nts_min_mm, options.nts_max_mm)
            projector.warp = _build_nts_warp(
                all_points, projector.scale,
                options.nts_min_mm, options.nts_max_mm, tiny_world, keep, rigid)
            projector.fit(all_points, draw_region)

    # Flat 45-degree runs project vertical, exactly like a riser; tilt them
    # off vertical so the sheet can be read. Checked against the projector as
    # already configured, so the warp is accounted for, and skipped entirely
    # when nothing on this sheet collides.
    skew_deg = max(0.0, min(20.0, options.skew_deg))
    if skew_deg and _draws_flat_run_vertically(projector, _straight_runs(comps)):
        projector.set_skew(skew_deg)
        projector.fit(all_points, draw_region)

    draw_centre = (draw_region[0] + draw_region[2] / 2,
                   draw_region[1] + draw_region[3] / 2)

    # ---- symbol pass -----------------------------------------------------
    flange_face_pts = _flange_faces(doc)
    _FLANGE_BANDS.clear()      # 每张图重置接带注册表（焊点吸附用）
    olet_header_pts = _olet_header_points(doc)
    olet_branch_conn = _olet_branch_connected(doc)
    flange_bar_mm = max(1.0, options.flange_bar_mm)
    anchors: dict[int, Vec2] = {}
    for idx, (comp, _ref) in enumerate(comps):
        try:
            anchor = _draw_component(canvas, projector, comp, doc,
                                     flange_face_pts, flange_bar_mm,
                                     olet_header_pts, olet_branch_conn,
                                     reducer_runs)
        except Exception as exc:  # defensive: a bad record must not kill the sheet
            warnings.append(f"component {idx} ({comp.type}): not drawn ({exc})")
            anchor = None
        if anchor is not None:
            anchors[idx] = anchor

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
            if -1.5 <= s <= 4.0 and lat < 2.0:
                dot = exitp
                break
        weld_dots.append(dot)
        _draw_weld_symbol(canvas, projector, weld, dot)
        nps = format_nps(bore_to_nps(weld.bore, doc.units_bore)) if weld.bore else "?"
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
    drawn_continuations: list[Continuation] = []
    for cont in options.continuations:
        if not cont.label:
            continue
        _draw_continuation(canvas, projector, cont, doc, options, draw_centre)
        drawn_continuations.append(cont)
        if not cont.resolved:
            warnings.append(
                f"continuation at {_coordinate_text(cont.point, doc, options)} "
                f"is labelled from the PCF's own reference ({cont.reference}) "
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
                                  weld_dots=weld_dots)

    # ---- furniture ----------------------------------------------------------
    _draw_border(canvas, page_w, page_h, border)
    _draw_north_arrow(canvas, projector, (draw_region[0] + 8, draw_region[1] + 10))
    bom_truncated = cut_truncated = False
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

    pdf.output(str(output_path))

    bom_lines = [
        BomLine(item=str(r["item"]), description=r["description"], nps=r["nps"],
                cls=r["class"], qty=r["qty"], length_in=round(r["length_in"], 3),
                weight=round(r["weight"], 3),
                ucis=";".join(sorted(r["_ucis"])), code=r["code"])
        for r in sorted(bom_rows.values(), key=lambda r: _item_sort_key(r["item"]))
    ]

    return RenderResult(
        pdf_path=output_path,
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
    inch = coord_to_inch_factor(doc.units_coords)

    def fmt(value: float) -> str:
        return format_feet_inches(value * inch, force_feet=True)

    return [f"N {fmt(north)}   E {fmt(east)}", f"EL {fmt(point[2])}"]


def _coordinate_text(point: Vec3, doc: PcfDocument,
                     options: RenderOptions) -> str:
    return "  ".join(_coordinate_lines(point, doc, options))


def _draw_continuation(canvas: _Canvas, proj: Projector, cont: Continuation,
                       doc: PcfDocument, options: RenderOptions,
                       draw_centre: Vec2) -> None:
    """One continuation mark: arrow along the run + the drawing/coords block."""
    at = proj.page(cont.point)
    out: Vec2 | None = None
    if cont.inboard is not None:
        out = _v2_norm(_v2_sub(at, proj.page(cont.inboard)))
    if not out or out == (0.0, 0.0):
        out = _v2_norm(_v2_sub(at, draw_centre)) or (0.7071, -0.7071)

    tip = _v2_add(at, _v2_scale(out, _CONT_STUB_MM))
    canvas.line(at, tip, width=0.35)
    _draw_arrowhead(canvas, tip, _v2_scale(out, -1.0))

    lines = [f"CONT ON {cont.label}", *_coordinate_lines(cont.point, doc, options)]
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
        if classify(comp.type) != "OLET" or not comp.centre_points:
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
    cls = classify(comp.type)
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
        if classify(comp.type) != "REDUCER" or len(comp.end_points) < 2:
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
        if classify(comp.type) != "REDUCER" or len(comp.end_points) < 2:
            continue
        if "ECCENTRIC" not in comp.type.upper():
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
             if classify(c.type) == "OLET" and c.branch1_points]
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


def _draw_component(canvas: _Canvas, proj: Projector, comp: Component,
                    doc: PcfDocument,
                    flange_face_pts: dict[int, Vec3] | None = None,
                    flange_bar_mm: float = 2 * DEFAULT_OFFSET_MM,
                    olet_header_pts: dict[int, Vec3] | None = None,
                    olet_branch_connected: set[int] | None = None,
                    reducer_runs: dict[int, tuple[Vec3, Vec3]] | None = None,
                    ) -> Vec2 | None:
    cls = classify(comp.type)
    eps = [proj.page(p.xyz) for p in comp.end_points]
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
    if cls in ("VALVE", "FLANGE", "ELBOW", "REDUCER", "COUPLING", "FILTER",
               "INSTRUMENT", "MULTIPORT") and len(eps) >= 2:
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
                    if classify(c.type) != "PIPE" or len(c.end_points) < 2:
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
                u_ax = _v2_norm(_v2_sub(sp1, sp0))
                off = _v2_scale(u_ax, _GASKET_GAP_MM)
                sp0 = _v2_add(sp0, off)
                sp1 = _v2_add(sp1, off)
        if _draw_symbol_2pt(canvas, comp.skey, sp0, sp1,
                            proj=proj, run3=run3,
                            wp=ep3[0].xyz if ep3 else None):
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
                    a2 = proj.page(wp)
                    b2 = proj.page((wp[0] + across[0], wp[1] + across[1],
                                    wp[2] + across[2]))
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
                if classify(fl.type) != "FLANGE":
                    continue
                bd = _flange_band(proj, doc, fl, flange_face_pts)
                if bd is None:
                    continue
                top, exitp, uu = bd
                for pt_name in ("a", "b"):
                    p = a if pt_name == "a" else b
                    d = _v2_sub(p, top)
                    s = d[0] * uu[0] + d[1] * uu[1]
                    lat = abs(d[0] * -uu[1] + d[1] * uu[0])
                    if -1.5 <= s <= 4.0 and lat < 2.0:
                        if pt_name == "a":
                            a = exitp
                        else:
                            b = exitp
            canvas.line(a, b, width=0.55)
            return _v2_mid(eps[0], eps[1])

    elif cls == "ELBOW":
        if len(eps) >= 2:
            corners = cps if cps else [_v2_mid(eps[0], eps[1])]
            chain = [eps[0], *corners, eps[1]]
            canvas.polyline(chain, width=0.55)
            branch = b1s[0] if b1s else None
            if branch is not None and corners:
                canvas.line(corners[0], branch, width=0.55)
            return corners[0]

    elif cls == "TEE":
        anchor = None
        if len(eps) >= 2:
            canvas.line(eps[0], eps[1], width=0.55)
            anchor = _v2_mid(eps[0], eps[1])
        centre = cps[0] if cps else anchor
        if centre is not None:
            for branch in (*b1s, *b2s):
                canvas.line(centre, branch, width=0.55)
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
                                    proj=proj, run3=r3, wp=c3):
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
        # 垫片不再画杠：法兰符号已外移 _GASKET_GAP_MM 表现垫片间隙，
        # 手画杠会压在阀体/法兰面 上（垫片模型厚度≈0，中点即贴合面）
        if len(eps) >= 2:
            return None  # gaskets are not ballooned on the face

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


def _flange_band(proj, doc, comp, flange_face_pts):
    """端点重合法兰（平焊等）的接带布局：(top, exit, uu) 或 None。

    top = 光面端（从配合面向管侧退 _GASKET_GAP_MM）；exit = 焊端（越过焊点
    0.6mm）；uu = 从接头指向管道远端的单位轴向。主管端点/焊点应修剪·吸附
    到 exit——实现"主管接焊端、阀门空开光面端"的接头布局。
    """
    eps3 = comp.end_points
    if len(eps3) < 2:
        return None
    e0 = proj.page(eps3[0].xyz)
    e1 = proj.page(eps3[1].xyz)
    if _v2_len(_v2_sub(e1, e0)) >= 0.5:
        return None
    best = None
    try:
        for c in doc.real_components():
            if classify(c.type) != "PIPE" or len(c.end_points) < 2:
                continue
            p0 = proj.page(c.end_points[0].xyz)
            p1 = proj.page(c.end_points[1].xyz)
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
        face = proj.page(fpt)
        top = _v2_add(face, _v2_scale(uu, _GASKET_GAP_MM))
    else:
        top = _v2_add(weld, _v2_scale((-uu[0], -uu[1]), 1.2))
    span = max(_v2_len(_v2_sub(weld, top)) + 0.6, 1.2)
    return top, _v2_add(top, _v2_scale(uu, span)), uu


def _draw_symbol_2pt(canvas: _Canvas, skey: str | None, p0: Vec2, p1: Vec2,
                     proj: "Projector | None" = None,
                     run3: tuple | None = None, wp: tuple | None = None,
                     width: float = 0.4) -> bool:
    """SKEY 符号库双接点对齐绘制（ASC 优先 → skeys.json 官方库）。

    符号 START(首点) → p0，EXIT(收笔点) → p1。
    轴向：符号自身轴向 e1=START→EXIT 先转正（偏心大小头离轴出口也能精确落点），
    再对准管轴 u=p0→p1，EXIT 精确落在 p1。
    横向 v：取"垂直于管轴的世界轴"的图面投影（等轴测中为 30° 斜向，与
    支撑符号同规则）——保证阀体两端面线与轴测坐标轴平行，而不是图面
    垂直方向。无投影信息时回退为图面垂直方向。
    未查到符号或几何退化时返回 False，调用方走手写回退分支。
    """
    try:
        from .asc_symbols import find_symbol_full
        found = find_symbol_full(skey or "")
    except Exception:
        return False
    if not found:
        return False
    sym, orient, stop = found
    diff = _v2_sub(p1, p0)
    d = (diff[0] * diff[0] + diff[1] * diff[1]) ** 0.5
    ex, ey = stop
    slen = (ex * ex + ey * ey) ** 0.5
    if d <= 1e-6 or slen <= 1e-6:
        return False
    e1 = (ex / slen, ey / slen)
    e2 = (-e1[1], e1[0])
    u = (diff[0] / d, diff[1] / d)
    v = _v2_perp(u)
    if proj is not None and run3 is not None and wp is not None:
        # 与支撑分支同规则：X 向管跨 Y，Y/Z 向管跨 X（跨管轴取水平世界轴）
        axis = max(range(3), key=lambda i: abs(run3[i]))
        across = (0.0, 1.0, 0.0) if axis == 0 else (1.0, 0.0, 0.0)
        a2 = proj.page(wp)
        b2 = proj.page((wp[0] + across[0], wp[1] + across[1], wp[2] + across[2]))
        vn = (b2[0] - a2[0]) ** 2 + (b2[1] - a2[1]) ** 2
        if vn > 1e-12:
            vn = vn ** 0.5
            vv = ((b2[0] - a2[0]) / vn, (b2[1] - a2[1]) / vn)
            # 与管轴近平行（异常投影）时保持图面垂直回退
            if abs(vv[0] * u[0] + vv[1] * u[1]) < 0.99:
                v = vv
    sx = d / slen
    bs = [c for seg in sym for c in
          (seg[0] * e2[0] + seg[1] * e2[1], seg[2] * e2[0] + seg[3] * e2[1])]
    perp = (max(bs) - min(bs)) if bs else 1.0
    sy = 4.5 / max(perp, 1e-6)
    # 双向保底：端距过小（平焊法兰两端点重合）不压扁；端距过大（盲板细长条
    # 符号遇长端距）不拉成宽扁形——拉伸比限制在等比的 2 倍以内
    if sx < sy:
        sx = sy
    elif sx > sy * 2.0:
        sx = sy * 2.0
    x0, y0 = p0
    for x1, y1, x2, y2 in sym:
        a1 = x1 * e1[0] + y1 * e1[1]
        b1 = x1 * e2[0] + y1 * e2[1]
        a2 = x2 * e1[0] + y2 * e1[1]
        b2 = x2 * e2[0] + y2 * e2[1]
        canvas.line(
            (x0 + a1 * sx * u[0] + b1 * sy * v[0],
             y0 + a1 * sx * u[1] + b1 * sy * v[1]),
            (x0 + a2 * sx * u[0] + b2 * sy * v[0],
             y0 + a2 * sx * u[1] + b2 * sy * v[1]),
            width=width)
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

    if "ECCENTRIC" in comp.type.upper():
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
            if _segs_cross(e[0], e[1], seg[0], seg[1]):
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
                   draw_centre: Vec2) -> tuple[dict | None, float]:
    """Best-scoring placement for a dimension between two 3D points.

    Offset candidates are the world axes NOT spanned by the measured
    segment (both signs), so the leader bars and dimension line stay on an
    isometric plane. Falls back to the screen perpendicular only for fully
    skew segments (all three axes spanned, e.g. an overridden diagonal).
    """
    a, b = proj.page(p3a), proj.page(p3b)
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


def _render_dim(canvas: _Canvas, geom: dict) -> None:
    for e in geom["ext"]:
        canvas.line(e[0], e[1], width=0.15, record=False)
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
            if classify(comp.type) == "PIPE" and len(comp.end_points) >= 2:
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
                    a2 = proj.page(wp)
                    b2 = proj.page((wp[0] + across[0], wp[1] + across[1],
                                    wp[2] + across[2]))
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
                             offset_mm, span_axes, draw_centre)
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

    inch = coord_to_inch_factor(units_coords)

    def leg_text(ax: int) -> str:
        return format_feet_inches(abs(delta[ax]) * inch)

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
    return classify(comp.type)


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
            nps = format_nps(bore_to_nps(bore, doc.units_bore)) if bore else ""
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
    """One square callout box, standing off the run with a leader to it.

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
    side = max(tw, _CALLOUT_CAP) + 2 * _CALLOUT_PAD_MM
    if options:
        anchor, pos = canvas.place_label_multi(
            options, side, side, draw_centre,
            min_leader=_CALLOUT_LEADER_MM, avoid_geometry=True,
            clearance=_CALLOUT_CLEARANCE_MM,
            anchor_cost=lambda p: _crowding_cost(p, occupied or []))
    else:
        pos = canvas.place_label(anchor, side, side, draw_centre,
                                 min_leader=_CALLOUT_LEADER_MM,
                                 avoid_geometry=True,
                                 clearance=_CALLOUT_CLEARANCE_MM)
    # Leader to the box *edge*, so the run of it is visible rather than
    # buried under the knockout, then the box over the top of it.
    edge = _box_edge_toward(pos, anchor, side, side)
    canvas.line(anchor, edge, width=0.15)
    with canvas.pdf.local_context():
        canvas.pdf.set_fill_color(255, 255, 255)
        canvas.pdf.set_line_width(0.3)
        canvas.pdf.rect(pos[0] - side / 2, pos[1] - side / 2, side, side,
                        style="DF")
    canvas.text((pos[0], pos[1] + _CALLOUT_CAP / 2), label, size=_CALLOUT_FONT)
    return anchor, pos


def _draw_callouts(canvas: _Canvas, proj: Projector, comps, anchors, bom_rows,
                   pipe_labels: dict[int, str], draw_centre,
                   every_instance: bool = True,
                   weld_dots: list[Vec2] | None = None) -> list[Callout]:
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
        if idx not in anchors or classify(comp.type) != "PIPE":
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
        anchor, pos = _draw_callout_box(
            canvas, anchor, label, draw_centre,
            options=_pipe_anchor_options(comp, proj) or [anchor],
            occupied=occupied)
        callouts.append(Callout(
            item=label, kind="pipe", x_mm=round(pos[0], 2), y_mm=round(pos[1], 2),
            anchor_x_mm=round(anchor[0], 2), anchor_y_mm=round(anchor[1], 2),
            leadered=True))

    # -- pass 2: material item-number callouts ------------------------------
    done: set[tuple] = set()
    for idx, (comp, _ref) in enumerate(comps):
        if idx not in anchors or classify(comp.type) == "PIPE":
            continue
        key = _bom_key(comp)
        row = bom_rows.get(key)
        if row is None:
            continue
        if not every_instance:
            if key in done:
                continue
            done.add(key)
        anchor = anchors[idx]
        _anchor, pos = _draw_callout_box(canvas, anchor, str(row["item"]),
                                         draw_centre)
        callouts.append(Callout(
            item=str(row["item"]), kind="material",
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
        qty = (format_feet_inches(row["length_in"]) if row["class"] == "PIPE"
               else str(row["qty"]))
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
