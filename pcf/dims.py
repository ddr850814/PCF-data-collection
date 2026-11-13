"""Dimensional callout inference with human override.

Auto-dimensioning follows the fabricator's convention: dimensions run
between *landmarks*, and the landmarks that matter are **fitting centres**.
An ell is dimensioned to its tangent-intersection corner and a tee/olet to
its branch centreline, because that is the point a fitter can actually lay
out. Pipe ends, nipple ends, and the weld ends of ells and tees are
deliberately *not* landmarks -- they are artifacts of how the model was
cut up, they sit a bore-radius away from the point that matters, and they
bury the real dimensions in noise. They earn a landmark only when they are
genuinely open (a spool terminus with nothing welded to it).

Landmark ranks, strongest first:

===== ==================================== ===============================
 rank  landmark                             source
===== ==================================== ===============================
 0     ELL CENTER, BRANCH CL                elbow/bend corner, tee/olet CL
 1     FLG FACE, VALVE CL, CAP              gasket face, valve CL, cap end
 2     LINE END                             genuinely open ends only
===== ==================================== ===============================

When two landmarks fall closer together than the readability threshold the
better-ranked one wins outright, so a centre is never displaced by an end
that happens to sit beside it. A component carrying a landmark of its own
never also contributes an open end, so no dimension measures a flange's or
a valve's own catalogue length.

Legs are *collinear runs*, not single pipes: every straight component axis
in the model is merged into a maximal straight line, coverage along that
line is walked, and the landmarks on it are chained. A run broken into a
pipe, a nipple and another pipe by intermediate welds still dimensions ell
centre to ell centre in one span, and a spool with no pipe in it at all --
two flanges bolted to an ell -- still has legs to chain along. When no leg
holds two landmarks (an eccentric reducer offsets the two faces onto
parallel centrelines) the spool's overall extent is dimensioned instead.

Inference is deterministic geometry, but the human always wins: every
render writes/refreshes a ``<name>.dims.json`` sidecar next to the PDF.
Each auto dimension carries a stable ``id`` (hash of its two 3D anchor
points). Edit the sidecar to suppress an auto dimension (``"suppressed":
true``), override its text (``"text"``), move it (``"offset_mm"``), or add
manual entries (``"kind": "manual"`` with ``from``/``to`` 3D points).
Re-rendering re-infers geometry but preserves any entry you touched.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path

from . import geometry
from .model import Component, PcfDocument
from .skeys import classify, is_annotation_symbol, looks_like_flange
from .units import format_length

Vec3 = tuple[float, float, float]

# Landmarks closer together than this fraction of the model diagonal do not
# get their own dimension (they'd be unreadable). When a pair is that close
# the better-ranked landmark absorbs the other rather than both surviving
# undimensioned.
_MIN_DIM_FRACTION = 0.01

# Standard perpendicular offset (page mm) of a dimension line from the run
# it measures. The renderer also uses it as the yardstick for the flange
# face bar, which is drawn twice this long.
DEFAULT_OFFSET_MM = 6.0

# Landmark ranks: lower wins when two landmarks coincide or crowd.
RANK_CENTRE = 0    # ELL CENTER / BRANCH CL -- what a fitter lays out to
RANK_FACE = 1      # FLG FACE / VALVE CL / CAP
RANK_END = 2       # LINE END -- genuinely open ends

# Component classes whose ends are never landmarks even when the
# connectivity graph calls them open: welds, bolting and gaskets have no
# meaning as a dimension target, and their "open" ends are modelling
# artifacts of root gaps and set-on attachments.
_NEVER_OPEN_END = {"WELD", "BOLT", "GASKET", "SUPPORT", "CONNECTION",
                   "MESSAGE", "GAP"}

# Which components carry a landmark of their own -- and so never contribute
# an open end on top of it -- is decided by _own_landmark().


@dataclass
class Dimension:
    id: str
    kind: str                    # auto | manual
    p_from: Vec3
    p_to: Vec3
    landmark_from: str
    landmark_to: str
    text: str | None = None      # explicit override; None = computed
    offset_mm: float = DEFAULT_OFFSET_MM   # offset of the dimension line
    suppressed: bool = False

    def computed_text(self, units_coords: str) -> str:
        if self.text:
            return self.text
        return format_length(geometry.dist(self.p_from, self.p_to), units_coords)


def _dim_id(a: Vec3, b: Vec3) -> str:
    lo, hi = sorted([a, b])
    raw = f"{lo[0]:.3f},{lo[1]:.3f},{lo[2]:.3f}|{hi[0]:.3f},{hi[1]:.3f},{hi[2]:.3f}"
    return hashlib.sha1(raw.encode()).hexdigest()[:10]


@dataclass
class _Landmark:
    point: Vec3
    name: str
    rank: int   # lower rank wins when landmarks coincide or crowd
    # id() of the component this landmark belongs to, when it has one. A
    # component contributing *two* landmarks (only reducers do) must never be
    # dimensioned across itself -- that is its catalogue length, not a shop
    # number -- so the chain walk refuses a consecutive same-owner pair.
    owner: int | None = None


def _collect_landmarks(doc: PcfDocument) -> list[_Landmark]:
    """Fitting centres first, faces second, genuinely open ends last."""
    tol = geometry.model_tolerance(doc)
    comps = doc.real_components()
    graph = geometry.ConnectivityGraph(comps, tol)

    landmarks: list[_Landmark] = []

    def add(point: Vec3, name: str, rank: int, owner: int | None = None) -> None:
        for existing in landmarks:
            if geometry.dist(existing.point, point) <= tol * 4:
                if rank < existing.rank:
                    existing.name = name
                    existing.rank = rank
                    existing.owner = owner
                return
        landmarks.append(_Landmark(point=point, name=name, rank=rank, owner=owner))

    # -- pass 1: the points a fitter lays out to ---------------------------
    speaks_for_itself: set[int] = set()
    for idx, comp in enumerate(comps):
        for point, name, rank in _own_landmarks(comp, comps, graph, idx, tol):
            add(point, name, rank, owner=id(comp))
            speaks_for_itself.add(id(comp))

    # -- pass 2: genuinely open ends, only where nothing better exists -----
    # add() already refuses to downgrade a coincident rank-0/1 landmark, so
    # a flange face or cap end that is also "open" keeps its own name.
    #
    # END-POSITION-OPEN/CLOSED records are the file *declaring* where the
    # line stops, and they are believed over the connectivity walk: a run
    # that dead-ends in a modelled field weld leaves that weld's outboard
    # end within tolerance of its own inboard end, so the graph sees it as
    # connected and the spool boundary would otherwise carry no landmark at
    # all. END-CONNECTION-* records are skipped -- those are continuations
    # into another pipeline, which lands mid-run on a merged sheet.
    #
    # Believed, but not blindly: see _is_run_terminus().
    cont = _continuity(doc, comps)
    for comp in doc.all_components():
        if not comp.type.upper().startswith("END-POSITION"):
            continue
        for pt in comp.co_ords:
            if not _is_run_terminus(pt.xyz, cont, tol):
                continue
            add(pt.xyz, "LINE END", RANK_END)

    # The connectivity walk needs the same judgement as the records above,
    # and for the same reason: a zero-length drafting mark dropped in the
    # middle of a run is joined to nothing, so the graph calls both of its
    # ends open and hands back a landmark that cuts the run in half.
    for comp, pt in geometry.open_ends(doc, tol):
        if classify(comp.type, comp.skey) in _NEVER_OPEN_END:
            continue
        if id(comp) in speaks_for_itself:
            continue  # its own centre/face already speaks for it
        if not _is_run_terminus(pt.xyz, cont, tol):
            continue
        add(pt.xyz, "LINE END", RANK_END)

    return landmarks


@dataclass
class _Continuity:
    """What material is doing at a point, for judging END-POSITION records.

    Built once per document: ``arrows`` are the points of drafting-mark
    records, ``bodies`` the straight stretches every other component spans.
    """
    arrows: list[Vec3]
    bodies: list[tuple[Vec3, Vec3]]


# An elbow's two weld ends are joined by the diagonal across the bend, which
# is not a line the drawing ever draws and not material anything sits inside.
# _leg_segments() splits it into arms for the same reason.
_NO_STRAIGHT_BODY = {"ELBOW", "BEND"}


def _continuity(doc: PcfDocument, comps: list[Component]) -> _Continuity:
    materials = doc.materials_descriptions()
    arrows: list[Vec3] = []
    bodies: list[tuple[Vec3, Vec3]] = []
    for comp in comps:
        desc = comp.description or materials.get(comp.item_code or "", "")
        if is_annotation_symbol(comp.type, desc, comp.skey):
            arrows.extend(p.xyz for p in comp.all_points())
            continue
        if classify(comp.type, comp.skey) in _NO_STRAIGHT_BODY:
            continue
        eps = comp.end_points
        if len(eps) >= 2:
            bodies.append((eps[0].xyz, eps[-1].xyz))
    return _Continuity(arrows=arrows, bodies=bodies)


def _is_run_terminus(point: Vec3, cont: _Continuity, tol: float) -> bool:
    """Whether an END-POSITION point is really where the line stops.

    The record says "the line ends here"; that is worth believing, but it is
    not the same claim as "no material continues past here", and two common
    exports break the second one:

    **Drafting marks.** The dominant exporter ships pipe flow arrows and slope
    arrows as zero-length ``MISC-COMPONENT`` records, each with its own
    coincident ``END-POSITION-OPEN``, dropped wherever the drafter wanted the
    arrow -- routinely the middle of a run. ``_build_bom`` already refuses to
    buy a slope arrow; believing its END-POSITION here cut a plain 20'-0"
    riser into two dimensions and turned a single sloped leg into three
    stacked offset boxes, because each fragment of a broken chain is an
    off-axis dimension that earns an offset box of its own.

    **Material passing straight through.** More generally, a point sitting
    strictly inside some component's body -- clear of both its ends -- has
    pipe running through it, whatever the record says. That catches exports
    that do not name the arrow in a description.

    Nothing here weakens the case the blanket trust was added for: a run
    dead-ending in a modelled field weld has material on one side only, so it
    is still a terminus and still gets its landmark.
    """
    if any(geometry.dist(p, point) <= tol * 4 for p in cont.arrows):
        return False
    for a, b in cont.bodies:
        axis = geometry.sub(b, a)
        span = geometry.length(axis)
        if span <= tol:
            continue
        unit = geometry.scale(axis, 1.0 / span)
        along = geometry.dot(geometry.sub(point, a), unit)
        if not (tol < along < span - tol):
            continue    # at or beyond an end, not inside the body
        foot = geometry.add(a, geometry.scale(unit, along))
        if geometry.dist(foot, point) <= tol * 4:
            return False
    return True


def _own_landmarks(comp: Component, comps: list[Component],
                   graph: geometry.ConnectivityGraph, idx: int,
                   tol: float) -> list[tuple[Vec3, str, int]]:
    """The rank-0/1 landmarks a component carries in its own right.

    A component that has one never also contributes a rank-2 open end. The
    ell corner, gasket face, valve centreline or cap end *is* the point a
    fitter lays out to; admitting the component's other end alongside it
    would dimension its own catalogue length -- a flange's 3 1/2", a valve's
    face-to-face -- which is a number from a catalogue, not from the shop.

    A reducer is the exception that carries *two*: it is where the line
    changes size, which is a point the shop has to lay out on both sides of,
    and its two ends sit on two different legs -- on an eccentric reducer,
    two parallel centrelines. One landmark would leave whichever leg it did
    not land on with nothing to chain to, which is exactly how the small-bore
    leg past a reducer went out undimensioned. The pair is tagged with a
    shared owner so the chain walk never dimensions the reducer's own body.
    """
    cls = classify(comp.type, comp.skey)
    if cls == "ELBOW":
        corner = _elbow_corner(comp)
        return [(corner, "ELL CENTER", RANK_CENTRE)] if corner is not None else []
    if cls in ("TEE", "OLET"):
        centre = branch_centre(comp)
        return [(centre, "BRANCH CL", RANK_CENTRE)] if centre is not None else []
    if looks_like_flange(comp.type, comp.description, comp.skey):
        face = flange_face(comp, comps, graph, idx, tol)
        return [(face, "FLG FACE", RANK_FACE)] if face is not None else []
    if cls == "VALVE":
        centre = _valve_centre(comp)
        return [(centre, "VALVE CL", RANK_FACE)] if centre is not None else []
    if cls == "CAP":
        end = _cap_end(comp, comps, graph, idx, tol)
        return [(end, "CAP", RANK_FACE)] if end is not None else []
    if cls == "REDUCER":
        return _reducer_ends(comp)
    return []


def _reducer_ends(comp: Component) -> list[tuple[Vec3, str, int]]:
    """Both weld ends of a reducer, large end first.

    Named apart so the sidecar says which end a dimension runs to: on an
    eccentric reducer they are an inch or more apart across the run, and a
    fitter reading "RED LG" knows which one.
    """
    pts = [p for p in comp.end_points][:2]
    if len(pts) < 2:
        return []
    big, small = sorted(pts, key=lambda p: p.bore or 0.0, reverse=True)
    if big.xyz == small.xyz:
        return []
    return [(big.xyz, "RED LG", RANK_FACE), (small.xyz, "RED SM", RANK_FACE)]


def _elbow_corner(comp: Component) -> Vec3 | None:
    """Tangent-intersection corner of an elbow/bend.

    The CENTRE-POINT in real exports is the corner (tangent intersection)
    itself; when it is missing, the midpoint of the two ends is the best
    available stand-in.
    """
    if comp.centre_points:
        return comp.centre_points[0].xyz
    eps = comp.end_points
    if len(eps) >= 2:
        return geometry.midpoint(eps[0].xyz, eps[1].xyz)
    return None


def branch_centre(comp: Component) -> Vec3 | None:
    """Tee/olet branch centreline on the header."""
    if comp.centre_points:
        return comp.centre_points[0].xyz
    eps = comp.end_points
    if len(eps) >= 2:
        return geometry.midpoint(eps[0].xyz, eps[1].xyz)
    return None


def _valve_centre(comp: Component) -> Vec3 | None:
    if comp.centre_points:
        return comp.centre_points[0].xyz
    eps = comp.end_points
    if len(eps) >= 2:
        return geometry.midpoint(eps[0].xyz, eps[1].xyz)
    return None


def _cap_end(comp: Component, comps: list[Component],
             graph: geometry.ConnectivityGraph, idx: int,
             tol: float) -> Vec3 | None:
    """The closed end of a cap -- the one not welded to the run."""
    eps = comp.end_points
    if not eps:
        return None
    if len(eps) == 1:
        return eps[0].xyz
    welded = _welded_end_index(eps, comps, graph, idx, tol)
    if welded is None:
        return eps[-1].xyz
    return eps[1 - welded].xyz


# Neighbours that mean "the run is welded on here", strongest first. A bare
# coupling is weaker evidence than a length of pipe, and a valve weaker still
# because it is as often bolted on as welded.
_WELDED_BY_RANK = (
    ("PIPE", "ELBOW", "TEE", "OLET", "REDUCER"),
    ("COUPLING", "WELD", "VALVE"),
)


def _welded_end_index(eps, comps: list[Component],
                      graph: geometry.ConnectivityGraph, idx: int,
                      tol: float) -> int | None:
    """Which of the first two end-points is joined to the run.

    A flange neighbour never counts, whichever end it lands on: flanges bolt
    to each other, they do not weld. That matters because the dominant
    exporter ships weld-neck flanges as COUPLING records, so on a bolted pair
    each half saw the other as run -- and whichever was reached first through
    the adjacency put its face bar on its *welded* end, at the back of the
    flange instead of the gasket.

    Ranking rather than first-hit for the same reason: with both ends
    claimed, the answer must not depend on adjacency order.
    """
    best: tuple[int, int] | None = None
    # 垫片只存在于法兰（螺栓）连接侧：一端贴着垫片时另一端必为焊端。
    # 这条结构证据优先于 rank 表——否则隔着垫片 1.5mm 的邻件（如 TEE，
    # rank-0）会抢先认领垫片侧端点，把焊端判反（法兰面杠画向管道）。
    gasket_ends: set[int] = set()
    for other_idx in graph.neighbours(idx):
        other = comps[other_idx]
        if classify(other.type, other.skey) != "GASKET":
            continue
        for pt in other.connection_points():
            for n, ep in enumerate(eps[:2]):
                if ep.distance_to(pt) <= tol:
                    gasket_ends.add(n)
    if len(gasket_ends) == 1:
        return 1 - next(iter(gasket_ends))
    for other_idx in graph.neighbours(idx):
        other = comps[other_idx]
        if looks_like_flange(other.type, other.description, other.skey):
            continue
        cls = classify(other.type, other.skey)
        rank = next((r for r, names in enumerate(_WELDED_BY_RANK)
                     if cls in names), None)
        if rank is None or (best is not None and rank >= best[0]):
            continue
        for pt in other.connection_points():
            for n, ep in enumerate(eps[:2]):
                if ep.distance_to(pt) <= tol:
                    best = (rank, n)
                    break
            else:
                continue
            break
    return None if best is None else best[1]


def flange_face(comp: Component, comps: list[Component],
                graph: geometry.ConnectivityGraph, idx: int,
                tol: float) -> Vec3 | None:
    """The gasket-side face of a flange, in native coordinates.

    Exports mark it directly often enough to trust: an ``FL`` end-point
    connection type is the flanged face and a weldable one (``BW``/``PL``/
    ``SW``) is the pipe side. When the codes are ambiguous -- plenty of real
    files stamp both ends ``PL`` -- fall back to connectivity and take the
    end that nothing is welded to.
    """
    eps = comp.end_points
    if not eps:
        return None
    if len(eps) == 1:
        return eps[0].xyz
    face_idx = _flange_face_index(comp, comps, graph, idx, tol)
    return eps[face_idx].xyz


_WELDABLE_ENDS = {"BW", "PL", "SW"}


def _flange_face_index(comp: Component, comps: list[Component],
                       graph: geometry.ConnectivityGraph, idx: int,
                       tol: float) -> int:
    """Index into ``comp.end_points`` of the gasket face (0 or 1)."""
    eps = comp.end_points
    types = [(ep.connection_type or "").upper() for ep in eps[:2]]
    if "FL" in types and not all(t == "FL" for t in types):
        return types.index("FL")
    flagged = [n for n, t in enumerate(types) if t in _WELDABLE_ENDS]
    if len(flagged) == 1:
        return 1 - flagged[0]
    if comp.type.upper().replace("_", "-") in ("FLANGE-BLIND", "BLIND-FLANGE"):
        # 盲法兰没有焊端：配合面是与垫片/阀门相接的一端，另一端为自由外侧。
        # 不能沿用普通法兰的默认末端，否则符号会朝阀门内部偏移。
        neighbours = graph.adjacency.get(idx, ())
        nearest = []
        for ep in eps[:2]:
            distances = [ep.distance_to(p) for j in neighbours
                         for p in comps[j].all_points()]
            nearest.append(min(distances) if distances else float("inf"))
        if nearest[0] != nearest[1]:
            return 0 if nearest[0] < nearest[1] else 1
    welded = _welded_end_index(eps, comps, graph, idx, tol)
    if welded is not None:
        return 1 - welded
    return 1


@dataclass
class _Leg:
    """A maximal straight line through the routing, in native coordinates."""
    origin: Vec3
    unit: Vec3


def _canonical(unit: Vec3) -> Vec3:
    """Flip a direction to a canonical sign so a line has one representation."""
    for value in unit:
        if abs(value) > 1e-9:
            return unit if value > 0 else geometry.scale(unit, -1.0)
    return unit


def _off_line(leg: _Leg, point: Vec3) -> tuple[float, float]:
    """(distance along the line, distance from it) for a point."""
    rel = geometry.sub(point, leg.origin)
    t = geometry.dot(rel, leg.unit)
    foot = geometry.add(leg.origin, geometry.scale(leg.unit, t))
    return t, geometry.dist(foot, point)


# Classes whose two end-points bound a straight length of the routing, and
# so can start a leg. Pipe is the usual case, but a flange, valve or reducer
# bolted straight to a fitting is a length of line too: a spool with no pipe
# in it at all -- two flanges on an elbow, ordinary on short tie-ins and
# ship-loose assemblies -- had no leg to hang its dimensions on and came out
# with none. Fittings that turn or branch (elbows, olets) are excluded: the
# straight line between their ends is not something the drawing ever draws.
# On a normal spool this adds nothing, because a flange or valve welded into
# a run is collinear with the pipe beside it and merges into its leg.
_LEG_CLASSES = {"PIPE", "FLANGE", "VALVE", "REDUCER", "CAP", "COUPLING",
                "TEE", "INSTRUMENT", "FILTER"}


def _leg_segments(comp: Component) -> list[tuple[Vec3, Vec3]]:
    """Straight stretches of routing this component establishes.

    Straight bodies give their own two ends. An **elbow gives its two arms**
    -- each weld end out to the tangent-intersection corner -- never its
    end-to-end chord, which is the diagonal across the bend and is not a
    line the drawing ever draws.

    The arms matter because a leg has to exist before anything can be
    dimensioned along it. A run built entirely out of fittings -- two 6"
    elbows welded back to back, an 18" offset with no pipe in it at all --
    established no leg, so the dimension between its two ell centres was
    never generated and the leg went out blank. On an ordinary spool the
    arms are collinear with the pipe they continue and merge into its leg,
    so this adds nothing.
    """
    eps = comp.end_points
    if len(eps) < 2:
        return []
    if classify(comp.type, comp.skey) in ("ELBOW", "BEND"):
        if not comp.centre_points:
            return []
        corner = comp.centre_points[0].xyz
        return [(eps[0].xyz, corner), (eps[-1].xyz, corner)]
    if classify(comp.type, comp.skey) in _LEG_CLASSES or looks_like_flange(
            comp.type, comp.description, comp.skey):
        return [(eps[0].xyz, eps[1].xyz)]
    return []


def _collect_legs(comps: list[Component], off_tol: float) -> list[_Leg]:
    """Merge every straight component axis into maximal collinear lines.

    A run chopped into pipe / nipple / pipe by intermediate welds is one
    leg, so its bounding ell centres get a single dimension between them
    instead of one dimension per modelled pipe piece.
    """
    legs: list[_Leg] = []
    for comp in comps:
        for a, b in _leg_segments(comp):
            axis = geometry.sub(b, a)
            if geometry.length(axis) <= off_tol:
                continue
            unit = _canonical(geometry.normalize(axis))
            for leg in legs:
                if abs(geometry.dot(unit, leg.unit)) < 0.9999:
                    continue
                if (_off_line(leg, a)[1] <= off_tol
                        and _off_line(leg, b)[1] <= off_tol):
                    break
            else:
                legs.append(_Leg(origin=a, unit=unit))
    return legs


def _leg_intervals(leg: _Leg, comps: list[Component],
                   off_tol: float, gap_tol: float) -> list[tuple[float, float]]:
    """Contiguous stretches of the leg line that carry material.

    A component contributes when at least two of its points sit on the line:
    that is true of pipes (both ends), elbows (the end on this line plus the
    corner, which is exactly how coverage reaches past the weld to the ell
    centre), tees on their run, caps, valves and flanges. Two separate runs
    that happen to be collinear -- a hairpin returning on its own axis --
    stay separate intervals, so nothing is dimensioned across thin air.
    """
    spans: list[tuple[float, float]] = []
    for comp in comps:
        ts = [t for t, off in (_off_line(leg, p.xyz) for p in comp.all_points())
              if off <= off_tol]
        if len(ts) >= 2:
            spans.append((min(ts), max(ts)))
    if not spans:
        return []
    spans.sort()
    merged = [spans[0]]
    for lo, hi in spans[1:]:
        if lo <= merged[-1][1] + gap_tol:
            merged[-1] = (merged[-1][0], max(merged[-1][1], hi))
        else:
            merged.append((lo, hi))
    return merged


def _thin_by_rank(chain: list[tuple[float, _Landmark]],
                  min_span: float) -> list[tuple[float, _Landmark]]:
    """Collapse crowded landmarks into the better-ranked one beside them.

    This is what makes centres win: a pipe end or open stub sitting within
    an unreadable distance of an ell centre is absorbed by it instead of
    stealing the dimension or breaking the chain.

    Two landmarks of the *same* rank are both kept however close they sit.
    Back-to-back elbows a bore apart are a real 45-degree offset and the
    fitter needs the number; suppressing one of them because the pair is
    under 1% of the model diagonal was how ell centres went missing on long
    lines. Short spans are legible anyway on a not-to-scale sheet, where the
    warp stretches every visible run to at least ``nts_min_mm``.
    """
    kept: list[tuple[float, _Landmark]] = []
    for t, lm in chain:
        if kept and t - kept[-1][0] < min_span:
            if lm.rank < kept[-1][1].rank:
                kept[-1] = (t, lm)      # a centre displaces a lesser landmark
            elif lm.rank == kept[-1][1].rank:
                kept.append((t, lm))    # peers: both are real fittings
            continue
        kept.append((t, lm))
    return kept


def infer_dimensions(doc: PcfDocument) -> list[Dimension]:
    """Dimension every straight run between consecutive landmarks."""
    landmarks = _collect_landmarks(doc)
    if len(landmarks) < 2:
        return []

    comps = doc.real_components()
    pts = [p.xyz for c in comps for p in c.all_points()]
    box = geometry.bounding_box(pts)
    diag = geometry.dist(box[0], box[1]) if box else 1.0
    min_span = diag * _MIN_DIM_FRACTION
    off_tol = max(diag * 0.002, geometry.model_tolerance(doc))

    dims: list[Dimension] = []
    seen: set[str] = set()

    legs = _collect_legs(comps, off_tol)
    for leg in legs:
        on_leg = []
        for lm in landmarks:
            t, off = _off_line(leg, lm.point)
            if off <= off_tol:
                on_leg.append((t, lm))
        if len(on_leg) < 2:
            continue
        on_leg.sort(key=lambda item: item[0])

        for lo, hi in _leg_intervals(leg, comps, off_tol, off_tol * 4):
            if hi - lo < min_span:
                continue
            window = [(t, lm) for t, lm in on_leg
                      if lo - off_tol <= t <= hi + off_tol]
            chain = _thin_by_rank(window, min_span)
            for (_t1, lm1), (_t2, lm2) in zip(chain, chain[1:]):
                if lm1.owner is not None and lm1.owner == lm2.owner:
                    continue  # a component's own body length is a catalogue number
                dim_id = _dim_id(lm1.point, lm2.point)
                if dim_id in seen:
                    continue
                seen.add(dim_id)
                dims.append(Dimension(
                    id=dim_id, kind="auto",
                    p_from=lm1.point, p_to=lm2.point,
                    landmark_from=lm1.name, landmark_to=lm2.name,
                ))

    dims.extend(_reducer_body_dimensions(comps, legs, dims, off_tol))
    dims.extend(_overall_dimensions(comps, landmarks, dims, off_tol, min_span))
    return dims


def _reducer_body_dimensions(comps: list[Component], legs: list[_Leg],
                             dims: list[Dimension],
                             off_tol: float) -> list[Dimension]:
    """The face-to-face length of every reducer, along the run.

    The chain walk will not produce these. It refuses any pair of landmarks
    sharing an ``owner`` -- a component's own body length is a catalogue
    number, not something the shop lays out -- and that is the right rule for
    a flange or a valve. A reducer is different: it is where the line changes
    size, the fitter has to lay the change out, and on an eccentric reducer
    the walk never even gets the chance, because the two ends sit on parallel
    centrelines an inch apart and no single leg holds both.

    The dimension is **axial**: from the large face to the small face
    *projected onto the run*, which is the longest of the two legs an
    eccentric reducer's ends make and the number a fitter measures. The
    diagonal between the faces is neither. The run direction comes from a leg
    through the large end that is not the reducer's own chord -- on a
    concentric reducer no such leg exists and the chord is the axis anyway.

    No ``min_span`` floor: a reducer body is short by nature and would fail
    it on any sheet of a decent size. The renderer still declines to draw a
    callout it has no room for, and says so.
    """
    out: list[Dimension] = []
    seen = {d.id for d in dims}
    for comp in comps:
        if classify(comp.type, comp.skey) != "REDUCER":
            continue
        ends = _reducer_ends(comp)
        if len(ends) < 2:
            continue
        big, small = ends[0][0], ends[1][0]
        delta = geometry.sub(small, big)
        if geometry.length(delta) <= off_tol:
            continue
        chord = geometry.normalize(delta)
        axis = next((leg.unit for leg in legs
                     if _off_line(leg, big)[1] <= off_tol
                     and abs(geometry.dot(leg.unit, chord)) < 0.9999), chord)
        axial = geometry.dot(delta, axis)
        if abs(axial) <= off_tol:
            continue
        far = geometry.add(big, geometry.scale(axis, axial))
        dim_id = _dim_id(big, far)
        if dim_id in seen:
            continue
        seen.add(dim_id)
        out.append(Dimension(
            id=dim_id, kind="auto", p_from=big, p_to=far,
            landmark_from="RED LG", landmark_to="RED SM",
        ))
    return out


def _overall_dimensions(comps: list[Component], landmarks: list[_Landmark],
                        dims: list[Dimension], off_tol: float,
                        min_span: float) -> list[Dimension]:
    """Overall extent of each spool the leg walk could not dimension.

    Landmarks are chained *along a leg*, which needs them to share a
    straight line. They do not always: put an eccentric reducer between two
    flanges and the two faces sit on parallel centrelines an inch apart, so
    neither leg holds both and the spool went out with no dimension on it.
    For such a group the two landmarks furthest apart are dimensioned. The
    renderer decomposes a pair that is not axis-aligned into a dashed offset
    box, so an eccentric spool prints its overall length *and* its offset --
    which is what a fitter needs from it.

    This runs **per connected group**, not once for the sheet. A sheet
    carrying a laid-out run plus a detached flange-and-reducer stub used to
    get nothing for the stub twice over: the fallback only fired when the
    whole sheet had produced no dimension at all, and it refused outright
    whenever more than one group was present. Each group is still dimensioned
    only within itself -- nothing is ever measured across the gap between two
    disjoint spools -- and only when the span is worth printing. A ship-loose
    item, with one landmark or none, still gets nothing; there is nothing on
    it to lay out.
    """
    graph = geometry.ConnectivityGraph(comps, off_tol)
    out: list[Dimension] = []
    seen = {d.id for d in dims}
    for group in graph.connected_sets():
        pts = [p.xyz for idx in group for p in comps[idx].all_points()]
        if not pts:
            continue
        box = geometry.bounding_box(pts)
        # Landmarks belonging to this group: those inside its own extent.
        lo, hi = box
        mine = [lm for lm in landmarks
                if all(lo[i] - off_tol <= lm.point[i] <= hi[i] + off_tol
                       for i in range(3))]
        if len(mine) < 2:
            continue
        if _group_is_dimensioned(dims, mine):
            continue
        a, b = max(
            ((p, q) for i, p in enumerate(mine) for q in mine[i + 1:]),
            key=lambda pair: geometry.dist(pair[0].point, pair[1].point),
            default=(None, None),
        )
        if a is None or geometry.dist(a.point, b.point) < min_span:
            continue
        dim_id = _dim_id(a.point, b.point)
        if dim_id in seen:
            continue
        seen.add(dim_id)
        out.append(Dimension(
            id=dim_id, kind="auto",
            p_from=a.point, p_to=b.point,
            landmark_from=a.name, landmark_to=b.name,
        ))
    return out


def _group_is_dimensioned(dims: list[Dimension],
                          mine: list[_Landmark]) -> bool:
    """True when some dimension already runs between two of these landmarks."""
    pts = {lm.point for lm in mine}
    return any(d.p_from in pts and d.p_to in pts for d in dims)


# ---------------------------------------------------------------------------
# Sidecar persistence


def sidecar_path(pdf_path: str | Path) -> Path:
    # Build the name by hand rather than chaining with_suffix(): line numbers
    # and spool ids routinely put dots inside the stem ("12.5878 - BREWSTER
    # ..._SHT_1.pdf"), and with_suffix would replace everything after the
    # *first* dot -- collapsing every sheet of a job onto one "12.dims.json".
    pdf_path = Path(pdf_path)
    return pdf_path.with_name(pdf_path.stem + ".dims.json") \
        if pdf_path.suffix == ".pdf" else Path(str(pdf_path) + ".dims.json")


def load_sidecar(path: str | Path) -> list[dict]:
    path = Path(path)
    if not path.is_file():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return data.get("dimensions", [])


def merge_with_sidecar(inferred: list[Dimension], sidecar_entries: list[dict]) -> list[Dimension]:
    """Apply user edits from the sidecar onto freshly inferred dimensions."""
    by_id = {d.id: d for d in inferred}
    out = list(inferred)
    for entry in sidecar_entries:
        entry_id = entry.get("id", "")
        if entry.get("kind") == "manual":
            p_from = tuple(entry["from"])
            p_to = tuple(entry["to"])
            out.append(Dimension(
                id=entry_id or _dim_id(p_from, p_to),
                kind="manual",
                p_from=p_from, p_to=p_to,
                landmark_from=entry.get("landmark_from", "MANUAL"),
                landmark_to=entry.get("landmark_to", "MANUAL"),
                text=entry.get("text"),
                offset_mm=float(entry.get("offset_mm", DEFAULT_OFFSET_MM)),
                suppressed=bool(entry.get("suppressed", False)),
            ))
            continue
        target = by_id.get(entry_id)
        if target is None:
            continue
        if "suppressed" in entry:
            target.suppressed = bool(entry["suppressed"])
        if entry.get("text"):
            target.text = entry["text"]
        if "offset_mm" in entry:
            target.offset_mm = float(entry["offset_mm"])
    return out


def write_sidecar(path: str | Path, dims: list[Dimension], units_coords: str) -> Path:
    path = Path(path)
    payload = {
        "comment": (
            "Dimensional callouts for the matching PDF. Edit and re-render: "
            "set suppressed=true to hide an auto dimension, set text to "
            "override the printed value, adjust offset_mm to move the line, "
            "or append kind=manual entries with from/to [x,y,z] in the PCF's "
            "native coordinates."
        ),
        "units_coords": units_coords,
        "dimensions": [
            {
                "id": d.id,
                "kind": d.kind,
                "from": list(d.p_from),
                "to": list(d.p_to),
                "landmark_from": d.landmark_from,
                "landmark_to": d.landmark_to,
                "value": d.computed_text(units_coords),
                "text": d.text,
                "offset_mm": d.offset_mm,
                "suppressed": d.suppressed,
            }
            for d in dims
        ],
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path
