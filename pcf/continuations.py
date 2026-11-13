"""Continuation callouts: where the piping leaves this drawing.

A fabrication isometric has to say, at every point where the line runs off
the sheet, **which drawing it runs onto**. That name is a property of the
*drawing set*, not of the PCF: a sheet may be one spool, or a whole
multi-spool region of a distance split, so the spool/pipeline reference the
file carries is the wrong answer -- it names the neighbouring *piping*, not
the neighbouring *sheet*, and on a merged sheet the two are not the same
thing. The reference is kept only as the fallback for a continuation that
resolves to no drawing in the set (piping that continues beyond everything
being rendered).

Resolution is therefore done across the whole set of documents about to be
rendered together, matching each sheet's exit points against every other
sheet's connection points:

    build_continuation_map([(name, doc), ...]) -> {name: [Continuation, ...]}

**A point is a hand-off when another sheet's piping lands on it.** That,
rather than whether the end looks open on this sheet, is the definition --
and it has to be, because a spool cut at a shop weld ends in a coupling
whose outboard end sits one modelled 1/16" root gap from its own inboard
end. That is inside the connection tolerance, so the connectivity walk reads
the boundary as *connected* and it never shows up as an open end at all.

``END-CONNECTION-PIPELINE`` records are a source of **names, not of exit
points**. Real exports declare them in pairs -- a CO-ORDS at the junction
plus the reference of the line that carries on -- and both halves survive
splitting, since ``_build_parts`` copies every connection record into each
part of its pipeline. Taken as exits they put a "continues on..." callout in
the middle of an unbroken run whenever both pipelines landed on one sheet.
They label a genuinely open end whose far side is on no drawing in the set:
the unresolved case, which the renderer reports as a warning because a line
reference must not be read as a sheet number.

``END-POSITION-OPEN``/``-CLOSED`` records are the file saying the line
*stops* there, so a point they mark never gets a continuation, and neither
does the closed end of a cap.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import geometry
from .model import PcfDocument
from .skeys import classify

Vec3 = tuple[float, float, float]

# Component classes whose open end is the end of the line, not a hand-off.
_TERMINATES = {"CAP", "BOLT", "GASKET", "SUPPORT", "MESSAGE", "CONNECTION"}

# Attributes an END-CONNECTION record names its continuing line in, in order
# of preference.
_REFERENCE_KEYS = ("PIPELINE-REFERENCE", "CONNECTION-REFERENCE",
                   "CONTINUATION-REFERENCE")


@dataclass
class Continuation:
    """One place where the piping leaves this drawing for another."""

    point: Vec3                     # the hand-off point, native coordinates
    drawing: str = ""               # the NEXT RENDERED DRAWING, when resolved
    reference: str = ""             # the PCF's own line reference (fallback)
    inboard: Vec3 | None = None     # a point up the run, to aim the arrow
    bore: float | None = None
    sheets: list[str] = field(default_factory=list)   # every drawing matched

    @property
    def resolved(self) -> bool:
        """True when a drawing in the set owns the other side of the joint."""
        return bool(self.drawing)

    @property
    def label(self) -> str:
        """What the callout prints: the drawing name, else the file's ref."""
        return self.drawing or self.reference


def _round(point: Vec3, quantum: float) -> tuple[int, int, int]:
    return (round(point[0] / quantum), round(point[1] / quantum),
            round(point[2] / quantum))


def reference_names(ref: str) -> list[str]:
    """An ISOGEN reference split into the names it shows, one per line.

    In PCF text ``$`` is ISOGEN's line-break within a printed field, so a
    connection reference like ``2105\\21053000$2105-BR3R-30001`` displays as
    two lines -- a sheet/area reference and the line it hands over to. Each
    line is stripped of any ``path\\sheet`` prefix so what survives is usable
    both as a pipeline name to match against and as clean display text.
    """
    names: list[str] = []
    for part in ref.split("$"):
        part = part.strip().rsplit("\\", 1)[-1].strip()
        if part:
            names.append(part)
    return names


class _PointIndex:
    """Spatial hash of every sheet's connection points, for cross matching.

    A job is routinely thirty sheets of three hundred points; matching each
    exit point against all of them pairwise is a hundred million distance
    tests per render pass, so points are bucketed by a tolerance-sized cell
    and only the 27 cells around a query are examined.
    """

    def __init__(self, tolerance: float):
        self.tolerance = tolerance
        self.cells: dict[tuple[int, int, int], list[tuple[Vec3, str]]] = {}

    def add(self, point: Vec3, sheet: str) -> None:
        self.cells.setdefault(_round(point, self.tolerance), []).append(
            (point, sheet))

    def near(self, point: Vec3):
        """Every (point, sheet) entry within tolerance of ``point``."""
        key = _round(point, self.tolerance)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for dz in (-1, 0, 1):
                    cell = (key[0] + dx, key[1] + dy, key[2] + dz)
                    for entry in self.cells.get(cell, ()):
                        if geometry.dist(point, entry[0]) <= self.tolerance:
                            yield entry

    def holds(self, point: Vec3) -> bool:
        return any(True for _ in self.near(point))

    def sheets_near(self, point: Vec3, exclude: str) -> list[str]:
        return sorted({sheet for _pt, sheet in self.near(point)
                       if sheet != exclude})

    def nearest(self, point: Vec3, exclude: str) -> float:
        """Distance to the closest entry not belonging to ``exclude``."""
        return min((geometry.dist(point, pt) for pt, sheet in self.near(point)
                    if sheet != exclude), default=float("inf"))


def _inboard_point(doc: PcfDocument, point: Vec3, tol: float) -> Vec3 | None:
    """A point up the run from ``point``, so the callout can aim outwards.

    The nearest fitting centre wins when there is one: at the weld end of an
    elbow the corner is up the leg the pipe actually runs along, while the
    elbow's far end is the diagonal across the bend and points nowhere the
    pipe goes. Straight bodies fall back to their far end.
    """
    best: Vec3 | None = None
    best_rank = (9, 0.0)
    for comp in doc.real_components():
        if not any(geometry.dist(pt.xyz, point) <= tol
                   for pt in comp.connection_points()):
            continue
        for rank, candidates in enumerate((
                [p.xyz for p in comp.centre_points],
                [p.xyz for p in comp.end_points])):
            for cand in candidates:
                span = geometry.dist(cand, point)
                if span <= tol:
                    continue
                if (rank, -span) < best_rank:
                    best, best_rank = cand, (rank, -span)
    return best


def _declared_references(doc: PcfDocument, tol: float) -> _PointIndex:
    """Where END-CONNECTION-PIPELINE records say the line carries on.

    Indexed by point so an open end can look up the name the file gives the
    line beyond it -- the fallback when no drawing in the set turns out to
    own the other side. Only ``-PIPELINE`` records qualify: a connection to
    equipment is where the line *ends*, not where a drawing continues.

    These are a source of *names*, never of exit points. A declared
    connection between two pipelines that both live on this sheet is not a
    continuation at all -- the piping runs straight through it -- so exits
    are found elsewhere (a point another sheet lands on, or an open end) and
    these only put a name to one.
    """
    index = _PointIndex(tol * 4)
    for comp in doc.all_components():
        if not comp.type.upper().startswith("END-CONNECTION-PIPELINE"):
            continue
        reference = ""
        for key in _REFERENCE_KEYS:
            value = (comp.get(key) or "").strip()
            if value:
                reference = value
                break
        if not reference:
            continue
        for pt in comp.co_ords:
            index.add(pt.xyz, reference)
    return index


def _stop_points(doc: PcfDocument) -> list[Vec3]:
    """Where the file declares the line ends (END-POSITION-OPEN/CLOSED)."""
    return [pt.xyz for comp in doc.all_components()
            if comp.type.upper().startswith("END-POSITION")
            for pt in comp.co_ords]


def _open_exits(doc: PcfDocument, tol: float) -> list[Vec3]:
    """Open ends that could hand off to another drawing."""
    out: list[Vec3] = []
    for comp, pt in geometry.open_ends(doc, tol):
        if classify(comp.type, comp.skey) in _TERMINATES:
            continue
        out.append(pt.xyz)
    return out


def build_continuation_map(
    sheets: list[tuple[str, PcfDocument]],
    *,
    tolerance: float | None = None,
) -> dict[str, list[Continuation]]:
    """Resolve every sheet's continuations against the rest of the set.

    ``sheets`` is ``[(drawing name, document), ...]`` -- the names that will
    appear on the finished drawings, which is what a continuation callout
    has to print. A single-sheet set is legal and still returns the
    continuations the file declares for itself, unresolved.
    """
    tol = tolerance
    if tol is None:
        tol = max((geometry.model_tolerance(doc) for _name, doc in sheets),
                  default=0.11)

    index = _PointIndex(tol * 4)
    for name, doc in sheets:
        for comp in doc.real_components():
            for pt in comp.connection_points():
                index.add(pt.xyz, name)

    out: dict[str, list[Continuation]] = {}
    for name, doc in sheets:
        doc_tol = geometry.model_tolerance(doc)
        stops = _stop_points(doc)
        declared = _declared_references(doc, doc_tol)
        found: list[Continuation] = []

        # A hand-off is defined by the *set*, not by one sheet: any point of
        # this sheet's piping that another sheet's piping also lands on is
        # where the two drawings meet, whatever the local connectivity walk
        # makes of it. That matters, because a spool cut at a shop weld ends
        # in a coupling whose outboard end sits one root gap -- less than the
        # connection tolerance -- from its own inboard end, so the graph
        # reads the boundary as connected and it is never an open end at all.
        #
        # Genuinely open ends are added on top, but only where the file
        # declares what carries on: that is the unresolved case, and without
        # a name there is nothing to print.
        #
        # Sorted closest to the neighbouring sheet first, so that when a joint
        # offers several coincident points the one that survives the
        # de-duplication below is the one on the boundary itself -- and both
        # sheets then print the same coordinates for the same joint.
        candidates: list[Vec3] = sorted(
            {pt.xyz for comp in doc.real_components()
             for pt in comp.connection_points()
             if index.sheets_near(pt.xyz, exclude=name)}
            | {point for point in _open_exits(doc, doc_tol)
               if declared.holds(point)},
            key=lambda p: (round(index.nearest(p, exclude=name), 6), p))

        accepted: list[Vec3] = []
        for point in candidates:
            # Matched at the connectivity tolerance, not a loose multiple of
            # it: an END-POSITION record names one exact end, and a spool can
            # easily carry another end half an inch from it that continues.
            if any(geometry.dist(point, stop) <= doc_tol for stop in stops):
                continue   # the file says the line ends here
            # One joint can offer several coincident points -- a coupling
            # end, the weld on it, the pipe beyond -- and they are one mark.
            if any(geometry.dist(point, seen) <= doc_tol * 4
                   for seen in accepted):
                continue
            matched = index.sheets_near(point, exclude=name)
            references = sorted({ref for _pt, ref in declared.near(point)})
            if not matched and not references:
                continue   # nothing continues here that anyone can name
            accepted.append(point)
            found.append(Continuation(
                point=point,
                drawing=" / ".join(matched),
                reference=" / ".join(references),
                inboard=_inboard_point(doc, point, doc_tol * 4),
                bore=_bore_at(doc, point, doc_tol * 4),
                sheets=matched,
            ))

        out[name] = sorted(found, key=lambda c: c.point)
    return out


def _bore_at(doc: PcfDocument, point: Vec3, tol: float) -> float | None:
    for comp in doc.real_components():
        for pt in comp.connection_points():
            if geometry.dist(pt.xyz, point) <= tol and pt.bore:
                return pt.bore
    return None
