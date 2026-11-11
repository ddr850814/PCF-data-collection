"""Deterministic pipe cut-length calculation.

Cut lengths come straight from END-POINT distance arithmetic -- no
inference. The rules, in order:

1. ``CUT-PIECE-LENGTH`` on the component wins outright (explicit override).
2. Otherwise length = |EP1 - EP2| in native units.
3. ``CUT-PIECE-ALLOWANCE EP1|EP2 <v>`` records are added per end.
4. An optional uniform ``extra_gap`` per welded end can be subtracted for
   dialects that model welds as zero-length records (classic Isogen); in the
   dominant Revit dialect the root gap is already modelled as a short "gap
   coupling" between the pipe ends, so the geometric length is already the
   cut length. Adjacent weld gaps found at each end are reported either way
   so nothing is silent.

Each cut piece also reports what sits at each end (description + connection)
and any set-on branches landing along the pipe (position from end 1,
bore, direction) so a fitter can lay out the piece without the drawing.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import geometry
from .model import Component, PcfDocument, Point
from .skeys import classify
from .units import bore_to_nps, coord_to_inch_factor, format_feet_inches, format_nps
from .welds import _is_gap_coupling


@dataclass
class EndInfo:
    connection: str | None          # PL/BW/SW/FL/... from the END-POINT record
    adjacent: str | None            # description of what the pipe joins
    adjacent_type: str | None
    weld_gap: float | None          # native units; gap modelled at this end


@dataclass
class BranchInfo:
    description: str | None
    bore: float | None
    nps: str
    distance_from_end1: float       # native units, along the pipe axis
    direction: str                  # compass/elevation direction of the branch


@dataclass
class CutPiece:
    label: str                      # e.g. "5A" = item-code 5, longest piece
    item_code: str | None
    description: str | None
    pipeline: str
    spool: str | None
    nps: str
    bore: float | None
    end1: Point
    end2: Point
    raw_length: float               # native units, |EP1-EP2|
    cut_length: float               # native units, after overrides/allowances
    length_source: str              # geometry | cut-piece-length | allowance
    direction: str
    end1_info: EndInfo | None = None
    end2_info: EndInfo | None = None
    branches: list[BranchInfo] = field(default_factory=list)
    uci: str | None = None
    comp_index: int | None = None   # position in doc.real_components(), for
    #                                 matching a pipe back to its drawn symbol

    def cut_length_text(self, units_coords: str) -> str:
        factor = coord_to_inch_factor(units_coords)
        return format_feet_inches(self.cut_length * factor)


def _end_info(
    pipe: Component,
    end: Point,
    comps: list[Component],
    graph: geometry.ConnectivityGraph,
    pipe_idx: int,
    tolerance: float,
    inch_factor: float,
) -> EndInfo:
    adjacent = None
    adjacent_type = None
    weld_gap = None
    # Walk through any weld/gap-coupling at this end to the real neighbour.
    for other_idx in sorted(graph.neighbours(pipe_idx)):
        other = comps[other_idx]
        touching = any(end.distance_to(op) <= tolerance for op in other.connection_points())
        if not touching:
            continue
        if classify(other.type) == "WELD":
            eps = other.end_points
            if len(eps) >= 2:
                gap = eps[0].distance_to(eps[1])
                weld_gap = gap if gap > 1e-9 else 0.0
            else:
                weld_gap = 0.0
            continue
        if _is_gap_coupling(other, inch_factor):
            eps = other.end_points
            weld_gap = eps[0].distance_to(eps[1])
            continue
        if adjacent is None:
            adjacent = other.description or other.type
            adjacent_type = other.type
    return EndInfo(
        connection=end.connection_type,
        adjacent=adjacent,
        adjacent_type=adjacent_type,
        weld_gap=weld_gap,
    )


def _branches_on(
    pipe: Component, comps: list[Component], units_bore: str, tolerance: float
) -> list[BranchInfo]:
    eps = pipe.end_points
    if len(eps) < 2:
        return []
    a, b = eps[0].xyz, eps[1].xyz
    axis = geometry.sub(b, a)
    axis_len = geometry.length(axis)
    if axis_len < 1e-9:
        return []
    unit = geometry.normalize(axis)

    out: list[BranchInfo] = []
    for comp in comps:
        if classify(comp.type) != "OLET" or not comp.centre_points:
            continue
        centre = comp.centre_points[0].xyz
        t = geometry.dot(geometry.sub(centre, a), unit)
        if t < -tolerance or t > axis_len + tolerance:
            continue
        on_axis = geometry.add(a, geometry.scale(unit, t))
        if geometry.dist(on_axis, centre) > tolerance * 4:
            continue
        branch_pt = comp.branch1_points[0] if comp.branch1_points else None
        direction = "NONE"
        bore = comp.bore
        if branch_pt is not None:
            direction = geometry.direction_name(geometry.sub(branch_pt.xyz, centre))
            bore = branch_pt.bore if branch_pt.bore is not None else bore
        out.append(BranchInfo(
            description=comp.description,
            bore=bore,
            nps=format_nps(bore_to_nps(bore, units_bore)) if bore is not None else "?",
            distance_from_end1=max(t, 0.0),
            direction=direction,
        ))
    out.sort(key=lambda item: item.distance_from_end1)
    return out


def build_cutlist(doc: PcfDocument, *, extra_gap: float = 0.0) -> list[CutPiece]:
    """Compute cut pieces for every PIPE-class fabrication component.

    extra_gap: native-unit gap subtracted per end that has a zero-length
    WELD record (classic Isogen dialect). Leave at 0 when gaps are already
    modelled in the geometry (the default and the dominant dialect).
    """
    inch_factor = coord_to_inch_factor(doc.units_coords)
    tolerance = geometry.model_tolerance(doc)

    comps: list[Component] = []
    pipeline_of: list[str] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append(comp)
            pipeline_of.append(pipeline.reference)

    graph = geometry.ConnectivityGraph(comps, tolerance)

    pieces: list[CutPiece] = []
    for idx, comp in enumerate(comps):
        if classify(comp.type) != "PIPE":
            continue
        eps = comp.end_points
        if len(eps) < 2:
            continue
        end1, end2 = eps[0], eps[1]
        raw_length = end1.distance_to(end2)

        length_source = "geometry"
        cut_length = raw_length

        override = comp.get("CUT-PIECE-LENGTH")
        if override is not None:
            try:
                cut_length = float(override.split()[0])
                length_source = "cut-piece-length"
            except (ValueError, IndexError):
                pass

        if length_source == "geometry":
            for allowance in comp.get_all("CUT-PIECE-ALLOWANCE"):
                tokens = allowance.split()
                if len(tokens) >= 2:
                    try:
                        cut_length += float(tokens[1])
                        length_source = "allowance"
                    except ValueError:
                        pass

        info1 = _end_info(comp, end1, comps, graph, idx, tolerance, inch_factor)
        info2 = _end_info(comp, end2, comps, graph, idx, tolerance, inch_factor)

        if extra_gap and length_source == "geometry":
            for info in (info1, info2):
                if info.weld_gap is not None and info.weld_gap <= 1e-9:
                    cut_length -= extra_gap
                    length_source = "geometry-minus-gap"

        bore = end1.bore if end1.bore is not None else comp.bore
        nps = format_nps(bore_to_nps(bore, doc.units_bore)) if bore is not None else "?"

        pieces.append(CutPiece(
            label="",
            item_code=comp.item_code,
            description=comp.description,
            pipeline=pipeline_of[idx],
            spool=comp.get("SPOOL-IDENTIFIER"),
            nps=nps,
            bore=bore,
            end1=end1,
            end2=end2,
            raw_length=raw_length,
            cut_length=cut_length,
            length_source=length_source,
            direction=geometry.direction_name(geometry.sub(end2.xyz, end1.xyz)),
            end1_info=info1,
            end2_info=info2,
            branches=_branches_on(comp, comps, doc.units_bore, tolerance),
            uci=comp.uci,
            comp_index=idx,
        ))

    _assign_labels(pieces)
    return pieces


def _assign_labels(pieces: list[CutPiece]) -> None:
    """Label pieces '<item-code><letter>', letters longest-first per code."""
    by_code: dict[str, list[CutPiece]] = {}
    for piece in pieces:
        by_code.setdefault(piece.item_code or "?", []).append(piece)
    for code, group in by_code.items():
        ordered = sorted(group, key=lambda p: -p.cut_length)
        if len(ordered) == 1:
            ordered[0].label = code
            continue
        for n, piece in enumerate(ordered):
            suffix = _letter_suffix(n)
            piece.label = f"{code}{suffix}"


def _letter_suffix(n: int) -> str:
    letters = ""
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters
