"""Weld derivation and deterministic numbering.

Real-world PCF exports encode welds three ways, all handled here:

1. Explicit ``WELD`` components.
2. "Gap couplings": COUPLING/UNION/CONNECTOR/CLAMP records whose description
   contains "weld" and whose end-points are at most ~1/4" apart -- the span
   between the end-points *is* the root gap (the dominant Revit-exporter
   dialect).
3. Set-on branch attachments (TEE-SET-ON / OLET / ELBOLET): the fitting is
   welded to the header at its CENTRE-POINT.

Optionally, coincident end-points of two weldable components with no weld
record between them are reported as field-fit welds (FFW).

Numbering is deterministic: welds are sorted by (pipeline order, component
order, position) and numbered 1..N, so the same file always yields the same
weld numbers.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import geometry
from .model import Component, PcfDocument
from .skeys import WELD_JOINT_TYPES, classify, joint_from_skey
from .units import coord_to_inch_factor

# A coupling this short (in inches) whose description mentions "weld" is a
# weld with a modelled root gap, not a physical coupling.
GAP_COUPLING_MAX_INCHES = 0.26

# End-point pairs closer than this (inches) with no weld between them are
# candidate field-fit welds.
PROXIMITY_INCHES = 0.05

_WELDABLE_CLASSES = {"PIPE", "ELBOW", "TEE", "REDUCER", "CAP", "OLET", "VALVE", "FILTER"}
_NON_WELD_CONNECTIONS = {"FL", "SC", "CP"}


@dataclass
class Weld:
    number: int
    label: str                    # SW | FW | FFW
    joint: str | None             # BW | SW | FW | THD ... best-effort
    source: str                   # explicit | coupling | set-on | proximity
    x: float
    y: float
    z: float
    bore: float | None
    gap: float | None             # native units, None when not modelled
    pipeline: str
    description: str | None
    uci: str | None
    neighbors: list[str] = field(default_factory=list)
    # Geometry the renderer needs to draw a socket weld the right way round.
    # `insert_toward` is a point on the side the pipe/nipple/swage enters
    # from -- a socket weld symbol is a bracket that opens towards the
    # inserted member. None when no neighbour is pipe-like, or for a set-on
    # branch, in which case the symbol falls back to a plain dot.
    insert_toward: tuple[float, float, float] | None = None
    # Populated only when derive_welds is given a spool_key: the weld's spool
    # ("SP-1"), or both spools for an inter-spool joint ("SP-1 / SP-2"), in
    # which case notes carries the "Field Weld" declaration.
    spool: str = ""
    notes: str = ""

    @property
    def xyz(self) -> tuple[float, float, float]:
        return (self.x, self.y, self.z)


def _weld_label(description: str | None) -> str:
    if description and "field" in description.lower():
        return "FW"
    return "SW"


def _is_gap_coupling(comp: Component, inch_factor: float) -> bool:
    if classify(comp.type, comp.skey) != "COUPLING":
        return False
    desc = (comp.description or "").lower()
    if "weld" not in desc:
        return False
    eps = comp.end_points
    if len(eps) < 2:
        return False
    return eps[0].distance_to(eps[1]) * inch_factor <= GAP_COUPLING_MAX_INCHES


def _neighbor_names(
    doc_components: list[Component],
    graph: geometry.ConnectivityGraph,
    idx: int,
) -> list[str]:
    names = []
    for other_idx in sorted(graph.neighbours(idx)):
        other = doc_components[other_idx]
        if classify(other.type, other.skey) in ("COUPLING", "CONNECTION", "MESSAGE"):
            continue
        label = other.description or other.type
        if label not in names:
            names.append(label)
    return names[:2]


def _insert_toward(
    doc_components: list[Component],
    graph: geometry.ConnectivityGraph,
    idx: int,
    comp: Component,
) -> tuple[float, float, float] | None:
    """Which end of this joint the pipe/nipple/swage enters from.

    A socket weld is drawn as a bracket opening towards the inserted member,
    so the symbol is only right if we know which side that is. We look for a
    neighbouring PIPE and return whichever of this component's two ends is
    nearest to it. Returns None when neither side is pipe -- a fitting welded
    straight to a fitting has no inserted member to point at.
    """
    eps = comp.end_points
    if len(eps) < 2:
        return None
    pipe_pts: list[tuple[float, float, float]] = []
    for other_idx in graph.neighbours(idx):
        other = doc_components[other_idx]
        if classify(other.type, other.skey) != "PIPE":
            continue
        pipe_pts.extend(p.xyz for p in other.connection_points())
    if not pipe_pts:
        return None
    a, b = eps[0].xyz, eps[-1].xyz
    da = min(geometry.dist(a, p) for p in pipe_pts)
    db = min(geometry.dist(b, p) for p in pipe_pts)
    if abs(da - db) < 1e-9:
        return None
    return a if da < db else b


def component_spools(
    doc: PcfDocument,
    spool_key: str,
    *,
    graph: geometry.ConnectivityGraph | None = None,
) -> list[str | None]:
    """Per-real-component spool value, flooded so unlabelled components
    (bare welds, gap couplings) take their nearest labelled neighbour's.

    ``spool_key`` is a grouping attribute (SPOOL-IDENTIFIER, a line number)
    or the PIPELINE-REFERENCE pseudo-key. A document with no values at all
    yields all None -- callers degrade gracefully rather than fail.
    """
    key_u = spool_key.strip().upper()
    comps: list[Component] = []
    values: list[str | None] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append(comp)
            if key_u in ("PIPELINE-REFERENCE", "PIPELINE"):
                raw = pipeline.reference
            else:
                raw = comp.get(key_u)
            raw = raw.strip() if raw else ""
            values.append(raw or None)
    if not any(v is not None for v in values):
        return values
    return geometry.flood_values(
        comps, values, geometry.model_tolerance(doc), graph=graph)


def derive_welds(
    doc: PcfDocument, *, detect_ffw: bool = False, spool_key: str | None = None
) -> list[Weld]:
    inch_factor = coord_to_inch_factor(doc.units_coords)
    tolerance = geometry.model_tolerance(doc)

    comps: list[Component] = []
    pipeline_of: list[str] = []
    for pipeline in doc.pipelines:
        for comp in pipeline.real_components():
            comps.append(comp)
            pipeline_of.append(pipeline.reference)

    graph = geometry.ConnectivityGraph(comps, tolerance)

    raw: list[tuple[tuple, dict]] = []  # (sort_key, weld kwargs)

    last_skey: str | None = None
    for idx, comp in enumerate(comps):
        ctype = classify(comp.type, comp.skey)
        skey = comp.skey
        if skey and ctype != "WELD" and not _is_gap_coupling(comp, inch_factor):
            last_skey = skey

        if ctype == "WELD":
            if comp.material_excluded():
                continue  # 边界焊（EXCLUDE+DOTTED-UNDIMENSIONED）：对侧所有，不编号不进焊表
            eps = comp.end_points
            if eps:
                pos = geometry.midpoint(eps[0].xyz, eps[-1].xyz) if len(eps) > 1 else eps[0].xyz
                gap = eps[0].distance_to(eps[-1]) if len(eps) > 1 else 0.0
                bore = eps[0].bore
            else:
                cos = comp.co_ords
                if not cos:
                    continue
                pos, gap, bore = cos[0].xyz, None, cos[0].bore
            joint = joint_from_skey(skey) or joint_from_skey(last_skey)
            if joint not in WELD_JOINT_TYPES:
                joint = "BW"
            raw.append(((idx, 0), dict(
                label=_weld_label(comp.description), joint=joint, source="explicit",
                x=pos[0], y=pos[1], z=pos[2], bore=bore,
                gap=gap if gap and gap > 1e-9 else None,
                pipeline=pipeline_of[idx], description=comp.description, uci=comp.uci,
                neighbors=_neighbor_names(comps, graph, idx),
                insert_toward=_insert_toward(comps, graph, idx, comp),
            )))
            continue

        if _is_gap_coupling(comp, inch_factor):
            eps = comp.end_points
            pos = geometry.midpoint(eps[0].xyz, eps[1].xyz)
            gap = eps[0].distance_to(eps[1])
            desc = (comp.description or "").lower()
            if "socket" in desc:
                joint = "SW"
            elif "slip-on" in desc or "slip on" in desc:
                joint = "FW"
            else:
                carried = joint_from_skey(last_skey)
                joint = carried if carried in WELD_JOINT_TYPES else "BW"
            raw.append(((idx, 0), dict(
                label=_weld_label(comp.description), joint=joint, source="coupling",
                x=pos[0], y=pos[1], z=pos[2], bore=eps[0].bore,
                gap=gap if gap > 1e-9 else None,
                pipeline=pipeline_of[idx], description=comp.description, uci=comp.uci,
                neighbors=_neighbor_names(comps, graph, idx),
                insert_toward=_insert_toward(comps, graph, idx, comp),
            )))
            continue

        if ctype == "OLET" and comp.centre_points:
            if comp.material_excluded():
                # 对侧支管台副本（CONTINUATION/EXCLUDE）：其 set-on 焊属于
                # 归属图（那边有显式 WELD 记录），本图不编号不画点
                continue
            centre = comp.centre_points[0]
            branch_bore = comp.branch1_points[0].bore if comp.branch1_points else comp.bore
            raw.append(((idx, 0), dict(
                label="SW", joint="FW", source="set-on",
                x=centre.x, y=centre.y, z=centre.z, bore=branch_bore, gap=None,
                pipeline=pipeline_of[idx], description=comp.description, uci=comp.uci,
                neighbors=_neighbor_names(comps, graph, idx),
            )))

    welds_positions = [(kw["x"], kw["y"], kw["z"]) for _, kw in raw]

    if detect_ffw:
        prox_tol = max(PROXIMITY_INCHES / inch_factor, tolerance)
        for idx, comp in enumerate(comps):
            if classify(comp.type, comp.skey) not in _WELDABLE_CLASSES:
                continue
            for pt in comp.end_points:
                if pt.connection_type in _NON_WELD_CONNECTIONS:
                    continue
                for other_idx in graph.neighbours(idx):
                    if other_idx <= idx:
                        continue
                    other = comps[other_idx]
                    if classify(other.type, other.skey) not in _WELDABLE_CLASSES:
                        continue
                    for opt in other.end_points:
                        if opt.connection_type in _NON_WELD_CONNECTIONS:
                            continue
                        if pt.distance_to(opt) > prox_tol:
                            continue
                        mid = geometry.midpoint(pt.xyz, opt.xyz)
                        near_existing = any(
                            geometry.dist(mid, wp) <= prox_tol * 4 for wp in welds_positions
                        )
                        if near_existing:
                            continue
                        welds_positions.append(mid)
                        raw.append(((idx, 1), dict(
                            label="FFW", joint="BW", source="proximity",
                            x=mid[0], y=mid[1], z=mid[2], bore=pt.bore, gap=None,
                            pipeline=pipeline_of[idx], description=None, uci=None,
                            neighbors=[comp.description or comp.type,
                                       other.description or other.type],
                        )))

    raw.sort(key=lambda item: item[0])
    welds = [Weld(number=n, **kw) for n, (_, kw) in enumerate(raw, start=1)]

    if spool_key:
        spool_of = component_spools(doc, spool_key, graph=graph)
        for weld, (sort_key, _kw) in zip(welds, raw):
            idx = sort_key[0]
            # The spools this joint touches: the weld record's own flooded
            # value plus every neighbour's. Two distinct spools meeting at
            # one joint is the definition of a field weld between spools.
            touching = {spool_of[idx]}
            touching.update(spool_of[n] for n in graph.neighbours(idx))
            touching.discard(None)
            if len(touching) > 1:
                weld.spool = " / ".join(sorted(touching))
                weld.notes = "Field Weld"
                if weld.label != "FFW":
                    weld.label = "FW"
            else:
                weld.spool = next(iter(touching), "") or ""
    return welds
