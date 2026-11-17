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

import math
from dataclasses import dataclass, field, replace

from . import geometry
from .model import Component, PcfDocument, Point
from .skeys import classify
from .units import coord_to_inch_factor, format_bore, format_length
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
    raw_length: float               # native units; 弯制合并件 = 直段长 + 弯段弧长
    cut_length: float               # native units, after overrides/allowances
    length_source: str              # geometry | cut-piece-length | allowance
    direction: str
    end1_info: EndInfo | None = None
    end2_info: EndInfo | None = None
    branches: list[BranchInfo] = field(default_factory=list)
    uci: str | None = None
    sheet: str = ""      # 分页页号（"p1"；单页输出为空）
    comp_index: int | None = None   # position in doc.real_components(), for
    #                                 matching a pipe back to its drawn symbol

    def cut_length_text(self, units_coords: str) -> str:
        return format_length(self.cut_length, units_coords)


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
        if classify(other.type, other.skey) == "WELD":
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
        if classify(comp.type, comp.skey) != "OLET" or not comp.centre_points:
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
            nps=format_bore(bore, units_bore) if bore is not None else "?",
            distance_from_end1=max(t, 0.0),
            direction=direction,
        ))
    out.sort(key=lambda item: item.distance_from_end1)
    return out


def bend_arc_len(comp: Component) -> float:
    """弯段弧长（原生单位）：R·θ。END-POINT=切点、CENTRE-POINT=圆心、
    ANGLE 属性=角度×100（"9000"→90°）。缺角度回退弦长反解圆心角，
    缺圆心/半径回退弦长。"""
    eps = comp.end_points
    if len(eps) < 2:
        return 0.0
    ctr = comp.centre_points
    radius = ctr[0].distance_to(eps[0]) if ctr else None
    try:
        angle = float((comp.get("ANGLE") or "").split()[0]) / 100.0
    except (ValueError, IndexError, AttributeError):
        angle = None
    if radius and radius > 1e-9:
        if angle:
            return radius * math.radians(angle)
        chord = eps[0].distance_to(eps[1])
        if chord <= 2 * radius:
            return radius * 2.0 * math.asin(chord / (2.0 * radius))
    return eps[0].distance_to(eps[1])


class _UnionFind:
    def __init__(self, n: int) -> None:
        self.parent = list(range(n))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def _merge_bent_piece(
    members: list[int],
    bends: list[int],
    comps: list[Component],
    pipeline: str,
    graph: geometry.ConnectivityGraph,
    doc: PcfDocument,
    tolerance: float,
    inch_factor: float,
) -> CutPiece | None:
    """把同一根弯制整管的直段与弯段合成一件下料（CutPiece）。

    从一个自由端（不与组内弯段相接的管端）沿"直段→弯段→直段"走链到另一
    自由端，raw_length = Σ直段弦长 + Σ弯段弧长；CUT-PIECE-LENGTH 整根值
    优先（各记录共享同一值，取一次即可，逐记录累加会重复计料）。走链失败
    返回 None，调用方回退逐记录出件。
    """
    pipe_eps = {i: comps[i].end_points for i in members}
    bend_eps = {i: comps[i].end_points for i in bends}
    if any(len(eps) < 2 for eps in pipe_eps.values()) \
            or any(len(eps) < 2 for eps in bend_eps.values()):
        return None

    def near(p: Point, q: Point) -> bool:
        return p.distance_to(q) <= tolerance

    # 自由端：不与组内任何弯段端点相接的管端
    free = [(i, ep) for i in members for ep in pipe_eps[i]
            if not any(near(ep, bep) for b in bends for bep in bend_eps[b])]
    if len(free) != 2:
        return None

    cur_i, start_pt = free[0]
    used_pipes = {cur_i}
    used_bends: set[int] = set()
    chain: list[tuple[str, int, Point, Point]] = []
    while True:
        eps = pipe_eps[cur_i]
        if near(eps[0], start_pt if not chain else chain[-1][3]):
            seg_from, seg_to = eps[0], eps[1]
        elif near(eps[1], start_pt if not chain else chain[-1][3]):
            seg_from, seg_to = eps[1], eps[0]
        else:
            return None  # 走链脱节
        chain.append(("pipe", cur_i, seg_from, seg_to))
        # 接在这条直段远端的组内弯段
        nxt_bend = next((b for b in bends if b not in used_bends
                         and (near(bend_eps[b][0], seg_to)
                              or near(bend_eps[b][1], seg_to))), None)
        if nxt_bend is None:
            end_pt = seg_to  # 到达整管另一自由端
            break
        used_bends.add(nxt_bend)
        beps = bend_eps[nxt_bend]
        bend_to = beps[1] if near(beps[0], seg_to) else beps[0]
        chain.append(("bend", nxt_bend, seg_to, bend_to))
        # 弯段另一端接下一直段
        nxt_pipe = next((i for i in members if i not in used_pipes
                         and (near(pipe_eps[i][0], bend_to)
                              or near(pipe_eps[i][1], bend_to))), None)
        if nxt_pipe is None:
            return None
        cur_i = nxt_pipe
        used_pipes.add(cur_i)
    if used_pipes != set(members) or used_bends != set(bends):
        return None  # 组内有成员没走到（分叉/悬空）：不硬并

    raw_length = 0.0
    branches: list[BranchInfo] = []
    for kind, i, frm, to in chain:
        if kind == "pipe":
            # 分支位置从链方向上的近端量起，换算到整管自由端 1
            base = frm.distance_to(comps[i].end_points[0])
            for br in _branches_on(comps[i], comps, doc.units_bore, tolerance):
                branches.append(replace(
                    br, distance_from_end1=raw_length + base + br.distance_from_end1))
            raw_length += frm.distance_to(to)
        else:
            raw_length += bend_arc_len(comps[i])
    branches.sort(key=lambda br: br.distance_from_end1)

    length_source = "geometry"
    cut_length = raw_length
    # 整根下料长：任一记录的 CUT-PIECE-LENGTH（同 UCI 各记录共享同一值）
    for i in (*members, *bends):
        override = comps[i].get("CUT-PIECE-LENGTH")
        if override is not None:
            try:
                cut_length = float(override.split()[0])
                length_source = "cut-piece-length"
                break
            except (ValueError, IndexError):
                pass
    if length_source == "geometry":
        added = False
        for i in members:
            for allowance in comps[i].get_all("CUT-PIECE-ALLOWANCE"):
                tokens = allowance.split()
                if len(tokens) >= 2:
                    try:
                        cut_length += float(tokens[1])
                        added = True
                    except ValueError:
                        pass
        if added:
            length_source = "allowance"

    first_i, last_i = chain[0][1], chain[-1][1]
    info1 = _end_info(comps[first_i], chain[0][2], comps, graph, first_i,
                      tolerance, inch_factor)
    info2 = _end_info(comps[last_i], chain[-1][3], comps, graph, last_i,
                      tolerance, inch_factor)

    # 方向/锚定记录取最长直段（弯制件整体方向无意义，锚在图上最长一段）
    longest = max(members, key=lambda i: pipe_eps[i][0].distance_to(pipe_eps[i][1]))
    lcomp = comps[longest]
    bore = chain[0][2].bore if chain[0][2].bore is not None else lcomp.bore
    return CutPiece(
        label="",
        item_code=lcomp.item_code,
        description=lcomp.description,
        pipeline=pipeline,
        spool=lcomp.get("SPOOL-IDENTIFIER"),
        nps=format_bore(bore, doc.units_bore) if bore is not None else "?",
        bore=bore,
        end1=chain[0][2],
        end2=end_pt,
        raw_length=raw_length,
        cut_length=cut_length,
        length_source=length_source,
        direction=geometry.direction_name(geometry.sub(
            pipe_eps[longest][1].xyz, pipe_eps[longest][0].xyz)),
        end1_info=info1,
        end2_info=info2,
        branches=branches,
        uci=next((comps[i].uci for i in (*members, *bends) if comps[i].uci), None),
        comp_index=longest,
    )


def _bend_piece(
    comp: Component,
    idx: int,
    pipeline: str,
    comps: list[Component],
    graph: geometry.ConnectivityGraph,
    doc: PcfDocument,
    tolerance: float,
    inch_factor: float,
) -> CutPiece:
    """孤立弯制段单独出件：默认长度取弧长 R·θ，CUT-PIECE-LENGTH 优先。"""
    eps = comp.end_points
    end1, end2 = eps[0], eps[1]
    raw_length = bend_arc_len(comp)
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
    bore = end1.bore if end1.bore is not None else comp.bore
    return CutPiece(
        label="",
        item_code=comp.item_code,
        description=comp.description,
        pipeline=pipeline,
        spool=comp.get("SPOOL-IDENTIFIER"),
        nps=format_bore(bore, doc.units_bore) if bore is not None else "?",
        bore=bore,
        end1=end1,
        end2=end2,
        raw_length=raw_length,
        cut_length=cut_length,
        length_source=length_source,
        direction=geometry.direction_name(geometry.sub(end2.xyz, end1.xyz)),
        end1_info=_end_info(comp, end1, comps, graph, idx, tolerance, inch_factor),
        end2_info=_end_info(comp, end2, comps, graph, idx, tolerance, inch_factor),
        branches=_branches_on(comp, comps, doc.units_bore, tolerance),
        uci=comp.uci,
        comp_index=idx,
    )


def build_cutlist(doc: PcfDocument, *, extra_gap: float = 0.0) -> list[CutPiece]:
    """Compute cut pieces for every PIPE-class fabrication component, plus
    standalone BEND segments (bent-tube records with no same-code pipe on
    either side emit their own row so no material is lost).

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

    # 弯制整管合并：SP3D 把一根整料弯管按弯拆成多条 PIPE/BEND 记录，共享
    # 同一 UCI 与同一 CUT-PIECE-LENGTH（整根下料长、含弧段），段间无焊缝
    # （弯制非焊接）。同料号、同 SPOOL、经弯段相连的直管并成一件下料
    # （直段长 + 弯段弧长）；独立弯头（料号不属管材）照旧单列。
    pipe_codes = {c.item_code for c in comps
                  if classify(c.type, c.skey) == "PIPE" and c.item_code}
    uf = _UnionFind(len(comps))
    for idx, comp in enumerate(comps):
        if classify(comp.type, comp.skey) != "ELBOW":
            continue
        if comp.material_excluded() or not comp.item_code \
                or comp.item_code not in pipe_codes:
            continue  # 真弯头或边界副本：不并
        spool = comp.get("SPOOL-IDENTIFIER")
        for nidx in graph.neighbours(idx):
            other = comps[nidx]
            if classify(other.type, other.skey) != "PIPE" \
                    or other.material_excluded() \
                    or other.item_code != comp.item_code \
                    or other.get("SPOOL-IDENTIFIER") != spool:
                continue
            uf.union(idx, nidx)
    groups: dict[int, list[int]] = {}
    for i in range(len(comps)):
        groups.setdefault(uf.find(i), []).append(i)

    pieces: list[CutPiece] = []
    emitted: set[int] = set()
    for idx, comp in enumerate(comps):
        cls = classify(comp.type, comp.skey)
        if cls == "ELBOW":
            # 没并进任何整料件的弯制段（两边没有同料号直管，或走链回退）：
            # 单独一行出件——它是弯制管材不是采购弯头，下料必须有它。
            if idx in emitted or comp.material_excluded():
                continue
            if comp.type.upper() != "BEND":
                continue  # 采购弯头（ELBOW/ELBW 等）：管件不下料，只进 BOM
            eps = comp.end_points
            if len(eps) < 2:
                continue
            pieces.append(_bend_piece(comp, idx, pipeline_of[idx], comps,
                                      graph, doc, tolerance, inch_factor))
            continue
        if cls != "PIPE" or idx in emitted:
            continue
        if comp.material_excluded():
            continue  # 对侧边界直管副本（CONTINUATION/EXCLUDE），不进本图切割表
        eps = comp.end_points
        if len(eps) < 2:
            continue
        group = groups[uf.find(idx)]
        if len(group) > 1:
            cls_of = lambda i: classify(comps[i].type, comps[i].skey)
            merged = _merge_bent_piece(
                [i for i in group if cls_of(i) == "PIPE"],
                [i for i in group if cls_of(i) == "ELBOW"],
                comps, pipeline_of[idx], graph, doc, tolerance, inch_factor)
            if merged is not None:
                pieces.append(merged)
                emitted.update(group)
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
        nps = format_bore(bore, doc.units_bore) if bore is not None else "?"

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
