"""Continuation callouts: where the piping leaves this drawing.

Every callout comes from the PCF's own ``END-CONNECTION-PIPELINE``
declarations, verbatim: the record's CO-ORDS is the hand-off point, its
CONNECTION-REFERENCE names the line that carries on, and the callout
prints "SEE ISO <that name>" with the point's plant coordinates.
Nothing is inferred from geometry -- where the exporter says the sheet
hands over is where the drawing says it hands over, so a wrong PCF
prints a wrong callout. Re-deriving the boundary from coordinates was
tried and only multiplied marks: distance-split exports carry the same
boundary run (pipe, weld, reducer) in both files, so every joint of the
duplicated span matched.

The render set is consulted only for what a PCF cannot know: which
rendered sheet the declared neighbour actually is. The reference names a
*pipeline*; the sheet bearing it, and its SHT number when one pipeline
renders as several sheets, belong to this render pass -- that resolution
feeds the split-junction phantom (drawn from the neighbour document) and
the sheet tags. A declared neighbour no sheet in the set carries stays
as a name-only callout from the declaration text, and the renderer warns
about it. A declaration whose neighbour turns out to live on this very
sheet is not a continuation at all -- the piping runs straight through
it -- so it prints nothing.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import geometry
from .model import PcfDocument

Vec3 = tuple[float, float, float]

# Attributes an END-CONNECTION record names its continuing line in, in
# order of preference.
_REFERENCE_KEYS = ("CONNECTION-REFERENCE", "PIPELINE-REFERENCE",
                   "CONTINUATION-REFERENCE")


@dataclass
class Continuation:
    """One place where the piping leaves this drawing for another."""

    point: Vec3                     # the declared hand-off point (CO-ORDS)
    drawing: str = ""               # the NEXT RENDERED DRAWING, when resolved
    reference: str = ""             # the declared CONNECTION-REFERENCE text
    inboard: Vec3 | None = None     # a point up the run, to aim the arrow
    bore: float | None = None
    sheets: list[str] = field(default_factory=list)   # rendered sheets carrying the neighbour

    @property
    def resolved(self) -> bool:
        """True when a drawing in the set owns the other side of the joint."""
        return bool(self.drawing)

    @property
    def label(self) -> str:
        """What the callout prints: the drawing name, else the file's ref."""
        return self.drawing or self.reference


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


def _pipeline_refs(doc: PcfDocument) -> set[str]:
    """本图全部管线号（PIPELINE-REFERENCE 属性，去空白）。"""
    return {p.reference.strip() for p in doc.pipelines
            if p.reference and p.reference.strip()}


def _declared_handoffs(doc: PcfDocument) -> list[tuple[Vec3, str]]:
    """每条 END-CONNECTION-PIPELINE 声明原样列出：(CO-ORDS, 引用串)。

    只有 ``-PIPELINE`` 记录算接图：接设备（``-EQUIPMENT``）是管线
    到头，不是交给另一张图。一条记录可以带多个 CO-ORDS，逐点列出。
    """
    out: list[tuple[Vec3, str]] = []
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
        out.extend((pt.xyz, reference) for pt in comp.co_ords)
    return out


def build_continuation_map(
    sheets: list[tuple[str, PcfDocument]],
) -> dict[str, list[Continuation]]:
    """每张图的接图点，逐字来自各文件的 END-CONNECTION-PIPELINE 声明。

    ``sheets`` 是 ``[(drawing name, document), ...]``——整套渲染集。图
    名与各图的管线号只用于把声明的对侧线号对到渲染集里实际画着它
    的图张（分图 SHT 标签、幻影件素材、未解析告警都要它）；接图点
    本身一个都不来自几何匹配。单图渲染也合法：声明对不上任何图张
    时按未解析处理，标注照打，渲染端给 warning。
    """
    # 渲染集承载表：图名 / 管线号 → 画着它的图名集合
    carriers: dict[str, set[str]] = {}
    for name, doc in sheets:
        carriers.setdefault(name, set()).add(name)
        for ref in _pipeline_refs(doc):
            carriers.setdefault(ref, set()).add(name)

    out: dict[str, list[Continuation]] = {}
    for name, doc in sheets:
        tol = geometry.model_tolerance(doc) * 4
        # 同一点的重复声明（tol 内）合并为一个标注
        groups: list[list[tuple[Vec3, str]]] = []
        for point, reference in _declared_handoffs(doc):
            for group in groups:
                if geometry.dist(group[0][0], point) <= tol:
                    group.append((point, reference))
                    break
            else:
                groups.append([(point, reference)])

        found: list[Continuation] = []
        for group in groups:
            point = group[0][0]
            refs: list[str] = []
            for _pt, reference in group:
                if reference not in refs:
                    refs.append(reference)
            # 声明线号 → 渲染集图张
            known = False
            matched: set[str] = set()
            for reference in refs:
                for nm in reference_names(reference):
                    if nm in carriers:
                        known = True
                        matched |= carriers[nm]
            matched -= {name}
            if known and not matched:
                # 声明的对侧就画在本图上：管线直通，不是接图
                continue
            found.append(Continuation(
                point=point,
                drawing=" / ".join(sorted(matched)),
                reference=" / ".join(refs),
                inboard=_inboard_point(doc, point, tol),
                bore=_bore_at(doc, point, tol),
                sheets=sorted(matched),
            ))
        out[name] = sorted(found, key=lambda c: c.point)
    return out


def _bore_at(doc: PcfDocument, point: Vec3, tol: float) -> float | None:
    for comp in doc.real_components():
        for pt in comp.connection_points():
            if geometry.dist(pt.xyz, point) <= tol and pt.bore:
                return pt.bore
    return None
