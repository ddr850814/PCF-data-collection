"""Splitting and merging PCF documents.

Split never mutates the source document: it returns new documents whose
components are re-parsed copies. A split point may land on a joint between
two components (the connection is severed there) or inside a PIPE (the pipe
is cut into two pieces with new END-POINTs at the cut).

Merge concatenates documents, reconciles the MATERIALS sections (re-coding
colliding ITEM-CODEs), and reports whether the merged models actually touch;
merging non-adjacent piping is allowed but warned about.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field

from . import geometry
from .model import POINT_KEYS, AttrLine, Component, PcfDocument, Pipeline, Point
from .parser import parse_document, write_document

Vec3 = tuple[float, float, float]

# Pseudo-key accepted by split_by_key/survey: group by the owning
# PIPELINE-REFERENCE block rather than a per-component attribute.
PIPELINE_KEY = "PIPELINE-REFERENCE"


class SplitError(ValueError):
    pass


class MergeError(ValueError):
    pass


@dataclass
class SplitResult:
    parts: list[PcfDocument]
    notes: list[str] = field(default_factory=list)
    # For key-based splits: the grouping value behind each part, aligned
    # with ``parts`` (None for geometric splits, which have no labels).
    labels: list[str] | None = None


@dataclass
class MergeResult:
    document: PcfDocument
    warnings: list[str] = field(default_factory=list)
    item_code_remaps: dict[str, dict[str, str]] = field(default_factory=dict)


def _copy_document(doc: PcfDocument) -> PcfDocument:
    return parse_document(write_document(doc), source_path=doc.source_path)


def _format_coord(value: float) -> str:
    return f"{value:.4f}"


def _point_line(key: str, xyz: Vec3, bore: float | None, flags: tuple[str, ...]) -> str:
    parts = [f"    {key}"]
    parts.extend(_format_coord(v) for v in xyz)
    if bore is not None:
        parts.append(_format_coord(bore))
    parts.extend(flags)
    return "  ".join(parts)


def _split_pipe(comp: Component, cut: Vec3) -> tuple[Component, Component]:
    """Cut a PIPE component at a point on its axis into two pipes."""
    eps = comp.end_points
    end1, end2 = eps[0], eps[1]

    def clone_with_ends(a: Point | Vec3, b: Point | Vec3) -> Component:
        new = Component(type=comp.type)
        ep_index = 0
        for attr in comp.attributes:
            if attr.key == "END-POINT":
                source = a if ep_index == 0 else b
                if isinstance(source, Point):
                    new.attributes.append(AttrLine(attr.raw))
                else:
                    template = end1 if ep_index == 0 else end2
                    new.attributes.append(AttrLine(
                        _point_line("END-POINT", source, template.bore, template.flags)
                    ))
                ep_index += 1
            else:
                new.attributes.append(AttrLine(attr.raw))
        return new

    first = clone_with_ends(end1, cut)
    second = clone_with_ends(cut, end2)
    # The second pipe's first END-POINT template used end1's raw line; rebuild
    # it from the cut point explicitly.
    for attr in second.attributes:
        if attr.key == "END-POINT":
            attr.raw = _point_line("END-POINT", cut, end1.bore, end1.flags)
            break
    return first, second


def _point_on_pipe(comp: Component, point: Vec3, tolerance: float) -> Vec3 | None:
    """Return the projection of point onto the pipe axis if strictly inside."""
    eps = comp.end_points
    if len(eps) < 2:
        return None
    a, b = eps[0].xyz, eps[1].xyz
    axis = geometry.sub(b, a)
    axis_len = geometry.length(axis)
    if axis_len < tolerance:
        return None
    unit = geometry.normalize(axis)
    t = geometry.dot(geometry.sub(point, a), unit)
    if t <= tolerance or t >= axis_len - tolerance:
        return None
    on_axis = geometry.add(a, geometry.scale(unit, t))
    if geometry.dist(on_axis, point) > tolerance * 4:
        return None
    return on_axis


def split_document(
    doc: PcfDocument,
    cut_points: list[Vec3],
    *,
    tolerance: float | None = None,
) -> SplitResult:
    """Split a document at the given points into connected parts."""
    if not cut_points:
        raise SplitError("no cut points given")

    work = _copy_document(doc)
    tol = tolerance if tolerance is not None else geometry.model_tolerance(work)
    notes: list[str] = []

    # First pass: cut any pipes that a cut point lands inside.
    for cut in cut_points:
        for pipeline in work.pipelines:
            for pos, comp in enumerate(pipeline.components):
                if comp.is_connection_record:
                    continue
                on_axis = _point_on_pipe(comp, cut, tol)
                if on_axis is None:
                    continue
                first, second = _split_pipe(comp, on_axis)
                pipeline.components[pos:pos + 1] = [first, second]
                notes.append(
                    f"cut {comp.type} (item {comp.item_code or '?'}) at "
                    f"({on_axis[0]:.4f}, {on_axis[1]:.4f}, {on_axis[2]:.4f})"
                )
                break
            else:
                continue
            break

    comps: list[Component] = []
    owner: list[tuple[int, int]] = []  # (pipeline idx, component idx)
    for p_idx, pipeline in enumerate(work.pipelines):
        for c_idx, comp in enumerate(pipeline.components):
            if comp.is_connection_record:
                continue
            comps.append(comp)
            owner.append((p_idx, c_idx))

    graph = geometry.ConnectivityGraph(comps, tol)

    # Sever adjacency at every cut point using a cut plane: components near
    # the cut are classified to one side or the other along the joint axis,
    # and only edges crossing the plane are severed. (Pairwise severing would
    # isolate the short weld coupling that typically spans the joint.)
    for cut in cut_points:
        at_cut = geometry.components_at_point(comps, cut, tol * 4)
        if not at_cut:
            continue
        axis = _joint_axis(comps, at_cut, cut)
        side: dict[int, float] = {}
        for i in at_cut:
            pts = [p.xyz for p in comps[i].all_points()]
            centroid = (
                sum(p[0] for p in pts) / len(pts),
                sum(p[1] for p in pts) / len(pts),
                sum(p[2] for p in pts) / len(pts),
            )
            projection = geometry.dot(geometry.sub(centroid, cut), axis)
            if abs(projection) <= tol:
                # Spans the cut (the weld itself): use its first point.
                first = pts[0]
                projection = geometry.dot(geometry.sub(first, cut), axis) or -1.0
            side[i] = 1.0 if projection >= 0 else -1.0
        for i in at_cut:
            for j in list(graph.adjacency.get(i, ())):
                if j in side and side[j] != side[i]:
                    graph.adjacency[i].discard(j)
                    graph.adjacency[j].discard(i)

    groups = graph.connected_sets()
    if len(groups) == 1:
        raise SplitError(
            "cut points did not separate the model; verify each point sits on "
            "a joint between components or inside a pipe"
        )

    # Primary groups are the ones the cuts actually touched; any other group
    # (piping that was already disconnected in the source, floating
    # instruments, ...) rides along with the nearest primary group instead of
    # becoming its own file.
    touched: set[int] = set()
    for cut in cut_points:
        touched.update(geometry.components_at_point(comps, cut, tol * 4))
    primaries = [g for g in groups if g & touched]
    riders = [g for g in groups if not (g & touched)]
    if len(primaries) < 2:
        primaries = groups
        riders = []
    for rider in riders:
        best = min(primaries, key=lambda g: _group_distance(comps, rider, g))
        best |= rider
        notes.append(
            f"attached {len(rider)} already-disconnected component(s) to the "
            "nearest part"
        )
    groups = primaries
    notes.append(f"model separated into {len(groups)} parts")

    # Deterministic part order: by lowest component index in each group.
    groups.sort(key=min)

    return SplitResult(parts=_build_parts(work, groups, owner), notes=notes)


def _build_parts(
    work: PcfDocument,
    groups: list[set[int]],
    owner: list[tuple[int, int]],
) -> list[PcfDocument]:
    """One new document per component group (header/materials shared)."""
    # 管口等连接记录不参与连通分组和材料统计，但它们是图面的一部分：
    # 按坐标跟随最近的实组件，且只进入其所在分图，不能复制到每一张图。
    component_group = {
        owner[i]: group_idx
        for group_idx, group in enumerate(groups)
        for i in group
    }
    positions = list(component_group)
    real_points = {
        pos: [p.xyz for p in work.pipelines[pos[0]].components[pos[1]].all_points()]
        for pos in positions
    }
    record_groups: dict[tuple[int, int], int] = {}
    for p_idx, pipeline in enumerate(work.pipelines):
        for c_idx, comp in enumerate(pipeline.components):
            if not comp.is_connection_record:
                continue
            rec_points = [p.xyz for p in comp.all_points()]
            best_position: tuple[int, int] | None = None
            if rec_points:
                best_distance = float("inf")
                for pos in positions:
                    points = real_points[pos]
                    if not points:
                        continue
                    distance = min(
                        geometry.dist(a, b) for a in rec_points for b in points
                    )
                    if distance < best_distance:
                        best_distance = distance
                        best_position = pos
            if best_position is None:
                # 无坐标的 MESSAGE 等记录跟随同管段中文件位置最近的实组件。
                candidates = [pos for pos in positions if pos[0] == p_idx] or positions
                if candidates:
                    best_position = min(
                        candidates,
                        key=lambda pos: (abs(pos[1] - c_idx), pos[0], pos[1]),
                    )
            if best_position is not None:
                record_groups[(p_idx, c_idx)] = component_group[best_position]

    parts: list[PcfDocument] = []
    for group_idx, group in enumerate(groups):
        part = PcfDocument(
            header_lines=list(work.header_lines),
            newline=work.newline,
            source_path=None,
        )
        part._had_trailing_newline = True  # type: ignore[attr-defined]
        keep = {owner[i] for i in group}
        kept_positions = {owner[i][0] for i in group}
        for p_idx, pipeline in enumerate(work.pipelines):
            if p_idx not in kept_positions:
                continue
            new_pipeline = Pipeline(
                reference=pipeline.reference,
                header_lines=[AttrLine(a.raw) for a in pipeline.header_lines],
                raw_line=pipeline.raw_line,
            )
            for c_idx, comp in enumerate(pipeline.components):
                if comp.is_connection_record:
                    if record_groups.get((p_idx, c_idx)) != group_idx:
                        continue
                elif (p_idx, c_idx) not in keep:
                    continue
                new_pipeline.components.append(Component(
                    type=comp.type,
                    attributes=[AttrLine(a.raw) for a in comp.attributes],
                ))
            part.pipelines.append(new_pipeline)
        part.materials_lines = list(work.materials_lines)
        parts.append(part)
    return parts


def default_split_distance(units_coords: str) -> float:
    """The stock sheet-separation distance: 1 inch, or 25 mm for metric files,
    expressed in the file's native coordinate units."""
    from .units import coord_to_inch_factor, coord_to_mm_factor

    if coord_to_inch_factor(units_coords) == 1.0:
        return 1.0
    return round(25.0 / coord_to_mm_factor(units_coords), 9)


def split_by_distance(
    doc: PcfDocument,
    distance: float | None = None,
) -> SplitResult:
    """Group components whose geometry lies within ``distance`` of each other
    (native units); each group becomes its own document ("sheet").

    Deterministic: same file + same distance -> same grouping and order.
    One group means no split is needed (the single part is still returned so
    callers can preview either way).
    """
    work = _copy_document(doc)
    if distance is None:
        distance = default_split_distance(work.units_coords)
    if distance <= 0:
        raise SplitError(f"split distance must be positive, got {distance}")

    comps: list[Component] = []
    owner: list[tuple[int, int]] = []
    for p_idx, pipeline in enumerate(work.pipelines):
        for c_idx, comp in enumerate(pipeline.components):
            if comp.is_connection_record:
                continue
            comps.append(comp)
            owner.append((p_idx, c_idx))
    if not comps:
        raise SplitError("document contains no components")

    graph = geometry.ConnectivityGraph(comps, distance)
    groups = graph.connected_sets()
    groups.sort(key=min)
    notes = [
        f"{len(groups)} group(s) at separation distance "
        f"{distance:g} {work.units_coords}"
    ]
    return SplitResult(parts=_build_parts(work, groups, owner), notes=notes)


@dataclass
class KeySurvey:
    """Census of one scalar attribute across a document's real components."""

    key: str
    present: int            # components carrying a non-empty value
    total: int              # real components in the document
    distinct: int           # distinct values
    values: dict[str, int]  # value -> component count
    # Fraction of value-groups that are a single connected island (1.0 =
    # every value maps to one contiguous piece of piping -- the signature
    # of a true spool/line grouping). None when not computed.
    cohesion: float | None = None

    @property
    def coverage(self) -> float:
        return self.present / self.total if self.total else 0.0

    @property
    def is_candidate(self) -> bool:
        """Usable as a split key: on every component, groups, not unique."""
        return self.present == self.total and 1 < self.distinct < self.total


def _collect_key_values(doc: PcfDocument) -> dict[str, dict[int, str]]:
    """Map attribute key -> {component index -> value} in one text pass."""
    comps = doc.real_components()
    assign: dict[str, dict[int, str]] = {}
    for i, comp in enumerate(comps):
        seen_here: set[str] = set()
        for attr in comp.attributes:
            k = attr.key
            if not k or k in POINT_KEYS or k.startswith("PORT-POINT"):
                continue
            if k in seen_here:
                continue  # count a repeated key once per component
            value = attr.value.strip()
            if not value:
                continue  # flag attributes (FABRICATION-ITEM...) group nothing
            seen_here.add(k)
            assign.setdefault(k, {})[i] = value
    # Pseudo-key: the owning PIPELINE-REFERENCE block always covers 100%.
    by_pipeline: dict[int, str] = {}
    idx = 0
    for pipeline in doc.pipelines:
        for _comp in pipeline.real_components():
            by_pipeline[idx] = pipeline.reference or "(unnamed)"
            idx += 1
    assign[PIPELINE_KEY] = by_pipeline
    return assign


def _cohesion(graph: geometry.ConnectivityGraph, mapping: dict[int, str]) -> float:
    """Fraction of value-groups forming a single connected island."""
    groups: dict[str, set[int]] = defaultdict(set)
    for i, value in mapping.items():
        groups[value].add(i)
    coherent = 0
    for members in groups.values():
        start = min(members)
        seen = {start}
        frontier = [start]
        while frontier:
            cur = frontier.pop()
            for nxt in graph.neighbours(cur):
                if nxt in members and nxt not in seen:
                    seen.add(nxt)
                    frontier.append(nxt)
        if seen == members:
            coherent += 1
    return coherent / len(groups)


def survey_split_keys(
    doc: PcfDocument, *, with_cohesion: bool = True
) -> list[KeySurvey]:
    """Census every scalar attribute as a potential split key.

    The census itself is a pure text pass (milliseconds even on large
    files). With ``with_cohesion`` one connectivity graph is built and each
    candidate key's value-groups are scored for spatial coherence, which is
    what separates real grouping keys (SPOOL-IDENTIFIER, line number) from
    keys that merely repeat (ITEM-CODE, PIPING-SPEC). Candidates sort
    first, best cohesion / fewest sheets on top.
    """
    comps = doc.real_components()
    if not comps:
        return []
    total = len(comps)
    graph: geometry.ConnectivityGraph | None = None
    surveys: list[KeySurvey] = []
    for key, mapping in _collect_key_values(doc).items():
        counter = Counter(mapping.values())
        survey = KeySurvey(key=key, present=len(mapping), total=total,
                           distinct=len(counter), values=dict(counter))
        if with_cohesion and survey.is_candidate:
            if graph is None:
                graph = geometry.ConnectivityGraph(
                    comps, geometry.model_tolerance(doc))
            survey.cohesion = _cohesion(graph, mapping)
        surveys.append(survey)
    # On a tie, a real component attribute beats the PIPELINE-REFERENCE
    # pseudo-key: it is the export's own grouping, not a structural echo.
    surveys.sort(key=lambda s: (
        not s.is_candidate, -(s.cohesion if s.cohesion is not None else -1.0),
        s.distinct, s.key == PIPELINE_KEY, s.key,
    ))
    return surveys


def split_by_key(
    doc: PcfDocument,
    key: str,
    *,
    tolerance: float | None = None,
) -> SplitResult:
    """Split into one document per distinct value of a grouping attribute.

    ``key`` is a per-component scalar attribute (SPOOL-IDENTIFIER, a line
    number field...) or the pseudo-key PIPELINE-REFERENCE (group by owning
    pipeline block). Components without a value inherit the value of their
    nearest connected labelled neighbour, so welds and gap couplings that
    lack the attribute travel with their spool. Deterministic: same file +
    same key -> same parts in the same order (file order of first use).
    """
    work = _copy_document(doc)
    key_u = key.strip().upper()

    comps: list[Component] = []
    owner: list[tuple[int, int]] = []
    for p_idx, pipeline in enumerate(work.pipelines):
        for c_idx, comp in enumerate(pipeline.components):
            if comp.is_connection_record:
                continue
            comps.append(comp)
            owner.append((p_idx, c_idx))
    if not comps:
        raise SplitError("document contains no components")

    values: list[str | None] = []
    for i, comp in enumerate(comps):
        if key_u in (PIPELINE_KEY, "PIPELINE"):
            raw = work.pipelines[owner[i][0]].reference
        else:
            raw = comp.get(key_u)
        raw = raw.strip() if raw else ""
        values.append(raw or None)

    if not any(v is not None for v in values):
        raise SplitError(f"no component carries a {key_u} value")

    notes: list[str] = []
    if any(v is None for v in values):
        tol = tolerance if tolerance is not None else geometry.model_tolerance(work)
        values = geometry.flood_values(comps, values, tol)
        unlabelled = sum(1 for v in values if v is None)
        assert unlabelled == 0
        notes.append(
            f"components without {key_u} were attached to their nearest "
            "labelled neighbour"
        )

    groups_map: dict[str, set[int]] = {}
    for i, value in enumerate(values):
        groups_map.setdefault(value, set()).add(i)  # type: ignore[arg-type]
    groups = sorted(groups_map.values(), key=min)
    labels = [values[min(g)] for g in groups]
    notes.append(f"{len(groups)} group(s) by {key_u}")
    return SplitResult(
        parts=_build_parts(work, groups, owner),
        notes=notes,
        labels=labels,  # type: ignore[arg-type]
    )


def _joint_axis(comps: list[Component], at_cut: list[int], cut: Vec3) -> Vec3:
    """Direction of the pipe run through a cut point."""
    from .skeys import classify

    # Prefer the weld/coupling spanning the joint, then any two-ended
    # component whose end is at the cut.
    best: Vec3 | None = None
    for i in at_cut:
        comp = comps[i]
        eps = comp.end_points
        if len(eps) < 2:
            continue
        direction = geometry.normalize(geometry.sub(eps[1].xyz, eps[0].xyz))
        if direction == (0.0, 0.0, 0.0):
            continue
        if classify(comp.type, comp.skey) in ("WELD", "COUPLING"):
            return direction
        if best is None:
            best = direction
    return best if best is not None else (1.0, 0.0, 0.0)


def _group_distance(comps: list[Component], a: set[int], b: set[int]) -> float:
    pts_a = [p.xyz for i in a for p in comps[i].all_points()]
    pts_b = [p.xyz for i in b for p in comps[i].all_points()]
    if not pts_a or not pts_b:
        return float("inf")
    return min(geometry.dist(pa, pb) for pa in pts_a for pb in pts_b)


def _renumber_code(existing: set[str], code: str) -> str:
    try:
        candidate = int(code)
        while str(candidate) in existing:
            candidate += 1
        return str(candidate)
    except ValueError:
        n = 2
        while f"{code}-{n}" in existing:
            n += 1
        return f"{code}-{n}"


def merge_documents(docs: list[PcfDocument]) -> MergeResult:
    if len(docs) < 2:
        raise MergeError("need at least two documents to merge")

    warnings: list[str] = []

    base_units = docs[0].units_coords
    for doc in docs[1:]:
        if doc.units_coords != base_units:
            raise MergeError(
                f"UNITS-CO-ORDS mismatch: {base_units} vs {doc.units_coords} "
                f"({doc.source_path}); convert before merging"
            )
        if doc.units_bore != docs[0].units_bore:
            warnings.append(
                f"UNITS-BORE differs: {docs[0].units_bore} vs {doc.units_bore}"
            )

    merged = _copy_document(docs[0])
    materials: dict[str, str] = merged.materials_descriptions()
    remaps: dict[str, dict[str, str]] = {}

    for doc in docs[1:]:
        incoming = _copy_document(doc)
        remap: dict[str, str] = {}
        for code, desc in incoming.materials_descriptions().items():
            if code in materials:
                if materials[code].strip().lower() == desc.strip().lower():
                    continue  # identical item, share the code
                new_code = _renumber_code(set(materials), code)
                remap[code] = new_code
                materials[new_code] = desc
            else:
                materials[code] = desc

        if remap:
            remaps[doc.source_path or f"doc{docs.index(doc)}"] = remap
            for comp in incoming.all_components():
                code = comp.item_code
                if code in remap:
                    comp.set("ITEM-CODE", remap[code])

        merged.pipelines.extend(incoming.pipelines)

    # Rebuild the MATERIALS section deterministically.
    lines = ["MATERIALS"]
    def _sort_key(code: str):
        try:
            return (0, int(code), code)
        except ValueError:
            return (1, 0, code)
    for code in sorted(materials, key=_sort_key):
        lines.append(f"ITEM-CODE  {code}")
        if materials[code]:
            lines.append(f"    DESCRIPTION  {materials[code]}")
    merged.materials_lines = lines

    # Adjacency check: does each subsequent document touch the growing model?
    tol = geometry.model_tolerance(merged)
    placed = [docs[0]]
    for doc in docs[1:]:
        min_dist = _min_distance(placed, doc)
        if min_dist is None:
            warnings.append(f"{doc.source_path or 'document'}: no coordinates to check")
        elif min_dist > tol:
            warnings.append(
                f"NOT ADJACENT: {doc.source_path or 'document'} is "
                f"{min_dist:.3f} {merged.units_coords} from the rest of the merged model "
                f"(tolerance {tol:.3f}); the merged file will contain disconnected piping"
            )
        placed.append(doc)

    return MergeResult(document=merged, warnings=warnings, item_code_remaps=remaps)


def _min_distance(placed: list[PcfDocument], doc: PcfDocument) -> float | None:
    other_points = [
        p.xyz for d in placed for c in d.real_components() for p in c.connection_points()
    ]
    own_points = [p.xyz for c in doc.real_components() for p in c.connection_points()]
    if not other_points or not own_points:
        return None
    return min(geometry.dist(a, b) for a in own_points for b in other_points)
