"""Vector math, direction naming, and component connectivity.

All geometry works in the file's native coordinate units (East, North,
Elevation). Connectivity is positional: two components are adjacent when a
connection point of one lies within tolerance of a connection point of the
other, which is how PCF expresses topology.
"""

from __future__ import annotations

import math
from collections import defaultdict
from typing import Iterable

from .model import Component, PcfDocument, Point

Vec3 = tuple[float, float, float]

# Positional tolerance as a fraction of the model's bounding-box diagonal,
# with an absolute floor to survive tiny models.
_REL_TOLERANCE = 1e-4


def sub(a: Vec3, b: Vec3) -> Vec3:
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def add(a: Vec3, b: Vec3) -> Vec3:
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def scale(a: Vec3, s: float) -> Vec3:
    return (a[0] * s, a[1] * s, a[2] * s)


def length(a: Vec3) -> float:
    return math.sqrt(a[0] ** 2 + a[1] ** 2 + a[2] ** 2)


def dist(a: Vec3, b: Vec3) -> float:
    return length(sub(a, b))


def normalize(a: Vec3) -> Vec3:
    mag = length(a)
    if mag < 1e-12:
        return (0.0, 0.0, 0.0)
    return (a[0] / mag, a[1] / mag, a[2] / mag)


def dot(a: Vec3, b: Vec3) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def midpoint(a: Vec3, b: Vec3) -> Vec3:
    return ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2, (a[2] + b[2]) / 2)


def point_segment_distance(p: Vec3, a: Vec3, b: Vec3) -> float:
    ab = sub(b, a)
    ab_len = length(ab)
    if ab_len < 1e-12:
        return dist(p, a)
    t = max(0.0, min(ab_len, dot(sub(p, a), normalize(ab))))
    return dist(p, add(a, scale(normalize(ab), t)))


def direction_name(vec: Vec3) -> str:
    """Human-readable compass/elevation direction of a vector.

    Axis convention: +X = East, +Y = North, +Z = Up. Off-axis runs list the
    dominant components, e.g. "NORTH & UP" for a sloped riser.
    """
    unit = normalize(vec)
    if unit == (0.0, 0.0, 0.0):
        return "NONE"
    names = []
    axes = (
        (unit[0], "EAST", "WEST"),
        (unit[1], "NORTH", "SOUTH"),
        (unit[2], "UP", "DOWN"),
    )
    for value, positive, negative in axes:
        if abs(value) > 0.05:
            names.append((abs(value), positive if value > 0 else negative))
    names.sort(reverse=True)
    return " & ".join(name for _, name in names) if names else "NONE"


def bounding_box(points: Iterable[Vec3]) -> tuple[Vec3, Vec3] | None:
    points = list(points)
    if not points:
        return None
    xs, ys, zs = zip(*points)
    return (min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs))


def model_tolerance(doc: PcfDocument) -> float:
    """Connection tolerance scaled to the model size.

    The absolute floor is just over a 1/16" weld root gap (in native units):
    real exports offset adjacent components by up to the modelled gap, and
    those joints must still register as connected.
    """
    from .units import coord_to_inch_factor

    floor = 0.11 / coord_to_inch_factor(doc.units_coords)
    pts = [p.xyz for c in doc.all_components() for p in c.all_points()]
    box = bounding_box(pts)
    if box is None:
        return floor
    diagonal = dist(box[0], box[1])
    return max(diagonal * _REL_TOLERANCE, floor)


class ConnectivityGraph:
    """Adjacency between components based on coincident connection points."""

    def __init__(self, components: list[Component], tolerance: float):
        self.components = components
        self.tolerance = tolerance
        self.adjacency: dict[int, set[int]] = defaultdict(set)
        self._build()

    def _quantize(self, point: Vec3) -> tuple[int, int, int]:
        q = self.tolerance
        return (round(point[0] / q), round(point[1] / q), round(point[2] / q))

    def _build(self) -> None:
        buckets: dict[tuple[int, int, int], list[int]] = defaultdict(list)
        for idx, comp in enumerate(self.components):
            for pt in comp.connection_points():
                key = self._quantize(pt.xyz)
                # Register in the 27 neighbouring cells so near-boundary
                # points still meet.
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        for dz in (-1, 0, 1):
                            buckets[(key[0] + dx, key[1] + dy, key[2] + dz)].append(idx)

        for idx, comp in enumerate(self.components):
            for pt in comp.connection_points():
                key = self._quantize(pt.xyz)
                for other in buckets.get(key, ()):
                    if other == idx:
                        continue
                    if self._touches(comp, self.components[other]):
                        self.adjacency[idx].add(other)
                        self.adjacency[other].add(idx)

        # Set-on attachments (olets, supports): the rider's CENTRE-POINT sits
        # mid-pipe, far from any pipe endpoint, so buckets never pair them.
        from .skeys import classify

        pipe_idxs = [i for i, c in enumerate(self.components)
                     if classify(c.type) == "PIPE"]
        for idx, comp in enumerate(self.components):
            if not comp.centre_points and not comp.co_ords:
                continue
            for pipe_idx in pipe_idxs:
                if pipe_idx == idx or pipe_idx in self.adjacency[idx]:
                    continue
                if self._on_axis(self.components[pipe_idx], comp):
                    self.adjacency[idx].add(pipe_idx)
                    self.adjacency[pipe_idx].add(idx)

    def _touches(self, a: Component, b: Component) -> bool:
        for pa in a.connection_points():
            for pb in b.connection_points():
                if pa.distance_to(pb) <= self.tolerance:
                    return True
        # Set-on attachments: a CENTRE-POINT (or connection point) landing on
        # the other component's pipe axis counts as connected.
        return self._on_axis(a, b) or self._on_axis(b, a)

    def _on_axis(self, host: Component, rider: Component) -> bool:
        from .skeys import classify

        if classify(host.type) != "PIPE":
            return False
        eps = host.end_points
        if len(eps) < 2:
            return False
        a_xyz, b_xyz = eps[0].xyz, eps[1].xyz
        attach = list(rider.centre_points) or list(rider.co_ords) \
            or rider.connection_points()
        for pt in attach:
            if point_segment_distance(pt.xyz, a_xyz, b_xyz) <= self.tolerance:
                return True
        return False

    def neighbours(self, idx: int) -> set[int]:
        return self.adjacency.get(idx, set())

    def connected_sets(self) -> list[set[int]]:
        seen: set[int] = set()
        groups: list[set[int]] = []
        for idx in range(len(self.components)):
            if idx in seen:
                continue
            group = {idx}
            frontier = [idx]
            while frontier:
                cur = frontier.pop()
                for nxt in self.adjacency.get(cur, ()):
                    if nxt not in group:
                        group.add(nxt)
                        frontier.append(nxt)
            seen |= group
            groups.append(group)
        return groups


def flood_values(
    components: list[Component],
    values: list[str | None],
    tolerance: float,
    *,
    graph: ConnectivityGraph | None = None,
) -> list[str | None]:
    """Fill unlabelled entries with the value of their nearest labelled peer.

    Connected components inherit through the adjacency graph (lowest-index
    labelled neighbour wins, so the result is deterministic); disconnected
    unlabelled islands take the value of the nearest labelled component by
    geometry. A list with no labels at all comes back unchanged -- callers
    decide whether that is an error.
    """
    values = list(values)
    pending = {i for i, v in enumerate(values) if v is None}
    if not pending or len(pending) == len(values):
        return values
    if graph is None:
        graph = ConnectivityGraph(components, tolerance)
    changed = True
    while pending and changed:
        changed = False
        for i in sorted(pending):
            labelled = sorted(
                n for n in graph.neighbours(i) if values[n] is not None)
            if labelled:
                values[i] = values[labelled[0]]
                pending.discard(i)
                changed = True
    for i in sorted(pending):
        nearest = min(
            (j for j, v in enumerate(values) if v is not None),
            key=lambda j: _component_distance(components, i, j),
        )
        values[i] = values[nearest]
    return values


def _component_distance(components: list[Component], i: int, j: int) -> float:
    pts_i = [p.xyz for p in components[i].all_points()]
    pts_j = [p.xyz for p in components[j].all_points()]
    if not pts_i or not pts_j:
        return float("inf")
    return min(dist(a, b) for a in pts_i for b in pts_j)


def components_at_point(
    components: list[Component], point: Vec3, tolerance: float
) -> list[int]:
    hits = []
    for idx, comp in enumerate(components):
        for pt in comp.connection_points():
            if dist(pt.xyz, point) <= tolerance:
                hits.append(idx)
                break
    return hits


def open_ends(doc: PcfDocument, tolerance: float | None = None) -> list[tuple[Component, Point]]:
    """Connection points not shared with any other component (line ends)."""
    comps = doc.real_components()
    tol = tolerance if tolerance is not None else model_tolerance(doc)
    graph = ConnectivityGraph(comps, tol)
    out: list[tuple[Component, Point]] = []
    for idx, comp in enumerate(comps):
        for pt in comp.connection_points():
            shared = False
            for other_idx in graph.neighbours(idx):
                other = comps[other_idx]
                if any(pt.distance_to(op) <= tol for op in other.connection_points()):
                    shared = True
                    break
            if not shared:
                out.append((comp, pt))
    return out
