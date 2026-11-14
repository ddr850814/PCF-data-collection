"""Bridge derived weld positions to the external weld-tracking system.

pcf2isometric deliberately owns **no weld identity**. Every render assigns a
provisional ordinal (``weld_number``, 1..N in file order) so the same file is
stable across re-renders, but the real, human-facing weld number lives in the
weld-tracking system -- the same way ``.weldb`` welds do. Two operations tie
the two worlds together through ``_welds.csv``:

* **record** (pull-back): once welds have been uploaded as map items and the
  QC team has assigned/edited their numbers, match each tracking map item back
  to its ``_welds.csv`` row and stamp the tracking system's stable
  ``map_item_id`` onto the row. Matching is by our uploaded label first (fast
  exact path while numbers are untouched) then by nearest page position (which
  survives relabelling, because the marker stays where we drew it). The CSV
  becomes the durable crosswalk ``world_xyz <-> map_item_id``; the mutable
  weld number and QC status are *not* copied -- they stay in the tracking
  system and are fetched live by ``map_item_id`` when needed, so the CSV can
  never go stale.

* **reconcile** (revision): when the model changes and the drawing is
  re-rendered, the fresh CSV has new ordinals but the same ``world_xyz`` for
  every weld that did not move -- and ``world_xyz`` is the only identity that
  survives a re-render (file order, scale, and quadrant may all change). Join
  the previous, stamped CSV to the fresh one by ``world_xyz`` within a
  tolerance, carry ``map_item_id`` forward for unchanged welds, and report
  welds **added** (need a tracking number) and **removed** (their map item
  should be retired). That report is the context an agent needs to revise a
  drawing without disturbing already-signed-off welds.

Both operations are pure functions over plain row dicts (the columns written
by ``scripts/_common.write_welds_csv``); the ``scripts/`` CLIs handle CSV and
MCP I/O around them.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from . import geometry

# The column added to a _welds.csv row by record/reconcile.
MAP_ITEM_ID_COLUMN = "map_item_id"

# Default revision join tolerance, in the CSV's native coordinate units: unmoved welds
# share identical world coordinates (distance 0); the slack only forgives minor nudges.
DEFAULT_RECONCILE_TOLERANCE = 0.1

# Default pull-back position tolerance, in canvas pixels: we placed the map items
# ourselves, so an untouched marker sits within a pixel of its row; slack forgives nudges.
DEFAULT_MATCH_TOLERANCE_PX = 30.0


@dataclass
class MapItem:
    """A weld map item as it stands in the tracking system."""
    map_item_id: str
    label: str
    x_px: float
    y_px: float


@dataclass
class RecordMatch:
    weld_number: str
    map_item_id: str
    method: str          # "label" | "position"
    distance_px: float


@dataclass
class RecordResult:
    matches: list[RecordMatch] = field(default_factory=list)
    unmatched_welds: list[str] = field(default_factory=list)   # rows with no map item
    unmatched_items: list[str] = field(default_factory=list)   # map items with no row

    @property
    def map_by_weld(self) -> dict[str, str]:
        return {m.weld_number: m.map_item_id for m in self.matches}


@dataclass
class CarriedWeld:
    old_weld_number: str
    new_weld_number: str
    map_item_id: str
    distance: float


@dataclass
class RemovedWeld:
    old_weld_number: str
    map_item_id: str


@dataclass
class AddedWeld:
    new_weld_number: str
    world_xyz: tuple[float, float, float]


@dataclass
class ReconcileResult:
    carried: list[CarriedWeld] = field(default_factory=list)
    added: list[AddedWeld] = field(default_factory=list)
    removed: list[RemovedWeld] = field(default_factory=list)


def _row_xyz(row: dict) -> tuple[float, float, float]:
    return (float(row["world_x"]), float(row["world_y"]), float(row["world_z"]))


def match_map_items(
    weld_rows: list[dict],
    map_items: list[MapItem],
    canvas_width: float,
    tolerance_px: float = DEFAULT_MATCH_TOLERANCE_PX,
) -> RecordResult:
    """Match tracking map items back to _welds.csv rows for the same drawing.

    ``canvas_width`` is the drawing image's pixel width (from ``get_drawing``);
    it scales each row's ``x_mm`` to the map item's pixel space
    (``scale = canvas_width / page_width_mm``, uniform, no flip -- the same
    transform ``weld_positions_to_canvas`` used to place the items).
    """
    result = RecordResult()
    if not weld_rows:
        result.unmatched_items = [it.map_item_id for it in map_items]
        return result

    page_w = float(weld_rows[0]["page_width_mm"])
    scale = canvas_width / page_w if page_w else 1.0
    targets = [(row, float(row["x_mm"]) * scale, float(row["y_mm"]) * scale)
               for row in weld_rows]

    used_items: set[str] = set()
    matched_welds: set[str] = set()

    # Pass 1: exact label identity (our uploaded numbers, still untouched). A
    # coincidental far-away label clash is rejected by a generous distance cap.
    label_cap = max(tolerance_px * 4, 60.0)
    by_label: dict[str, list[MapItem]] = {}
    for it in map_items:
        by_label.setdefault(str(it.label), []).append(it)
    for row, ex, ey in targets:
        wn = str(row["weld_number"])
        cands = [it for it in by_label.get(wn, []) if it.map_item_id not in used_items]
        if len(cands) == 1:
            it = cands[0]
            d = math.hypot(it.x_px - ex, it.y_px - ey)
            if d <= label_cap:
                used_items.add(it.map_item_id)
                matched_welds.add(wn)
                result.matches.append(RecordMatch(wn, it.map_item_id, "label", round(d, 1)))

    # Pass 2: nearest unused map item within tolerance, one-to-one by distance.
    candidates: list[tuple[float, str, MapItem]] = []
    for row, ex, ey in targets:
        wn = str(row["weld_number"])
        if wn in matched_welds:
            continue
        for it in map_items:
            if it.map_item_id in used_items:
                continue
            d = math.hypot(it.x_px - ex, it.y_px - ey)
            if d <= tolerance_px:
                candidates.append((d, wn, it))
    candidates.sort(key=lambda c: c[0])
    for d, wn, it in candidates:
        if wn in matched_welds or it.map_item_id in used_items:
            continue
        used_items.add(it.map_item_id)
        matched_welds.add(wn)
        result.matches.append(RecordMatch(wn, it.map_item_id, "position", round(d, 1)))

    result.matches.sort(key=lambda m: _weld_sort_key(m.weld_number))
    result.unmatched_welds = [str(r["weld_number"]) for r in weld_rows
                              if str(r["weld_number"]) not in matched_welds]
    result.unmatched_items = [it.map_item_id for it in map_items
                              if it.map_item_id not in used_items]
    return result


def reconcile_welds(
    old_rows: list[dict],
    new_rows: list[dict],
    tolerance: float = DEFAULT_RECONCILE_TOLERANCE,
) -> ReconcileResult:
    """Join a stamped previous CSV to a freshly rendered CSV by world position.

    ``old_rows`` are expected to carry a ``map_item_id`` column (from a prior
    ``record`` pull-back); rows without one still reconcile geometrically, they
    just carry an empty id forward.
    """
    pairs: list[tuple[float, int, int]] = []
    for i, o in enumerate(old_rows):
        oxyz = _row_xyz(o)
        for j, n in enumerate(new_rows):
            d = geometry.dist(oxyz, _row_xyz(n))
            if d <= tolerance:
                pairs.append((d, i, j))
    pairs.sort(key=lambda p: p[0])

    used_old: set[int] = set()
    used_new: set[int] = set()
    result = ReconcileResult()
    for d, i, j in pairs:
        if i in used_old or j in used_new:
            continue
        used_old.add(i)
        used_new.add(j)
        o, n = old_rows[i], new_rows[j]
        result.carried.append(CarriedWeld(
            old_weld_number=str(o["weld_number"]),
            new_weld_number=str(n["weld_number"]),
            map_item_id=str(o.get(MAP_ITEM_ID_COLUMN, "") or ""),
            distance=round(d, 4),
        ))

    result.carried.sort(key=lambda c: _weld_sort_key(c.new_weld_number))
    result.added = [AddedWeld(new_weld_number=str(n["weld_number"]),
                              world_xyz=_row_xyz(n))
                    for j, n in enumerate(new_rows) if j not in used_new]
    result.added.sort(key=lambda a: _weld_sort_key(a.new_weld_number))
    result.removed = [RemovedWeld(old_weld_number=str(o["weld_number"]),
                                  map_item_id=str(o.get(MAP_ITEM_ID_COLUMN, "") or ""))
                      for i, o in enumerate(old_rows) if i not in used_old]
    result.removed.sort(key=lambda r: _weld_sort_key(r.old_weld_number))
    return result


def _weld_sort_key(number: str):
    try:
        return (0, int(number), "")
    except (TypeError, ValueError):
        return (1, 0, str(number))
