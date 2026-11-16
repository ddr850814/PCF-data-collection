"""Data model for Piping Component Files.

The model is deliberately *lossless*: every line of the source file is kept
verbatim (``AttrLine.raw``) so a parsed document can be re-emitted
byte-identically. Typed accessors (points, SKEY, item codes...) parse on
demand from the raw lines; edits go through helpers that rewrite the raw
line so the round-trip guarantee survives modification.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

INDENT = "    "

# Component types whose records occupy a single point in space.
CO_ORDS_KEY = "CO-ORDS"
POINT_KEYS = (
    "END-POINT",
    "CENTRE-POINT",
    "BRANCH1-POINT",
    "BRANCH2-POINT",
    "CO-ORDS",
)
_PORT_POINT_RE = re.compile(r"^PORT-POINT\d*$", re.IGNORECASE)

# END-POINT trailing connection-type codes (PCF Reference Guide V13).
CONNECTION_TYPES = {"BW", "CP", "FL", "PL", "SC", "SW", "LN", "MP", "LC", "LR"}

# Record keywords that start a pseudo-component inside a pipeline (they may
# legally appear at column 1 and carry indented CO-ORDS children).
CONNECTION_RECORD_PREFIXES = (
    "END-CONNECTION-",
    "END-POSITION-",
    "TAP-CONNECTION",
    "NOZZLE",
    "MESSAGE",
    "TML",
)


@dataclass(frozen=True)
class Point:
    """A parsed coordinate record: position, optional bore, trailing flags."""

    key: str
    x: float
    y: float
    z: float
    bore: float | None = None
    flags: tuple[str, ...] = ()

    @property
    def xyz(self) -> tuple[float, float, float]:
        return (self.x, self.y, self.z)

    @property
    def connection_type(self) -> str | None:
        for flag in self.flags:
            if flag.upper() in CONNECTION_TYPES:
                return flag.upper()
        return None

    def distance_to(self, other: "Point | tuple[float, float, float]") -> float:
        ox, oy, oz = other.xyz if isinstance(other, Point) else other
        return math.dist((self.x, self.y, self.z), (ox, oy, oz))


def parse_point(raw: str) -> Point | None:
    """Parse an indented coordinate attribute line into a Point."""
    tokens = raw.split()
    if len(tokens) < 4:
        return None
    key = tokens[0].upper()
    try:
        x, y, z = float(tokens[1]), float(tokens[2]), float(tokens[3])
    except ValueError:
        return None
    if not (math.isfinite(x) and math.isfinite(y) and math.isfinite(z)):
        # NAN/INF 脏数据：静默入模型会让容差桶的 round(nan) 直接崩
        return None
    bore: float | None = None
    flags: list[str] = []
    for tok in tokens[4:]:
        if bore is None:
            try:
                bore = float(tok)
                continue
            except ValueError:
                pass
        flags.append(tok)
    if bore is not None and not math.isfinite(bore):
        bore = None
    return Point(key=key, x=x, y=y, z=z, bore=bore, flags=tuple(flags))


@dataclass
class AttrLine:
    """One attribute line of a component or pipeline header, kept verbatim."""

    raw: str

    @property
    def tokens(self) -> list[str]:
        return self.raw.split()

    @property
    def key(self) -> str:
        toks = self.tokens
        return toks[0].upper() if toks else ""

    @property
    def value(self) -> str:
        stripped = self.raw.strip()
        parts = stripped.split(None, 1)
        return parts[1] if len(parts) > 1 else ""

    def set_value(self, value: str) -> None:
        key_text = self.raw.strip().split(None, 1)[0]
        self.raw = f"{INDENT}{key_text}  {value}".rstrip()


@dataclass
class Component:
    """One component (or connection/information) record and its attributes."""

    type: str
    attributes: list[AttrLine] = field(default_factory=list)

    # -- typed accessors -------------------------------------------------
    # 坐标解析结果缓存：渲染链一次出图要反复取同一批点（ConnectivityGraph
    # 每文档重建多次），逐次重解析属性字符串是纯浪费；set/remove 时失效。

    def _cached_points(self, kind: str, compute) -> list[Point]:
        cache = getattr(self, "_pcache", None)
        if cache is None:
            cache = {}
            self._pcache = cache
        pts = cache.get(kind)
        if pts is None:
            pts = compute()
            cache[kind] = pts
        return pts

    def _points(self, key: str) -> list[Point]:
        return list(self._cached_points(key, lambda: self._parse_points(key)))

    def _parse_points(self, key: str) -> list[Point]:
        out = []
        for attr in self.attributes:
            if attr.key == key:
                pt = parse_point(attr.raw)
                if pt is not None:
                    out.append(pt)
        return out

    @property
    def end_points(self) -> list[Point]:
        return self._points("END-POINT")

    @property
    def centre_points(self) -> list[Point]:
        return self._points("CENTRE-POINT")

    @property
    def branch1_points(self) -> list[Point]:
        return self._points("BRANCH1-POINT")

    @property
    def branch2_points(self) -> list[Point]:
        return self._points("BRANCH2-POINT")

    @property
    def co_ords(self) -> list[Point]:
        return self._points("CO-ORDS")

    @property
    def port_points(self) -> list[Point]:
        return list(self._cached_points(
            "_PORT_", lambda: self._filter_points(_PORT_POINT_RE.match)))

    def all_points(self) -> list[Point]:
        """Every coordinate this component carries, in file order."""
        return list(self._cached_points("_ALL_", lambda: self._filter_points(
            lambda k: k in POINT_KEYS or _PORT_POINT_RE.match(k))))

    def _filter_points(self, want) -> list[Point]:
        out = []
        for attr in self.attributes:
            if want(attr.key):
                pt = parse_point(attr.raw)
                if pt is not None:
                    out.append(pt)
        return out

    def connection_points(self) -> list[Point]:
        """External keypoints: where this component joins its neighbours."""
        return list(self._cached_points("_CONN_", lambda: self._compute_conn_points()))

    def _compute_conn_points(self) -> list[Point]:
        out = []
        for attr in self.attributes:
            if attr.key in ("END-POINT", "BRANCH1-POINT", "BRANCH2-POINT") or _PORT_POINT_RE.match(
                attr.key
            ):
                pt = parse_point(attr.raw)
                if pt is not None:
                    out.append(pt)
        if not out:
            out = self._parse_points("CO-ORDS")
        return out

    # -- scalar attribute helpers ----------------------------------------

    def get(self, key: str, default: str | None = None) -> str | None:
        key = key.upper()
        for attr in self.attributes:
            if attr.key == key:
                return attr.value
        return default

    def get_all(self, key: str) -> list[str]:
        key = key.upper()
        return [attr.value for attr in self.attributes if attr.key == key]

    def has(self, key: str) -> bool:
        return self.get(key) is not None

    def material_excluded(self) -> bool:
        """MATERIAL-LIST EXCLUDE：本图只画不计料。

        距离拆分导出把边界段（支管台/三通/直管及其法兰副料、焊缝）双写
        进两侧文件，客侧副本带此标记——ISOGEN 惯例是画示意但不进本图
        BOM/切割表/焊表，归属侧无标记、正常计量。
        """
        return (self.get("MATERIAL-LIST") or "").strip().upper() == "EXCLUDE"

    def is_continuation_copy(self) -> bool:
        """CONTINUATION：对侧图纸边界件在本图的副本（画虚线示意）。"""
        return self.has("CONTINUATION")

    def set(self, key: str, value: str) -> None:
        """Set (or add) a scalar attribute, preserving other lines verbatim."""
        key_u = key.upper()
        for attr in self.attributes:
            if attr.key == key_u:
                attr.set_value(value)
                self._pcache = None  # 坐标缓存失效
                return
        self.attributes.append(AttrLine(f"{INDENT}{key.upper()}  {value}"))
        self._pcache = None

    def remove(self, key: str) -> int:
        key_u = key.upper()
        before = len(self.attributes)
        self.attributes = [a for a in self.attributes if a.key != key_u]
        self._pcache = None
        return before - len(self.attributes)

    @property
    def skey(self) -> str | None:
        return self.get("SKEY")

    @property
    def item_code(self) -> str | None:
        return self.get("ITEM-CODE")

    @property
    def description(self) -> str | None:
        return self.get("ITEM-DESCRIPTION")

    @property
    def uci(self) -> str | None:
        return self.get("UNIQUE-COMPONENT-IDENTIFIER") or self.get("UCI")

    @property
    def weight(self) -> float | None:
        raw = self.get("WEIGHT")
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    @property
    def angle_degrees(self) -> float | None:
        raw = self.get("ANGLE")
        if raw is None:
            return None
        try:
            return float(raw) / 100.0
        except ValueError:
            return None

    @property
    def bore(self) -> float | None:
        """Largest bore found on any keypoint (main-run size)."""
        bores = [p.bore for p in self.all_points() if p.bore is not None]
        return max(bores) if bores else None

    @property
    def is_connection_record(self) -> bool:
        return self.type.upper().startswith(CONNECTION_RECORD_PREFIXES)

    @property
    def is_fabrication(self) -> bool:
        if self.has("FABRICATION-ITEM"):
            return True
        category = self.get("CATEGORY")
        return category is not None and category.upper().startswith("FAB")

    @property
    def is_erection(self) -> bool:
        if self.has("ERECTION-ITEM") or self.has("OFFSHORE-ITEM"):
            return True
        category = self.get("CATEGORY")
        return category is not None and category.upper().startswith(("EREC", "OFF"))

    def __repr__(self) -> str:  # pragma: no cover - debugging nicety
        return f"Component({self.type!r}, skey={self.skey!r}, item={self.item_code!r})"


@dataclass
class Pipeline:
    """One PIPELINE-REFERENCE block: header attributes plus its components."""

    reference: str
    header_lines: list[AttrLine] = field(default_factory=list)
    components: list[Component] = field(default_factory=list)
    raw_line: str | None = None

    def get(self, key: str, default: str | None = None) -> str | None:
        key = key.upper()
        for attr in self.header_lines:
            if attr.key == key:
                return attr.value
        return default

    def set(self, key: str, value: str) -> None:
        key_u = key.upper()
        for attr in self.header_lines:
            if attr.key == key_u:
                attr.set_value(value)
                return
        self.header_lines.append(AttrLine(f"{INDENT}{key.upper()}  {value}"))

    def real_components(self) -> list[Component]:
        """Components that are physical items (not END-CONNECTION/MESSAGE...)."""
        return [c for c in self.components if not c.is_connection_record]


@dataclass
class PcfDocument:
    """A whole PCF file: header, pipelines, and the MATERIALS section."""

    header_lines: list[str] = field(default_factory=list)
    pipelines: list[Pipeline] = field(default_factory=list)
    materials_lines: list[str] = field(default_factory=list)
    newline: str = "\r\n"
    source_path: str | None = None

    def header_get(self, key: str, default: str | None = None) -> str | None:
        key = key.upper()
        for line in self.header_lines:
            tokens = line.split()
            if tokens and tokens[0].upper() == key:
                return line.strip().split(None, 1)[1] if len(tokens) > 1 else ""
        return default

    @property
    def units_coords(self) -> str:
        return (self.header_get("UNITS-CO-ORDS") or "INCH").upper()

    @property
    def units_bore(self) -> str:
        return (self.header_get("UNITS-BORE") or "INCH").upper()

    def all_components(self) -> list[Component]:
        return [c for p in self.pipelines for c in p.components]

    def real_components(self) -> list[Component]:
        return [c for p in self.pipelines for c in p.real_components()]

    def materials_descriptions(self) -> dict[str, str]:
        """Map ITEM-CODE (or MATERIAL-IDENTIFIER) -> DESCRIPTION."""
        out: dict[str, str] = {}
        current: str | None = None
        for raw in self.materials_lines:
            stripped = raw.strip()
            if not stripped:
                continue
            tokens = stripped.split(None, 1)
            key = tokens[0].upper()
            value = tokens[1].strip() if len(tokens) > 1 else ""
            if key in ("ITEM-CODE", "MATERIAL-IDENTIFIER"):
                current = value
                out.setdefault(current, "")
            elif key == "DESCRIPTION" and current is not None:
                out[current] = value
        return out
