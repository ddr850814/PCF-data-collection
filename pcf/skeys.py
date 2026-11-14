"""SKEY interpretation and joint-type inference.

SKEYs (Isogen symbol keys) usually encode the end connection in their final
two characters. Joint inference follows the proven wmdb heuristics:
weld-neck implies butt weld, slip-on implies fillet weld, and a weld record
with no usable SKEY inherits the previous component's SKEY.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

# End-connection suffix -> joint type.
_SUFFIX_JOINTS = {
    "BW": "BW",   # butt weld
    "SW": "SW",   # socket weld
    "SC": "THD",  # screwed / threaded
    "FL": "FLG",  # flanged
    "PL": "PL",   # plain end
    "CP": "CP",   # compression
    "WN": "BW",   # weld neck flange -> butt weld
    "SO": "FW",   # slip-on flange -> fillet weld
}

WELD_JOINT_TYPES = {"BW", "SW", "FW"}


def joint_from_skey(skey: str | None) -> str | None:
    """Joint type implied by a SKEY's trailing two characters."""
    if not skey or len(skey) < 2:
        return None
    return _SUFFIX_JOINTS.get(skey[-2:].upper())


def connection_from_skey(skey: str | None) -> str | None:
    """End-connection code (BW/SW/SC/FL/PL/CP) implied by a SKEY."""
    if not skey or len(skey) < 2:
        return None
    suffix = skey[-2:].upper()
    if suffix in ("BW", "SW", "SC", "FL", "PL", "CP"):
        return suffix
    if suffix == "WN":
        return "BW"
    if suffix == "SO":
        return "FL"
    return None


# -- SKEY -> PCF 类型反查（skeys.json "skey_to_types" 段） -------------------
# 官方映射表：skey（含 ** 通配，** 占两个字符的端接代码）-> PCF 组件类型
# 列表。仅当 COMPONENT-TYPE 分类落到 OTHER 时作为兜底——出口器把管件写成
# 非标类型但保留标准 SKEY 时（如 TYPE=FITTING + SKEY=RCBW 的同心变径），
# 绘制、尺寸、报表仍按真实类别处理；标准数据零影响。

_SKEY_TYPES_EXACT: dict[str, str] = {}
_SKEY_TYPES_WILD: list[tuple[re.Pattern[str], str]] = []
_SKEY_TYPES_LOADED = False


def _load_skey_types() -> None:
    global _SKEY_TYPES_LOADED
    dirs = []
    env = os.environ.get("PCF_SYMBOL_DIR")
    if env:
        dirs.append(Path(env))
    dirs.append(Path(__file__).resolve().parent.parent / "symbols")
    for d in dirs:
        path = d / "skeys.json"
        if not path.is_file():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for pattern, types in (data.get("skey_to_types") or {}).items():
            resolved = next((t for t in types if t != "END-PREP"), None)
            if not resolved:
                continue
            if "**" in pattern:
                rx = re.compile("^" + pattern.upper().replace("**", "..") + "$")
                _SKEY_TYPES_WILD.append((rx, resolved))
            else:
                _SKEY_TYPES_EXACT.setdefault(pattern.upper(), resolved)
        _SKEY_TYPES_LOADED = True
        return
    # 没找到可用的 skeys.json（缺失/损坏）：同样置位缓存空表。否则
    # type_from_skey 每次调用都重扫磁盘——它被建图循环反复调，二次方 IO。
    _SKEY_TYPES_LOADED = True


def type_from_skey(skey: str | None) -> str | None:
    """SKEY -> PCF 组件类型（先精确后 ** 通配，如 RCBW -> REDUCER-CONCENTRIC）。"""
    if not skey:
        return None
    if not _SKEY_TYPES_LOADED:
        _load_skey_types()
    s = skey.upper()
    resolved = _SKEY_TYPES_EXACT.get(s)
    if resolved:
        return resolved
    for rx, t in _SKEY_TYPES_WILD:
        if rx.match(s):
            return t
    return None


def reducer_is_eccentric(component_type: str, skey: str | None = None) -> bool:
    """偏心变径判定：类型串明示优先，类型不明时按 SKEY 反查的类型判断。"""
    t = component_type.upper().replace("_", "-")
    if "ECCENTRIC" in t:
        return True
    if "CONCENTRIC" in t:
        return False
    resolved = type_from_skey(skey) if skey else None
    return bool(resolved and "ECCENTRIC" in resolved.upper())


# Component type -> short display class used by the renderer and reports.
def _classify_type(component_type: str) -> str:
    t = component_type.upper().replace("_", "-")
    if t.startswith("PIPE") or t == "NIPPLE":
        return "PIPE"
    if t.startswith("ELBOW") or t.startswith("BEND"):
        return "ELBOW"
    if t.startswith(("TEE-SET-ON", "TEE-STUB", "OLET", "ELBOLET", "CROSS-SET-ON", "CROSS-STUB")):
        return "OLET"
    # Branch bosses come out of real exporters under their trade names --
    # WELDOLET, SOCKOLET, THREDOLET, NIPOLET, LATROLET -- none of which
    # *start* with "OLET", so they used to fall through to OTHER and draw as
    # a bare line. Hyphens and underscores are already normalised above, and
    # stripping them catches THREAD-O-LET too.
    if "OLET" in t.replace("-", ""):
        return "OLET"
    if t.startswith("TEE") or t.startswith("CROSS") or t.startswith("Y-PIECE"):
        return "TEE"
    if t.startswith("VALVE"):
        return "VALVE"
    # SP3D 标准为 FLANGE-BLIND，部分导出也写成 BLIND-FLANGE。
    if (t.startswith("FLANGE") or t.startswith("BLIND-FLANGE")
            or t.startswith("LAPJOINT") or t.startswith("LAP-JOINT")):
        return "FLANGE"
    if t.startswith("REDUCER"):
        return "REDUCER"
    if t in ("GASKET", "SAFETY-DISC") or t.startswith("MISC-HYGIENIC"):
        return "GASKET"
    if t == "WELD":
        return "WELD"
    if t in ("BOLT", "NUT"):
        return "BOLT"
    if t in ("SUPPORT", "REINFORCEMENT-PAD", "PENETRATION-PLATE"):
        return "SUPPORT"
    if t.startswith("INSTRUMENT") or t.startswith("MISC-COMPONENT"):
        return "INSTRUMENT"
    if t.startswith("FILTER") or t.startswith("TRAP"):
        return "FILTER"
    if t == "CAP":
        return "CAP"
    if t in ("COUPLING", "UNION", "CONNECTOR", "CLAMP"):
        return "COUPLING"
    if t == "GAP":
        return "GAP"
    if t == "MULTI-PORT-COMPONENT":
        return "MULTIPORT"
    if t.startswith(("END-CONNECTION", "END-POSITION", "TAP-CONNECTION")):
        return "CONNECTION"
    if t.startswith("MESSAGE"):
        return "MESSAGE"
    return "OTHER"


def classify(component_type: str, skey: str | None = None) -> str:
    """按 COMPONENT-TYPE 分类；类型不认识（OTHER）时用 SKEY 反查兜底。"""
    cls = _classify_type(component_type)
    if cls == "OTHER" and skey:
        resolved = type_from_skey(skey)
        if resolved:
            return _classify_type(resolved)
    return cls


_VALVE_SUBTYPE_KEYWORDS = (
    ("ball", "BALL"),
    ("butterfly", "BUTTERFLY"),
    ("bfly", "BUTTERFLY"),
    ("check", "CHECK"),
    ("globe", "GLOBE"),
    ("needle", "NEEDLE"),
    ("plug", "PLUG"),
    ("pinch", "PINCH"),
    ("diaphragm", "DIAPHRAGM"),
    ("relief", "RELIEF"),
    ("safety", "RELIEF"),
    ("gate", "GATE"),
)


def valve_subtype(description: str | None, skey: str | None) -> str:
    """Classify a valve for symbol decoration from description + SKEY."""
    haystack = f"{description or ''} {skey or ''}".lower()
    for needle, subtype in _VALVE_SUBTYPE_KEYWORDS:
        if needle in haystack:
            return subtype
    if skey:
        prefix = skey[:2].upper()
        mapping = {"VB": "BALL", "VC": "CHECK", "VG": "GATE", "VT": "BUTTERFLY",
                   "VP": "PLUG", "VN": "NEEDLE", "VY": "RELIEF"}
        if prefix in mapping:
            return mapping[prefix]
    return "GATE"


def is_slip_on_flange(description: str | None) -> bool:
    if not description:
        return False
    hay = description.lower()
    return "slip-on" in hay or "slip on" in hay or "(slip-on)" in hay or " so " in f" {hay} "


# Real exports are inconsistent about flanges: plenty of them arrive as
# COUPLING/OTHER records whose only evidence of being a flange is the
# description ("Sch.80 Weld Neck Flange RF #300") or a FL- SKEY. Symbol
# dispatch and dimension landmarks need to see those as flanges; the BOM,
# cut list, and weld derivation keep using classify() unchanged so material
# accounting is never altered by a display heuristic.
_FLANGE_WORDS = ("flange", "flg", "blind")


def looks_like_flange(component_type: str, description: str | None = None,
                      skey: str | None = None) -> bool:
    """Display-level flange test: the class, the SKEY, or the description."""
    if classify(component_type) == "FLANGE":
        return True
    if classify(component_type) not in ("COUPLING", "OTHER"):
        return False
    if skey and skey.upper().startswith("FL"):
        return True
    hay = (description or "").lower()
    return any(word in hay for word in _FLANGE_WORDS)


def is_weld_coupling(component_type: str, description: str | None = None,
                     skey: str | None = None) -> bool:
    """A COUPLING record that is really a weld, not a piece of material.

    The tell is "weld" in the description ("Shop Weld - .0625"), but that
    word also opens "Weld Neck Flange RF STD #150" -- and the dominant
    exporter dialect ships weld-neck flanges as COUPLING records. Matching
    on the word alone dropped those flanges off the BOM and off the sheet's
    callouts while they were still drawn, so a flange test comes first: a
    flange is material and is always accounted for.
    """
    if classify(component_type) != "COUPLING":
        return False
    if looks_like_flange(component_type, description, skey):
        return False
    return "weld" in (description or "").lower()


# Annotation records an exporter ships as components: pipe flow arrows and
# slope arrows. They are drafting marks, not things anyone buys or welds.
_ANNOTATION_WORDS = ("flow arrow", "slope arrow")


def is_annotation_symbol(component_type: str, description: str | None = None,
                         skey: str | None = None) -> bool:
    """A record that is a drawing symbol rather than a piece of material.

    Exporters emit pipe flow arrows and slope arrows as MISC-COMPONENT
    records carrying a single repeated point and zero weight. They classify
    as INSTRUMENT, so they used to earn a BOM row, an item number and a
    balloon apiece -- a fabricator was being asked to buy a slope arrow --
    while drawing nothing, because a symbol whose two ends are the same
    point has no axis to draw a body along.

    Matched on the description, which is unambiguous ("6 inch pipe slope
    arrow"): no real material is named a flow or slope arrow. The SKEY these
    arrive under (CHPL) is not specific enough to key on by itself. Pass the
    *resolved* description -- the component's own, falling back to the
    MATERIALS table -- because plenty of exports name the part in only one
    of the two.
    """
    hay = (description or "").lower()
    return any(word in hay for word in _ANNOTATION_WORDS)
