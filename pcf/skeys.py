"""SKEY interpretation and joint-type inference.

SKEYs (Isogen symbol keys) usually encode the end connection in their final
two characters. Joint inference follows the proven wmdb heuristics:
weld-neck implies butt weld, slip-on implies fillet weld, and a weld record
with no usable SKEY inherits the previous component's SKEY.
"""

from __future__ import annotations

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


# Component type -> short display class used by the renderer and reports.
def classify(component_type: str) -> str:
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
    if t.startswith("FLANGE") or t.startswith("LAPJOINT") or t.startswith("LAP-JOINT"):
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
