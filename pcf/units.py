"""Unit handling and shop-style dimension formatting.

PCF headers declare units independently for coordinates, bore, bolts and
weights; a parser must never assume bore units match coordinate units.
"""

from __future__ import annotations

import sys
from fractions import Fraction

# Multiply a native coordinate value by this factor to get inches.
_COORD_TO_INCH = {
    "INCH": 1.0,
    "INCHES": 1.0,
    "MM": 1.0 / 25.4,
    "MM.": 1.0 / 25.4,           # 部分出口器单位带尾点
    "MILLIMETRE": 1.0 / 25.4,     # 英式拼写
    "MILLIMETER": 1.0 / 25.4,     # 美式拼写
    "MM-HUNDREDTHS": 0.01 / 25.4,
    "CM": 1.0 / 2.54,
    "DM": 10.0 / 2.54,
    "M": 1000.0 / 25.4,
    "METRE": 1000.0 / 25.4,
    "METER": 1000.0 / 25.4,
}

# Multiply a native coordinate value by this factor to get millimetres.
_COORD_TO_MM = {key: factor * 25.4 for key, factor in _COORD_TO_INCH.items()}

# 未识别单位只告警一次：这些函数在渲染循环里被高频调用
_UNKNOWN_UNITS_WARNED: set[str] = set()


def _warn_unknown_units(units: str) -> None:
    u = units.upper()
    if u not in _UNKNOWN_UNITS_WARNED:
        _UNKNOWN_UNITS_WARNED.add(u)
        print(f"警告: 未识别的单位声明 {units!r}，按英寸处理", file=sys.stderr)


def coord_to_inch_factor(units_co_ords: str) -> float:
    f = _COORD_TO_INCH.get(units_co_ords.upper())
    if f is None:
        # 未知单位静默按英寸会把 mm 文件坐标/口径错 25 倍，必须出声
        _warn_unknown_units(units_co_ords)
        return 1.0
    return f


def coord_to_mm_factor(units_co_ords: str) -> float:
    f = _COORD_TO_MM.get(units_co_ords.upper())
    if f is None:
        _warn_unknown_units(units_co_ords)
        return 25.4
    return f


def is_metric_units(units_coords: str) -> bool:
    """英寸原生文件按车间惯例显示英尺-英寸，其余单位一律显示毫米。"""
    return coord_to_inch_factor(units_coords) != 1.0


def format_length(value: float, units_coords: str, *, force_feet: bool = False) -> str:
    """原生坐标长度 → 图面显示：跟随 UNITS-CO-ORDS（英寸 ft-in，其余 mm）。"""
    if is_metric_units(units_coords):
        return format_mm(value * coord_to_mm_factor(units_coords))
    return format_feet_inches(value * coord_to_inch_factor(units_coords),
                              force_feet=force_feet)


def bore_to_inch(bore: float, units_bore: str) -> float:
    units = units_bore.upper().rstrip(".")
    if units in ("INCH", "INCHES"):
        return bore
    if units == "INCH-SIXTEENTH":
        return bore / 16.0
    if units in ("MM", "MILLIMETRE", "MILLIMETER"):
        return bore / 25.4
    _warn_unknown_units(units_bore)
    return bore


_NPS_FROM_MM = {
    6: 0.125, 8: 0.25, 10: 0.375, 15: 0.5, 20: 0.75, 25: 1.0, 32: 1.25,
    40: 1.5, 50: 2.0, 65: 2.5, 80: 3.0, 90: 3.5, 100: 4.0, 125: 5.0,
    150: 6.0, 200: 8.0, 250: 10.0, 300: 12.0, 350: 14.0, 400: 16.0,
    450: 18.0, 500: 20.0, 600: 24.0, 750: 30.0, 900: 36.0,
}


def bore_to_nps(bore: float, units_bore: str) -> float:
    """Nominal pipe size in inches, mapping metric DN to NPS when close."""
    units = units_bore.upper()
    if units == "MM":
        dn = round(bore)
        if dn in _NPS_FROM_MM:
            return _NPS_FROM_MM[dn]
        return bore / 25.4
    return bore_to_inch(bore, units_bore)


def format_nps(nps_inches: float) -> str:
    """Format a nominal size the way fabricators write it (e.g. 1-1/2")."""
    frac = Fraction(nps_inches).limit_denominator(16)
    whole = int(frac)
    rem = frac - whole
    if rem == 0:
        return f'{whole}"'
    if whole == 0:
        return f'{rem.numerator}/{rem.denominator}"'
    return f'{whole}-{rem.numerator}/{rem.denominator}"'


def format_bore(bore: float, units_bore: str) -> str:
    """口径显示跟随 UNITS-BORE：MM 文件显示 DN，英寸文件保持 NPS 写法。"""
    if units_bore.upper() == "MM":
        dn = round(bore)
        if dn in _NPS_FROM_MM:
            return f"DN{dn}"
    return format_nps(bore_to_nps(bore, units_bore))


def format_feet_inches(
    length_inches: float, *, denominator: int = 16, force_feet: bool = False
) -> str:
    """Format a length as shop feet-inches, e.g. 12'-3 5/8" or 9 1/2".

    Rounds to the nearest 1/denominator inch. Lengths under 24 inches are
    written in inches only unless force_feet is set.
    """
    if length_inches < 0:
        return "-" + format_feet_inches(-length_inches, denominator=denominator,
                                        force_feet=force_feet)
    total_sixteenths = round(length_inches * denominator)
    inches_whole, sixteenths = divmod(total_sixteenths, denominator)

    frac_text = ""
    if sixteenths:
        frac = Fraction(sixteenths, denominator)
        frac_text = f" {frac.numerator}/{frac.denominator}"

    if inches_whole < 24 and not force_feet:
        if inches_whole == 0 and frac_text:
            return f'{frac_text.strip()}"'
        return f'{inches_whole}{frac_text}"'

    feet, inches = divmod(inches_whole, 12)
    return f"{feet}'-{inches}{frac_text}\""


def format_mm(length_mm: float) -> str:
    return f"{length_mm:.0f}"
