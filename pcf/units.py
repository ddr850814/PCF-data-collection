"""Unit handling and shop-style dimension formatting.

PCF headers declare units independently for coordinates, bore, bolts and
weights; a parser must never assume bore units match coordinate units.
"""

from __future__ import annotations

from fractions import Fraction

# Multiply a native coordinate value by this factor to get inches.
_COORD_TO_INCH = {
    "INCH": 1.0,
    "INCHES": 1.0,
    "MM": 1.0 / 25.4,
    "MM-HUNDREDTHS": 0.01 / 25.4,
    "CM": 1.0 / 2.54,
    "DM": 10.0 / 2.54,
    "M": 1000.0 / 25.4,
    "METRE": 1000.0 / 25.4,
    "METER": 1000.0 / 25.4,
}

# Multiply a native coordinate value by this factor to get millimetres.
_COORD_TO_MM = {key: factor * 25.4 for key, factor in _COORD_TO_INCH.items()}


def coord_to_inch_factor(units_co_ords: str) -> float:
    return _COORD_TO_INCH.get(units_co_ords.upper(), 1.0)


def coord_to_mm_factor(units_co_ords: str) -> float:
    return _COORD_TO_MM.get(units_co_ords.upper(), 25.4)


def bore_to_inch(bore: float, units_bore: str) -> float:
    units = units_bore.upper()
    if units in ("INCH", "INCHES"):
        return bore
    if units == "INCH-SIXTEENTH":
        return bore / 16.0
    if units == "MM":
        return bore / 25.4
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
