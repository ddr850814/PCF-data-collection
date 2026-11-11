"""Split-option statistics -- the data behind the four-way split menu.

"How many drawings would each split scheme produce, and how many components
land on each drawing" is pure analysis: nothing is written, nothing is
rendered. It lives here so the stats script (split_stats.py), the one-shot
pipeline (process_pcf.py), and library users all share one implementation.

The four options, in the order they are numbered for the user:

  1. none      no split -- one drawing per master file
  2. key       one sheet per value of a grouping attribute
  3. distance  geometric sheets at the separation distance
  4. hybrid    the distance sheets, plus spool labels and Field Weld
               declarations between spools (requires a grouping key)

``docs`` throughout is a list of (source path, PcfDocument) pairs.
"""

from __future__ import annotations

from .geometry import ConnectivityGraph
from .model import PcfDocument
from .splitmerge import default_split_distance, survey_split_keys
from .welds import derive_welds

OPTION_NUMBERS = {"none": 1, "key": 2, "distance": 3, "hybrid": 4}
# Tie-break preference when several options score equally against the
# target band: simplest first, but hybrid beats plain distance because it
# costs nothing extra and keeps spool traceability.
_PREFERENCE = {"none": 0, "key": 1, "hybrid": 2, "distance": 3}
# Soft components-per-drawing target used to rank options; the user can
# override it per run (--target) or per job (project.json).
DEFAULT_TARGET = (10, 20)


def survey_report(doc: PcfDocument, source, *, with_cohesion: bool = True) -> dict:
    """Census every scalar attribute of one document as a split candidate."""
    comps = doc.real_components()
    surveys = survey_split_keys(doc, with_cohesion=with_cohesion)
    return {
        "file": str(source),
        "component_count": len(comps),
        "pipeline_count": len(doc.pipelines),
        "keys": [
            {
                "key": s.key,
                "present": s.present,
                "coverage": round(s.coverage, 4),
                "distinct": s.distinct,
                "candidate": s.is_candidate,
                "cohesion": None if s.cohesion is None else round(s.cohesion, 4),
                "avg_group_size": round(s.present / s.distinct, 2) if s.distinct else 0,
            }
            for s in surveys
        ],
    }


def aggregate_candidates(reports: list[dict]) -> list[dict]:
    """Keys that are split candidates in EVERY surveyed file, ranked."""
    per_key: dict[str, list[dict]] = {}
    for report in reports:
        for entry in report["keys"]:
            per_key.setdefault(entry["key"], []).append(entry)
    out = []
    for key, entries in per_key.items():
        if len(entries) < len(reports):
            continue
        if not all(e["candidate"] for e in entries):
            continue
        cohesions = [e["cohesion"] for e in entries if e["cohesion"] is not None]
        out.append({
            "key": key,
            "files": len(entries),
            "total_sheets": sum(e["distinct"] for e in entries),
            "min_cohesion": min(cohesions) if cohesions else None,
            "avg_group_size": round(
                sum(e["present"] for e in entries)
                / sum(e["distinct"] for e in entries), 2),
        })
    out.sort(key=lambda e: (
        -(e["min_cohesion"] if e["min_cohesion"] is not None else -1.0),
        e["total_sheets"], e["key"] == "PIPELINE-REFERENCE", e["key"],
    ))
    return out


def cohesive_keys(reports: list[dict]) -> list[dict]:
    """Grouping keys usable on every file, junk repeats filtered out.

    Cohesion >= 50% separates real spool/line keys (75-100% on production
    exports) from keys that merely repeat (ITEM-DESCRIPTION at 8-20%).
    Group size is only screened at the near-singleton level: averaging
    under 2 components per sheet is the tell of a repeat (WEIGHT,
    ITEM-CODE), but a real spool key averaging 2-3 stays on the menu --
    its low average is exactly what the options table shows the user.
    """
    return [
        c for c in aggregate_candidates(reports)
        if c["min_cohesion"] is not None and c["min_cohesion"] >= 0.5
        and c["avg_group_size"] >= 2
    ]


def best_split_key(docs) -> str | None:
    """The best grouping key usable across all documents, if any."""
    reports = [survey_report(doc, source) for source, doc in docs]
    cohesive = cohesive_keys(reports)
    return cohesive[0]["key"] if cohesive else None


def distance_drawing_count(docs, distance: float | None) -> int:
    """Sheets a distance split would produce, without building the parts."""
    total = 0
    for _source, doc in docs:
        d = distance if distance is not None \
            else default_split_distance(doc.units_coords)
        comps = doc.real_components()
        total += len(ConnectivityGraph(comps, d).connected_sets()) if comps else 1
    return total


def _entry(mode: str, drawings: int, total_comps: int, **extra) -> dict:
    return {
        "id": mode, "option": OPTION_NUMBERS[mode], "drawings": drawings,
        "avg_components": round(total_comps / drawings, 1) if drawings else 0.0,
        **extra,
    }


def _key_rows(cohesive: list[dict], total_comps: int, dist_drawings: int,
              distance: float | None, field_welds: int) -> list[dict]:
    """The key and hybrid entries for a table whose best key is known."""
    best = cohesive[0]
    return [
        _entry("key", best["total_sheets"], total_comps, key=best["key"],
               alternates=[{"key": c["key"], "drawings": c["total_sheets"],
                            "avg_components": c["avg_group_size"]}
                           for c in cohesive[1:5]]),
        _entry("hybrid", dist_drawings, total_comps, distance=distance,
               spool_key=best["key"], field_welds=field_welds),
    ]


def count_field_welds(doc: PcfDocument, spool_key: str) -> int:
    """Welds joining two different spool values (hybrid's declarations)."""
    return sum(1 for w in derive_welds(doc, spool_key=spool_key)
               if w.notes == "Field Weld")


def build_split_options(
    docs,
    distance: float | None = None,
    *,
    reports: list[dict] | None = None,
) -> list[dict]:
    """The four split options with predicted drawings and avg components.

    Pass precomputed ``reports`` (one survey_report per doc, same order)
    to avoid re-surveying when the caller already has them.
    """
    if reports is None:
        reports = [survey_report(doc, source) for source, doc in docs]
    total_comps = sum(len(doc.real_components()) for _, doc in docs)
    cohesive = cohesive_keys(reports)
    dist_drawings = distance_drawing_count(docs, distance)

    options = [
        _entry("none", len(docs), total_comps),
        _entry("distance", dist_drawings, total_comps, distance=distance),
    ]
    if cohesive:
        field_welds = sum(count_field_welds(doc, cohesive[0]["key"])
                          for _source, doc in docs)
        options.extend(_key_rows(cohesive, total_comps, dist_drawings,
                                 distance, field_welds))
    options.sort(key=lambda o: o["option"])
    return options


def combine_file_options(
    docs,
    reports: list[dict],
    per_file_options: list[list[dict]],
    distance: float | None = None,
) -> list[dict]:
    """The job-level option table, derived from already-computed per-file
    tables without re-running any geometry pass.

    Distance drawings sum exactly (distance groups never span files). The
    key/hybrid rows use the best key common to EVERY file; a file whose
    own best key differs gets its field welds re-counted under the common
    key (rare -- mixed export families), everything else is reused.
    """
    total_comps = sum(r["component_count"] for r in reports)
    dist_drawings = sum(
        next(o["drawings"] for o in options if o["id"] == "distance")
        for options in per_file_options
    )
    combined = [
        _entry("none", len(docs), total_comps),
        _entry("distance", dist_drawings, total_comps, distance=distance),
    ]
    cohesive = cohesive_keys(reports)
    if cohesive:
        common_key = cohesive[0]["key"]
        field_welds = 0
        for (_source, doc), options in zip(docs, per_file_options):
            hybrid = next((o for o in options if o["id"] == "hybrid"), None)
            if hybrid is not None and hybrid.get("spool_key") == common_key:
                field_welds += hybrid["field_welds"]
            else:
                field_welds += count_field_welds(doc, common_key)
        combined.extend(_key_rows(cohesive, total_comps, dist_drawings,
                                  distance, field_welds))
    combined.sort(key=lambda o: o["option"])
    return combined


def options_are_trivial(options: list[dict], file_count: int) -> bool:
    """True when every option is the same drawing set (nothing to decide):
    no grouping key exists and each file is one connected group."""
    return (len(options) == 2
            and all(o["drawings"] == file_count for o in options))


def recommend_option(options: list[dict], tmin: int, tmax: int) -> dict:
    def band_miss(o):
        avg = o["avg_components"]
        if tmin <= avg <= tmax:
            return 0.0
        return tmin - avg if avg < tmin else avg - tmax

    winner = min(options, key=lambda o: (band_miss(o), _PREFERENCE[o["id"]]))
    avg = winner["avg_components"]
    if band_miss(winner) == 0.0:
        reason = (f"averages {avg:g} components/drawing, inside the "
                  f"{tmin}-{tmax} target")
    else:
        reason = (f"averages {avg:g} components/drawing, closest of the four "
                  f"to the {tmin}-{tmax} target")
    if winner["id"] == "hybrid":
        reason += (f"; keeps spool traceability with labels and declares "
                   f"{winner.get('field_welds', 0)} field weld(s) between spools")
    elif winner["id"] == "key":
        reason += f" (one sheet per {winner['key']})"
    return {"id": winner["id"], "option": winner["option"], "reason": reason}
