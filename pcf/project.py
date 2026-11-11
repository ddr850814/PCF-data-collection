"""Project workspace conventions.

A PCF project is a folder that keeps originals safe and derived artifacts
organised::

    <project>/
        project.json    render settings: north axis, quadrant, logo, paper,
                        company block, requirements/quality notes, drawing
                        notes, drawing properties
        wiki.md         running log of what was done to which file and why
        source/         original PCFs -- never modified, never deleted
        working/        editable copies, split parts, merge results
        drawings/       rendered PDFs + <name>_welds.csv + <name>.dims.json

Tools read PCFs from anywhere but only ever *write* PCFs into working/ and
artifacts into drawings/. Archive-don't-delete: superseded files move into
an archive/ subfolder rather than being removed.
"""

from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from .render import RenderOptions

PROJECT_FILE = "project.json"
WIKI_FILE = "wiki.md"
SOURCE_DIR = "source"
WORKING_DIR = "working"
DRAWINGS_DIR = "drawings"
ARCHIVE_DIR = "archive"
# Default name for a workspace created beside/inside a CAD-managed export
# folder (the folder itself is never touched -- outputs live here).
DEFAULT_WORKSPACE_NAME = "pcf2isometric"

# Split sheets are derived files: <parent stem>_SHT_<n>.pcf for geometric
# splits, <parent stem>_SPL_<value>.pcf for key-based (spool/line) splits.
# They are never sources, never snapshot targets, and never part of the
# source audit.
SHEET_RE = re.compile(r"_SHT_(\d+)$", re.IGNORECASE)
KEY_SHEET_RE = re.compile(r"_SPL_.+$", re.IGNORECASE)

# Zips are transport, never a deliverable. An archive found inside a
# workspace or a CAD export folder is always a defect: either outputs were
# zipped up "for convenience" instead of being left on disk as the folder
# tree, or a transport zip was delivered and never unpacked. Both hide the
# real deliverable and break the source audit, so every entry point that
# touches a workspace scans for these and says so loudly.
ARCHIVE_SUFFIXES = (
    ".zip", ".7z", ".rar", ".tar", ".tgz", ".tbz2", ".txz",
    ".tar.gz", ".tar.bz2", ".tar.xz",
)
ZIP_RULE = ("zips are transport, never a deliverable -- hand over the "
            "workspace folder tree, never an archive of it")
# Directories with no bearing on the deliverable; walking them is wasted
# work on a large CAD folder.
_WALK_SKIP = {".git", "__pycache__", ".venv", "venv", "node_modules",
              ".pytest_cache", ".ruff_cache"}


def is_archive(path: str | Path) -> bool:
    """True for any file whose name ends in an archive extension."""
    name = Path(path).name.lower()
    return any(name.endswith(suffix) for suffix in ARCHIVE_SUFFIXES)


def find_stray_archives(*roots: str | Path | None) -> list[Path]:
    """Every archive file sitting anywhere in the given trees.

    Pass the workspace root and (when configured) the CAD src folder. A
    non-empty result is a defect to report, not a condition to work
    around: unpack it with unpack_workspace.py, delete the archive, and
    deliver the tree.
    """
    found: list[Path] = []
    seen: set[Path] = set()
    for root in roots:
        if root is None:
            continue
        root = Path(root)
        if not root.is_dir():
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in _WALK_SKIP]
            for name in sorted(filenames):
                if not is_archive(name):
                    continue
                path = Path(dirpath) / name
                resolved = path.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    found.append(path)
    return sorted(found)


def stray_archive_problem(path: str | Path) -> str:
    """The one-line defect report for a stray archive."""
    return (f"archive in the workspace/CAD tree ({ZIP_RULE}): {path} -- "
            f"unpack it (unpack_workspace.py \"{path}\" --delete-zip) or "
            f"delete it, then deliver the folder tree")


def is_sheet_pcf(path: str | Path) -> bool:
    """True for derived split-sheet files (_SHT_<n> or _SPL_<value>)."""
    stem = Path(path).stem
    return SHEET_RE.search(stem) is not None or KEY_SHEET_RE.search(stem) is not None


def sheet_parent_stem(path: str | Path) -> str:
    """The originating file's stem for a sheet path (identity otherwise)."""
    return KEY_SHEET_RE.sub("", SHEET_RE.sub("", Path(path).stem))


def sanitize_sheet_value(value: str) -> str:
    """A grouping value made filename-safe for _SPL_<value> sheet names."""
    safe = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip()).strip("-_.")
    return (safe or "NONE")[:48]


@dataclass
class ProjectConfig:
    north_axis: str = "Y"
    quadrant: str = "NE"
    paper: str = "tabloid"
    logo_path: str | None = None
    company_lines: list[str] = field(default_factory=list)
    project: str = ""
    job: str = ""
    requirements: list[str] = field(default_factory=list)
    # Standing drawing notes (numbered NOTES block on every sheet) and
    # drawing properties (title-block key/value column), applied to every
    # render in this job.
    notes: list[str] = field(default_factory=list)
    properties: dict[str, str] = field(default_factory=dict)
    detect_ffw: bool = False
    # Balloon crowding preference; adjusted semantically via
    # render_iso.py --spread-welds / --pack-welds and remembered per job.
    balloon_spread: float = 1.0
    # The CAD-managed folder the customer's PCF exports live in (absolute, or
    # relative to the project root). The tools only ever READ from it;
    # check_source.py compares it against source/ and pulls copies in.
    src_dir: str | None = None
    # Sheet-splitting policy: split_by_distance.py separates piping groups
    # further apart than split_distance (native units; null = 1 inch / 25 mm
    # by the file's units). With auto_split true the agent applies a needed
    # split without asking first.
    split_distance: float | None = None
    auto_split: bool = False
    # Key-based splitting: split_by_key.py splits masters into one sheet per
    # value of this attribute (e.g. SPOOL-IDENTIFIER). When set, it is the
    # job's preferred split and the pre-render check runs against it instead
    # of the distance grouping.
    split_key: str | None = None
    # The job's split policy, chosen from process_pcf.py's four options:
    # "none" (full-file render), "key" (one sheet per split_key value),
    # "distance" (geometric sheets), or "hybrid" (distance sheets with spool
    # labels and Field Weld declarations between spools). None = not chosen
    # yet; a legacy split_key alone implies "key".
    split_mode: str | None = None
    # Hybrid mode's labelling attribute (usually SPOOL-IDENTIFIER): spools
    # are labelled on each sheet and inter-spool welds become field welds.
    spool_key: str | None = None
    # Soft target for components per drawing, used only to rank the split
    # options presented to the user; any option can still be chosen.
    target_min_components: int = 10
    target_max_components: int = 20
    # Drawings are schematic by default: pipe runs are compressed/stretched
    # into a readable range and the sheet is stamped NOT TO SCALE. Rendering
    # only -- cut lengths, BOM and dimension text always stay true.
    not_to_scale: bool = True
    # Degrees of visual skew given to piping that is flat to the ground but
    # runs on a 45/135/225/315-degree bearing, which a true isometric would
    # draw as a vertical line on top of the risers. Set 0 for a strict
    # isometric. Rendering only.
    skew_deg: float = RenderOptions().skew_deg
    # Page length (mm) of the bar drawn across a flange's gasket face.
    # Defaults to twice the standard dimension offset; raise it if a shop
    # wants flanges to shout louder on a busy sheet.
    flange_bar_mm: float = RenderOptions().flange_bar_mm

    def render_options(self, **overrides) -> RenderOptions:
        options = RenderOptions(
            paper=self.paper,
            north_axis=self.north_axis,
            quadrant=self.quadrant,
            logo_path=self.logo_path,
            company_lines=list(self.company_lines),
            project=self.project,
            job=self.job,
            requirements=list(self.requirements),
            notes=list(self.notes),
            properties=dict(self.properties),
            detect_ffw=self.detect_ffw,
            balloon_spread=self.balloon_spread,
            not_to_scale=self.not_to_scale,
            skew_deg=self.skew_deg,
            flange_bar_mm=self.flange_bar_mm,
            spool_key=self.spool_key if self.split_mode == "hybrid" else None,
        )
        # Standing notes/properties are additive: a per-render --note or
        # --property must never silently drop the job's standing entries.
        extra_notes = overrides.pop("notes", None)
        if extra_notes:
            options.notes.extend(n for n in extra_notes if n not in options.notes)
        extra_properties = overrides.pop("properties", None)
        if extra_properties:
            options.properties.update(extra_properties)
        for key, value in overrides.items():
            if value is not None and hasattr(options, key):
                setattr(options, key, value)
        return options


def init_project(root: str | Path, config: ProjectConfig | None = None) -> Path:
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    for sub in (SOURCE_DIR, WORKING_DIR, DRAWINGS_DIR):
        (root / sub).mkdir(exist_ok=True)
    config_path = root / PROJECT_FILE
    if not config_path.exists():
        save_config(root, config or ProjectConfig())
    wiki = root / WIKI_FILE
    if not wiki.exists():
        wiki.write_text(
            "# Project log\n\n"
            "One entry per change. Newest first. Record what was done, to "
            "which file, and why, so drawing revisions are traceable.\n\n",
            encoding="utf-8",
        )
    return root


def find_project_root(start: str | Path) -> Path | None:
    """Walk up from start looking for a project.json."""
    current = Path(start).resolve()
    if current.is_file():
        current = current.parent
    for candidate in (current, *current.parents):
        if (candidate / PROJECT_FILE).is_file():
            return candidate
    return None


def load_config(root: str | Path) -> ProjectConfig:
    path = Path(root) / PROJECT_FILE
    if not path.is_file():
        return ProjectConfig()
    data = json.loads(path.read_text(encoding="utf-8"))
    config = ProjectConfig()
    for key, value in data.items():
        if hasattr(config, key):
            setattr(config, key, value)
    return config


def save_config(root: str | Path, config: ProjectConfig) -> Path:
    path = Path(root) / PROJECT_FILE
    path.write_text(json.dumps(asdict(config), indent=2), encoding="utf-8")
    return path


def resolve_src_dir(root: str | Path, config: ProjectConfig | None = None) -> Path | None:
    """The CAD-managed export folder this project mirrors, if configured."""
    config = config if config is not None else load_config(root)
    if not config.src_dir:
        return None
    src = Path(config.src_dir)
    if not src.is_absolute():
        src = Path(root) / src
    return src


def log_to_wiki(root: str | Path, message: str) -> None:
    """Prepend a timestamped entry under the wiki header."""
    wiki = Path(root) / WIKI_FILE
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    entry = f"- **{stamp}** - {message}"
    if wiki.is_file():
        lines = wiki.read_text(encoding="utf-8").splitlines()
        insert_at = next(
            (n for n, line in enumerate(lines) if line.startswith("- **")),
            len(lines),
        )
        lines.insert(insert_at, entry)
        wiki.write_text("\n".join(lines) + "\n", encoding="utf-8")
    else:
        wiki.write_text(f"# Project log\n\n{entry}\n", encoding="utf-8")


def archive_file(path: str | Path) -> Path:
    """Move a file into an archive/ folder beside it (never delete)."""
    path = Path(path)
    archive = path.parent / ARCHIVE_DIR
    archive.mkdir(exist_ok=True)
    target = archive / path.name
    n = 1
    while target.exists():
        target = archive / f"{path.stem}_{n}{path.suffix}"
        n += 1
    shutil.move(str(path), str(target))
    return target


def drawings_dir_for(pcf_path: str | Path) -> Path:
    """Where derived artifacts for this PCF belong."""
    root = find_project_root(pcf_path)
    if root is not None:
        out = root / DRAWINGS_DIR
        out.mkdir(exist_ok=True)
        return out
    return Path(pcf_path).parent


def working_dir_for(pcf_path: str | Path) -> Path:
    """Where new/modified PCFs belong (never overwrite source/)."""
    root = find_project_root(pcf_path)
    if root is not None:
        out = root / WORKING_DIR
        out.mkdir(exist_ok=True)
        return out
    return Path(pcf_path).parent
