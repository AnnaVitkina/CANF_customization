"""
Empty the CANF working folders: input/, output/, partly_df/

Safe for Google Colab when run as:
    exec(open("clean_folders.py").read())

If auto-detection fails, set before running:
    import os
    os.environ["CANF_PROJECT_ROOT"] = "/content/CANF_customization"  # your clone path
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Iterable, List, Optional

_FOLDER_NAMES = ("input", "output", "partly_df")


def _project_root() -> Path:
    env = os.environ.get("CANF_PROJECT_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    cwd = Path.cwd().resolve()
    for rel in (
        Path("CANF_customization"),
        Path("CANF-customization"),
        Path("Apple CANF customization"),
    ):
        p = (cwd / rel).resolve()
        if p.is_dir() and (p / "shipment_input.py").is_file():
            return p
    return cwd


def _empty_directory(path: Path) -> List[str]:
    """Remove all contents of path; create path if missing. Returns list of removed names."""
    removed: List[str] = []
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
        return removed
    if not path.is_dir():
        raise NotADirectoryError(f"Not a directory: {path}")

    for child in sorted(path.iterdir(), key=lambda p: str(p).lower()):
        rel = child.name
        try:
            if child.is_symlink() or child.is_file():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                shutil.rmtree(child)
            removed.append(rel)
        except OSError as e:
            print(f"   ⚠️ Could not remove {child}: {e}")
    return removed


def clean_canf_folders(project_root: Optional[os.PathLike] = None) -> Path:
    """
    Delete all files and subfolders inside input/, output/, and partly_df/ under project_root.
    The three folders are recreated empty.

    Args:
        project_root: Repo root (folder containing shipment_input.py). Default: auto-detect.

    Returns:
        Resolved project root path.
    """
    root = Path(project_root).resolve() if project_root is not None else _project_root()
    print(f"📁 Cleaning CANF folders under: {root}")

    for name in _FOLDER_NAMES:
        target = root / name
        removed = _empty_directory(target)
        target.mkdir(parents=True, exist_ok=True)
        if removed:
            print(f"   ✓ {name}/  cleared ({len(removed)} item(s))")
        else:
            print(f"   ✓ {name}/  (empty or created)")

    print("✅ Done: input, output, and partly_df are empty.")
    return root


def main(argv: Optional[Iterable[str]] = None) -> int:
    argv = list(argv) if argv is not None else sys.argv[1:]
    root = None
    if argv and argv[0] in ("-h", "--help"):
        print(__doc__)
        print("Usage: python clean_folders.py [PROJECT_ROOT]")
        return 0
    if argv:
        root = argv[0]
    clean_canf_folders(project_root=root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
