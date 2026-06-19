"""preview.py — publish-bot lite.

Converts a single DOCX draft to HTML and opens it in the browser.
No git push, no Google Sheets update, no archiving.

Usage:
  python preview.py                       # opens a file picker dialog
  python preview.py path\\to\\report.docx  # skip the dialog
"""
from __future__ import annotations

import os
import shutil
import sys
import webbrowser
from pathlib import Path

# --- path resolution: handles both script mode and PyInstaller frozen exe ---
_FROZEN = getattr(sys, "frozen", False)

if _FROZEN:
    # Bundled exe: assets live in the PyInstaller temp extraction dir
    _BUNDLE_DIR = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    _EXE_DIR = Path(sys.executable).parent
    os.environ["PYPANDOC_PANDOC"] = str(_BUNDLE_DIR / "pandoc.exe")
    _CSS_DIR = _BUNDLE_DIR / "style"
    _HEADER_ASSETS = _BUNDLE_DIR / "style" / "rs_header"
    _OUTPUT_ROOT = _EXE_DIR / "preview_output"
else:
    # Normal script: paths relative to repo root
    _REPO_ROOT = Path(__file__).resolve().parent
    if str(_REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(_REPO_ROOT))
    _CSS_DIR = _REPO_ROOT / "style"
    _HEADER_ASSETS = _REPO_ROOT / "style" / "rs_header"
    _OUTPUT_ROOT = _REPO_ROOT / "preview_output"

from utils.conversion import (
    convert_docx_to_html,
    ensure_css_assets,
    parse_document_metadata,
    patch_html_after_conversion,
    resolve_publication_display_date,
    resolve_publication_year,
)


def _pick_docx() -> Path:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        chosen = filedialog.askopenfilename(
            title="Select a DOCX draft to preview",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        root.destroy()
        if not chosen:
            print("No file selected.")
            sys.exit(0)
        return Path(chosen)
    except Exception as exc:
        print(f"File picker unavailable ({exc}).")
        print("Drag and drop a .docx file onto FPC_Preview.exe instead.")
        sys.exit(1)


def _copy_header_assets(output_dir: Path) -> None:
    dest = output_dir / "assets" / "rs_header"
    dest.mkdir(parents=True, exist_ok=True)
    for f in _HEADER_ASSETS.glob("*"):
        if f.is_file():
            shutil.copy2(f, dest / f.name)


def run_preview(docx_path: Path) -> None:
    docx_path = docx_path.resolve()
    if not docx_path.exists():
        print(f"File not found: {docx_path}")
        sys.exit(1)
    if docx_path.suffix.lower() != ".docx":
        print(f"Expected a .docx file, got: {docx_path.name}")
        sys.exit(1)

    output_dir = _OUTPUT_ROOT / docx_path.stem
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    print(f"Converting {docx_path.name} ...")

    metadata = parse_document_metadata(docx_path)
    year = resolve_publication_year("")
    display_date = resolve_publication_display_date("")

    ensure_css_assets(_CSS_DIR, output_dir, None)
    _copy_header_assets(output_dir)

    html_path = output_dir / "index.html"
    convert_docx_to_html(docx_path, html_path, output_dir / "assets", metadata.title, metadata.author)
    patch_html_after_conversion(
        html_path=html_path,
        input_docx=docx_path,
        author=metadata.author,
        publication_year=year,
        publication_display_date=display_date,
        folder_name=docx_path.stem,
        header_relative_path="assets/rs_header",
        web_root="",
    )

    print(f"Done. Output: {html_path}")
    webbrowser.open(html_path.as_uri())


def main() -> None:
    if len(sys.argv) > 1:
        run_preview(Path(sys.argv[1]))
    else:
        run_preview(_pick_docx())


if __name__ == "__main__":
    main()
