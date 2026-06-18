from __future__ import annotations

import webbrowser
from pathlib import Path

from utils.config import load_config
from utils.processor import process_documents, validate_documents


def main() -> int:
    """Run the publish-bot entrypoint.

    Entry behavior:
    1. Load configuration from `config.json`.
    2. If `validate_first` is enabled, run validation and print every issue.
    3. If validation passes, optionally ask the operator whether to continue.
    4. Execute the full processing workflow.
    5. Print a concise summary for each processed file.

    Returns:
    - `0` for success or user-cancelled-after-validation.
    - `1` when validation fails before processing begins.
    """
    config = load_config(Path(__file__).resolve().parent / "config.json")
    if config.validate_first:
        input_files, validation = validate_documents(config)
        print(f"Validation checked {len(input_files)} file(s).")
        if validation.issues:
            for issue in validation.issues:
                print(f"[{issue.level.upper()}] {issue.source_file}: {issue.message}")
        else:
            print("Validation passed with no issues.")

        if not validation.ok:
            print("Validation failed. Halting before processing.")
            return 1

        reply = input("Validation passed. Proceed with processing? (y/n): ").strip().lower()
        if reply != "y":
            print("Operation cancelled after validation.")
            return 0

    results = process_documents(config)
    for result in results:
        print(f"Processed {result.source_file.name} -> {result.archive_dir}")
        if result.live_dir:
            print(f"Published web copy -> {result.live_dir}")
        if result.duplicate_message:
            print(result.duplicate_message)

    if config.dry_run and config.preview:
        for result in results:
            html_file = result.archive_dir / "index.html"
            if html_file.exists():
                print(f"Opening preview: {html_file}")
                webbrowser.open(html_file.as_uri())

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
