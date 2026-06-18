from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """Resolved application configuration used by the runtime pipeline.

    This dataclass is the single normalized source of configuration after
    `config.json` has been read and path strings have been converted into
    absolute `Path` objects. Keeping this as a strongly typed container gives
    the rest of the codebase a stable contract and avoids repeating path
    resolution logic across modules.

    Important behavior:
    - All filesystem paths stored here are absolute.
    - `dry_run` influences where archive batches are created and whether live
      publishing and git push behavior occur.
    - `dry_run_meta_save` controls whether Google Sheets metadata is updated
      during a dry run.
    - `validate_first` determines whether `main.py` performs a preflight check
      and asks the operator to confirm before continuing.
    """
    input: Path
    complete: Path
    dry_run: bool
    dry_run_meta_save: bool
    validate_first: bool
    meta_table: str
    meta_sheet: str
    google_api: Path
    custom_name: str
    custom_date: str
    custom_css: Path | None
    commit_message: str
    preview: bool
    keep_input: bool
    repo_root: Path
    members_repo_root: Path
    live_reports_root: Path
    header_assets_root: Path
    dry_run_root: Path
    default_css_dir: Path

    @property
    def archive_root(self) -> Path:
        """Return the root directory where batch archives should be created.

        Dry runs are intentionally isolated inside `complete/dry_run` so the
        test archive trail does not mix with production archives. Live runs
        archive directly into `complete`.
        """
        return self.dry_run_root if self.dry_run else self.complete


def _resolve_optional_path(raw_value: str, repo_root: Path) -> Path | None:
    """Resolve an optional config path into an absolute `Path`.

    Parameters:
    - `raw_value`: The string value loaded from `config.json`. This may be an
      empty string, a relative path, or an absolute path.
    - `repo_root`: The publish-bot repository root. Relative paths are resolved
      against this directory.

    Returns:
    - `None` when the config field is blank.
    - An absolute `Path` when the value is present.

    Why this exists:
    Optional config fields such as `custom_css` should be easy to leave blank
    without forcing downstream code to special-case empty strings everywhere.
    """
    value = (raw_value or "").strip()
    if not value:
        return None
    path = Path(value)
    if not path.is_absolute():
        path = repo_root / path
    return path.resolve()


def load_config(config_path: Path) -> AppConfig:
    """Load `config.json`, resolve all paths, and build the runtime config.

    Parameters:
    - `config_path`: Absolute or relative path to the JSON config file.

    Returns:
    - A fully populated `AppConfig` instance with absolute paths and derived
      directory locations for the members repo, archive roots, and CSS assets.

    Workflow notes:
    1. The JSON file is read from disk.
    2. The publish-bot repo root is inferred from the config file location.
    3. The sibling `members` repo paths are derived automatically so the user
       does not need to duplicate them in config.
    4. Dry-run archive paths are derived from the configured `complete` path.

    This function is intentionally conservative: it resolves paths but does not
    create directories or validate existence. Structural validation happens
    later in the processing layer.
    """
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    repo_root = config_path.resolve().parent
    members_root = repo_root.parent / "members"
    dry_run_root = Path(payload["complete"]).resolve() / "dry_run"

    return AppConfig(
        input=Path(payload["input"]).resolve(),
        complete=Path(payload["complete"]).resolve(),
        dry_run=bool(payload["dry_run"]),
        dry_run_meta_save=bool(payload.get("dry_run_meta_save", True)),
        validate_first=bool(payload.get("validate_first", False)),
        meta_table=str(payload["meta_table"]),
        meta_sheet=str(payload["meta_sheet"]),
        google_api=Path(payload["google_api"]).resolve(),
        custom_name=str(payload.get("custom_name", "")).strip(),
        custom_date=str(payload.get("custom_date", "")).strip(),
        custom_css=_resolve_optional_path(str(payload.get("custom_css", "")), repo_root),
        commit_message=str(payload.get("commit_message", "")).strip(),
        preview=bool(payload.get("preview", False)),
        keep_input=bool(payload.get("keep_input", False)),
        repo_root=repo_root,
        members_repo_root=members_root.resolve(),
        live_reports_root=(members_root / "research_summaries").resolve(),
        header_assets_root=(repo_root / "style" / "rs_header").resolve(),
        dry_run_root=dry_run_root,
        default_css_dir=(repo_root / "style").resolve(),
    )
