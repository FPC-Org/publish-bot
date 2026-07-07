from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from utils.config import AppConfig
from utils.conversion import (
    build_folder_name,
    convert_docx_to_html,
    ensure_css_assets,
    iter_docx_paragraphs,
    normalize_text,
    parse_document_metadata,
    patch_html_after_conversion,
    resolve_publication_display_date,
    resolve_publication_year,
    sanitize_name_part,
)
from utils.git_ops import push_report_changes
from utils.gsheets import append_report_row


@dataclass(frozen=True)
class ProcessResult:
    """Summary of one processed document after the pipeline completes."""
    source_file: Path
    folder_name: str
    archive_dir: Path
    live_dir: Path | None
    duplicate: bool
    duplicate_message: str | None
    web_link: str


@dataclass(frozen=True)
class ValidationIssue:
    """A single validation message produced during preflight checks."""
    level: str
    source_file: Path
    message: str


@dataclass(frozen=True)
class ValidationResult:
    """Aggregate result of validating one or more input documents."""
    ok: bool
    issues: list[ValidationIssue]


def _iter_input_files(input_path: Path) -> list[Path]:
    """Resolve the configured input into a concrete list of `.docx` files.

    Accepted input forms:
    - a single `.docx` file path
    - a directory containing one or more `.docx` files
    """
    if input_path.is_file():
        return [input_path]
    return sorted(path for path in input_path.glob("*.docx") if path.is_file())


def _ensure_runtime_directories(config: AppConfig) -> None:
    """Create the top-level directories needed for the current run."""
    config.complete.mkdir(parents=True, exist_ok=True)
    config.archive_root.mkdir(parents=True, exist_ok=True)
    config.live_reports_root.mkdir(parents=True, exist_ok=True)


def _copy_header_assets(report_dir: Path, config: AppConfig) -> None:
    """Copy shared header assets into one archived report folder.

    Each archived report is designed to be self-contained. That means the logo
    and any future shared header assets are copied into the report folder rather
    than being referenced from a central archive location.
    """
    header_dir = report_dir / "assets" / "rs_header"
    header_dir.mkdir(parents=True, exist_ok=True)
    for path in config.header_assets_root.glob("*"):
        if path.is_file():
            shutil.copy2(path, header_dir / path.name)


def _allocate_live_folder(live_root: Path, preferred_name: str) -> tuple[str, bool, str | None]:
    """Choose the live folder name, resolving duplicates with numeric suffixes.

    Returns:
    - final folder name
    - whether a duplicate was detected
    - an informational message including the previous deployment timestamp
    """
    candidate = live_root / preferred_name
    if not candidate.exists():
        return preferred_name, False, None

    timestamp = datetime.fromtimestamp(candidate.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    index = 1
    while True:
        indexed_name = f"{preferred_name}_{index}"
        indexed = live_root / indexed_name
        if not indexed.exists():
            message = f"Duplicate folder detected for {preferred_name}. Last deployment timestamp: {timestamp}."
            return indexed_name, True, message
        index += 1


def _normalize_status(status: str) -> str:
    lowered = status.strip().lower()
    if "updated" in lowered:
        return "Update"
    if "new" in lowered:
        return "New"
    return status.strip()


def _split_authors(author: str) -> list[str]:
    """Split a raw author string into individual name parts."""
    parts = [p.strip() for p in re.split(r"\s+and\s+|,", author) if p.strip()]
    return parts


def _last_name(name: str) -> str:
    """Return the last word of a name as the surname."""
    words = name.strip().split()
    return words[-1] if words else name


def _first_author_last_name(author: str) -> str:
    """Return the first author's last name (lowercase) for sort key use."""
    parts = _split_authors(author)
    return _last_name(parts[0]).lower() if parts else author.lower()


def _format_sheet_author(author: str) -> str:
    """Format author string for the sheet: last name only, two last names, or et al."""
    parts = _split_authors(author)
    last_names = [_last_name(p) for p in parts]
    if len(last_names) == 1:
        return last_names[0]
    if len(last_names) == 2:
        return f"{last_names[0]} and {last_names[1]}"
    return f"{last_names[0]} et al."


def _build_sheet_row(title: str, author: str, year: str, section: str, status: str, studies: str, web_link: str, duplicate: bool) -> list[str | bool]:
    """Build the Google Sheets row in the expected column order."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return [title, year, _format_sheet_author(author), section, _normalize_status(status), studies, web_link, timestamp, duplicate]


def _write_batch_log(batch_dir: Path, lines: list[str]) -> None:
    """Write or overwrite the batch-level run log for the current execution.

    The log is intentionally plain text so it remains easy to inspect without
    any extra tooling. The processor writes it once near the start of a run and
    then again at the end with processed-report details appended.
    """
    log_path = batch_dir / "run.log"
    payload = "\n".join(lines).rstrip() + "\n"
    log_path.write_text(payload, encoding="utf-8")


def _write_report_audit_file(
    archive_dir: Path,
    *,
    run_timestamp: str,
    batch_id: str,
    folder_name: str,
    source_name: str,
    dry_run: bool,
    duplicate: bool,
    web_link: str,
) -> None:
    """Create a per-report audit file inside the archived report folder.

    This file is meant for humans doing manual traceability work later. It
    captures the core facts needed to answer:
    - when was this report processed?
    - which batch created it?
    - was it a dry run?
    - what source document did it come from?
    """
    audit_path = archive_dir / "publish_audit.txt"
    lines = [
        f"Run Timestamp: {run_timestamp}",
        f"Batch ID: {batch_id}",
        f"Report Folder: {folder_name}",
        f"Source File: {source_name}",
        f"Dry Run: {dry_run}",
        f"Duplicate Folder Name: {duplicate}",
        f"Web Link: {web_link}",
    ]
    audit_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _remove_processed_inputs(files: list[Path], keep_input: bool) -> None:
    """Delete processed input files after the batch has completed successfully."""
    if keep_input:
        return
    for path in files:
        if path.exists():
            path.unlink()


def _create_batch_dir(base_root: Path, report_count: int, dry_run: bool) -> Path:
    """Create a uniquely indexed batch directory for the current run.

    Naming rules:
    - live: `YYYYMMDD_{ReportCount}_{Index}`
    - dry: `YYYYMMDD_dry_{ReportCount}_{Index}`

    The incrementing index prevents collisions when multiple runs happen on the
    same date with the same number of reports.
    """
    stamp = datetime.now().strftime("%Y%m%d")
    base_name = f"{stamp}_dry_{report_count}" if dry_run else f"{stamp}_{report_count}"
    index = 1
    while True:
        batch_dir = base_root / f"{base_name}_{index}"
        if not batch_dir.exists():
            batch_dir.mkdir(parents=True, exist_ok=True)
            return batch_dir
        index += 1


def _metadata_map(docx_path: Path) -> dict[str, str]:
    """Return a raw metadata dictionary parsed from `Label: Value` paragraphs.

    This helper is used by validation, which needs direct access to the presence
    or absence of specific headers before applying any fallback logic.
    """
    values: dict[str, str] = {}
    for _, text in iter_docx_paragraphs(docx_path):
        match = re.match(r"^\s*([^:]+)\s*:\s*(.*?)\s*$", text, flags=re.IGNORECASE)
        if not match:
            continue
        key = normalize_text(match.group(1)).lower()
        values[key] = normalize_text(match.group(2))
    return values


def validate_documents(config: AppConfig) -> tuple[list[Path], ValidationResult]:
    """Run preflight validation against the configured input set.

    Validation checks include:
    - input `.docx` files exist
    - Google credential path exists
    - required header assets exist
    - CSS source exists
    - documents are readable
    - required metadata fields exist
    - body content exists after metadata is removed
    - filename normalization or duplicate folder conditions are surfaced as
      warnings

    Returns:
    - the list of input files that were discovered
    - a `ValidationResult` describing errors and warnings
    """
    issues: list[ValidationIssue] = []
    input_files = _iter_input_files(config.input)
    if not input_files:
        return input_files, ValidationResult(ok=False, issues=[ValidationIssue("error", config.input, f"No .docx files found in {config.input}")])

    if not config.google_api.exists():
        issues.append(ValidationIssue("error", config.google_api, f"Google credential file not found: {config.google_api}"))
    if not config.header_assets_root.exists():
        issues.append(ValidationIssue("error", config.header_assets_root, f"Header assets path not found: {config.header_assets_root}"))
    if not config.default_css_dir.exists() and not config.custom_css:
        issues.append(ValidationIssue("error", config.default_css_dir, f"CSS directory not found: {config.default_css_dir}"))
    if config.custom_css and not config.custom_css.exists():
        issues.append(ValidationIssue("error", config.custom_css, f"Custom CSS file not found: {config.custom_css}"))

    year = resolve_publication_year(config.custom_date)
    required_fields = {
        "title": "Title",
        "author": "Author",
        "new or updated": "New or Updated",
        "species": "Species",
    }

    for docx_path in input_files:
        try:
            paragraphs = [text for _, text in iter_docx_paragraphs(docx_path)]
        except Exception as exc:
            issues.append(ValidationIssue("error", docx_path, f"Unreadable .docx file: {exc}"))
            continue

        if not paragraphs:
            issues.append(ValidationIssue("error", docx_path, "Document contains no readable text content."))
            continue

        metadata = _metadata_map(docx_path)
        for key, label in required_fields.items():
            if not metadata.get(key):
                issues.append(ValidationIssue("error", docx_path, f"Missing required metadata field: {label}"))

        body_paragraphs = []
        for paragraph in paragraphs:
            lowered = normalize_text(paragraph).lower()
            if any(lowered.startswith(prefix) for prefix in ("title:", "author:", "new or updated:", "species:", "rw studies used:", "date:", "published:")):
                continue
            body_paragraphs.append(paragraph)

        if not body_paragraphs:
            issues.append(ValidationIssue("error", docx_path, "Document body content is missing."))

        doc_meta = parse_document_metadata(docx_path)
        preferred_name = build_folder_name(doc_meta.author, year, doc_meta.title, config.custom_name)
        cleaned_title = sanitize_name_part(doc_meta.title, "Untitled")
        if '"' in doc_meta.title or "'" in doc_meta.title:
            issues.append(ValidationIssue("warning", docx_path, "Quotes were removed from the output folder name."))
        if cleaned_title != normalize_text(doc_meta.title).replace(" ", "_"):
            issues.append(ValidationIssue("warning", docx_path, f"Folder name normalized to: {preferred_name}"))
        duplicate_name, duplicate, duplicate_message = _allocate_live_folder(config.live_reports_root, preferred_name)
        if duplicate and duplicate_message:
            issues.append(ValidationIssue("warning", docx_path, f"{duplicate_message} Next folder name will be: {duplicate_name}."))

    ok = not any(issue.level == "error" for issue in issues)
    return input_files, ValidationResult(ok=ok, issues=issues)


def _copy_live_report(archive_dir: Path, live_dir: Path) -> None:
    """Copy an archived report to the live site folder and strip source docs.

    The archive copy intentionally keeps the original `.docx`, but the live site
    should contain only web-facing artifacts.
    """
    if live_dir.exists():
        shutil.rmtree(live_dir)
    shutil.copytree(archive_dir, live_dir)
    for docx_file in live_dir.rglob("*.docx"):
        docx_file.unlink()


def _should_save_metadata(config: AppConfig) -> bool:
    """Decide whether this run should write metadata to Google Sheets."""
    if not config.dry_run:
        return True
    return config.dry_run_meta_save


def process_documents(config: AppConfig) -> list[ProcessResult]:
    """Execute the full publish-bot workflow for all discovered input files.

    High-level workflow:
    1. Ensure runtime directories exist.
    2. Validate the input set.
    3. Create a unique batch archive folder.
    4. For each document:
       - scrape metadata
       - determine final folder name
       - build the archived report folder
       - copy CSS and header assets
       - copy the source `.docx` into the archive
       - convert the DOCX to HTML
       - patch the HTML into publish format
       - optionally copy the report into the live members repo
       - optionally append a metadata row to Google Sheets
       - write per-report audit info
    5. Update the batch log with run details.
    6. Remove processed files from the input folder.
    7. If this is not a dry run, commit and push changes from the members repo.

    Returns:
    - A list of `ProcessResult` records, one for each processed input file.
    """
    _ensure_runtime_directories(config)
    input_files, validation = validate_documents(config)
    if not input_files:
        raise FileNotFoundError(f"No .docx files found in {config.input}")
    if not validation.ok:
        messages = [f"{issue.source_file}: {issue.message}" for issue in validation.issues if issue.level == "error"]
        raise ValueError("Validation failed:\n" + "\n".join(messages))

    run_start = datetime.now()
    results: list[ProcessResult] = []
    year = resolve_publication_year(config.custom_date)
    display_date = resolve_publication_display_date(config.custom_date)
    reports_root_url = "https://members.forestproductivity.org/rs/"
    batch_dir = _create_batch_dir(config.archive_root, len(input_files), config.dry_run)
    run_timestamp = run_start.strftime("%Y-%m-%d %H:%M:%S")
    log_lines = [
        f"Run Timestamp: {run_timestamp}",
        f"Batch ID: {batch_dir.name}",
        f"Report Count: {len(input_files)}",
        "",
        "=== Configuration ===",
        f"Input: {config.input}",
        f"Complete: {config.complete}",
        f"Dry Run: {config.dry_run}",
        f"Keep Input: {config.keep_input}",
        f"Preview: {config.preview}",
        f"Dry Run Meta Save: {config.dry_run_meta_save}",
        f"Validate First: {config.validate_first}",
        f"Meta Table: {config.meta_table}",
        f"Meta Sheet: {config.meta_sheet}",
        f"Custom Name: {config.custom_name or '(none)'}",
        f"Custom Date: {config.custom_date or '(none)'}",
        f"Custom CSS: {config.custom_css or '(none)'}",
        f"Commit Message: {config.commit_message or '(none)'}",
        "",
    ]
    if validation.issues:
        log_lines.append("Validation Issues:")
        for issue in validation.issues:
            log_lines.append(f"[{issue.level.upper()}] {issue.source_file}: {issue.message}")
    _write_batch_log(batch_dir, log_lines)

    pending_sheet_rows: list[tuple[str, list]] = []

    for docx_path in input_files:
        metadata = parse_document_metadata(docx_path)
        preferred_name = build_folder_name(metadata.author, year, metadata.title, config.custom_name)
        folder_name, duplicate, duplicate_message = _allocate_live_folder(config.live_reports_root, preferred_name)
        archive_dir = batch_dir / folder_name
        archive_dir.mkdir(parents=True, exist_ok=True)
        ensure_css_assets(config.default_css_dir, archive_dir, config.custom_css)
        _copy_header_assets(archive_dir, config)
        shutil.copy2(docx_path, archive_dir / docx_path.name)

        html_path = archive_dir / "index.html"
        convert_docx_to_html(docx_path, html_path, archive_dir / "assets", metadata.title, metadata.author)
        patch_html_after_conversion(
            html_path=html_path,
            input_docx=docx_path,
            author=metadata.author,
            publication_year=year,
            publication_display_date=display_date,
            folder_name=folder_name,
            header_relative_path="assets/rs_header",
            web_root=reports_root_url,
        )

        live_dir = None
        if not config.dry_run:
            live_dir = config.live_reports_root / folder_name
            _copy_live_report(archive_dir, live_dir)

        web_link = f"{reports_root_url}{folder_name}/"
        _write_report_audit_file(
            archive_dir,
            run_timestamp=run_timestamp,
            batch_id=batch_dir.name,
            folder_name=folder_name,
            source_name=docx_path.name,
            dry_run=config.dry_run,
            duplicate=duplicate,
            web_link=web_link,
        )
        if _should_save_metadata(config):
            pending_sheet_rows.append((
                _first_author_last_name(metadata.author),
                _build_sheet_row(
                    metadata.title,
                    metadata.author,
                    year,
                    metadata.section,
                    metadata.status,
                    metadata.studies,
                    web_link,
                    duplicate,
                ),
            ))

        results.append(
            ProcessResult(
                source_file=docx_path,
                folder_name=folder_name,
                archive_dir=archive_dir,
                live_dir=live_dir,
                duplicate=duplicate,
                duplicate_message=duplicate_message,
                web_link=web_link,
            )
        )

    # Insert sheet rows in reverse alpha order so the sheet ends up A→Z top to bottom
    for _, row_values in sorted(pending_sheet_rows, key=lambda x: x[0], reverse=True):
        append_report_row(
            credentials_path=config.google_api,
            spreadsheet_name=config.meta_table,
            sheet_name=config.meta_sheet,
            row_values=row_values,
        )

    log_lines.append("Processed Reports:")
    for result in results:
        log_lines.append(f"{result.source_file.name} -> {result.archive_dir}")
        if result.live_dir:
            log_lines.append(f"Published: {result.live_dir}")
        if result.duplicate_message:
            log_lines.append(f"Note: {result.duplicate_message}")
    runtime_seconds = (datetime.now() - run_start).total_seconds()
    log_lines.append(f"Total Runtime: {runtime_seconds:.1f}s")
    _write_batch_log(batch_dir, log_lines)
    _remove_processed_inputs(input_files, config.keep_input)

    if not config.dry_run:
        base_message = f"publish-bot deployment {datetime.now():%Y-%m-%d %H:%M:%S}"
        commit_msg = f"{base_message} — {config.commit_message}" if config.commit_message else base_message
        push_report_changes(config.members_repo_root, commit_msg)

    return results
