# publish-bot v2.0

`publish-bot` is a local Windows automation tool for turning Word research summary documents into publishable HTML report folders, archiving each batch run, optionally updating Google Sheets metadata, and optionally pushing the live site updates to the `members` repo.

This version is intentionally more modular than the legacy script in `publish-bot_depreciated`, but it still reuses the same core ideas that produced the desired HTML output style:

- DOCX metadata scraping from labeled paragraphs
- Pandoc-based DOCX to HTML conversion
- HTML post-processing to enforce the report layout
- deterministic folder naming
- duplicate handling for already-deployed reports
- structured archive creation for auditability

## Repository Layout

```text
publish-bot/
  .env/
    Gsheet-creds.json
  complete/
  css/
    publication.css
  input/
  utils/
    config.py
    conversion.py
    git_ops.py
    gsheets.py
    processor.py
  config.json
  main.py
  requirements.txt
  README.md
```

## High-Level Workflow

Every run follows this general sequence:

1. `main.py` loads `config.json`.
2. If `validate_first` is `true`, the script performs a preflight validation pass and prints warnings/errors.
3. If validation succeeds and interactive confirmation is enabled via `validate_first`, the user is asked whether to continue.
4. The processor discovers all `.docx` files from the configured `input`.
5. A batch folder is created inside the proper archive root.
6. For each document:
   - metadata is parsed
   - a final folder name is generated
   - an archive report folder is created
   - CSS and header assets are copied into that report folder
   - the original source `.docx` is copied into the archive copy
   - Pandoc converts the DOCX to `index.html`
   - HTML is patched to match the report style
   - an audit text file is written into the archived report folder
   - if this is a live run, the report folder is copied to the `members` repo and the `.docx` file is removed from that live copy
   - if metadata saving is enabled for this run, a row is added to Google Sheets
7. A batch log is written or updated inside the batch folder.
8. Processed `.docx` files are removed from the input location.
9. For live runs only, git add / commit / push is executed from the `members` repo.

## Required External Dependencies

The current code expects:

- Python 3.10+
- Pandoc available on the machine
- The Python package `pypandoc`
- The Python package `gspread`
- Access to the `members` repo at `C:\2_Scripts\members`
- A Google service account credential JSON file

Install Python dependencies with:

```powershell
pip install -r requirements.txt
```

Current `requirements.txt`:

```text
pypandoc
gspread
```

## Required Neighbor Directories

The script derives some paths automatically and expects this local layout:

```text
C:\2_Scripts\
  publish-bot\
  members\
    assets\
      rs_header\
    research_summaries\
```

### Required `members` paths

- `C:\2_Scripts\members\assets\rs_header`
- `C:\2_Scripts\members\research_summaries`

## Configuration

All runtime behavior is driven by `config.json`.

### Current config shape

```json
{
  "input": "C:\\2_Scripts\\publish-bot\\input",
  "complete": "C:\\2_Scripts\\publish-bot\\complete",
  "dry_run": true,
  "dry_run_meta_save": true,
  "validate_first": false,
  "meta_table": "fpc-reports",
  "meta_sheet": "test",
  "google_api": "C:\\2_Scripts\\publish-bot\\.env\\Gsheet-creds.json",
  "custom_name": "",
  "custom_date": "",
  "custom_css": ""
}
```

## Configuration Parameters

### `input`

Type: string path

Required: yes

Accepted forms:

- full path to a single `.docx` file
- full path to a folder containing `.docx` files

Effect:

- Determines which source documents will be processed in the current run.
- If it points to a folder, every `.docx` file in that folder is processed.

### `complete`

Type: string path

Required: yes

Effect:

- Defines the archive root for finished runs.
- Live runs archive into `complete\`.
- Dry runs archive into `complete\dry_run\`.

### `dry_run`

Type: boolean

Required: yes

Effect when `true`:

- No git push occurs.
- Reports are archived under `complete\dry_run\...`.
- Live site folders are not copied into `C:\2_Scripts\members\research_summaries`.

Effect when `false`:

- Reports archive into `complete\...`.
- Reports are copied into the live site folder.
- Git commit/push is attempted inside the `members` repo.

### `dry_run_meta_save`

Type: boolean

Required: no

Default: `true`

Effect:

- Only matters when `dry_run` is `true`.
- If `true`, Google Sheets metadata rows are still inserted during dry runs.
- If `false`, dry runs skip Google Sheets updates.

### `validate_first`

Type: boolean

Required: no

Default: `false`

Effect:

- If `true`, the script performs a validation pass before any output work begins.
- Validation results are printed to the terminal.
- If validation succeeds, the user is asked whether to proceed.
- If validation fails, processing halts before any batch folder is created.

### `meta_table`

Type: string

Required: yes if metadata saving is enabled

Effect:

- The Google Spreadsheet name opened through `gspread`.
- Example: `fpc-reports`

### `meta_sheet`

Type: string

Required: yes if metadata saving is enabled

Effect:

- The worksheet name inside the spreadsheet.
- Example: `test`

### `google_api`

Type: string path

Required: yes if metadata saving is enabled

Effect:

- Path to the Google service account JSON credential file.
- This should live under `.env/` so it stays out of git.

### `custom_name`

Type: string

Required: no

Default: empty string

Effect:

- If blank, the script uses the standard folder naming pattern:
  `AuthorLastName_Year_Title`
- If populated, this value completely overrides the standard folder name after sanitization.

### `custom_date`

Type: string

Required: no

Expected format: `YYYYMMDD`

Effect:

- If valid, the first four digits are used as the publication year in naming and citation metadata.
- If blank or invalid, the current year is used.

### `custom_css`

Type: string path

Required: no

Effect:

- If blank, `css/publication.css` is copied into each report folder.
- If provided, that file is copied into each report folder as `publication.css`.

## Input Document Requirements

The current validation and processing logic expects:

- a readable `.docx` file
- readable paragraph content inside the DOCX XML
- required metadata labels present as readable paragraphs
- actual body content after metadata lines are removed

### Required metadata headers

These are mandatory:

- `Title:`
- `Author:`
- `New or Updated:`
- `Species:`

### Optional metadata header

- `RW Studies Used:`

### Body content requirement

The report must contain non-metadata body content. A document that only contains metadata lines and no real body paragraphs will fail validation.

## Validation Behavior

When `validate_first` is enabled, the validator checks:

- that input files exist
- that the Google credential file exists
- that the `rs_header` asset folder exists
- that the CSS source exists
- that each `.docx` can be read
- that required metadata fields are present
- that body content exists
- whether the final folder name needed normalization
- whether a duplicate live folder already exists

### Validation result levels

#### Errors

Errors stop the run. Examples:

- missing required metadata
- unreadable `.docx`
- missing credential file
- missing header asset directory
- missing CSS source
- missing body content

#### Warnings

Warnings do not stop the run by themselves. Examples:

- quotes removed from folder names
- general filename normalization
- duplicate live folder detected, causing `_1`, `_2`, etc.

## Folder Naming

Default naming pattern:

```text
AuthorLastName_Year_Title
```

Rules applied:

- Title is limited to the first 8 words for the folder name.
- Quotes are removed.
- Other punctuation is normalized into underscores.
- Repeated underscores are collapsed.
- If the live site already contains that folder name, an incrementing suffix is added:
  - `_1`
  - `_2`
  - etc.

Example:

```text
Trlica_2026_Live_Crown_Response_in_Loblolly_Pine
```

## Batch Folder Naming

Each run creates a unique batch archive folder.

### Live run pattern

```text
YYYYMMDD_{ReportCount}_{Index}
```

Example:

```text
20260429_5_1
```

### Dry run pattern

```text
YYYYMMDD_dry_{ReportCount}_{Index}
```

Example:

```text
20260429_dry_5_1
```

## Triple-Save / Distribution Model

### 1. Master archive

The archive copy is the source of truth for what happened during the run.

Each archived report folder contains:

- `index.html`
- `publication.css`
- `assets\...`
- original source `.docx`
- `publish_audit.txt`

### 2. Live site copy

For non-dry runs, the archived report folder is copied to:

```text
C:\2_Scripts\members\research_summaries\[FolderName]
```

Then any `.docx` file is removed from that live copy so only web-ready artifacts remain.

### 3. Input cleanup

After a successful batch, processed `.docx` files are removed from the configured input location.

## Logging and Audit Files

### Batch log

Each batch folder gets a plain text log:

```text
run.log
```

This log includes:

- run timestamp
- batch ID
- dry run status
- dry-run metadata-save status
- validate-first status
- input source
- report count
- validation issues
- processed report summary

### Per-report audit file

Each archived report folder gets:

```text
publish_audit.txt
```

This file currently records:

- run timestamp
- batch ID
- final report folder name
- source `.docx` filename
- dry run status
- duplicate folder flag
- final web link

## HTML Conversion Workflow

For each report, publish-bot performs:

1. Copy stylesheet into the report folder.
2. Copy header assets into `assets\rs_header`.
3. Copy the source `.docx` into the archive report folder.
4. Run Pandoc conversion to create standalone HTML and extract media.
5. Parse authoritative figure captions from DOCX table XML.
6. Replace eligible HTML tables with semantic figure markup.
7. Remove metadata paragraphs from visible HTML.
8. Convert absolute asset paths to relative paths.
9. Insert the official FPC logo/header block.
10. Inject JavaScript-driven lightbox, theme toggle, print behavior, and figure download behavior.
11. Ensure author and year are present in the byline.
12. Rebuild the top-of-page header into the standard publication layout.

## Metadata / Google Sheets Behavior

When metadata saving is enabled for the current run, the script inserts a row at row 2 of the configured worksheet.

Current row order:

1. Title
2. Author
3. Published
4. Section
5. Report Status
6. Regionwide Studies Involved
7. Web Link
8. Timestamp
9. Duplicate

Note:

- The code currently sends an empty value for the sheet timestamp column and relies on the sheet itself if it has formulas or automatic timestamp behavior configured.

## Git Behavior

Git operations only run when:

- `dry_run` is `false`

Git operations are executed only in:

```text
C:\2_Scripts\members
```

The publish-bot repo itself does not receive archive commits. The entire `complete/` tree is ignored in the publish-bot repo.

## Running the Script

From the publish-bot repo root:

```powershell
python main.py
```

If `validate_first` is `true`, you will see a validation summary before processing starts.

### Typical safe test mode

```json
"dry_run": true,
"dry_run_meta_save": false,
"validate_first": true
```

Effect:

- archives to `complete\dry_run`
- does not publish to live site
- does not push git
- does not write to Google Sheets
- performs validation first and asks whether to continue

### Typical live mode

```json
"dry_run": false,
"dry_run_meta_save": true,
"validate_first": true
```

Effect:

- archives to `complete`
- copies reports to the live site repo
- updates Google Sheets
- pushes the members repo
- still gives you a validation checkpoint first

## Module Responsibilities

### `main.py`

Owns:

- application startup
- config loading
- optional preflight interaction
- user-facing console summary

### `utils/config.py`

Owns:

- reading `config.json`
- normalizing string paths into `Path` objects
- deriving members-repo and archive locations

### `utils/conversion.py`

Owns:

- DOCX text extraction
- metadata parsing
- folder-name construction
- Pandoc conversion
- figure caption extraction
- HTML rewriting and page enhancement

### `utils/processor.py`

Owns:

- batch orchestration
- validation policy
- archive folder creation
- live-copy logic
- logging and audit files
- metadata-save gating
- input cleanup

### `utils/gsheets.py`

Owns:

- Google Sheets authentication via `gspread`
- row insertion into the configured sheet

### `utils/git_ops.py`

Owns:

- commit and push behavior for the `members` repo

## Known Operational Assumptions

- The Google credential file is a valid service-account JSON.
- The service account already has access to the target spreadsheet.
- Pandoc is installed and reachable in the Python environment used to run the script.
- The neighboring `members` repo exists in the expected location.
- The report template continues to use the current metadata label names.

## Good Next Improvements

If you keep extending the system, the next high-value upgrades are probably:

- explicit rollback support for failed live publishes
- manifest JSON output per batch
- automated HTML sanity checks after conversion
- richer logging around Google Sheets and git failures
- optional deduplication based on source file hash
