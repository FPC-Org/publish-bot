from __future__ import annotations

import os
import re
import shutil
import urllib.parse
import zipfile
from dataclasses import dataclass
from datetime import datetime
from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W_VAL = f"{{{W_NS}}}val"

PROGRAM_NAME = "Forest Productivity Cooperative Research Summaries"
RESEARCH_SUMMARIES_URL = "https://members.forestproductivity.org/rs/"
HOME_URL = "https://members.forestproductivity.org/rs"
MEMBER_DASHBOARD_URL = "https://www.forestproductivity.org/memberdashboard"
AUTHOR_CONTACTS_URL = "https://www.forestproductivity.org/people"

_ICON_DASHBOARD = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"'
    ' fill="none" stroke="#990000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path stroke="none" d="M0 0h24v24H0z" fill="none"/>'
    '<path d="M5 4h4a1 1 0 0 1 1 1v6a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-6a1 1 0 0 1 1 -1"/>'
    '<path d="M5 16h4a1 1 0 0 1 1 1v2a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-2a1 1 0 0 1 1 -1"/>'
    '<path d="M15 12h4a1 1 0 0 1 1 1v6a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-6a1 1 0 0 1 1 -1"/>'
    '<path d="M15 4h4a1 1 0 0 1 1 1v2a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-2a1 1 0 0 1 1 -1"/></svg>'
)
_ICON_REPORT = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"'
    ' fill="none" stroke="#990000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M14 3v4a1 1 0 0 0 1 1h4"/>'
    '<path d="M17 21h-10a2 2 0 0 1 -2 -2v-14a2 2 0 0 1 2 -2h7l5 5v11a2 2 0 0 1 -2 2"/>'
    '<path d="M8 13h6"/><path d="M8 16h6"/><path d="M8 19h4"/></svg>'
)
_ICON_MAIL = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"'
    ' fill="none" stroke="#990000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path stroke="none" d="M0 0h24v24H0z" fill="none"/>'
    '<path d="M3 7a2 2 0 0 1 2 -2h14a2 2 0 0 1 2 2v10a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2v-10"/>'
    '<path d="M3 7l9 6l9 -6"/></svg>'
)
_ICON_PRINTER = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"'
    ' fill="none" stroke="#990000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path stroke="none" d="M0 0h24v24H0z" fill="none"/>'
    '<path d="M17 17h2a2 2 0 0 0 2 -2v-4a2 2 0 0 0 -2 -2h-14a2 2 0 0 0 -2 2v4a2 2 0 0 0 2 2h2"/>'
    '<path d="M17 9v-4a2 2 0 0 0 -2 -2h-6a2 2 0 0 0 -2 2v4"/>'
    '<path d="M7 15a2 2 0 0 1 2 -2h6a2 2 0 0 1 2 2v4a2 2 0 0 1 -2 2h-6a2 2 0 0 1 -2 -2l0 -4"/></svg>'
)


@dataclass(frozen=True)
class DocumentMetadata:
    """Normalized metadata scraped from the Word document template.

    Field meanings:
    - `title`: Human-readable report title used in HTML and naming logic.
    - `author`: Human-readable author string used in HTML and naming logic.
    - `status`: Value scraped from `New or Updated:`.
    - `section`: Value scraped from `Species:`.
    - `studies`: Value scraped from `RW Studies Used:`. This may be blank.
    """
    title: str
    author: str
    status: str
    section: str
    studies: str


def normalize_text(text: str | None) -> str:
    """Collapse whitespace and trim leading/trailing space.

    This helper is used everywhere metadata or DOCX text is parsed because Word
    XML often contains odd spacing, line breaks, and tabs that should behave as
    simple spaces in filenames, metadata fields, and validation checks.
    """
    return re.sub(r"\s+", " ", text or "").strip()


def iter_docx_paragraphs(docx_path: Path):
    """Yield readable paragraph text from a `.docx` file in document order.

    Parameters:
    - `docx_path`: Path to the Word document being analyzed.

    Yields:
    - `(style, text)` tuples where `style` is the Word paragraph style ID and
      `text` is normalized visible text content.

    Why this exists:
    This is the foundational DOCX reader for the project. Instead of relying on
    a heavyweight document library, the function reads `word/document.xml`
    directly from the DOCX zip package and reconstructs user-visible text from
    the XML nodes we care about.
    """
    with zipfile.ZipFile(docx_path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)

    for paragraph in root.findall(".//w:body/w:p", NS):
        style_el = paragraph.find("./w:pPr/w:pStyle", NS)
        style = (style_el.attrib.get(W_VAL, "") if style_el is not None else "").strip()
        chunks: list[str] = []
        for node in paragraph.iter():
            local = node.tag.rsplit("}", 1)[-1]
            if local == "t" and node.text:
                chunks.append(node.text)
            elif local == "tab":
                chunks.append("\t")
            elif local in {"br", "cr"}:
                chunks.append("\n")
        text = normalize_text("".join(chunks))
        if text:
            yield style, text


def parse_document_metadata(docx_path: Path) -> DocumentMetadata:
    """Parse report metadata fields from the source Word document.

    Required labels expected by the current template:
    - `Title:`
    - `Author:`
    - `New or Updated:`
    - `Species:`

    Optional label:
    - `RW Studies Used:`

    Fallback behavior:
    - If `Title:` is missing, the first readable paragraph becomes the title.
    - If `Author:` is missing, the second readable paragraph becomes the author.

    Returns:
    - A `DocumentMetadata` object with best-effort values.

    Validation for required fields is handled elsewhere. This parser stays
    permissive so the validation layer can make policy decisions cleanly.
    """
    labels = {
        "title": "title",
        "author": "author",
        "new or updated": "status",
        "species": "section",
        "rw studies used": "studies",
    }
    values = {
        "title": "",
        "author": "",
        "status": "",
        "section": "",
        "studies": "",
    }

    paragraphs = [text for _, text in iter_docx_paragraphs(docx_path)]
    for text in paragraphs:
        match = re.match(r"^\s*([^:]+)\s*:\s*(.*?)\s*$", text, flags=re.IGNORECASE)
        if not match:
            continue
        key = normalize_text(match.group(1)).lower()
        value = normalize_text(match.group(2))
        mapped = labels.get(key)
        if mapped and not values[mapped]:
            values[mapped] = value

    if not values["title"] and paragraphs:
        values["title"] = paragraphs[0]
    if not values["author"] and len(paragraphs) > 1:
        values["author"] = paragraphs[1]

    return DocumentMetadata(
        title=values["title"] or docx_path.stem,
        author=values["author"] or "Unknown Author",
        status=values["status"],
        section=values["section"],
        studies=values["studies"],
    )


def resolve_publication_display_date(custom_date: str) -> str:
    """Resolve a human-readable Month YYYY date for display in the report header.

    Parameters:
    - `custom_date`: Optional config override in `YYYYMMDD` format.

    Returns:
    - "Month YYYY" derived from `custom_date` when it matches the required format.
    - "Month YYYY" of the current date otherwise.
    """
    clean = normalize_text(custom_date)
    if re.fullmatch(r"\d{8}", clean):
        year = int(clean[:4])
        month = int(clean[4:6])
        try:
            return datetime(year, month, 1).strftime("%B %Y")
        except ValueError:
            pass
    now = datetime.now()
    return now.strftime("%B %Y")


def resolve_publication_year(custom_date: str) -> str:
    """Resolve the publication year used in folder naming and citations.

    Parameters:
    - `custom_date`: Optional config override expected in `YYYYMMDD` format.

    Returns:
    - The first four digits of `custom_date` when it matches the required
      format.
    - The current year when no valid override is provided.
    """
    clean = normalize_text(custom_date)
    if re.fullmatch(r"\d{8}", clean):
        return clean[:4]
    return str(datetime.now().year)


def sanitize_name_part(text: str, fallback: str) -> str:
    """Convert arbitrary human text into a filesystem-friendly name segment.

    Behavior:
    - Trims and normalizes whitespace.
    - Removes single and double quotes explicitly.
    - Replaces remaining punctuation and spacing runs with underscores.
    - Collapses repeated underscores.
    - Falls back to `fallback` if the result becomes empty.

    This is used for folder naming, where stability and path safety matter more
    than preserving punctuation exactly.
    """
    cleaned = normalize_text(text)
    cleaned = cleaned.replace('"', "").replace("'", "")
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned or fallback


def limit_title_words(title: str, max_words: int = 8) -> str:
    """Cap a title to the first `max_words` words for folder naming.

    The visible HTML title remains the full title. This truncation only affects
    directory naming so path lengths and slug readability stay manageable.
    """
    words = normalize_text(title).split()
    return " ".join(words[:max_words]) if words else title


def extract_author_last_name(author: str) -> str:
    """Extract a best-effort last-name token from an author string.

    The current naming convention is `AuthorLastName_Year_Title`, so we need a
    compact stable author fragment even when the source contains credentials,
    commas, or multiple authors.
    """
    primary = normalize_text(author.split(";", 1)[0].split(",", 1)[0])
    tokens = re.findall(r"[A-Za-z][A-Za-z'\\-]*", primary)
    return tokens[-1] if tokens else "UnknownAuthor"


def build_folder_name(author: str, year: str, title: str, custom_name: str = "") -> str:
    """Build the canonical output folder name for a report.

    Parameters:
    - `author`: Human-readable author string from document metadata.
    - `year`: Resolved publication year.
    - `title`: Human-readable report title.
    - `custom_name`: Optional override from config. If present, it replaces the
      normal naming pattern entirely after sanitization.

    Returns:
    - A sanitized folder name in `AuthorLastName_Year_Title` form unless
      `custom_name` is supplied.
    """
    override = normalize_text(custom_name)
    if override:
        return sanitize_name_part(override, "publication")
    limited_title = limit_title_words(title)
    return "_".join(
        [
            sanitize_name_part(extract_author_last_name(author), "UnknownAuthor"),
            sanitize_name_part(year, str(datetime.now().year)),
            sanitize_name_part(limited_title, "Untitled"),
        ]
    )


def ensure_css_assets(default_css_dir: Path, report_dir: Path, custom_css: Path | None) -> None:
    """Place the correct stylesheet into the generated report folder.

    Rules:
    - If `custom_css` is configured, it is copied into the report folder and
      renamed to `publication.css` so the generated HTML can refer to a stable
      filename.
    - Otherwise the default repo stylesheet at `style/publication.css` is copied.
    """
    report_dir.mkdir(parents=True, exist_ok=True)
    if custom_css:
        shutil.copy2(custom_css, report_dir / "publication.css")
        return

    default_css_dir.mkdir(parents=True, exist_ok=True)
    default_css = default_css_dir / "publication.css"
    if default_css.exists():
        shutil.copy2(default_css, report_dir / "publication.css")


def _convert_emf_in_media(html_path: Path, media_dir: Path) -> None:
    """Convert any .emf files in the media directory to PNG for browser display.

    Browsers cannot render Windows Enhanced Metafile (.emf) images. This
    function uses Windows GDI via PowerShell (.NET System.Drawing) to render
    each EMF to a PNG, patches the HTML src references, then removes the
    original EMF files.
    """
    import os
    import subprocess
    import tempfile

    emf_files = list(media_dir.glob("*.emf"))
    if not emf_files:
        return

    renames: dict[str, str] = {}
    for emf_path in emf_files:
        png_path = emf_path.with_suffix(".png")
        emf_str = str(emf_path).replace("'", "''")
        png_str = str(png_path).replace("'", "''")
        ps_lines = [
            "Add-Type -AssemblyName System.Drawing",
            f"$emf = New-Object System.Drawing.Imaging.Metafile('{emf_str}')",
            "$aspect = if ($emf.Width -gt 0) { $emf.Height / $emf.Width } else { 0.75 }",
            "$dstW = 1200",
            "$dstH = [Math]::Max(1, [int]($dstW * $aspect))",
            "$bmp = New-Object System.Drawing.Bitmap($dstW, $dstH)",
            "$g = [System.Drawing.Graphics]::FromImage($bmp)",
            "$g.Clear([System.Drawing.Color]::White)",
            "$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias",
            "$g.DrawImage($emf, 0, 0, $dstW, $dstH)",
            "$g.Dispose()",
            f"$bmp.Save('{png_str}', [System.Drawing.Imaging.ImageFormat]::Png)",
            "$bmp.Dispose()",
            "$emf.Dispose()",
        ]
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as tf:
            tf.write("\n".join(ps_lines))
            script_path = tf.name
        try:
            result = subprocess.run(
                ["powershell", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script_path],
                capture_output=True, text=True, timeout=30,
            )
            if result.returncode == 0 and png_path.exists():
                renames[emf_path.name] = png_path.name
                emf_path.unlink()
        finally:
            os.unlink(script_path)

    if renames:
        html = html_path.read_text(encoding="utf-8")
        for old_name, new_name in renames.items():
            html = html.replace(old_name, new_name)
        html_path.write_text(html, encoding="utf-8")


def convert_docx_to_html(input_docx: Path, html_path: Path, extract_media_dir: Path, title: str, author: str) -> None:
    """Convert a Word document into standalone HTML using Pandoc.

    Parameters:
    - `input_docx`: Source `.docx` file.
    - `html_path`: Final `index.html` output path.
    - `extract_media_dir`: Directory where Pandoc should unpack embedded images.
    - `title`: Title metadata injected into the generated HTML.
    - `author`: Author metadata injected into the generated HTML.

    Implementation notes:
    - The function imports `pypandoc` lazily so the module can still be loaded
      in environments where Pandoc conversion is unavailable.
    - The generated HTML is standalone and references `publication.css`.
    """
    try:
        import pypandoc
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pypandoc is required for DOCX to HTML conversion. Install it in the Python environment used for publish-bot.") from exc

    extra_args = [
        "--standalone",
        "--wrap=none",
        f"--metadata=title:{title}",
        f"--metadata=author:{author}",
        f"--extract-media={extract_media_dir}",
        "--css=publication.css",
    ]
    pypandoc.convert_file(str(input_docx), "html", outputfile=str(html_path), extra_args=extra_args)
    _convert_emf_in_media(html_path, extract_media_dir / "media")


def _local_name(tag: str) -> str:
    """Return the XML local tag name without its namespace wrapper."""
    return tag.rsplit("}", 1)[-1]


def _render_fldsimple_text(fld_node: ET.Element) -> str:
    """Render visible text from a Word `fldSimple` element.

    Word captions often contain field-generated numbers. Pandoc can lose some
    of that information when flattening tables, so this helper extracts the
    field's human-visible value directly from the XML.
    """
    visible_parts: list[str] = []
    for node in fld_node.iter():
        local = _local_name(node.tag)
        if local == "t" and node.text:
            visible_parts.append(node.text)
        elif local == "tab":
            visible_parts.append("\t")
        elif local in {"br", "cr"}:
            visible_parts.append("\n")
    return normalize_text("".join(visible_parts))


def _collect_cell_text_from_docx(cell_node: ET.Element) -> str:
    """Collect visible text from a Word table cell, including field output.

    This function exists primarily so figure captions authored inside tables can
    be reconstructed faithfully before the HTML post-processing stage.
    """
    parts: list[str] = []

    def walk(node: ET.Element) -> None:
        for child in list(node):
            local = _local_name(child.tag)
            if local == "fldSimple":
                field_text = _render_fldsimple_text(child)
                if field_text:
                    parts.append(field_text)
                continue
            if local == "t" and child.text:
                parts.append(child.text)
            elif local == "tab":
                parts.append("\t")
            elif local in {"br", "cr"}:
                parts.append("\n")
            walk(child)

    walk(cell_node)
    return normalize_text("".join(parts))


def extract_figure_table_captions(docx_path: Path) -> list[list[str]]:
    """Extract figure captions from authoritative two-row Word tables.

    Expected table shape:
    - Row 1: image cells
    - Row 2: caption cells

    Returns:
    - A nested list where each outer item is one figure table and each inner
      item is the caption text for a figure column.

    This lets the HTML patcher restore captions even when Pandoc drops field
    prefixes like `Figure 1`.
    """
    with zipfile.ZipFile(docx_path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)

    def _row_has_image(row) -> bool:
        return any(
            cell.find(".//w:drawing", NS) is not None or cell.find(".//w:pict", NS) is not None
            for cell in row.findall("./w:tc", NS)
        )

    extracted: list[list[str]] = []
    for table in root.findall(".//w:body/w:tbl", NS):
        rows = table.findall("./w:tr", NS)

        if len(rows) == 2:
            # Standard layout: row 0 = images, row 1 = captions (or colspan caption)
            first_cells = rows[0].findall("./w:tc", NS)
            second_cells = rows[1].findall("./w:tc", NS)
            colspan_caption = len(second_cells) == 1 and len(first_cells) > 1
            if not first_cells or (not colspan_caption and len(first_cells) != len(second_cells)):
                continue
            image_indices = [
                idx for idx, cell in enumerate(first_cells)
                if cell.find(".//w:drawing", NS) is not None or cell.find(".//w:pict", NS) is not None
            ]
            if not image_indices:
                continue
            if colspan_caption:
                shared = _collect_cell_text_from_docx(second_cells[0])
                extracted.append([shared for _ in image_indices])
            else:
                extracted.append([_collect_cell_text_from_docx(second_cells[idx]) for idx in image_indices])

        elif len(rows) >= 4 and len(rows) % 2 == 0:
            # Alternating layout: row 0 = image, row 1 = caption, row 2 = image, row 3 = caption …
            if not all(
                _row_has_image(rows[i]) and not _row_has_image(rows[i + 1])
                for i in range(0, len(rows), 2)
            ):
                continue
            captions = [
                _collect_cell_text_from_docx(rows[i].findall("./w:tc", NS)[0])
                for i in range(1, len(rows), 2)
            ]
            extracted.append(captions)

    return extracted


def _extract_row_cells(tr_html: str) -> list[str]:
    """Extract the raw inner HTML for cells inside one HTML table row."""
    return [
        cell.strip()
        for cell in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr_html, flags=re.IGNORECASE | re.DOTALL)
    ]


def _strip_outer_paragraph(text: str) -> str:
    """Remove a single wrapping `<p>` when a fragment contains only one block.

    Pandoc often wraps table cell content in paragraph tags. Figure captions are
    cleaner and easier to style when the `<figcaption>` contains inline content
    directly instead of a nested paragraph.
    """
    cleaned = text.strip()
    match = re.match(r"^<p[^>]*>(.*?)</p>$", cleaned, flags=re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else cleaned


def _inject_attrs_into_img(image_html: str, attrs: dict[str, str]) -> str:
    """Inject CSS classes and data attributes into the first image tag found.

    The generated figure HTML uses `data-*` attributes to power lightbox and
    download behavior without needing a server or extra metadata file.
    """
    def replace_img(match: re.Match[str]) -> str:
        tag_attrs = match.group(1)
        self_closing = False
        if tag_attrs.rstrip().endswith("/"):
            self_closing = True
            tag_attrs = re.sub(r"/\s*$", "", tag_attrs)
        if re.search(r'class\s*=', tag_attrs, flags=re.IGNORECASE):
            tag_attrs = re.sub(
                r'(class\s*=\s*")([^"]*)"',
                lambda m: f'{m.group(1)}{m.group(2)} pub-figure-img"',
                tag_attrs,
                count=1,
                flags=re.IGNORECASE,
            )
        else:
            tag_attrs += ' class="pub-figure-img"'
        for key, value in attrs.items():
            tag_attrs += f' {key}="{escape(value, quote=True)}"'
        closing = " />" if self_closing else ">"
        return f"<img{tag_attrs}{closing}"

    return re.sub(r"<img\b([^>]*)>", replace_img, image_html, count=1, flags=re.IGNORECASE | re.DOTALL)


def _make_figure_block(
    image_cell_html: str,
    caption_cell_html: str,
    *,
    caption_override: str | None,
    citation_author: str,
    citation_year: str,
    citation_title: str,
    figure_index: int,
    folder_name: str,
) -> str:
    """Build one semantic `<figure>` block from image and caption fragments.

    The result includes:
    - a normalized image tag with metadata attributes
    - a `<figcaption>` element
    - a client-side download button for figure export
    """
    caption_content = _strip_outer_paragraph(caption_cell_html)
    if caption_override and normalize_text(caption_override):
        caption_content = escape(caption_override)
    caption_plain = normalize_text(re.sub(r"<[^>]+>", "", caption_content))
    image_with_attrs = _inject_attrs_into_img(
        image_cell_html,
        {
            "data-caption": caption_plain,
            "data-author": citation_author,
            "data-year": citation_year,
            "data-title": citation_title,
            "data-program": PROGRAM_NAME,
            "data-download-name": f"{folder_name}-figure-{figure_index}",
        },
    )
    return (
        '<figure class="figure-block">\n'
        f"{image_with_attrs}\n"
        f'<figcaption class="pub-figure-caption">{caption_content}</figcaption>\n'
        '<button type="button" class="pub-figure-download" aria-label="Download image with caption">Download</button>\n'
        "</figure>"
    )


def _convert_figure_tables(
    html: str,
    docx_captions: list[list[str]],
    *,
    citation_author: str,
    citation_year: str,
    citation_title: str,
    folder_name: str,
) -> str:
    """Replace eligible HTML tables with semantic figure markup.

    The DOCX authoring convention is treated as authoritative: only tables with
    exactly two rows and at least one image cell are rewritten. This keeps
    ordinary tables intact while upgrading figure tables into cleaner HTML.
    """
    table_pattern = re.compile(r"<table\b[^>]*>.*?</table>", flags=re.IGNORECASE | re.DOTALL)
    figure_table_index = 0
    figure_counter = 0

    def _html_row_has_image(tr_html: str) -> bool:
        return bool(re.search(r"<img\b", tr_html, flags=re.IGNORECASE))

    def replace_table(match: re.Match[str]) -> str:
        nonlocal figure_table_index, figure_counter
        table_html = match.group(0)
        row_matches = list(re.finditer(r"<tr\b[^>]*>.*?</tr>", table_html, flags=re.IGNORECASE | re.DOTALL))

        # Determine layout: standard 2-row or alternating image/caption pairs
        side_by_side = False
        image_caption_pairs: list[tuple[list[str], list[str], bool]] = []

        if len(row_matches) == 2:
            first_row_cells = _extract_row_cells(row_matches[0].group(0))
            second_row_cells = _extract_row_cells(row_matches[1].group(0))
            colspan_caption = len(second_row_cells) == 1 and len(first_row_cells) > 1
            if not first_row_cells or (not colspan_caption and len(first_row_cells) != len(second_row_cells)):
                return table_html
            if not any(re.search(r"<img\b", c, flags=re.IGNORECASE) for c in first_row_cells):
                return table_html
            image_caption_pairs = [(first_row_cells, second_row_cells, colspan_caption)]
            side_by_side = True

        elif len(row_matches) >= 4 and len(row_matches) % 2 == 0:
            # Alternating: odd rows = images, even rows = captions
            if not all(
                _html_row_has_image(row_matches[i].group(0)) and not _html_row_has_image(row_matches[i + 1].group(0))
                for i in range(0, len(row_matches), 2)
            ):
                return table_html
            image_caption_pairs = [
                (_extract_row_cells(row_matches[i].group(0)), _extract_row_cells(row_matches[i + 1].group(0)), False)
                for i in range(0, len(row_matches), 2)
            ]

        else:
            return table_html

        figure_blocks: list[str] = []
        docx_caption_row = docx_captions[figure_table_index] if figure_table_index < len(docx_captions) else []
        docx_caption_idx = 0

        for img_cells, cap_cells, colspan_cap in image_caption_pairs:
            for idx, image_cell in enumerate(img_cells):
                if re.search(r"<img\b", image_cell, flags=re.IGNORECASE) is None:
                    continue
                caption_override = docx_caption_row[docx_caption_idx] if docx_caption_idx < len(docx_caption_row) else None
                caption_cell = cap_cells[0] if (colspan_cap or len(cap_cells) == 1) else cap_cells[idx] if idx < len(cap_cells) else cap_cells[0]
                figure_counter += 1
                figure_blocks.append(
                    _make_figure_block(
                        image_cell,
                        caption_cell,
                        caption_override=caption_override,
                        citation_author=citation_author,
                        citation_year=citation_year,
                        citation_title=citation_title,
                        figure_index=figure_counter,
                        folder_name=folder_name,
                    )
                )
                docx_caption_idx += 1

        if not figure_blocks:
            return table_html
        figure_table_index += 1
        if len(figure_blocks) == 1:
            return figure_blocks[0]
        if side_by_side:
            return f'<div class="figure-grid figure-cols-{len(figure_blocks)}">\n' + "\n".join(figure_blocks) + "\n</div>"
        return "\n".join(figure_blocks)

    return table_pattern.sub(replace_table, html)


def _remove_metadata_paragraphs(html: str) -> str:
    """Strip visible metadata paragraphs that should not appear in body HTML.

    The document template stores operational metadata inside the Word document.
    Once the values have been scraped, those paragraphs should not remain in the
    visible report body.
    """
    patterns = [
        r"<p[^>]*>\s*(?:<em>\s*)?Title\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?Author\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?New or Updated\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?Species\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?RW Studies Used\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?Date\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*(?:<em>\s*)?Published\s*:\s*.*?(?:\s*</em>)?\s*</p>\s*",
        r"<p[^>]*>\s*-{10,}\s*</p>\s*",
    ]
    for pattern in patterns:
        html = re.sub(pattern, "", html, flags=re.IGNORECASE | re.DOTALL)
    return html


def _relativize_paths(html: str, report_dir: Path) -> str:
    """Convert absolute filesystem links in generated HTML into relative links.

    Pandoc may emit absolute local paths for images or linked assets. Relative
    paths are required for portable deployment inside the archive folder and the
    site repo.
    """
    def replace_path(match: re.Match[str]) -> str:
        attr = match.group(1)
        raw_value = match.group(2)
        if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", raw_value) or raw_value.startswith("data:"):
            return match.group(0)

        absolute_path: Path | None = None
        if raw_value.lower().startswith("file:///"):
            parsed = urllib.parse.urlparse(raw_value)
            decoded = urllib.parse.unquote(parsed.path)
            if re.match(r"^/[a-zA-Z]:", decoded):
                decoded = decoded[1:]
            absolute_path = Path(decoded)
        elif re.match(r"^[a-zA-Z]:[\\/]", raw_value) or raw_value.startswith("\\\\"):
            absolute_path = Path(raw_value)

        if absolute_path is None:
            return match.group(0)
        try:
            relative = os.path.relpath(absolute_path.resolve(), report_dir.resolve()).replace("\\", "/")
            return f'{attr}="{relative}"'
        except Exception:
            return match.group(0)

    return re.sub(r'(src|href)="([^"]+)"', replace_path, html)


def _build_header_block(header_rel_prefix: str) -> str:
    """Build the canonical header/logo HTML block for a report page."""
    header_logo = f"{header_rel_prefix}/fpc_logo.png".replace("\\", "/")
    return (
        '<div class="logo-header" data-fpc-logo="true">'
        f'<a class="fpc-logo-link" href="https://www.forestproductivity.org/" target="_blank" rel="noopener noreferrer">'
        f'<img class="fpc-logo" src="{header_logo}" alt="Forest Productivity Cooperative logo">'
        "</a>"
        "</div>"
    )


def _remove_existing_logo_instances(html: str) -> str:
    """Remove older or duplicate logo insertions from the HTML.

    This keeps the system idempotent. If the patcher is run more than once, the
    page should still end up with exactly one official logo block.
    """
    patterns = [
        r"<div[^>]*class=\"[^\"]*logo-header[^\"]*\"[^>]*>.*?</div>\s*",
        r"<p[^>]*>\s*(?:<a[^>]*>\s*)?<img[^>]*fpc_logo\.png[^>]*>(?:\s*</a>)?\s*</p>\s*",
        r"<a[^>]*>\s*<img[^>]*fpc_logo\.png[^>]*>\s*</a>\s*",
        r"<img[^>]*fpc_logo\.png[^>]*>\s*",
    ]
    for pattern in patterns:
        html = re.sub(pattern, "", html, flags=re.IGNORECASE | re.DOTALL)
    return html


def _insert_header_once(html: str, header_rel_prefix: str) -> str:
    """Insert the canonical logo/header block into the HTML exactly once."""
    if 'data-fpc-logo="true"' in html:
        return html
    html = _remove_existing_logo_instances(html)
    header_block = _build_header_block(header_rel_prefix)
    title_match = re.search(r"<h1[^>]*class=\"title\"[^>]*>.*?</h1>", html, flags=re.IGNORECASE | re.DOTALL)
    if title_match:
        return html[:title_match.start()] + header_block + "\n" + html[title_match.start():]
    body_match = re.search(r"<body[^>]*>", html, flags=re.IGNORECASE)
    if body_match:
        return html[:body_match.end()] + "\n" + header_block + html[body_match.end():]
    return header_block + "\n" + html


def _inject_interactivity(html: str, report_web_root: str) -> str:
    """Inject client-side UI features into the generated HTML.

    Features added:
    - figure image lightbox
    - figure PNG download
    - dark mode toggle
    - print button wiring
    - back-to-top button
    - footer link back to the research summaries page

    Everything is embedded inline so the final report remains a static, portable
    artifact with no JavaScript build step.
    """
    if 'id="pub-image-lightbox"' in html:
        return html
    payload = f"""
<div class="pub-footer-actions no-print">
  <button id="scrollTopBtn" class="pub-footer-button" type="button">Scroll to Top</button>
  <a class="pub-footer-button" href="{HOME_URL}">Back to Research Summaries</a>
</div>
<div id="pub-image-lightbox" class="pub-lightbox" role="dialog" aria-modal="true" aria-hidden="true">
  <div class="pub-lightbox-backdrop" data-lightbox-close="true"></div>
  <div class="pub-lightbox-panel">
    <button type="button" class="pub-lightbox-close" aria-label="Close image">X</button>
    <img class="pub-lightbox-img" alt="">
    <p class="pub-lightbox-caption"></p>
  </div>
</div>
<script>
(() => {{
  const lightbox = document.getElementById("pub-image-lightbox");
  if (!lightbox) return;
  const lightboxImg = lightbox.querySelector(".pub-lightbox-img");
  const lightboxCaption = lightbox.querySelector(".pub-lightbox-caption");
  const closeButton = lightbox.querySelector(".pub-lightbox-close");
  const printPdfBtn = document.getElementById("printPdfBtn");
  const scrollTopBtn = document.getElementById("scrollTopBtn");
  const themeToggle = document.getElementById("themeToggle");
  const themeStorageKey = "pubTheme";
  const applyTheme = (mode) => {{
    const isDark = mode === "dark";
    document.documentElement.classList.toggle("dark-mode", isDark);
    if (themeToggle) themeToggle.checked = isDark;
  }};
  try {{
    applyTheme(localStorage.getItem(themeStorageKey) === "dark" ? "dark" : "light");
  }} catch (_) {{
    applyTheme("light");
  }}
  if (themeToggle) {{
    themeToggle.addEventListener("change", () => {{
      const mode = themeToggle.checked ? "dark" : "light";
      applyTheme(mode);
      try {{ localStorage.setItem(themeStorageKey, mode); }} catch (_) {{}}
    }});
  }}
  if (printPdfBtn) printPdfBtn.addEventListener("click", () => window.print());
  if (scrollTopBtn) {{
    scrollTopBtn.addEventListener("click", () => window.scrollTo({{ top: 0, behavior: "smooth" }}));
  }}
  const closeLightbox = () => {{
    lightbox.classList.remove("is-open");
    lightbox.setAttribute("aria-hidden", "true");
    lightboxImg.removeAttribute("src");
    lightboxCaption.textContent = "";
  }};
  const openLightbox = (imgEl, captionText) => {{
    lightboxImg.src = imgEl.getAttribute("src") || "";
    lightboxImg.alt = imgEl.getAttribute("alt") || captionText || "Figure image";
    lightboxCaption.textContent = captionText;
    lightbox.classList.add("is-open");
    lightbox.setAttribute("aria-hidden", "false");
    closeButton.focus();
  }};
  lightbox.addEventListener("click", (event) => {{
    if (event.target === lightbox || event.target.closest("[data-lightbox-close='true']")) closeLightbox();
  }});
  closeButton.addEventListener("click", closeLightbox);
  document.addEventListener("keydown", (event) => {{
    if (event.key === "Escape" && lightbox.classList.contains("is-open")) closeLightbox();
  }});
  const wrapCanvasText = (ctx, text, maxWidth) => {{
    const words = (text || "").split(/\\s+/).filter(Boolean);
    if (!words.length) return [""];
    const lines = [];
    let current = words[0];
    for (let i = 1; i < words.length; i += 1) {{
      const next = current + " " + words[i];
      if (ctx.measureText(next).width <= maxWidth) current = next;
      else {{ lines.push(current); current = words[i]; }}
    }}
    lines.push(current);
    return lines;
  }};
  const withPeriod = (value) => {{
    const cleaned = (value || "").trim();
    if (!cleaned) return "";
    return /[.!?]$/.test(cleaned) ? cleaned : `${{cleaned}}.`;
  }};
  const downloadFigureImage = (imgEl, captionEl) => {{
    const src = imgEl?.getAttribute("src");
    if (!src) return;
    const captionText = (captionEl?.textContent || imgEl.dataset.caption || "").trim();
    const citationAuthor = (imgEl.dataset.author || "Unknown Author").trim();
    const citationYear = (imgEl.dataset.year || "").trim();
    const citationTitle = (imgEl.dataset.title || "").trim();
    const program = (imgEl.dataset.program || "").trim();
    const citation = [withPeriod(citationAuthor), withPeriod(citationYear), withPeriod(citationTitle), program].filter(Boolean).join(" ");
    const imageLoader = new Image();
    imageLoader.decoding = "async";
    imageLoader.onload = () => {{
      const naturalWidth = imageLoader.naturalWidth || imgEl.naturalWidth || 1200;
      const naturalHeight = imageLoader.naturalHeight || imgEl.naturalHeight || 900;
      const exportWidth = Math.min(naturalWidth, 1600);
      const imageScale = exportWidth / naturalWidth;
      const exportImageHeight = Math.max(1, Math.round(naturalHeight * imageScale));
      const padding = Math.max(16, Math.round(exportWidth * 0.03));
      const captionFontSize = Math.max(16, Math.round(exportWidth * 0.022));
      const citationFontSize = Math.max(14, Math.round(exportWidth * 0.018));
      const lineGap = Math.max(6, Math.round(captionFontSize * 0.35));
      const usableTextWidth = exportWidth - padding * 2;
      const measureCanvas = document.createElement("canvas");
      const measureCtx = measureCanvas.getContext("2d");
      if (!measureCtx) return;
      measureCtx.font = `italic ${{captionFontSize}}px Georgia, "Times New Roman", Times, serif`;
      const captionLines = wrapCanvasText(measureCtx, captionText, usableTextWidth);
      measureCtx.font = `${{citationFontSize}}px Georgia, "Times New Roman", Times, serif`;
      const citationLines = wrapCanvasText(measureCtx, citation, usableTextWidth);
      const captionLineHeight = Math.round(captionFontSize * 1.35);
      const citationLineHeight = Math.round(citationFontSize * 1.35);
      const footerHeight = padding + captionLines.length * captionLineHeight + lineGap + citationLines.length * citationLineHeight + padding;
      const canvas = document.createElement("canvas");
      canvas.width = exportWidth;
      canvas.height = exportImageHeight + footerHeight;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(imageLoader, 0, 0, exportWidth, exportImageHeight);
      let y = exportImageHeight + padding + captionLineHeight;
      ctx.fillStyle = "#000000";
      ctx.font = `italic ${{captionFontSize}}px Georgia, "Times New Roman", Times, serif`;
      for (const line of captionLines) {{ ctx.fillText(line, padding, y); y += captionLineHeight; }}
      y += lineGap;
      ctx.font = `${{citationFontSize}}px Georgia, "Times New Roman", Times, serif`;
      for (const line of citationLines) {{ ctx.fillText(line, padding, y); y += citationLineHeight; }}
      canvas.toBlob((blob) => {{
        if (!blob) return;
        const objectUrl = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        const fileBase = (imgEl.dataset.downloadName || "figure-download").replace(/[^a-zA-Z0-9_-]/g, "-");
        anchor.href = objectUrl;
        anchor.download = `${{fileBase}}.png`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        setTimeout(() => URL.revokeObjectURL(objectUrl), 2000);
      }}, "image/png");
    }};
    imageLoader.src = src;
  }};
  document.querySelectorAll("figure.figure-block").forEach((figureEl) => {{
    const imgEl = figureEl.querySelector("img.pub-figure-img");
    const captionEl = figureEl.querySelector(".pub-figure-caption");
    const downloadButton = figureEl.querySelector(".pub-figure-download");
    if (!imgEl) return;
    const captionText = (captionEl?.textContent || imgEl.dataset.caption || "").trim();
    imgEl.setAttribute("tabindex", "0");
    imgEl.setAttribute("role", "button");
    imgEl.setAttribute("aria-label", "Open image in lightbox");
    imgEl.addEventListener("click", () => openLightbox(imgEl, captionText));
    imgEl.addEventListener("keydown", (event) => {{
      if (event.key === "Enter" || event.key === " ") {{
        event.preventDefault();
        openLightbox(imgEl, captionText);
      }}
    }});
    if (downloadButton) downloadButton.addEventListener("click", () => downloadFigureImage(imgEl, captionEl));
  }});
  window.addEventListener("beforeprint", () => {{
    Array.from(document.body.children).forEach(el => {{
      if (window.getComputedStyle(el).position === "fixed") {{
        el.dataset.msPrintHide = "1";
        el.style.setProperty("display", "none", "important");
      }}
    }});
  }});
  window.addEventListener("afterprint", () => {{
    document.querySelectorAll("[data-ms-print-hide]").forEach(el => {{
      el.style.removeProperty("display");
      delete el.dataset.msPrintHide;
    }});
  }});
}})();
</script>
""".strip()
    body_close = re.search(r"</body>", html, flags=re.IGNORECASE)
    return html[: body_close.start()] + "\n" + payload + "\n" + html[body_close.start() :] if body_close else html + "\n" + payload


def _style_caption_labels(html: str) -> str:
    """Bold the 'Figure X:' / 'Table X:' label in all captions, leaving the description italic-only.

    Applied to two caption locations:
    - <figcaption> elements (already italic via CSS; <strong> makes the label bold+italic)
    - <th> cells whose content is a caption row (overrides default th bold, adds italic)
    """
    _LABEL_RE = re.compile(r"((?:Figure|Table)\s+[\w.]+\s*:)", flags=re.IGNORECASE)

    def _wrap_label(content: str) -> str:
        return _LABEL_RE.sub(r"<strong>\1</strong>", content, count=1)

    def _fix_figcaption(m: re.Match) -> str:
        return f'<figcaption class="pub-figure-caption">{_wrap_label(m.group(1))}</figcaption>'

    html = re.sub(
        r"<figcaption[^>]*>(.*?)</figcaption>",
        _fix_figcaption,
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )

    def _fix_th_caption(m: re.Match) -> str:
        attrs, content = m.group(1), m.group(2)
        if not _LABEL_RE.search(content):
            return m.group(0)
        if re.search(r'class\s*=', attrs, flags=re.IGNORECASE):
            attrs = re.sub(r'(class\s*=\s*")([^"]*)"', r'\1\2 pub-table-caption"', attrs, count=1, flags=re.IGNORECASE)
        else:
            attrs += ' class="pub-table-caption"'
        return f"<th{attrs}>{_wrap_label(content)}</th>"

    html = re.sub(
        r"<th\b([^>]*)>(.*?)</th>",
        _fix_th_caption,
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )

    return html


_MEMBERSPACE_SNIPPET = (
    '<script> var MemberSpace = window.MemberSpace || {"cookieDomain":"forestproductivity.org","subdomain":"fpc"};'
    " (function(d){ var s = d.createElement(\"script\");"
    ' s.src = "https://cdn.memberspace.com/scripts/widgets.js";'
    " var e = d.getElementsByTagName(\"script\")[0];"
    " e.parentNode.insertBefore(s,e); }(document)); </script>"
)


def _insert_memberspace_script(html: str) -> str:
    """Inject the MemberSpace protection script as the first element in <head>."""
    if "cdn.memberspace.com" in html:
        return html
    return html.replace("<head>", f"<head>\n  {_MEMBERSPACE_SNIPPET}", 1)


def _insert_copyright_footer(html: str) -> str:
    """Insert an FPC copyright notice at the bottom of the report body."""
    if 'class="pub-copyright"' in html:
        return html
    year = datetime.now().year
    notice = f'<p class="pub-copyright no-print">&#169; {year} Forest Productivity Cooperative. All Rights Reserved.</p>'
    body_close = re.search(r"</body>", html, flags=re.IGNORECASE)
    if body_close:
        return html[: body_close.start()] + "\n" + notice + "\n" + html[body_close.start() :]
    return html + "\n" + notice


def _apply_publication_header_layout(html: str, web_root: str) -> str:
    """Rebuild the top title block into the publish-bot header layout.

    The final layout places title/byline content on the left and operator-facing
    actions such as print and theme toggle on the right. This gives the reports
    a consistent visual structure regardless of how Pandoc initially arranged
    the title block.
    """
    if 'class="pub-header"' in html:
        return html
    header_match = re.search(r"(<header[^>]*id=\"title-block-header\"[^>]*>)(.*?)(</header>)", html, flags=re.IGNORECASE | re.DOTALL)
    if not header_match:
        return html
    header_open, header_inner, header_close = header_match.group(1), header_match.group(2), header_match.group(3)
    logo_match = re.search(r"<div[^>]*class=\"[^\"]*logo-header[^\"]*\"[^>]*>.*?</div>\s*", header_inner, flags=re.IGNORECASE | re.DOTALL)
    logo_html = logo_match.group(0).strip() if logo_match else ""
    inner_no_logo = re.sub(
        r"<div[^>]*class=\"[^\"]*logo-header[^\"]*\"[^>]*>.*?</div>\s*",
        "",
        header_inner,
        count=1,
        flags=re.IGNORECASE | re.DOTALL,
    )
    title_match = re.search(r"<h1[^>]*class=\"title\"[^>]*>.*?</h1>\s*", inner_no_logo, flags=re.IGNORECASE | re.DOTALL)
    if not title_match:
        return html
    title_html = title_match.group(0).strip()
    author_match = re.search(r"<p[^>]*class=\"[^\"]*author[^\"]*\"[^>]*>.*?</p>\s*", inner_no_logo, flags=re.IGNORECASE | re.DOTALL)
    date_match = re.search(r"<p[^>]*class=\"[^\"]*pub-date[^\"]*\"[^>]*>.*?</p>\s*", inner_no_logo, flags=re.IGNORECASE | re.DOTALL)
    author_html = author_match.group(0).strip() if author_match else ""
    date_html = date_match.group(0).strip() if date_match else ""
    remainder = inner_no_logo
    for snippet in (title_html, author_html, date_html):
        if snippet:
            remainder = remainder.replace(snippet, "", 1)
    remainder = remainder.strip()
    text_parts = [title_html]
    if author_html:
        text_parts.append(author_html)
    if date_html:
        text_parts.append(date_html)
    pub_header = (
        '<div class="pub-header"><div class="pub-header-text">'
        + "\n".join(text_parts)
        + '</div><div class="pub-header-actions no-print">'
        + f'<a class="pub-header-button" href="{MEMBER_DASHBOARD_URL}" target="_blank" rel="noopener">{_ICON_DASHBOARD} Member Dashboard</a>'
        + f'<a class="pub-header-button" href="{HOME_URL}">{_ICON_REPORT} All Research Summaries</a>'
        + f'<a class="pub-header-button" href="{AUTHOR_CONTACTS_URL}" target="_blank" rel="noopener">{_ICON_MAIL} Author Contacts</a>'
        + f'<button class="pub-header-button" type="button" id="printPdfBtn">{_ICON_PRINTER} Print / Save as PDF</button>'
        + '<label class="pub-theme-toggle no-print" for="themeToggle">'
        + '<input type="checkbox" id="themeToggle" aria-label="Dark Mode">'
        + '<span class="pub-theme-slider" aria-hidden="true"></span>'
        + '<span class="pub-theme-label">Dark Mode</span>'
        + "</label></div></div>"
    )
    new_inner = [logo_html] if logo_html else []
    new_inner.append(pub_header)
    if remainder:
        new_inner.append(remainder)
    rebuilt_header = header_open + "\n" + "\n".join(new_inner) + "\n" + header_close
    return html[: header_match.start()] + rebuilt_header + html[header_match.end() :]


def patch_html_after_conversion(
    html_path: Path,
    input_docx: Path,
    author: str,
    publication_year: str,
    publication_display_date: str,
    folder_name: str,
    header_relative_path: str,
    web_root: str,
) -> None:
    """Apply all HTML post-processing required by the publish workflow.

    Steps performed:
    1. Read the raw Pandoc HTML.
    2. Rebuild figure tables as semantic figure markup.
    3. Remove visible metadata paragraphs from the body.
    4. Rewrite asset links to relative paths.
    5. Insert the official header/logo block.
    6. Inject interactive figure and page controls.
    7. Ensure author and publication year are visible.
    8. Apply the final structured header layout.
    9. Write the patched HTML back to disk.
    """
    html = html_path.read_text(encoding="utf-8")
    metadata = parse_document_metadata(input_docx)
    html = _convert_figure_tables(
        html,
        docx_captions=extract_figure_table_captions(input_docx),
        citation_author=author,
        citation_year=publication_year,
        citation_title=metadata.title,
        folder_name=folder_name,
    )
    html = _remove_metadata_paragraphs(html)
    html = _style_caption_labels(html)
    html = _relativize_paths(html, html_path.parent)
    html = _insert_header_once(html, header_relative_path)
    html = _inject_interactivity(html, web_root)

    if 'class="author"' not in html:
        title_match = re.search(r"(<h1[^>]*class=\"title\"[^>]*>.*?</h1>)", html, flags=re.IGNORECASE | re.DOTALL)
        if title_match:
            html = html[: title_match.end()] + f'\n<p class="author byline">{escape(author)}</p>' + html[title_match.end() :]
    if 'class="pub-date"' not in html:
        author_match = re.search(r"(<p[^>]*class=\"[^\"]*author[^\"]*\"[^>]*>.*?</p>)", html, flags=re.IGNORECASE | re.DOTALL)
        date_html = f'<p class="pub-date byline">{escape(publication_display_date)}</p>'
        if author_match:
            html = html[: author_match.end()] + "\n" + date_html + html[author_match.end() :]
        else:
            title_match = re.search(r"(<h1[^>]*class=\"title\"[^>]*>.*?</h1>)", html, flags=re.IGNORECASE | re.DOTALL)
            if title_match:
                html = html[: title_match.end()] + "\n" + date_html + html[title_match.end() :]

    html = _apply_publication_header_layout(html, web_root)
    html = _insert_copyright_footer(html)
    html = _insert_memberspace_script(html)
    html_path.write_text(html, encoding="utf-8")
