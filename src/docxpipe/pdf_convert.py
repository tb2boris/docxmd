"""PDF → Markdown по страницам, карта страниц, рисунки и CSV текстовых таблиц."""

from __future__ import annotations

import csv
import io
import re
import statistics
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import fitz

from docxpipe.convert import media_dirname
from docxpipe.ocr_engine import OcrUnavailable, recognize_page

PDF_DEPERS_NOTE = (
    "Обезличивается извлечённый текст и CSV таблиц. "
    "Картинки страниц и вынутые рисунки не закрашиваются."
)

TEXT_MIN_CHARS = 40
RASTER_MIN_RATIO = 0.60
OUTLINE_MIN_ITEMS = 25
HEADER_BAND = 0.08
EDGE_BAND = 0.08
HEADING_RATIO = 1.3
OCR_DPI = 300
LOW_CONFIDENCE = 45
LETTER_RATIO = 0.45
FIGURE_MIN_AREA = 800
FULL_PAGE_RATIO = 0.75

_DIGIT_RE = re.compile(r"\d+")
OcrPage = Callable[[bytes], tuple[str, float | None]]


class PdfOpenError(Exception):
    """Файл нельзя открыть: пароль или повреждённые данные."""


@dataclass
class TextLine:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class PdfBlock:
    kind: str
    text: str


@dataclass
class PdfTable:
    rows: list[list[str]]
    merges: list[str]
    recovered: bool
    csv_name: str | None = None


@dataclass
class PdfImage:
    file: str
    role: str
    page: int
    data: bytes


@dataclass
class PdfPage:
    number: int
    kind: str
    chars: int = 0
    layout: str | None = None
    blocks: list[PdfBlock] = field(default_factory=list)
    tables: list[PdfTable] = field(default_factory=list)
    form_fields: list[dict] = field(default_factory=list)
    ocr_confidence: float | None = None
    ocr_weak: bool = False
    error: str | None = None
    image_files: list[str] = field(default_factory=list)


@dataclass
class PdfDocument:
    source_name: str
    headers_footers: list[str] = field(default_factory=list)
    pages: list[PdfPage] = field(default_factory=list)
    images: list[PdfImage] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def load_pdf(
    path: Path,
    *,
    force_ocr: bool = False,
    auto: bool = True,
    save_page_images: bool = True,
    ocr_page: OcrPage | None = None,
) -> PdfDocument:
    """Читает PDF в структуру страниц. Исходный файл не изменяется."""
    document = PdfDocument(source_name=path.name)
    opened = _open(path)
    try:
        raw_pages = [_extract_lines(page) for page in opened]
        repeated, examples = _repeated_headers(raw_pages)
        document.headers_footers = [examples[key] for key in sorted(examples) if key in repeated]
        figure_no = 1
        recognizer = ocr_page or _default_ocr
        warned = False
        for index, page in enumerate(opened):
            number = index + 1
            try:
                model, blobs, figure_no, warning = _build_page(
                    page,
                    raw_pages[index],
                    repeated,
                    number=number,
                    force_ocr=force_ocr,
                    auto=auto,
                    save_page_images=save_page_images,
                    ocr_page=recognizer,
                    figure_no=figure_no,
                )
                if warning and not warned:
                    document.warnings.append(warning)
                    warned = True
            except Exception as exc:
                model = PdfPage(number=number, kind="error", error=str(exc))
                blobs = []
            document.pages.append(model)
            document.images.extend(blobs)
    finally:
        opened.close()
    return document


def render_pdf(document: PdfDocument, stem: str) -> tuple[str, dict[str, str], dict]:
    """Собирает Markdown, CSV и карту страниц. Ссылки считают папку результата известной."""
    folder = media_dirname(stem)
    csv_files: dict[str, str] = {}
    parts: list[str] = []
    page_payloads: list[dict] = []
    for page in document.pages:
        parts.append(f"## Страница {page.number}")
        parts.append("")
        if page.kind == "error":
            parts.append("Страница не разобрана.")
            parts.append("")
        elif page.kind == "empty":
            parts.append("Страница без текста.")
            parts.append("")
        elif page.ocr_weak:
            parts.append("Распознавание ненадёжно. Текст страницы не приведён.")
            parts.append("")
        elif page.kind in {"ocr", "outline"} and not page.blocks and _recognition_disabled(page):
            parts.append("Страница без текстового слоя. Распознавание выключено.")
            parts.append("")
        else:
            for block in page.blocks:
                if block.kind == "heading":
                    parts.append(f"### {block.text}")
                else:
                    parts.append(block.text)
                parts.append("")
        for table in page.tables:
            if table.recovered and table.csv_name and table.rows:
                csv_files[table.csv_name] = _csv_text(table.rows)
                parts.append(
                    f"Таблица на странице {page.number}: [{table.csv_name}]({stem}/{table.csv_name})"
                )
            else:
                parts.append(f"Таблица на странице {page.number}. Сетка не восстановлена.")
            parts.append("")
        if page.form_fields:
            parts.append("Поля формы:")
            parts.append("")
            for item in page.form_fields:
                parts.append(f"- {item['name']}: {item['value']}")
            parts.append("")
        for name in page.image_files:
            role = _image_role(document, name)
            alt = f"страница {page.number}" if role == "page" else "рисунок"
            parts.append(f"![{alt}]({folder}/{name})")
            parts.append("")
        page_payloads.append(_page_payload(page))
    markdown = "\n".join(parts).rstrip() + "\n"
    payload = {
        "source": document.source_name,
        "page_count": len(document.pages),
        "headers_footers": list(document.headers_footers),
        "pages": page_payloads,
    }
    return markdown, csv_files, payload


def anonymize_pdf_document(processor, document: PdfDocument) -> list[dict]:
    """Обезличивает текст страниц и ячейки таблиц. Имена файлов и растры не меняет."""
    from docxpipe.anonymize import anonymize_fragments

    slots: list[tuple[object, object, str]] = []
    for index, text in enumerate(document.headers_footers):
        slots.append((document.headers_footers, index, text))
    for page in document.pages:
        for block in page.blocks:
            slots.append((block, "text", block.text))
        for table in page.tables:
            for row in table.rows:
                for col, value in enumerate(row):
                    slots.append((row, col, value))
        for item in page.form_fields:
            slots.append((item, "value", str(item.get("value", ""))))
    if not slots:
        return []
    updated, records = anonymize_fragments(processor, [text for _holder, _key, text in slots])
    for (holder, key, _text), value in zip(slots, updated):
        if isinstance(holder, PdfBlock):
            holder.text = value
        else:
            holder[key] = value
    return records


def _recognition_disabled(page: PdfPage) -> bool:
    return page.kind in {"ocr", "outline"} and page.ocr_confidence is None and not page.ocr_weak and not page.error


def _image_role(document: PdfDocument, name: str) -> str:
    for image in document.images:
        if image.file == name:
            return image.role
    return "figure"


def _page_payload(page: PdfPage) -> dict:
    headings = [
        {"text": block.text, "heading_guess": True}
        for block in page.blocks
        if block.kind == "heading"
    ]
    tables = []
    for table in page.tables:
        tables.append({
            "file": table.csv_name if table.recovered else None,
            "merges": list(table.merges),
            "recovered": table.recovered,
        })
    return {
        "number": page.number,
        "kind": page.kind,
        "chars": page.chars,
        "layout": page.layout,
        "headings": headings,
        "tables": tables,
        "form_fields": list(page.form_fields),
        "images": list(page.image_files),
        "ocr_confidence": page.ocr_confidence,
        "error": page.error,
    }


def _csv_text(rows: list[list[str]]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerows(rows)
    return buffer.getvalue()


def _open(path: Path) -> fitz.Document:
    try:
        document = fitz.open(path)
    except Exception as exc:
        raise PdfOpenError(f"Не удалось открыть PDF: {exc}") from exc
    if document.needs_pass:
        document.close()
        raise PdfOpenError("PDF закрыт паролем. Файл не обработан.")
    return document


def _default_ocr(blob: bytes) -> tuple[str, float | None]:
    text, confidence = recognize_page(blob)
    return text, confidence


def _build_page(
    page: fitz.Page,
    raw: tuple[float, float, list[TextLine]],
    repeated: set[str],
    *,
    number: int,
    force_ocr: bool,
    auto: bool,
    save_page_images: bool,
    ocr_page: OcrPage,
    figure_no: int,
) -> tuple[PdfPage, list[PdfImage], int, str | None]:
    width, height, lines = raw
    chars = _char_count(lines)
    kind = _classify(page, chars, force_ocr=force_ocr)
    model = PdfPage(number=number, kind=kind, chars=chars)
    blobs: list[PdfImage] = []
    warning: str | None = None
    if kind == "text":
        kept = _without_headers(lines, height, repeated)
        table_rects, tables = _tables_from_text(page, number)
        model.tables = tables
        body_lines = [line for line in kept if not _inside(line, table_rects)]
        edge, body = _split_edge(body_lines, height)
        model.layout = _layout(body, width)
        ordered = _order_lines(body, width, model.layout)
        model.blocks = _blocks(ordered, edge)
        model.form_fields = _form_fields(page)
        figure_no, figures = _figures(page, number, figure_no)
        blobs.extend(figures)
        model.image_files.extend(image.file for image in figures)
        return model, blobs, figure_no, None

    if kind == "empty":
        model.blocks = []
        return model, blobs, figure_no, None

    run_ocr = force_ocr or auto
    if not run_ocr:
        return model, blobs, figure_no, None

    try:
        png = _render_png(page)
        png = deskew_png(png)
        text, confidence = ocr_page(png)
    except OcrUnavailable as exc:
        model.ocr_weak = True
        warning = str(exc)
        text, confidence, png = "", None, _render_png(page)
    if save_page_images and png:
        file_name = f"page_{number:03d}.png"
        blobs.append(PdfImage(file=file_name, role="page", page=number, data=png))
        model.image_files.append(file_name)
    model.ocr_confidence = confidence
    if warning or _unreliable(text, confidence):
        model.ocr_weak = True
        model.blocks = []
    else:
        model.blocks = _ocr_blocks(text or "")
    return model, blobs, figure_no, warning


def _classify(page: fitz.Page, chars: int, *, force_ocr: bool) -> str:
    if force_ocr:
        return "ocr"
    if chars >= TEXT_MIN_CHARS:
        return "text"
    if _raster_ratio(page) >= RASTER_MIN_RATIO:
        return "ocr"
    if _vector_items(page) >= OUTLINE_MIN_ITEMS:
        return "outline"
    return "empty"


def _char_count(lines: list[TextLine]) -> int:
    return sum(1 for line in lines for char in line.text if not char.isspace() and char != "\u00ad")


def _raster_ratio(page: fitz.Page) -> float:
    area = page.rect.width * page.rect.height
    if area <= 0:
        return 0.0
    best = 0.0
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 1:
            continue
        rect = fitz.Rect(block["bbox"])
        best = max(best, rect.get_area() / area)
    return best


def _vector_items(page: fitz.Page) -> int:
    total = 0
    for drawing in page.get_drawings():
        items = drawing.get("items") or []
        total += len(items) or 1
    return total


def _extract_lines(page: fitz.Page) -> tuple[float, float, list[TextLine]]:
    raw_lines: list[TextLine] = []
    payload = page.get_text("rawdict")
    for block in payload.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            chars = [char for span in line.get("spans", []) for char in span.get("chars", [])]
            if not chars:
                continue
            size = max((float(span.get("size") or 0) for span in line.get("spans", [])), default=12)
            text = _join_chars(chars, size)
            if not text:
                continue
            bbox = line.get("bbox") or chars[0]["bbox"]
            raw_lines.append(TextLine(text, bbox[0], bbox[1], bbox[2], bbox[3], size))
    return page.rect.width, page.rect.height, _merge_baselines(raw_lines)


def _join_chars(chars: list[dict], size: float) -> str:
    parts: list[str] = []
    previous_x1: float | None = None
    gap_limit = max(size * 0.22, 0.8)
    for char in chars:
        symbol = str(char.get("c", "")).replace("\xa0", " ")
        x0, _y0, x1, _y1 = char["bbox"]
        if previous_x1 is not None and symbol not in {" ", "\u00ad"} and x0 - previous_x1 > gap_limit:
            if not parts or not parts[-1].endswith(" "):
                parts.append(" ")
        if symbol == " ":
            if parts and not parts[-1].endswith(" "):
                parts.append(" ")
        elif symbol:
            parts.append(symbol)
        previous_x1 = x1
    return "".join(parts).strip()


def _merge_baselines(lines: list[TextLine]) -> list[TextLine]:
    """Склеивает куски одной строки. Широкий зазор — это колонка, а не пробел."""
    ordered = sorted(lines, key=lambda item: (round(item.y0, 1), item.x0))
    rows: list[list[TextLine]] = []
    for line in ordered:
        if rows and abs(line.y0 - rows[-1][0].y0) <= 2:
            rows[-1].append(line)
        else:
            rows.append([line])
    merged: list[TextLine] = []
    for row in rows:
        row.sort(key=lambda item: item.x0)
        phrase = [row[0]]
        for nxt in row[1:]:
            gap = nxt.x0 - phrase[-1].x1
            if gap <= max(phrase[-1].size * 2.5, 18):
                phrase.append(nxt)
            else:
                merged.append(_phrase(phrase))
                phrase = [nxt]
        merged.append(_phrase(phrase))
    return merged


def _phrase(parts: list[TextLine]) -> TextLine:
    text = parts[0].text
    x1 = parts[0].x1
    size = parts[0].size
    for nxt in parts[1:]:
        gap = nxt.x0 - x1
        if text.endswith(("\u00ad", "-")) and nxt.text[:1].islower():
            text = _join_hyphen(text, nxt.text)
        elif gap > max(size * 0.15, 1.0):
            text = text.rstrip() + " " + nxt.text.lstrip()
        else:
            text += nxt.text
        x1 = max(x1, nxt.x1)
        size = max(size, nxt.size)
    return TextLine(
        text,
        min(item.x0 for item in parts),
        min(item.y0 for item in parts),
        max(item.x1 for item in parts),
        max(item.y1 for item in parts),
        size,
    )


def _join_hyphen(left: str, right: str) -> str:
    right = right.lstrip()
    if left.endswith(("\u00ad", "-")) and right[:1].islower():
        return left[:-1] + right
    return left + " " + right


def header_key(text: str) -> str:
    cleaned = text.replace("\u00ad", "").replace("\xa0", " ")
    cleaned = _DIGIT_RE.sub("#", cleaned.strip().lower())
    return re.sub(r"\s+", " ", cleaned)


def _in_band(line: TextLine, height: float, band: float) -> bool:
    if height <= 0:
        return False
    if line.center_y <= height * band:
        return True
    return line.center_y >= height * (1 - band)


def _repeated_headers(pages: list[tuple[float, float, list[TextLine]]]) -> tuple[set[str], dict[str, str]]:
    page_count = len(pages)
    seen: dict[str, set[int]] = {}
    examples: dict[str, str] = {}
    for index, (_width, height, lines) in enumerate(pages):
        for line in lines:
            if not _in_band(line, height, HEADER_BAND):
                continue
            key = header_key(line.text)
            if len(key) < 2 and key != "#":
                continue
            seen.setdefault(key, set()).add(index)
            examples.setdefault(key, line.text)
    if page_count < 2:
        return set(), {}
    repeated = {
        key
        for key, pageset in seen.items()
        if len(pageset) >= 2 and len(pageset) > page_count / 2
    }
    return repeated, examples


def _without_headers(lines: list[TextLine], height: float, repeated: set[str]) -> list[TextLine]:
    kept: list[TextLine] = []
    for line in lines:
        if _in_band(line, height, HEADER_BAND) and header_key(line.text) in repeated:
            continue
        kept.append(line)
    return kept


def _split_edge(lines: list[TextLine], height: float) -> tuple[list[TextLine], list[TextLine]]:
    edge: list[TextLine] = []
    body: list[TextLine] = []
    for line in lines:
        if _in_band(line, height, EDGE_BAND):
            edge.append(line)
        else:
            body.append(line)
    return edge, body


def _layout(lines: list[TextLine], width: float) -> str:
    if width <= 0 or len(lines) < 6:
        return "single"
    narrow = [line for line in lines if line.width < width * 0.62]
    left = [line for line in narrow if line.center_x < width * 0.42]
    right = [line for line in narrow if line.center_x > width * 0.58]
    middle = [line for line in narrow if width * 0.42 <= line.center_x <= width * 0.58]
    if len(left) >= 3 and len(right) >= 3 and len(middle) >= 3:
        return "complex"
    if len(left) >= 3 and len(right) >= 3:
        gutter = min(line.x0 for line in right) - max(line.x1 for line in left)
        if gutter >= width * 0.03:
            return "columns"
    return "single"


def _order_lines(lines: list[TextLine], width: float, layout: str) -> list[TextLine]:
    if layout != "columns":
        return sorted(lines, key=lambda line: (line.y0, line.x0))
    full = [line for line in lines if line.width >= width * 0.62]
    columns = [line for line in lines if line not in full]
    left = sorted((line for line in columns if line.center_x < width * 0.5), key=lambda line: line.y0)
    right = sorted((line for line in columns if line.center_x >= width * 0.5), key=lambda line: line.y0)
    column_top = min((line.y0 for line in columns), default=0)
    above = sorted((line for line in full if line.y0 < column_top), key=lambda line: line.y0)
    below = sorted((line for line in full if line.y0 >= column_top), key=lambda line: line.y0)
    return above + left + right + below


def _blocks(lines: list[TextLine], edge: list[TextLine]) -> list[PdfBlock]:
    median = statistics.median([line.size for line in lines]) if lines else 12
    blocks: list[PdfBlock] = []
    buffer: list[TextLine] = []
    buffer_heading = False

    def flush() -> None:
        nonlocal buffer, buffer_heading
        if not buffer:
            return
        text = buffer[0].text
        for line in buffer[1:]:
            text = _join_hyphen(text, line.text)
        kind = "heading" if buffer_heading else "paragraph"
        if text.strip():
            blocks.append(PdfBlock(kind, text.strip()))
        buffer = []
        buffer_heading = False

    for line in lines:
        heading = _is_heading(line, median)
        if not buffer:
            buffer = [line]
            buffer_heading = heading
            continue
        gap = line.y0 - buffer[-1].y1
        same = gap >= -1 and gap <= max(buffer[-1].size * 0.9, 8) and not heading and not buffer_heading
        if same:
            buffer.append(line)
        else:
            flush()
            buffer = [line]
            buffer_heading = heading
    flush()
    for line in edge:
        if line.text.strip():
            blocks.append(PdfBlock("edge", line.text.strip()))
    return blocks


def _is_heading(line: TextLine, median: float) -> bool:
    if median <= 0 or len(line.text) > 140 or line.text.endswith(","):
        return False
    return line.size >= median * HEADING_RATIO


def _inside(line: TextLine, rects: list[fitz.Rect]) -> bool:
    for rect in rects:
        if rect.x0 - 1 <= line.center_x <= rect.x1 + 1 and rect.y0 - 1 <= line.center_y <= rect.y1 + 1:
            return True
    return False


def _tables_from_text(page: fitz.Page, number: int) -> tuple[list[fitz.Rect], list[PdfTable]]:
    found = _find_tables(page)
    rects: list[fitz.Rect] = []
    tables: list[PdfTable] = []
    recovered = 0
    for table in found:
        bbox = fitz.Rect(table.bbox)
        if not _table_has_grid(page, bbox):
            continue
        grid = _as_grid(table.extract())
        if not _meaningful(grid):
            tables.append(PdfTable(rows=[], merges=[], recovered=False))
            rects.append(bbox)
            continue
        rows, merges = fill_merges(grid)
        recovered += 1
        tables.append(PdfTable(
            rows=rows,
            merges=merges,
            recovered=True,
            csv_name=f"p{number}_t{recovered}.csv",
        ))
        rects.append(bbox)
    if not tables and _page_has_unparsed_grid(page):
        tables.append(PdfTable(rows=[], merges=[], recovered=False))
    return rects, tables


def _find_tables(page: fitz.Page) -> list:
    try:
        finder = page.find_tables()
    except Exception:
        return []
    return list(getattr(finder, "tables", []) or [])


def _as_grid(raw) -> list[list[str | None]]:
    grid: list[list[str | None]] = []
    for row in raw or []:
        cells: list[str | None] = []
        for cell in row:
            if cell is None:
                cells.append(None)
            else:
                cells.append(str(cell).replace("\n", " ").strip())
        grid.append(cells)
    return grid


def _meaningful(grid: list[list[str | None]]) -> bool:
    texts = [cell for row in grid for cell in row if cell]
    return len(grid) >= 2 and len(texts) >= 2


def fill_merges(grid: list[list[str | None]]) -> tuple[list[list[str]], list[str]]:
    """Повторяет значение объединённой ячейки и возвращает диапазоны вроде A1:B1."""
    rows = [list(row) for row in grid]
    width = max((len(row) for row in rows), default=0)
    for row in rows:
        row.extend([None] * (width - len(row)))
    spans = [[cell is None for cell in row] for row in rows]
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            if cell is not None:
                continue
            if j > 0 and row[j - 1] is not None:
                row[j] = row[j - 1]
            elif i > 0 and rows[i - 1][j] is not None:
                row[j] = rows[i - 1][j]
            else:
                row[j] = ""
    merges: list[str] = []
    for i, row in enumerate(rows):
        for j in range(width):
            if spans[i][j]:
                continue
            end_j = j
            while end_j + 1 < width and spans[i][end_j + 1] and rows[i][end_j + 1] == row[j]:
                end_j += 1
            end_i = i
            while end_i + 1 < len(rows) and all(
                spans[end_i + 1][col] and rows[end_i + 1][col] == row[j]
                for col in range(j, end_j + 1)
            ):
                end_i += 1
            if end_i > i or end_j > j:
                merges.append(f"{_col_name(j)}{i + 1}:{_col_name(end_j)}{end_i + 1}")
    filled = [[cell if cell is not None else "" for cell in row] for row in rows]
    return filled, merges


def _col_name(index: int) -> str:
    name = ""
    number = index + 1
    while number:
        number, rem = divmod(number - 1, 26)
        name = chr(65 + rem) + name
    return name


def _table_has_grid(page: fitz.Page, bbox: fitz.Rect) -> bool:
    horizontal, vertical = _grid_counts(page, bbox)
    return horizontal >= 2 and vertical >= 1


def _page_has_unparsed_grid(page: fitz.Page) -> bool:
    horizontal, vertical = _grid_counts(page, page.rect)
    return horizontal >= 3 and vertical >= 3


def _grid_counts(page: fitz.Page, limit: fitz.Rect) -> tuple[int, int]:
    horizontal = 0
    vertical = 0
    for drawing in page.get_drawings():
        for item in drawing.get("items") or []:
            kind = item[0]
            if kind == "re":
                rect = fitz.Rect(item[1])
                if rect.intersects(limit):
                    horizontal += 1
                    vertical += 1
            elif kind == "l":
                p1, p2 = item[1], item[2]
                segment = fitz.Rect(min(p1.x, p2.x), min(p1.y, p2.y), max(p1.x, p2.x), max(p1.y, p2.y))
                if not segment.intersects(limit):
                    continue
                if abs(p1.y - p2.y) <= 1.5:
                    horizontal += 1
                elif abs(p1.x - p2.x) <= 1.5:
                    vertical += 1
    return horizontal, vertical


def _form_fields(page: fitz.Page) -> list[dict]:
    fields: list[dict] = []
    try:
        widgets = page.widgets()
    except Exception:
        return fields
    if not widgets:
        return fields
    for widget in widgets:
        value = getattr(widget, "field_value", None)
        if value is None:
            continue
        text = str(value).strip()
        if not text or text == "Off":
            continue
        name = str(getattr(widget, "field_name", "") or "").strip()
        fields.append({"name": name, "value": text.replace("\n", " ")})
    return fields


def _figures(page: fitz.Page, number: int, figure_no: int) -> tuple[int, list[PdfImage]]:
    images: list[PdfImage] = []
    page_area = page.rect.width * page.rect.height
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 1:
            continue
        rect = fitz.Rect(block["bbox"])
        if page_area and rect.get_area() / page_area >= FULL_PAGE_RATIO:
            continue
        if rect.get_area() < FIGURE_MIN_AREA:
            continue
        blob = block.get("image")
        if not blob:
            continue
        ext = str(block.get("ext") or "png").lower()
        if ext == "jpg":
            ext = "jpeg"
        if ext not in {"png", "jpeg", "jp2", "tif", "tiff"}:
            ext = "png"
        file_name = f"img_{figure_no:03d}.{ext if ext != 'jpeg' else 'jpg'}"
        figure_no += 1
        images.append(PdfImage(file=file_name, role="figure", page=number, data=blob))
    return figure_no, images


def _render_png(page: fitz.Page) -> bytes:
    pixmap = page.get_pixmap(dpi=OCR_DPI, alpha=False)
    return pixmap.tobytes("png")


def deskew_png(blob: bytes) -> bytes:
    """Выправляет небольшой перекос. Прямой лист возвращается без поворота."""
    try:
        from PIL import Image
    except ImportError:
        return blob
    image = Image.open(io.BytesIO(blob)).convert("L")
    angle = _deskew_angle(image)
    if abs(angle) < 0.4:
        return blob
    rotated = image.rotate(angle, expand=True, fillcolor=255)
    output = io.BytesIO()
    rotated.convert("RGB").save(output, format="PNG")
    return output.getvalue()


def _deskew_angle(image) -> float:
    base = _projection_score(_resized(image))
    best_angle = 0.0
    best_score = base
    for step in range(-6, 7):
        if step == 0:
            continue
        angle = step * 0.5
        rotated = image.rotate(angle, expand=False, fillcolor=255)
        score = _projection_score(_resized(rotated))
        if score > best_score * 1.05:
            best_score = score
            best_angle = angle
    return best_angle


def _resized(image):
    width = 400
    if image.width <= width:
        return image
    height = max(1, int(image.height * width / image.width))
    return image.resize((width, height))


def _projection_score(image) -> float:
    pixels = image.load()
    width, height = image.size
    counts = []
    for y in range(height):
        dark = 0
        for x in range(0, width, 2):
            if pixels[x, y] < 160:
                dark += 1
        counts.append(dark)
    if not counts:
        return 0.0
    mean = sum(counts) / len(counts)
    return sum((count - mean) ** 2 for count in counts)


def _unreliable(text: str, confidence: float | None) -> bool:
    if confidence is None:
        return True
    stripped = text.strip()
    if not stripped or confidence < LOW_CONFIDENCE:
        return True
    symbols = [char for char in stripped if not char.isspace()]
    letters = [char for char in symbols if char.isalpha()]
    if not symbols:
        return True
    return len(letters) / len(symbols) < LETTER_RATIO


def _ocr_blocks(text: str) -> list[PdfBlock]:
    chunks = re.split(r"\n\s*\n", text.strip())
    return [PdfBlock("paragraph", re.sub(r"[ \t]*\n[ \t]*", " ", chunk).strip()) for chunk in chunks if chunk.strip()]
