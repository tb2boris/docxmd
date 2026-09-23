"""Сборка результата: Markdown, папка image_<имя>, опционально OCR и DP152."""

from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from docxpipe.anonymize import (
    RESTORE_WARNING,
    anonymize_fragments,
    anonymize_markdown,
    anonymize_plain,
    apply_restore,
    begin_pair_log,
    build_report,
    build_restore_document,
    load_processor,
)
from docxpipe.convert import convert_document, image_ocr_text, media_dirname
from docxpipe.ocr_engine import OcrUnavailable, recognize
from docxpipe.pdf_convert import PdfOpenError, anonymize_pdf_document, load_pdf, render_pdf
from docxpipe.xlsx_convert import iter_text_slots, read_workbook, render_workbook, write_slot

SKIP_NOTE = "Колонтитулы, сноски, примечания и правки не извлекаются."


@dataclass
class FileOutcome:
    source: Path
    markdown_path: Path | None = None
    media_dir: Path | None = None
    report_path: Path | None = None
    restore_path: Path | None = None
    saved: int = 0
    skipped: int = 0
    detections: int = 0
    warnings: list[str] = field(default_factory=list)
    error: str | None = None


def unique_stem(stem: str, used: set[str]) -> str:
    if stem not in used:
        used.add(stem)
        return stem
    number = 2
    while f"{stem}_{number}" in used:
        number += 1
    unique = f"{stem}_{number}"
    used.add(unique)
    return unique


def list_docx(path: Path, output_dir: Path) -> list[Path]:
    if path.is_file():
        if path.suffix.lower() != ".docx" or path.name.startswith("~$"):
            return []
        return [path]
    output = output_dir.resolve()
    found: list[Path] = []
    for item in sorted(path.rglob("*.docx")):
        if item.name.startswith("~$"):
            continue
        try:
            item.resolve().relative_to(output)
        except ValueError:
            found.append(item)
    return found


def list_pdf(path: Path, output_dir: Path) -> list[Path]:
    if path.is_file():
        if path.suffix.lower() != ".pdf" or path.name.startswith("~$"):
            return []
        return [path]
    output = output_dir.resolve()
    found: list[Path] = []
    for item in sorted(path.rglob("*.pdf")):
        if item.name.startswith("~$"):
            continue
        try:
            item.resolve().relative_to(output)
        except ValueError:
            found.append(item)
    return found


def list_xlsx(path: Path, output_dir: Path) -> list[Path]:
    if path.is_file():
        if path.suffix.lower() != ".xlsx" or path.name.startswith("~$"):
            return []
        return [path]
    output = output_dir.resolve()
    found: list[Path] = []
    for item in sorted(path.rglob("*.xlsx")):
        if item.name.startswith("~$"):
            continue
        try:
            item.resolve().relative_to(output)
        except ValueError:
            found.append(item)
    return found


def planned_names(sources: list[Path]) -> list[str]:
    used: set[str] = set()
    return [f"{unique_stem(item.stem, used)}.md" for item in sources]


def existing_outputs(sources: list[Path], output_dir: Path) -> list[Path]:
    return [output_dir / name for name in planned_names(sources) if (output_dir / name).exists()]


def convert_tree(
    source: Path,
    output_dir: Path,
    *,
    extract_images: bool,
    ocr: bool,
    depersonalize: bool,
    save_report: bool,
    save_restore_map: bool = False,
    seed: int | None,
    overwrite: bool,
    progress=None,
) -> list[FileOutcome]:
    output_dir.mkdir(parents=True, exist_ok=True)
    sources = list_docx(source, output_dir)
    if not sources:
        return [FileOutcome(source=source, error="Нет файлов .docx")]

    processor = load_processor(seed) if depersonalize else None
    outcomes: list[FileOutcome] = []
    used_stems: set[str] = set()
    ocr_note: str | None = None

    def ocr_fn(blob: bytes, extension: str) -> str:
        nonlocal ocr_note
        try:
            return recognize(blob, extension)
        except OcrUnavailable as exc:
            if ocr_note is None:
                ocr_note = str(exc)
            return ""

    for index, item in enumerate(sources, start=1):
        if progress:
            progress(index, len(sources), item.name)
        outcome = FileOutcome(source=item)
        out_stem = unique_stem(item.stem, used_stems)
        target = output_dir / f"{out_stem}.md"
        if target.exists() and not overwrite:
            outcome.error = f"Файл уже существует: {target.name}"
            outcomes.append(outcome)
            continue
        try:
            result = convert_document(
                item,
                extract_images=extract_images,
                ocr=ocr_fn if extract_images and ocr else None,
            )
            text = result.markdown
            records: list[dict] = []
            pair_log = None
            if processor is not None:
                if save_restore_map:
                    pair_log = begin_pair_log(processor)
                text, records = anonymize_markdown(processor, text)
                if ocr:
                    for image in result.images:
                        raw = image_ocr_text(image) or ""
                        cleaned, extra = anonymize_plain(processor, raw)
                        setattr(image, "ocr_text", cleaned)
                        for row in extra:
                            tagged = dict(row)
                            tagged["where"] = "ocr"
                            tagged["image"] = image.id
                            records.append(tagged)
            folder = media_dirname(out_stem)
            prefix = f"{result.media_dirname}/"
            if prefix != f"{folder}/":
                text = text.replace(prefix, f"{folder}/")
            for image in result.images:
                image.path = f"{folder}/{image.file}"
            target.write_text(text, encoding="utf-8")
            outcome.markdown_path = target
            outcome.saved = result.saved
            outcome.skipped = result.skipped
            outcome.detections = len(records)
            if extract_images and result.images:
                media = output_dir / folder
                if media.exists():
                    shutil.rmtree(media)
                media.mkdir(parents=True)
                for name, blob in result.blobs.items():
                    (media / name).write_bytes(blob)
                payload = {
                    "source": item.name,
                    "media_dir": folder,
                    "images": [
                        image.as_dict(ocr_text=image_ocr_text(image) if ocr else None)
                        for image in result.images
                    ],
                }
                (media / "images.json").write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.media_dir = media
            if processor is not None and save_report:
                report_path = output_dir / f"{out_stem}.report.json"
                report_path.write_text(
                    json.dumps(build_report(item.name, records), ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.report_path = report_path
            if processor is not None and save_restore_map and pair_log is not None:
                restore_path = output_dir / f"{out_stem}.restore.json"
                payload = build_restore_document(item.name, seed, pair_log.pairs)
                restore_path.write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.restore_path = restore_path
                outcome.warnings.append(RESTORE_WARNING)
            if ocr_note and ocr_note not in outcome.warnings:
                outcome.warnings.append(ocr_note)
        except Exception as exc:
            outcome.error = str(exc)
        outcomes.append(outcome)
    return outcomes


def convert_pdf_tree(
    source: Path,
    output_dir: Path,
    *,
    force_ocr: bool = False,
    auto: bool = True,
    save_page_images: bool = True,
    depersonalize: bool = False,
    save_report: bool = False,
    save_restore_map: bool = False,
    seed: int | None = None,
    overwrite: bool = False,
    ocr_page=None,
    progress=None,
) -> list[FileOutcome]:
    output_dir.mkdir(parents=True, exist_ok=True)
    sources = list_pdf(source, output_dir)
    if not sources:
        return [FileOutcome(source=source, error="Нет файлов .pdf")]

    processor = load_processor(seed) if depersonalize else None
    outcomes: list[FileOutcome] = []
    used_stems: set[str] = set()

    for index, item in enumerate(sources, start=1):
        if progress:
            progress(index, len(sources), item.name)
        outcome = FileOutcome(source=item)
        out_stem = unique_stem(item.stem, used_stems)
        target = output_dir / f"{out_stem}.md"
        if target.exists() and not overwrite:
            outcome.error = f"Файл уже существует: {target.name}"
            outcomes.append(outcome)
            continue
        before = item.read_bytes()
        try:
            document = load_pdf(
                item,
                force_ocr=force_ocr,
                auto=auto,
                save_page_images=save_page_images,
                ocr_page=ocr_page,
            )
            records: list[dict] = []
            pair_log = None
            if processor is not None:
                if save_restore_map:
                    pair_log = begin_pair_log(processor)
                records = anonymize_pdf_document(processor, document)
            markdown, csv_files, payload = render_pdf(document, out_stem)
            target.write_text(markdown, encoding="utf-8")
            outcome.markdown_path = target
            outcome.detections = len(records)
            outcome.warnings.extend(document.warnings)
            pages_path = output_dir / f"{out_stem}.pages.json"
            pages_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            if document.images:
                media = output_dir / media_dirname(out_stem)
                if media.exists():
                    shutil.rmtree(media)
                media.mkdir(parents=True)
                for image in document.images:
                    (media / image.file).write_bytes(image.data)
                image_payload = {
                    "source": item.name,
                    "media_dir": media.name,
                    "images": [
                        {"file": image.file, "role": image.role, "page": image.page}
                        for image in document.images
                    ],
                }
                (media / "images.json").write_text(
                    json.dumps(image_payload, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.media_dir = media
            if csv_files:
                csv_dir = output_dir / out_stem
                if csv_dir.exists():
                    shutil.rmtree(csv_dir)
                csv_dir.mkdir(parents=True)
                for name, text in csv_files.items():
                    (csv_dir / name).write_text(text, encoding="utf-8-sig")
                if outcome.media_dir is None:
                    outcome.media_dir = csv_dir
                else:
                    outcome.warnings.append(f"CSV: {csv_dir}")
            outcome.saved = len(document.images) + len(csv_files)
            if processor is not None and save_report:
                report_path = output_dir / f"{out_stem}.report.json"
                report_path.write_text(
                    json.dumps(build_report(item.name, records), ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.report_path = report_path
            if processor is not None and save_restore_map and pair_log is not None:
                restore_path = output_dir / f"{out_stem}.restore.json"
                restore_path.write_text(
                    json.dumps(build_restore_document(item.name, seed, pair_log.pairs), ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.restore_path = restore_path
                outcome.warnings.append(RESTORE_WARNING)
        except PdfOpenError as exc:
            outcome.error = str(exc)
        except Exception as exc:
            outcome.error = str(exc)
        if item.read_bytes() != before:
            outcome.error = "Исходный PDF был изменён"
        outcomes.append(outcome)
    return outcomes


def convert_xlsx_tree(
    source: Path,
    output_dir: Path,
    *,
    visible_only: bool = True,
    repeat_merges: bool = True,
    keep_formulas: bool = True,
    depersonalize: bool = False,
    save_report: bool = False,
    save_restore_map: bool = False,
    seed: int | None = None,
    overwrite: bool = False,
    progress=None,
) -> list[FileOutcome]:
    output_dir.mkdir(parents=True, exist_ok=True)
    sources = list_xlsx(source, output_dir)
    if not sources:
        return [FileOutcome(source=source, error="Нет файлов .xlsx")]

    processor = load_processor(seed) if depersonalize else None
    outcomes: list[FileOutcome] = []
    used_stems: set[str] = set()

    for index, item in enumerate(sources, start=1):
        if progress:
            progress(index, len(sources), item.name)
        outcome = FileOutcome(source=item)
        out_stem = unique_stem(item.stem, used_stems)
        target = output_dir / f"{out_stem}.md"
        if target.exists() and not overwrite:
            outcome.error = f"Файл уже существует: {target.name}"
            outcomes.append(outcome)
            continue
        try:
            data = read_workbook(
                item,
                visible_only=visible_only,
                repeat_merges=repeat_merges,
                keep_formulas=keep_formulas,
            )
            records: list[dict] = []
            pair_log = None
            if processor is not None:
                if save_restore_map:
                    pair_log = begin_pair_log(processor)
                slots = list(iter_text_slots(data))
                updated, records = anonymize_fragments(processor, [value for _slot, value in slots])
                for slot, value in zip((slot for slot, _value in slots), updated):
                    write_slot(slot, value)
            markdown, csv_files, payload = render_workbook(data, out_stem)
            target.write_text(markdown, encoding="utf-8")
            outcome.markdown_path = target
            outcome.saved = len(csv_files)
            outcome.skipped = sum(1 for sheet in data.sheets if sheet.hidden and sheet.csv_name is None)
            outcome.detections = len(records)
            if csv_files:
                csv_dir = output_dir / out_stem
                if csv_dir.exists():
                    shutil.rmtree(csv_dir)
                csv_dir.mkdir(parents=True)
                for name, text in csv_files.items():
                    (csv_dir / name).write_text(text, encoding="utf-8-sig")
                outcome.media_dir = csv_dir
            sheets_path = output_dir / f"{out_stem}.sheets.json"
            sheets_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            if processor is not None and save_report:
                report_path = output_dir / f"{out_stem}.report.json"
                report_path.write_text(
                    json.dumps(build_report(item.name, records), ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.report_path = report_path
            if processor is not None and save_restore_map and pair_log is not None:
                restore_path = output_dir / f"{out_stem}.restore.json"
                restore_path.write_text(
                    json.dumps(build_restore_document(item.name, seed, pair_log.pairs), ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                outcome.restore_path = restore_path
                outcome.warnings.append(RESTORE_WARNING)
        except Exception as exc:
            outcome.error = str(exc)
        outcomes.append(outcome)
    return outcomes


@dataclass
class RestoreOutcome:
    output_path: Path | None = None
    applied: int = 0
    missing: list[str] = field(default_factory=list)
    ambiguous: list[str] = field(default_factory=list)
    error: str | None = None


def restore_file(
    markdown_path: Path,
    map_path: Path,
    output_path: Path,
    *,
    overwrite: bool = False,
) -> RestoreOutcome:
    """Возвращает исходные значения в отредактированный Markdown по словарю замен."""
    if output_path.resolve() == markdown_path.resolve():
        return RestoreOutcome(error="Нельзя затереть файл правки. Укажите другое имя.")
    if output_path.exists() and not overwrite:
        return RestoreOutcome(error=f"Файл уже существует: {output_path.name}")
    try:
        payload = json.loads(map_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return RestoreOutcome(error=f"Не удалось прочитать словарь: {exc}")
    pairs = payload.get("pairs")
    if not isinstance(pairs, list):
        return RestoreOutcome(error="В словаре нет списка pairs.")
    text, applied, missing, ambiguous = apply_restore(
        markdown_path.read_text(encoding="utf-8"),
        pairs,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return RestoreOutcome(
        output_path=output_path,
        applied=len(applied),
        missing=[str(item.get("replacement", "")) for item in missing],
        ambiguous=[str(item.get("replacement", "")) for item in ambiguous],
    )
