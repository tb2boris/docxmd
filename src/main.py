"""Командная точка входа конвертации DOCX в Markdown."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docxpipe.pdf_convert import PDF_DEPERS_NOTE
from docxpipe.pipeline import SKIP_NOTE, convert_pdf_tree, convert_tree, convert_xlsx_tree, restore_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="docxmd", description="DOCX → Markdown")
    parser.add_argument("input", nargs="?", help="Файл .docx/.xlsx/.pdf или папка с такими файлами")
    parser.add_argument("--output", "-o", help="Папка результата или файл восстановления")
    parser.add_argument("--format", choices=["docx", "xlsx", "pdf"], default=None, help="Формат входа")
    parser.add_argument("--pdf-force-ocr", action="store_true", help="Распознать все страницы PDF")
    parser.add_argument("--no-pdf-page-images", action="store_true", help="Не сохранять растр распознанных страниц")
    parser.add_argument("--images", action="store_true", help="Извлечь рисунки и схемы")
    parser.add_argument("--ocr", action="store_true", help="Локальный OCR (нужен --images)")
    parser.add_argument("--depersonalize", action="store_true", help="Деперсонализация ядром DP152")
    parser.add_argument("--report", action="store_true", help="JSON-отчёт деперсонализации")
    parser.add_argument("--restore-map", action="store_true", help="Сохранить словарь замен .restore.json")
    parser.add_argument("--seed", type=int, default=None, help="Seed синтетических замен")
    parser.add_argument("--force", action="store_true", help="Перезаписать существующие .md")
    parser.add_argument("--restore", help="Отредактированный Markdown для возврата исходных значений")
    parser.add_argument("--map", help="Файл словаря замен .restore.json")
    return parser


def restore_cli(args: argparse.Namespace) -> int:
    if not args.map or not args.output:
        print("Ошибка: для восстановления нужны --map и --output", file=sys.stderr)
        return 1
    source = Path(args.restore)
    mapping = Path(args.map)
    if not source.is_file():
        print(f"Ошибка: файл не найден: {source}", file=sys.stderr)
        return 1
    if not mapping.is_file():
        print(f"Ошибка: словарь не найден: {mapping}", file=sys.stderr)
        return 1
    output = Path(args.output)
    if output.suffix.lower() != ".md":
        output = output / f"{source.stem}.restored.md"
    outcome = restore_file(source, mapping, output, overwrite=args.force)
    if outcome.error:
        print(f"Ошибка: {outcome.error}", file=sys.stderr)
        return 1
    print(f"Восстановлено: {outcome.output_path}")
    print(f"Пар применено: {outcome.applied}")
    for replacement in outcome.missing:
        print(f"В тексте нет замены «{replacement}»")
    for replacement in outcome.ambiguous:
        print(f"Пропущена неоднозначная замена «{replacement}»")
    return 0


def _resolve_format(source: Path, explicit: str | None) -> str | None:
    if source.is_file():
        suffix = source.suffix.lower().lstrip(".")
        if explicit and explicit != suffix:
            print(f"Ошибка: файл {source.name} не соответствует формату {explicit}", file=sys.stderr)
            return None
        if suffix in {"docx", "xlsx", "pdf"}:
            return explicit or suffix
        print(f"Ошибка: нужен .docx, .xlsx или .pdf, получен {source.name}", file=sys.stderr)
        return None
    if not explicit:
        print("Ошибка: для папки укажите --format docx, xlsx или pdf", file=sys.stderr)
        return None
    return explicit


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.restore:
        return restore_cli(args)
    if not args.input or not args.output:
        print("Ошибка: укажите входной файл и --output", file=sys.stderr)
        return 1
    source = Path(args.input)
    if not source.exists():
        print(f"Ошибка: путь не найден: {source}", file=sys.stderr)
        return 1
    fmt = _resolve_format(source, args.format)
    if fmt is None:
        return 1
    if fmt == "pdf":
        if args.depersonalize:
            print(PDF_DEPERS_NOTE)
        outcomes = convert_pdf_tree(
            source,
            Path(args.output),
            force_ocr=args.pdf_force_ocr,
            save_page_images=not args.no_pdf_page_images,
            depersonalize=args.depersonalize,
            save_report=args.report and args.depersonalize,
            save_restore_map=args.restore_map and args.depersonalize,
            seed=args.seed,
            overwrite=args.force,
        )
    elif fmt == "xlsx":
        outcomes = convert_xlsx_tree(
            source,
            Path(args.output),
            depersonalize=args.depersonalize,
            save_report=args.report and args.depersonalize,
            save_restore_map=args.restore_map and args.depersonalize,
            seed=args.seed,
            overwrite=args.force,
        )
    else:
        print(SKIP_NOTE)
        outcomes = convert_tree(
            source,
            Path(args.output),
            extract_images=args.images,
            ocr=args.ocr and args.images,
            depersonalize=args.depersonalize,
            save_report=args.report and args.depersonalize,
            save_restore_map=args.restore_map and args.depersonalize,
            seed=args.seed,
            overwrite=args.force,
        )
    failed = 0
    for outcome in outcomes:
        if outcome.error and outcome.markdown_path is None:
            print(f"Ошибка ({outcome.source.name}): {outcome.error}", file=sys.stderr)
            failed += 1
            continue
        if outcome.markdown_path:
            print(f"Markdown: {outcome.markdown_path}")
        if outcome.media_dir:
            label = "CSV" if fmt == "xlsx" else "Файлы" if fmt == "pdf" else "Рисунки"
            print(f"{label}: {outcome.saved} → {outcome.media_dir}")
        if outcome.report_path:
            print(f"Отчёт: {outcome.report_path}")
        if outcome.restore_path:
            print(f"Словарь замен: {outcome.restore_path}")
        for warning in outcome.warnings:
            print(warning)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
