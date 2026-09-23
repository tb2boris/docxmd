"""Деперсонализация готового Markdown ядром DP152."""

from __future__ import annotations

import importlib
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

_IMG_RE = re.compile(r"!\[[^\]]*\]\([^)\r\n]+\)")

RESTORE_WARNING = (
    "Словарь замен содержит исходные персональные данные. "
    "Не передавайте его вместе с обезличенным Markdown."
)


def dp152_src() -> Path:
    env = os.environ.get("DP152_SRC")
    if env:
        return Path(env)
    return Path(__file__).resolve().parents[3] / "dp152" / "src"


def load_processor(seed: int | None):
    src = dp152_src()
    if src.is_dir():
        src_text = str(src)
        if src_text not in sys.path:
            sys.path.insert(0, src_text)
    try:
        module = importlib.import_module("core.processor")
    except ModuleNotFoundError as exc:
        raise FileNotFoundError(f"Ядро DP152 недоступно. Ожидался каталог {src}") from exc
    return module.DocumentProcessor(seed=seed)


def mask_image_links(text: str) -> tuple[str, list[str]]:
    tokens: list[str] = []

    def repl(match: re.Match[str]) -> str:
        tokens.append(match.group(0))
        return f"<!-- docxmd-img-{len(tokens) - 1} -->"

    return _IMG_RE.sub(repl, text), tokens


def restore_image_links(text: str, tokens: list[str]) -> str:
    for index, token in enumerate(tokens):
        text = text.replace(f"<!-- docxmd-img-{index} -->", token)
    return text


class PairLog:
    """Пары исходное → замена, встретившиеся при обработке одного файла."""

    def __init__(self) -> None:
        self.pairs: list[dict] = []
        self._seen: set[tuple[str, str]] = set()

    def record(self, category: str, original: str, replacement: str) -> None:
        surface = original.strip()
        key = (category, surface.lower())
        if key in self._seen:
            return
        self._seen.add(key)
        self.pairs.append({
            "category": category,
            "original": surface,
            "replacement": replacement,
        })


def begin_pair_log(processor) -> PairLog:
    """Включает запись пар на время обработки одного файла."""
    _install_recorder(processor)
    log = PairLog()
    processor._docxmd_active_log = log
    return log


def _install_recorder(processor) -> None:
    if getattr(processor, "_docxmd_recorder_installed", False):
        return
    mapper = processor.mapper
    original_get = mapper.get_replacement

    def wrapped(original: str, category: str, generator_name: str | None = None) -> str:
        replacement = original_get(original, category, generator_name)
        log = getattr(processor, "_docxmd_active_log", None)
        if log is not None:
            log.record(category, original, replacement)
        return replacement

    mapper.get_replacement = wrapped
    processor._docxmd_recorder_installed = True


def build_restore_document(source_name: str, seed: int | None, pairs: list[dict]) -> dict:
    return {
        "version": 1,
        "source": source_name,
        "seed": seed,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "pairs": pairs,
    }


def apply_restore(text: str, pairs: list[dict]) -> tuple[str, list[dict], list[dict], list[dict]]:
    """Возвращает текст, применённые пары, ненайденные и неоднозначные.

    Неоднозначная замена — одна и та же строка замены у разных исходных значений.
    """
    grouped: dict[str, list[dict]] = {}
    for pair in pairs:
        replacement = str(pair.get("replacement") or "")
        original = str(pair.get("original") or "")
        if not replacement or not original:
            continue
        grouped.setdefault(replacement, []).append(pair)

    usable: list[dict] = []
    ambiguous: list[dict] = []
    for group in grouped.values():
        originals = {str(item["original"]) for item in group}
        if len(originals) > 1:
            ambiguous.extend(group)
        else:
            usable.append(group[0])

    usable.sort(key=lambda item: len(str(item["replacement"])), reverse=True)
    applied: list[dict] = []
    missing: list[dict] = []
    for pair in usable:
        replacement = str(pair["replacement"])
        original = str(pair["original"])
        pattern = re.compile(r"(?<!\w)" + re.escape(replacement) + r"(?!\w)")
        updated, count = pattern.subn(lambda _match: original, text)
        if count == 0:
            missing.append(pair)
            continue
        text = updated
        applied.append(pair)
    return text, applied, missing, ambiguous


def anonymize_markdown(processor, text: str) -> tuple[str, list[dict]]:
    masked, tokens = mask_image_links(text)
    updated, records = processor.process_text(masked, dry_run=False)
    return restore_image_links(updated, tokens), records


def anonymize_fragments(processor, fragments: list[str]) -> tuple[list[str], list[dict]]:
    """Деперсонализирует список строк одним проходом, сохраняя границы фрагментов."""
    if not fragments:
        return [], []
    parts: list[str] = []
    for index, fragment in enumerate(fragments):
        parts.append(fragment)
        parts.append(f"\n<!-- docxmd-cell-{index} -->\n")
    updated, records = processor.process_text("".join(parts), dry_run=False)
    pieces = re.split(r"\n<!-- docxmd-cell-\d+ -->\n", updated)
    if pieces and pieces[-1] == "":
        pieces = pieces[:-1]
    if len(pieces) != len(fragments):
        raise RuntimeError("Не удалось разобрать деперсонализированные ячейки")
    return pieces, records


def anonymize_plain(processor, text: str) -> tuple[str, list[dict]]:
    if not text.strip():
        return text, []
    return processor.process_text(text, dry_run=False)


def build_report(source_name: str, records: list[dict]) -> dict:
    categories: dict[str, dict] = {}
    unique: dict[str, set] = {}
    for record in records:
        category = record["category"]
        categories.setdefault(category, {"count": 0, "unique": 0})
        categories[category]["count"] += 1
        unique.setdefault(category, set())
        unique[category].add(record.get("original_hash") or record.get("matched_text", ""))
    for category, values in unique.items():
        categories[category]["unique"] = len(values)
    return {
        "file": source_name,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "total_detections": len(records),
        "categories": categories,
        "detections": records,
    }
