"""Локальный OCR через Tesseract. Бинарник в приложение не вшивается."""

from __future__ import annotations

import io
import shutil


class OcrUnavailable(Exception):
    """Tesseract не найден в PATH или не смог прочитать кадр."""


def tesseract_available() -> bool:
    return shutil.which("tesseract") is not None


def recognize_page(blob: bytes, *, lang: str = "rus+eng") -> tuple[str, float]:
    """Распознаёт страницу целиком: абзацы сохраняются, возвращается средняя уверенность.

    В отличие от recognize(), пробелы и переносы не схлопываются в одну строку.
    """
    if not tesseract_available():
        raise OcrUnavailable("Tesseract не найден в PATH. Страница не распознана.")
    try:
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        raise OcrUnavailable("Нужны пакеты Pillow и pytesseract.") from exc
    try:
        image = Image.open(io.BytesIO(blob))
        data = pytesseract.image_to_data(image, lang=lang, output_type=pytesseract.Output.DICT)
    except Exception as exc:
        raise OcrUnavailable(f"OCR не выполнен: {exc}") from exc

    groups: dict[tuple[int, int, int], list[str]] = {}
    order: list[tuple[int, int, int]] = []
    confidences: list[float] = []
    texts = data.get("text") or []
    for index, raw in enumerate(texts):
        word = str(raw or "").strip()
        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError, KeyError, IndexError):
            confidence = -1
        if confidence >= 0 and word:
            confidences.append(confidence)
        if not word:
            continue
        key = (
            int(data["block_num"][index]),
            int(data["par_num"][index]),
            int(data["line_num"][index]),
        )
        if key not in groups:
            order.append(key)
            groups[key] = []
        groups[key].append(word)

    parts: list[str] = []
    previous_par: tuple[int, int] | None = None
    for key in order:
        paragraph = (key[0], key[1])
        if previous_par is not None and paragraph != previous_par:
            parts.append("")
        parts.append(" ".join(groups[key]))
        previous_par = paragraph
    text = "\n".join(parts).strip()
    mean = sum(confidences) / len(confidences) if confidences else 0.0
    return text, mean


def recognize(blob: bytes, extension: str, *, lang: str = "rus+eng") -> str:
    if extension.lower() in {".emf", ".wmf", ".svg", ".bin"}:
        return ""
    if not tesseract_available():
        raise OcrUnavailable("Tesseract не найден в PATH. Текст рисунков не распознан.")
    try:
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        raise OcrUnavailable("Нужны пакеты Pillow и pytesseract.") from exc
    try:
        image = Image.open(io.BytesIO(blob))
        text = pytesseract.image_to_string(image, lang=lang)
    except Exception as exc:
        raise OcrUnavailable(f"OCR не выполнен: {exc}") from exc
    return " ".join(text.split())
