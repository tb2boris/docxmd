import base64
import json
from pathlib import Path

import fitz

from docxpipe.pdf_convert import fill_merges
from docxpipe.pipeline import convert_pdf_tree

FONT = next(
    path
    for path in (Path(r"C:\Windows\Fonts\arial.ttf"), Path(r"C:\Windows\Fonts\calibri.ttf"))
    if path.is_file()
)
FILLER = "Служебный абзац документа содержит достаточно букв для текстового слоя."
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _font() -> fitz.Font:
    return fitz.Font(fontfile=str(FONT))


def _write(page: fitz.Page, runs: list[tuple]) -> None:
    font = _font()
    writer = fitz.TextWriter(page.rect)
    for run in runs:
        x, y, text, size = run
        writer.append((x, y), text, font=font, fontsize=size)
    writer.write_text(page)


def _save(doc: fitz.Document, path: Path) -> Path:
    doc.save(path)
    doc.close()
    return path


def _text_page(runs: list[tuple], path: Path) -> Path:
    doc = fitz.open()
    page = doc.new_page()
    _write(page, runs)
    return _save(doc, path)


def _ocr(_blob: bytes) -> tuple[str, float]:
    return "Распознанный абзац приказа", 92.0


def test_fill_merges_repeats_span_and_records_range():
    rows, merges = fill_merges([["Название", None], ["Альфа", "Бета"]])
    assert rows == [["Название", "Название"], ["Альфа", "Бета"]]
    assert merges == ["A1:B1"]


def test_text_page_joins_gap_hyphen_and_heading(tmp_path: Path):
    source = _text_page(
        [
            (72, 150, "Раздел первый", 22),
            (72, 210, "Приказ", 12),
            (130, 210, "номер", 12),
            (72, 250, "конт-", 12),
            (72, 266, "роль", 12),
            (72, 320, FILLER, 12),
            (72, 345, "Вторая строка тела документа для устойчивого кегля.", 12),
            (72, 370, "Третья строка тела остаётся обычным абзацем.", 12),
        ],
        tmp_path / "приказ.pdf",
    )
    before = source.read_bytes()
    out = tmp_path / "out"
    outcomes = convert_pdf_tree(source, out, overwrite=True)
    assert outcomes[0].error is None
    assert source.read_bytes() == before
    text = (out / "приказ.md").read_text(encoding="utf-8")
    payload = json.loads((out / "приказ.pages.json").read_text(encoding="utf-8"))
    assert "Приказ номер" in text
    assert "контроль" in text
    assert "конт-" not in text
    assert "### Раздел первый" in text
    page = payload["pages"][0]
    assert page["kind"] == "text"
    assert page["headings"][0]["text"] == "Раздел первый"
    assert page["headings"][0]["heading_guess"] is True


def test_repeated_header_is_removed_from_body(tmp_path: Path):
    doc = fitz.open()
    for index in range(1, 4):
        page = doc.new_page()
        _write(
            page,
            [
                (72, 36, "Колонтитул приказа", 11),
                (72, 220, f"Текст листа {index}. {FILLER}", 12),
            ],
        )
    source = _save(doc, tmp_path / "журнал.pdf")
    out = tmp_path / "out"
    convert_pdf_tree(source, out, overwrite=True)
    text = (out / "журнал.md").read_text(encoding="utf-8")
    payload = json.loads((out / "журнал.pages.json").read_text(encoding="utf-8"))
    assert "Колонтитул приказа" not in text
    assert payload["headers_footers"] == ["Колонтитул приказа"]
    assert "Текст листа 1" in text
    assert "Текст листа 3" in text


def test_two_columns_are_not_mixed(tmp_path: Path):
    source = _text_page(
        [
            (72, 160, "ЛеваяПервая", 12),
            (72, 185, "ЛеваяВторая", 12),
            (72, 210, "ЛеваяТретья", 12),
            (360, 160, "ПраваяПервая", 12),
            (360, 185, "ПраваяВторая", 12),
            (360, 210, "ПраваяТретья", 12),
            (72, 420, FILLER, 12),
        ],
        tmp_path / "колонки.pdf",
    )
    out = tmp_path / "out"
    convert_pdf_tree(source, out, overwrite=True)
    text = (out / "колонки.md").read_text(encoding="utf-8")
    payload = json.loads((out / "колонки.pages.json").read_text(encoding="utf-8"))
    assert text.find("ЛеваяПервая") < text.find("ЛеваяТретья") < text.find("ПраваяПервая")
    assert payload["pages"][0]["layout"] == "columns"


def test_image_page_is_ocr_and_outline_is_marked(tmp_path: Path):
    doc = fitz.open()
    scanned = doc.new_page(width=320, height=420)
    scanned.insert_image(scanned.rect, stream=PNG)
    outlined = doc.new_page(width=320, height=420)
    for index in range(40):
        outlined.draw_line((40, 40 + index * 4), (280, 40 + index * 4))
    source = _save(doc, tmp_path / "скан.pdf")
    calls: list[int] = []

    def ocr(blob: bytes) -> tuple[str, float]:
        calls.append(len(blob))
        return _ocr(blob)

    out = tmp_path / "out"
    outcomes = convert_pdf_tree(source, out, overwrite=True, ocr_page=ocr)
    assert outcomes[0].error is None
    text = (out / "скан.md").read_text(encoding="utf-8")
    payload = json.loads((out / "скан.pages.json").read_text(encoding="utf-8"))
    assert payload["pages"][0]["kind"] == "ocr"
    assert payload["pages"][1]["kind"] == "outline"
    assert text.count("Распознанный абзац приказа") == 2
    assert (out / "image_скан" / "page_001.png").is_file()
    assert (out / "image_скан" / "page_002.png").is_file()
    assert len(calls) == 2
    assert not (out / "скан").exists()


def test_low_confidence_does_not_dump_garbage(tmp_path: Path):
    doc = fitz.open()
    page = doc.new_page(width=320, height=420)
    page.insert_image(page.rect, stream=PNG)
    source = _save(doc, tmp_path / "шум.pdf")
    out = tmp_path / "out"
    convert_pdf_tree(
        source,
        out,
        overwrite=True,
        ocr_page=lambda _blob: ("###$$$!!!", 12.0),
    )
    text = (out / "шум.md").read_text(encoding="utf-8")
    assert "Распознавание ненадёжно" in text
    assert "###$$$" not in text
    assert (out / "image_шум" / "page_001.png").is_file()


def test_text_table_csv_repeats_merge(tmp_path: Path):
    doc = fitz.open()
    page = doc.new_page()
    shape = page.new_shape()
    x0, y0, width, height = 72, 180, 150, 28
    shape.draw_rect(fitz.Rect(x0, y0, x0 + 2 * width, y0 + height))
    shape.draw_rect(fitz.Rect(x0, y0 + height, x0 + width, y0 + 2 * height))
    shape.draw_rect(fitz.Rect(x0 + width, y0 + height, x0 + 2 * width, y0 + 2 * height))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    _write(
        page,
        [
            (80, y0 + 18, "Название", 11),
            (80, y0 + height + 18, "Альфа", 11),
            (x0 + width + 8, y0 + height + 18, "Бета", 11),
            (72, 120, FILLER, 12),
        ],
    )
    source = _save(doc, tmp_path / "ведомость.pdf")
    out = tmp_path / "out"
    convert_pdf_tree(source, out, overwrite=True)
    text = (out / "ведомость.md").read_text(encoding="utf-8")
    payload = json.loads((out / "ведомость.pages.json").read_text(encoding="utf-8"))
    csv_path = out / "ведомость" / "p1_t1.csv"
    assert csv_path.is_file()
    body = csv_path.read_text(encoding="utf-8-sig")
    assert "Название,Название" in body.replace("\r", "")
    assert "Альфа,Бета" in body
    assert payload["pages"][0]["tables"][0]["merges"] == ["A1:B1"]
    assert "p1_t1.csv" in text
    assert "Название" not in text.split("Таблица на странице")[0]


def test_scan_table_does_not_write_csv(tmp_path: Path):
    doc = fitz.open()
    page = doc.new_page(width=320, height=420)
    page.insert_image(page.rect, stream=PNG)
    shape = page.new_shape()
    for step in range(4):
        shape.draw_line((40, 80 + step * 30), (280, 80 + step * 30))
        shape.draw_line((40 + step * 60, 80), (40 + step * 60, 200))
    shape.finish(color=(0, 0, 0), width=1)
    shape.commit()
    source = _save(doc, tmp_path / "скан-таблица.pdf")
    out = tmp_path / "out"
    convert_pdf_tree(source, out, overwrite=True, ocr_page=_ocr)
    assert not (out / "скан-таблица").exists()
    text = (out / "скан-таблица.md").read_text(encoding="utf-8")
    assert "p1_t1.csv" not in text
    payload = json.loads((out / "скан-таблица.pages.json").read_text(encoding="utf-8"))
    assert payload["pages"][0]["kind"] == "ocr"


def test_form_fields_go_to_map_and_markdown(tmp_path: Path):
    doc = fitz.open()
    page = doc.new_page()
    _write(page, [(72, 120, FILLER, 12)])
    widget = fitz.Widget()
    widget.field_name = "fio"
    widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
    widget.field_value = "Иванов Иван"
    widget.rect = fitz.Rect(72, 200, 320, 230)
    page.add_widget(widget)
    source = _save(doc, tmp_path / "анкета.pdf")
    out = tmp_path / "out"
    convert_pdf_tree(source, out, overwrite=True)
    text = (out / "анкета.md").read_text(encoding="utf-8")
    payload = json.loads((out / "анкета.pages.json").read_text(encoding="utf-8"))
    assert payload["pages"][0]["form_fields"] == [{"name": "fio", "value": "Иванов Иван"}]
    assert "fio: Иванов Иван" in text


def test_password_and_corrupt_do_not_write_markdown(tmp_path: Path):
    locked = fitz.open()
    locked.new_page().insert_text((72, 80), "secret")
    locked_path = tmp_path / "закрыт.pdf"
    locked.save(locked_path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="secret", owner_pw="owner")
    locked.close()
    broken = tmp_path / "битый.pdf"
    broken.write_bytes(b"%PDF-1.4\nthis is not a pdf")
    out = tmp_path / "out"
    locked_outcome = convert_pdf_tree(locked_path, out, overwrite=True)[0]
    broken_outcome = convert_pdf_tree(broken, out, overwrite=True)[0]
    assert locked_outcome.markdown_path is None
    assert "парол" in (locked_outcome.error or "")
    assert broken_outcome.markdown_path is None
    assert broken_outcome.error
    assert not (out / "закрыт.md").exists()
    assert not (out / "битый.md").exists()


def test_figure_on_text_page_and_depersonalize_keeps_bytes(tmp_path: Path):
    doc = fitz.open()
    page = doc.new_page()
    _write(page, [(72, 120, f"Почта user@example.com. {FILLER}", 12)])
    page.insert_image(fitz.Rect(72, 200, 220, 340), stream=PNG)
    source = _save(doc, tmp_path / "карточка.pdf")
    plain = tmp_path / "plain"
    masked = tmp_path / "masked"
    plain_outcome = convert_pdf_tree(source, plain, overwrite=True)[0]
    masked_outcome = convert_pdf_tree(
        source,
        masked,
        overwrite=True,
        depersonalize=True,
        save_report=True,
        seed=1,
    )[0]
    assert plain_outcome.error is None
    assert masked_outcome.error is None
    plain_image = plain / "image_карточка" / "img_001.png"
    masked_image = masked / "image_карточка" / "img_001.png"
    assert plain_image.is_file()
    assert plain_image.read_bytes() == masked_image.read_bytes()
    plain_text = (plain / "карточка.md").read_text(encoding="utf-8")
    masked_text = (masked / "карточка.md").read_text(encoding="utf-8")
    report = (masked / "карточка.report.json").read_text(encoding="utf-8")
    assert "user@example.com" in plain_text
    assert "user@example.com" not in masked_text
    assert "user@example.com" not in report
    assert "image_карточка/img_001.png" in masked_text
    assert "page_001" not in plain_text
