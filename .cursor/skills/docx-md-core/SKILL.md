---
name: docx-md-core
description: >-
  Converts a DOCX file to Markdown with headings, lists and tables, and optionally
  extracts only figures and schemes into image_<filename> with images.json and
  local Tesseract OCR. Use when converting Word to Markdown for docxmd, running
  the docxmd CLI, or refreshing image_<stem>/images.json from a .docx.
---

# DOCX → Markdown (docxmd)

Один проход. Исходный `.docx` не изменяется. Реализация — пакет `docxpipe` в `c:\_CURSOR\docxmd\src`, не копия скриптов Россетей.

## Запуск

Каталог — `c:\_CURSOR\docxmd\src`.

```powershell
python main.py "<файл.docx>" --output "<каталог>" --images
```

Флаги:

- `--images` — папка `image_<имя_файла>`, ссылки, `images.json`
- `--ocr` — только вместе с `--images`; Tesseract в PATH, языки `rus+eng`
- `--depersonalize` — ядро DP152, замена сразу
- `--report` — только вместе с `--depersonalize`
- `--seed` — целое число
- `--force` — перезаписать существующий `.md`

## Отбор изображений

Сохранять только то, что помечено как «рисунок», «рис.», «схема», «диаграмма», «график», «чертёж», либо то, на что есть отсылка в соседнем тексте («на рисунке», «как показано на схеме»). Иконки, логотипы и прочие кадры без такой пометки не выгружать.

Имена кадров: `img_001.png`. Папка и индекс привязаны к имени файла. Повтор имени в одном запуске — суффикс `_2`.

## Чего не делать

- Не писать результат поверх исходного docx.
- Не вызывать `convert_docx.py` и `convert_illustrated_docx.py` из навыков Россетей: у них другие пути и другой отбор картинок.
- Не добавлять режим «только обнаружение» DP152.
