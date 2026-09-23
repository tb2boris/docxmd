---
name: image-index-recognize
description: >-
  Reads image_<filename>/images.json produced by docxmd and writes a semantic
  index of the extracted figures: kind, title, description, suggested alt.
  Use after docxmd export when the user asks to распознать рисунки, дописать
  индекс изображений, or describe screenshots already saved next to a Markdown file.
---

# Смысловой индекс рисунков docxmd

Exe и CLI уже отобрали рисунки и схемы. Этот навык не конвертирует Word и не выгружает новые файлы. Он читает готовые кадры и дополняет индекс.

## Вход

- Папка `image_<имя>/` рядом с `.md`
- `image_<имя>/images.json`

Если папки нет — остановиться и сообщить, что выгрузка рисунков была выключена или ни один кадр не прошёл отбор.

## Шаги

1. Прочитать `images.json`. Не удалять поля `id`, `file`, `path`, `sha256`, `bytes`, `heading`, `caption`, `reason`, `preview`, `ocr_text`.
2. Прочитать каждый файл с `preview: true`. Кадры `preview: false` (EMF/WMF) не описывать по пикселям: в описании указать, что предпросмотр недоступен.
3. На кадр дописать в ту же запись:
   - `kind`: `diagram` | `screenshot` | `table` | `photo` | `other`
   - `title`
   - `description`
   - `suggested_alt`
4. Не выдумывать текст и объекты, которых нет на кадре. Нечитаемое помечать как неразобранное.
5. Записать обновлённый `images.json` и рядом `image-index.md`: таблица `id / заголовок раздела / подпись / тип / краткое описание`.

Иконки и логотипы в эту папку не попадают. Не добавлять их из исходного Word.

## Чего не делать

- Не запускать деперсонализацию заново.
- Не переименовывать `img_001` и не переносить папку.
- Не подменять структурный индекс пустым шаблоном, если кадр не открылся: оставить запись и указать ошибку чтения в `description`.
