"""Графический интерфейс конвертации DOCX в Markdown."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from docxpipe.anonymize import RESTORE_WARNING
from docxpipe.pdf_convert import PDF_DEPERS_NOTE
from docxpipe.pipeline import (
    SKIP_NOTE,
    convert_pdf_tree,
    convert_tree,
    convert_xlsx_tree,
    existing_outputs,
    list_docx,
    list_pdf,
    list_xlsx,
    restore_file,
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Документы → текст")
        self.geometry("760x640")
        self.minsize(640, 520)

        self._queue: queue.Queue[str] = queue.Queue()
        self._processing = False
        self._build_ui()
        self._poll_queue()

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 4}

        frm_input = ttk.LabelFrame(self, text="Вход", padding=8)
        frm_input.pack(fill="x", **pad)
        ttk.Label(frm_input, text="Файл или папка:").grid(row=0, column=0, sticky="w")
        self._var_input = tk.StringVar()
        ttk.Entry(frm_input, textvariable=self._var_input, width=62).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(frm_input, text="Файл…", width=8, command=self._browse_file).grid(row=0, column=2, padx=2)
        ttk.Button(frm_input, text="Папка…", width=8, command=self._browse_dir).grid(row=0, column=3, padx=2)
        ttk.Label(frm_input, text="Формат:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self._var_format = tk.StringVar(value="DOCX")
        self._fmt = ttk.Combobox(
            frm_input,
            textvariable=self._var_format,
            values=("DOCX", "XLSX", "PDF"),
            state="readonly",
            width=12,
        )
        self._fmt.grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
        self._fmt.bind("<<ComboboxSelected>>", lambda _event: self._apply_format())
        frm_input.columnconfigure(1, weight=1)

        frm_output = ttk.LabelFrame(self, text="Выход", padding=8)
        frm_output.pack(fill="x", **pad)
        ttk.Label(frm_output, text="Папка результата:").grid(row=0, column=0, sticky="w")
        self._var_output = tk.StringVar()
        ttk.Entry(frm_output, textvariable=self._var_output, width=62).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(frm_output, text="Обзор…", width=8, command=self._browse_output).grid(row=0, column=2, padx=2)
        frm_output.columnconfigure(1, weight=1)

        frm_opts = ttk.LabelFrame(self, text="Параметры", padding=8)
        frm_opts.pack(fill="x", **pad)

        self._var_images = tk.BooleanVar(value=False)
        self._chk_images = ttk.Checkbutton(
            frm_opts,
            text="Извлечь рисунки и схемы",
            variable=self._var_images,
            command=self._toggle_images,
        )
        self._chk_images.grid(row=0, column=0, sticky="w")

        self._frm_excel = ttk.Frame(frm_opts)
        self._frm_excel.grid(row=0, column=0, sticky="w")
        self._var_visible_only = tk.BooleanVar(value=True)
        self._var_repeat_merges = tk.BooleanVar(value=True)
        self._var_keep_formulas = tk.BooleanVar(value=True)
        ttk.Checkbutton(self._frm_excel, text="Только видимые листы", variable=self._var_visible_only).pack(anchor="w")
        ttk.Checkbutton(
            self._frm_excel,
            text="Повторять значение объединённой ячейки",
            variable=self._var_repeat_merges,
        ).pack(anchor="w")
        ttk.Checkbutton(
            self._frm_excel,
            text="Сохранять текст формул",
            variable=self._var_keep_formulas,
        ).pack(anchor="w")
        self._frm_excel.grid_remove()

        self._frm_pdf = ttk.Frame(frm_opts)
        self._frm_pdf.grid(row=0, column=0, sticky="w")
        self._var_pdf_auto = tk.BooleanVar(value=True)
        self._var_pdf_force = tk.BooleanVar(value=False)
        self._var_pdf_images = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self._frm_pdf,
            text="Разбирать страницы автоматически",
            variable=self._var_pdf_auto,
        ).pack(anchor="w")
        ttk.Checkbutton(
            self._frm_pdf,
            text="Распознать все страницы",
            variable=self._var_pdf_force,
        ).pack(anchor="w")
        ttk.Checkbutton(
            self._frm_pdf,
            text="Сохранять изображение распознанных страниц",
            variable=self._var_pdf_images,
        ).pack(anchor="w")
        ttk.Label(self._frm_pdf, text="Язык OCR: rus+eng").pack(anchor="w")
        self._frm_pdf.grid_remove()

        self._frm_ocr = ttk.Frame(frm_opts)
        self._frm_ocr.grid(row=1, column=0, sticky="w", padx=(24, 0))
        self._var_ocr = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self._frm_ocr,
            text="Локальный OCR (Tesseract, rus+eng)",
            variable=self._var_ocr,
        ).pack(anchor="w")
        self._frm_ocr.grid_remove()

        self._var_depers = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm_opts,
            text="Деперсонализация (152-ФЗ)",
            variable=self._var_depers,
            command=self._toggle_depers,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        self._frm_depers = ttk.Frame(frm_opts)
        self._frm_depers.grid(row=3, column=0, sticky="w", padx=(24, 0))
        self._var_report = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self._frm_depers,
            text="Сохранить JSON-отчёт",
            variable=self._var_report,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(self._frm_depers, text="Seed:").grid(row=0, column=1, sticky="e", padx=(16, 4))
        self._var_seed = tk.StringVar()
        ttk.Entry(self._frm_depers, textvariable=self._var_seed, width=12).grid(row=0, column=2, sticky="w")
        ttk.Label(self._frm_depers, text="(пусто = случайный)").grid(row=0, column=3, sticky="w", padx=4)
        self._var_restore_map = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self._frm_depers,
            text="Сохранить словарь замен",
            variable=self._var_restore_map,
        ).grid(row=1, column=0, columnspan=4, sticky="w")
        self._frm_depers.grid_remove()

        frm_btn = ttk.Frame(self, padding=4)
        frm_btn.pack(fill="x", **pad)
        self._btn_run = ttk.Button(frm_btn, text="Обработать", command=self._run)
        self._btn_run.pack(side="left", padx=4)
        self._btn_restore = ttk.Button(frm_btn, text="Восстановить", command=self._restore)
        self._btn_restore.pack(side="left", padx=4)

        frm_log = ttk.LabelFrame(self, text="Лог обработки", padding=4)
        frm_log.pack(fill="both", expand=True, **pad)
        self._log = ScrolledText(frm_log, height=14, state="disabled", wrap="word", font=("Consolas", 9))
        self._log.pack(fill="both", expand=True)

        self._var_status = tk.StringVar(value="Готово")
        ttk.Label(self, textvariable=self._var_status, relief="sunken", anchor="w", padding=4).pack(
            fill="x", side="bottom"
        )

    def _apply_format(self) -> None:
        kind = self._var_format.get()
        self._chk_images.grid_remove()
        self._frm_ocr.grid_remove()
        self._frm_excel.grid_remove()
        self._frm_pdf.grid_remove()
        if kind == "XLSX":
            self._var_ocr.set(False)
            self._frm_excel.grid()
        elif kind == "PDF":
            self._var_ocr.set(False)
            self._frm_pdf.grid()
        else:
            self._chk_images.grid()
            self._toggle_images()

    def _toggle_images(self) -> None:
        if self._var_images.get():
            self._frm_ocr.grid()
        else:
            self._var_ocr.set(False)
            self._frm_ocr.grid_remove()

    def _toggle_depers(self) -> None:
        if self._var_depers.get():
            self._frm_depers.grid()
        else:
            self._frm_depers.grid_remove()

    def _browse_file(self) -> None:
        kind = self._var_format.get()
        path = filedialog.askopenfilename(
            title=f"Выберите .{kind.lower()}",
            filetypes=[(kind, f"*.{kind.lower()}"), ("Все файлы", "*.*")],
        )
        if path:
            self._var_input.set(path)

    def _browse_dir(self) -> None:
        path = filedialog.askdirectory(title=f"Выберите папку с .{self._var_format.get().lower()}")
        if path:
            self._var_input.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку для результата")
        if path:
            self._var_output.set(path)

    def _run(self) -> None:
        input_path = self._var_input.get().strip()
        output_path = self._var_output.get().strip()
        if not input_path or not output_path:
            messagebox.showwarning("Документы → текст", "Укажите вход и папку результата.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("Документы → текст", f"Путь не найден:\n{input_path}")
            return
        fmt = self._var_format.get().lower()
        selected = Path(input_path)
        if selected.is_file() and selected.suffix.lower() != f".{fmt}":
            messagebox.showwarning("Документы → текст", "Файл не соответствует выбранному формату.")
            return
        seed_raw = self._var_seed.get().strip()
        if self._var_depers.get() and seed_raw and not seed_raw.isdigit():
            messagebox.showwarning("Документы → текст", "Seed должен быть целым числом или пустым.")
            return
        if self._processing:
            messagebox.showinfo("Документы → текст", "Обработка уже выполняется.")
            return

        if fmt == "xlsx":
            lister = list_xlsx
        elif fmt == "pdf":
            lister = list_pdf
        else:
            lister = list_docx
        sources = lister(selected, Path(output_path))
        occupied = existing_outputs(sources, Path(output_path))
        overwrite = False
        if occupied:
            names = "\n".join(item.name for item in occupied[:8])
            if not messagebox.askyesno("Документы → текст", f"Перезаписать существующие файлы?\n{names}"):
                return
            overwrite = True

        self._processing = True
        self._btn_run.state(["disabled"])
        self._btn_restore.state(["disabled"])
        self._clear_log()
        self._var_status.set("Обработка…")
        seed = int(seed_raw) if seed_raw else None
        worker = threading.Thread(
            target=self._worker,
            args=(
                input_path,
                output_path,
                self._var_images.get(),
                self._var_ocr.get(),
                self._var_depers.get(),
                self._var_report.get(),
                self._var_restore_map.get(),
                seed,
                overwrite,
                fmt,
                self._var_visible_only.get(),
                self._var_repeat_merges.get(),
                self._var_keep_formulas.get(),
                self._var_pdf_auto.get(),
                self._var_pdf_force.get(),
                self._var_pdf_images.get(),
            ),
            daemon=True,
        )
        worker.start()

    def _worker(
        self,
        input_path: str,
        output_path: str,
        images: bool,
        ocr: bool,
        depers: bool,
        save_report: bool,
        save_restore_map: bool,
        seed: int | None,
        overwrite: bool,
        fmt: str,
        visible_only: bool,
        repeat_merges: bool,
        keep_formulas: bool,
        pdf_auto: bool,
        pdf_force: bool,
        pdf_images: bool,
    ) -> None:
        try:
            if depers:
                self._log_msg("Инициализация детектора DP152…")
            if fmt == "docx":
                self._log_msg(SKIP_NOTE)
            if fmt == "pdf" and depers:
                self._log_msg(PDF_DEPERS_NOTE)

            def on_progress(current: int, total: int, name: str) -> None:
                self._log_msg(f"[{current}/{total}] {name}")

            if fmt == "pdf":
                outcomes = convert_pdf_tree(
                    Path(input_path),
                    Path(output_path),
                    force_ocr=pdf_force,
                    auto=pdf_auto,
                    save_page_images=pdf_images,
                    depersonalize=depers,
                    save_report=save_report and depers,
                    save_restore_map=save_restore_map and depers,
                    seed=seed,
                    overwrite=overwrite,
                    progress=on_progress,
                )
            elif fmt == "xlsx":
                outcomes = convert_xlsx_tree(
                    Path(input_path),
                    Path(output_path),
                    visible_only=visible_only,
                    repeat_merges=repeat_merges,
                    keep_formulas=keep_formulas,
                    depersonalize=depers,
                    save_report=save_report and depers,
                    save_restore_map=save_restore_map and depers,
                    seed=seed,
                    overwrite=overwrite,
                    progress=on_progress,
                )
            else:
                outcomes = convert_tree(
                    Path(input_path),
                    Path(output_path),
                    extract_images=images,
                    ocr=ocr and images,
                    depersonalize=depers,
                    save_report=save_report and depers,
                    save_restore_map=save_restore_map and depers,
                    seed=seed,
                    overwrite=overwrite,
                    progress=on_progress,
                )
            for outcome in outcomes:
                if outcome.error and outcome.markdown_path is None:
                    self._log_msg(f"Ошибка ({outcome.source.name}): {outcome.error}")
                    continue
                if outcome.markdown_path:
                    self._log_msg(f"Markdown: {outcome.markdown_path}")
                if fmt == "xlsx":
                    self._log_msg(f"Листы CSV: {outcome.saved}, скрытых пропущено: {outcome.skipped}")
                    if outcome.media_dir:
                        self._log_msg(f"CSV: {outcome.media_dir}")
                elif fmt == "pdf":
                    self._log_msg(f"Карта: {outcome.markdown_path.with_suffix('.pages.json')}")
                    self._log_msg(f"Картинки и таблицы: {outcome.saved}")
                    if outcome.media_dir:
                        self._log_msg(f"Файлы: {outcome.media_dir}")
                elif outcome.media_dir:
                    self._log_msg(
                        f"Рисунки: {outcome.saved} → {outcome.media_dir.name} (пропущено {outcome.skipped})"
                    )
                elif images:
                    self._log_msg(f"Рисунки не отобраны (пропущено {outcome.skipped})")
                if depers:
                    self._log_msg(f"Замен: {outcome.detections}")
                if outcome.report_path:
                    self._log_msg(f"Отчёт: {outcome.report_path}")
                if outcome.restore_path:
                    self._log_msg(f"Словарь замен: {outcome.restore_path}")
                for warning in outcome.warnings:
                    self._log_msg(warning)
        except Exception as exc:
            self._log_msg(f"\nОШИБКА: {exc}")
        finally:
            self._queue.put("__DONE__")

    def _restore(self) -> None:
        if self._processing:
            messagebox.showinfo("Документы → текст", "Обработка уже выполняется.")
            return
        markdown = filedialog.askopenfilename(
            title="Отредактированный Markdown",
            filetypes=[("Markdown", "*.md"), ("Все файлы", "*.*")],
        )
        if not markdown:
            return
        mapping = filedialog.askopenfilename(
            title="Словарь замен",
            filetypes=[("Словарь замен", "*.restore.json"), ("JSON", "*.json"), ("Все файлы", "*.*")],
        )
        if not mapping:
            return
        source = Path(markdown)
        destination = filedialog.asksaveasfilename(
            title="Куда сохранить восстановленный файл",
            initialdir=str(source.parent),
            initialfile=f"{source.stem}.restored.md",
            defaultextension=".md",
            filetypes=[("Markdown", "*.md")],
        )
        if not destination:
            return
        if Path(destination).resolve() == source.resolve():
            messagebox.showerror("Документы → текст", "Нельзя затереть файл правки. Укажите другое имя.")
            return

        self._processing = True
        self._btn_run.state(["disabled"])
        self._btn_restore.state(["disabled"])
        self._clear_log()
        self._var_status.set("Восстановление…")
        worker = threading.Thread(
            target=self._restore_worker,
            args=(source, Path(mapping), Path(destination)),
            daemon=True,
        )
        worker.start()

    def _restore_worker(self, markdown: Path, mapping: Path, destination: Path) -> None:
        try:
            outcome = restore_file(markdown, mapping, destination, overwrite=True)
            if outcome.error:
                self._log_msg(f"Ошибка: {outcome.error}")
                return
            self._log_msg(f"Восстановлено: {outcome.output_path}")
            self._log_msg(f"Пар применено: {outcome.applied}")
            for replacement in outcome.missing:
                self._log_msg(f"В тексте нет замены «{replacement}»")
            for replacement in outcome.ambiguous:
                self._log_msg(f"Пропущена неоднозначная замена «{replacement}»")
            self._log_msg(RESTORE_WARNING)
        except Exception as exc:
            self._log_msg(f"\nОШИБКА: {exc}")
        finally:
            self._queue.put("__DONE__")

    def _log_msg(self, text: str) -> None:
        self._queue.put(text + "\n")

    def _clear_log(self) -> None:
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _poll_queue(self) -> None:
        while True:
            try:
                msg = self._queue.get_nowait()
            except queue.Empty:
                break
            if msg.strip() == "__DONE__":
                self._processing = False
                self._btn_run.state(["!disabled"])
                self._btn_restore.state(["!disabled"])
                self._var_status.set("Готово")
                continue
            self._log.configure(state="normal")
            self._log.insert("end", msg)
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(100, self._poll_queue)


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
