import os
import queue
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog

from app_core import (
    run_pipeline, run_pipeline_multi_clinic, run_clinic_comparison, load_data,
    get_file_info, detect_clinics, ALL_CLINICS_LABEL, COMPARE_CLINICS_LABEL,
    COMPARE_CLINICS_KEY, validate_spec, save_compare_preset, load_compare_preset,
)
from sheet_builder import SheetBuilderDialog

ctk.set_appearance_mode("dark")          # "dark", "light" или "system"
ctk.set_default_color_theme("blue")      # blue / green / dark-blue

ACCENT = "#2CC98F"
DANGER = "#E35D4A"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🧬  Обработка микробиологических данных")
        self.geometry("980x760")
        self.minsize(860, 640)

        self.input_file = ctk.StringVar()
        self.output_file = ctk.StringVar()
        self._log_queue = queue.Queue()
        self._processing = False
        self._finished = threading.Event()
        self._finished_ok = False
        # Конструктор дополнительных листов
        self._sheet_specs = []          # спецификации из конструктора
        self._antibiotics_cache = []    # антибиотики последнего загруженного файла
        # Клиники последнего загруженного файла
        self._clinics_cache = []        # список клиник (без спецпунктов)
        self._clinic_vars = {}          # код клиники -> BooleanVar (режим сравнения)
        self._clinic_checkboxes = {}    # код клиники -> CTkCheckBox
        self._scan_token = 0            # отмена устаревших фоновых сканов файла

        self._build_ui()
        self.after(100, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Интерфейс
    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---------- Шапка ----------
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray86", "gray17"), height=74)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left", padx=24, pady=12)
        ctk.CTkLabel(title_box, text="Микробиология · Аналитика посевов",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_box, text="Антибиотикограммы · эпидемиология · резистентность",
                     font=ctk.CTkFont(size=12), text_color=("gray40", "gray65")).pack(anchor="w")

        theme_box = ctk.CTkFrame(header, fg_color="transparent")
        theme_box.pack(side="right", padx=24)
        ctk.CTkLabel(theme_box, text="Тема:", font=ctk.CTkFont(size=13)).pack(side="left", padx=(0, 8))
        self.theme_menu = ctk.CTkOptionMenu(
            theme_box, values=["System", "Dark", "Light"], width=110,
            command=self._change_theme)
        self.theme_menu.pack(side="right")
        self.theme_menu.set("Dark")

        # ---------- Основная область ----------
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=18, pady=(12, 6))

        # Файлы
        files_card = ctk.CTkFrame(body)
        files_card.pack(fill="x", pady=(0, 10))

        inner = ctk.CTkFrame(files_card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)
        inner.columnconfigure(1, weight=1)

        ctk.CTkLabel(inner, text="Исходный файл (Excel с результатами посевов)",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=3,
                                                                    sticky="w", pady=(0, 6))
        self.input_entry = ctk.CTkEntry(inner, textvariable=self.input_file,
                                        placeholder_text="Выберите .xlsx с результатами посевов…")
        self.input_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ctk.CTkButton(inner, text="Обзор…", width=110, command=self.select_input
                      ).grid(row=1, column=2, sticky="e")

        ctk.CTkLabel(inner, text="Результирующий файл",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, columnspan=3,
                                                                    sticky="w", pady=(14, 6))
        self.output_entry = ctk.CTkEntry(inner, textvariable=self.output_file,
                                         placeholder_text="Куда сохранить отчёт .xlsx…")
        self.output_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ctk.CTkButton(inner, text="Сохранить как…", width=110, command=self.select_output
                      ).grid(row=3, column=2, sticky="e")

        # ---------- Клиника: выбор режима обработки ----------
        clinic_card = ctk.CTkFrame(body)
        clinic_card.pack(fill="x", pady=(0, 10))

        cinner = ctk.CTkFrame(clinic_card, fg_color="transparent")
        cinner.pack(fill="x", padx=16, pady=14)
        cinner.columnconfigure(0, weight=1)

        top_row = ctk.CTkFrame(cinner, fg_color="transparent")
        top_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        top_row.columnconfigure(1, weight=1)
        ctk.CTkLabel(top_row, text="Клиника:",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0,
                                                                    sticky="w", padx=(0, 8))
        self.clinic_menu = ctk.CTkOptionMenu(
            top_row, values=[ALL_CLINICS_LABEL], width=260,
            command=self._on_clinic_mode_changed)
        self.clinic_menu.grid(row=0, column=1, sticky="w")
        self.clinic_menu.set(ALL_CLINICS_LABEL)

        # Панель отдельного режима «Сравнение клиник» (видна только в нём)
        self.compare_frame = ctk.CTkFrame(cinner, fg_color="transparent")
        self.compare_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.compare_frame.grid_remove()

        cf_inner = ctk.CTkFrame(self.compare_frame, fg_color="transparent")
        cf_inner.pack(fill="x")
        cf_inner.columnconfigure(1, weight=1)

        self.cmp_title_label = ctk.CTkLabel(cf_inner, text="Антибиотики для сравнения\n(ничего не отмечено = все):",
                                            font=ctk.CTkFont(size=12), justify="left")
        self.cmp_title_label.grid(row=0, column=0, sticky="nw", padx=(0, 8))
        self.cmp_ab_frame = ctk.CTkFrame(cf_inner, fg_color="transparent")
        self.cmp_ab_frame.grid(row=0, column=1, sticky="ew")
        self.cmp_ab_frame.columnconfigure(0, weight=1)

        cmp_list_row = ctk.CTkFrame(self.cmp_ab_frame, fg_color="transparent")
        cmp_list_row.grid(row=0, column=0, sticky="ew")
        cmp_list_row.columnconfigure(0, weight=1)
        self.cmp_search_var = ctk.StringVar()
        self.cmp_search_var.trace_add("write", lambda *_: self._refilter_cmp_antibiotics())
        ctk.CTkEntry(cmp_list_row, textvariable=self.cmp_search_var, height=28,
                     placeholder_text="Поиск антибиотика по названию…").grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        cmp_btns = ctk.CTkFrame(cmp_list_row, fg_color="transparent")
        cmp_btns.grid(row=0, column=1)
        ctk.CTkButton(cmp_btns, text="Выбрать все", width=92, height=26,
                      fg_color="transparent", border_width=1,
                      command=self._cmp_select_all_abs).pack(side="left", padx=(0, 4))
        ctk.CTkButton(cmp_btns, text="Снять всё", width=92, height=26,
                      fg_color="transparent", border_width=1,
                      command=self._cmp_clear_abs).pack(side="left")

        self.cmp_ab_listbox = ctk.CTkTextbox(self.cmp_ab_frame, height=120,
                                             wrap="none", activatescrollbars=True)
        # Tk Listbox внутри CTk-фрейма — многосписочный выбор с Ctrl/Shift
        import tkinter as tk
        self.cmp_ab_lb = tk.Listbox(self.cmp_ab_frame, selectmode="extended",
                                    height=6, exportselection=False,
                                    activestyle="none", font=("Segoe UI", 11),
                                    bg="#2b2b2b", fg="#e5e5e5",
                                    selectbackground="#3B82F6", selectforeground="white",
                                    highlightthickness=1, highlightbackground="#555")
        self.cmp_ab_lb.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.cmp_ab_hint = ctk.CTkLabel(self.cmp_ab_frame, text="",
                                       font=ctk.CTkFont(size=11),
                                       text_color=("gray45", "gray65"))
        self.cmp_ab_hint.grid(row=2, column=0, sticky="w", pady=(3, 0))

        # Чекбоксы клиник (заполняются после чтения исходного файла)
        clinic_row = ctk.CTkFrame(cf_inner, fg_color="transparent")
        clinic_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        clinic_row.columnconfigure(1, weight=1)
        self.cmp_clinic_title = ctk.CTkLabel(clinic_row,
                                             text="Клиники для сравнения\n(отметьте минимум две):",
                                             font=ctk.CTkFont(size=12), justify="left")
        self.cmp_clinic_title.grid(row=0, column=0, sticky="nw", padx=(0, 8))
        self.cmp_clinic_box = ctk.CTkFrame(clinic_row, fg_color="transparent")
        self.cmp_clinic_box.grid(row=0, column=1, sticky="ew")
        self._clinics_placeholder = ctk.CTkLabel(
            self.cmp_clinic_box,
            text="Клиники появятся здесь после выбора исходного файла.",
            font=ctk.CTkFont(size=12), text_color=("gray45", "gray65"))
        self._clinics_placeholder.grid(row=0, column=0, sticky="w")

        cmp_presets_row = ctk.CTkFrame(cf_inner, fg_color="transparent")
        cmp_presets_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ctk.CTkButton(cmp_presets_row, text="💾 Сохранить настройки сравнения…",
                      width=230, height=28, fg_color="transparent", border_width=1,
                      command=self._save_compare_settings).pack(side="left", padx=(0, 6))
        ctk.CTkButton(cmp_presets_row, text="📂 Загрузить настройки сравнения…",
                      width=230, height=28, fg_color="transparent", border_width=1,
                      command=self._load_compare_settings).pack(side="left")

        self.cmp_info = ctk.CTkLabel(cinner, text="",
                                     font=ctk.CTkFont(size=12),
                                     text_color=("gray40", "gray65"), justify="left")
        self.cmp_info.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # Конструктор листов + опции + кнопки запуска
        actions = ctk.CTkFrame(body, fg_color="transparent")
        actions.pack(fill="x", pady=(0, 10))

        self.builder_btn = ctk.CTkButton(actions, text="🧩 Конструктор листов…",
                                         width=200, height=40, corner_radius=10,
                                         fg_color="#3B82F6", hover_color="#2f6ad0",
                                         font=ctk.CTkFont(size=14, weight="bold"),
                                         command=self.open_sheet_builder)
        self.builder_btn.pack(side="left", padx=4)

        self.builder_info = ctk.CTkLabel(actions, text="доп. листов: 0",
                                         font=ctk.CTkFont(size=12),
                                         text_color=("gray40", "gray65"))
        self.builder_info.pack(side="left", padx=(2, 10))

        self.styled_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(actions, text="Красивое оформление Excel (шапки, зебра, подсветка %)",
                      variable=self.styled_var, font=ctk.CTkFont(size=13)
                      ).pack(side="right", padx=6)

        self.run_btn = ctk.CTkButton(actions, text="▶  Запустить обработку",
                                     width=220, height=40, corner_radius=10,
                                     font=ctk.CTkFont(size=15, weight="bold"),
                                     fg_color=ACCENT, hover_color="#25a878",
                                     command=self.run_processing)
        self.run_btn.pack(side="right", padx=4)

        ctk.CTkButton(actions, text="Открыть папку результата", width=170,
                      fg_color="transparent", border_width=1,
                      command=self.open_output_folder).pack(side="right", padx=8)

        # Прогресс
        self.progress = ctk.CTkProgressBar(body, mode="indeterminate", height=8)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(0, 10))

        # Журнал
        log_card = ctk.CTkFrame(body)
        log_card.pack(fill="both", expand=True)
        ctk.CTkLabel(log_card, text="Журнал обработки",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(10, 4))
        self.log_area = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=12),
                                       wrap="word")
        self.log_area.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.log_area.configure(state="disabled")

        # Строка состояния
        self.status = ctk.CTkLabel(self, text="Готово к работе", anchor="w",
                                   font=ctk.CTkFont(size=12),
                                   text_color=("gray40", "gray65"))
        self.status.pack(fill="x", padx=24, pady=(0, 8))

    # ------------------------------------------------------------------
    # Слоты интерфейса
    # ------------------------------------------------------------------
    def _change_theme(self, choice):
        mapping = {"System": "system", "Dark": "dark", "Light": "light"}
        ctk.set_appearance_mode(mapping[choice])

    def select_input(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if filename:
            self.input_file.set(filename)
            if not self.output_file.get():
                dirname = os.path.dirname(filename)
                basename = os.path.basename(filename)
                name, _ = os.path.splitext(basename)
                stamp = datetime.now().strftime("%Y%m%d_%H%M")
                self.output_file.set(os.path.join(dirname, f"{name}_обработанный_{stamp}.xlsx"))
            # фоновое определение антибиотиков — чтобы конструктор сразу знал список
            self._antibiotics_cache = []
            threading.Thread(target=self._scan_antibiotics, args=(filename,),
                             daemon=True).start()

    def _scan_antibiotics(self, path):
        try:
            _, _, _, antibiotics, _ = load_data(path, log_func=None)
            self._antibiotics_cache = list(antibiotics)
            self._log_queue.put(f"Список антибиотиков обновлён: {len(antibiotics)} шт. "
                                f"(для конструктора листов)")
        except Exception as e:
            self._antibiotics_cache = []
            self._log_queue.put(f"Не удалось заранее прочитать антибиотики: {e}")

    # ------------------------------------------------------------------
    # Конструктор дополнительных листов
    # ------------------------------------------------------------------
    def open_sheet_builder(self):
        dlg = SheetBuilderDialog(self,
                                 antibiotics_getter=lambda: self._antibiotics_cache)
        dlg.wait_window()
        if dlg.specs is not None:  # нажата кнопка «Применить»
            self._sheet_specs = dlg.specs
            n = len(self._sheet_specs)
            self.builder_info.configure(text=f"доп. листов: {n}")
            names = ", ".join(s["name"] for s in self._sheet_specs[:5])
            self.log(f"Конструктор: добавлено листов — {n} ({names})"
                     + ("…" if n > 5 else ""))

    def select_output(self):
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.output_file.set(filename)

    def open_output_folder(self):
        path = self.output_file.get().strip()
        folder = os.path.dirname(path) if path else ""
        if folder and os.path.isdir(folder):
            try:
                if os.name == "nt":
                    os.startfile(folder)  # noqa
                elif os.name == "posix":
                    opener = "open" if __import__("sys").platform == "darwin" else "xdg-open"
                    __import__("subprocess").Popen([opener, folder])
                self.log(f"Открыта папка: {folder}")
            except Exception as e:
                self.log(f"Не удалось открыть папку: {e}")
        else:
            self.log("Папка результата ещё не выбрана.")

    # ------------------------------------------------------------------
    # Журналирование (через очередь — безопасно из фонового потока)
    # ------------------------------------------------------------------
    def log(self, message):
        self._log_queue.put(str(message))

    # ------------------------------------------------------------------
    # Запуск обработки
    # ------------------------------------------------------------------
    def run_processing(self):
        if self._processing:
            return
        input_path = self.input_file.get().strip()
        output_path = self.output_file.get().strip()
        if not input_path or not os.path.isfile(input_path):
            self._error("Выберите существующий исходный файл.")
            return
        if not output_path:
            self._error("Укажите имя результирующего файла.")
            return

        styled = bool(self.styled_var.get())
        specs = list(self._sheet_specs)

        self._processing = True
        self.run_btn.configure(state="disabled", text="Обработка…", fg_color="gray45")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.progress.start()
        self.log("Начало обработки...")
        if specs:
            self.log(f"План: 3 базовых листа + {len(specs)} из конструктора: "
                     + ", ".join(s['name'] for s in specs))

        thread = threading.Thread(target=self._task,
                                  args=(input_path, output_path, styled, specs),
                                  daemon=True)
        thread.start()

    def _task(self, input_path, output_path, styled, specs):
        # Из фонового потока не трогаем Tk напрямую — только очередь и флаг.
        try:
            run_pipeline(input_path, output_path, styled=styled, log_func=self.log,
                         extra_sheets=specs)
            self._log_queue.put(f"Результат сохранён: {output_path}")
            self._finished_ok = True
        except Exception as e:
            self._log_queue.put(f"ОШИБКА: {e}")
            self._finished_ok = False
        finally:
            self._finished.set()

    def _poll_log_queue(self):
        try:
            while True:
                msg = self._log_queue.get_nowait()
                stamp = datetime.now().strftime("%H:%M:%S")
                self.log_area.configure(state="normal")
                if msg.startswith("ОШИБКА"):
                    line = f"[{stamp}] ⛔ {msg}\n"
                elif "✅" in msg or "завершена" in msg:
                    line = f"[{stamp}] ✅ {msg}\n"
                else:
                    line = f"[{stamp}]  {msg}\n"
                self.log_area.insert("end", line)
                self.log_area.see("end")
                self.log_area.configure(state="disabled")
                self.status.configure(text=msg[:110])
        except queue.Empty:
            pass
        # Завершение фоновой задачи обрабатываем в главном потоке
        if self._processing and self._finished.is_set():
            self.progress.stop()
            self.progress.set(0)
            self.run_btn.configure(state="normal", text="▶  Запустить обработку", fg_color=ACCENT)
            self._processing = False
            ok = getattr(self, "_finished_ok", False)
            self._finished.clear()
            if not ok:
                from tkinter import messagebox
                messagebox.showerror("Ошибка", "При обработке произошла ошибка.\nПодробности — в журнале.")
        self.after(100, self._poll_log_queue)

    def _error(self, text):
        from tkinter import messagebox
        messagebox.showerror("Ошибка", text)
        self.log(f"ОШИБКА: {text}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
