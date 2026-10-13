import os
import queue
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog

from app_core import (
    run_pipeline, run_pipeline_multi_clinic, run_clinic_comparison,
    get_file_info, count_rows_by_clinic, load_data,
    ALL_CLINICS_LABEL, COMPARE_CLINICS_LABEL, COMPARE_CLINICS_KEY,
    DEFAULT_CLINIC_LABEL,
    save_compare_preset, load_compare_preset,
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

        # Tk Listbox внутри CTk-фрейма — множественный выбор с Ctrl/Shift
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

        # Порог значимости (единый для всех режимов): число определённых
        # ответов S+I+R, ниже которого строка/ячейка не попадает в листы.
        thr_frame = ctk.CTkFrame(actions, fg_color="transparent")
        thr_frame.pack(side="right", padx=8)
        ctk.CTkLabel(thr_frame, text="Порог значимости n≥",
                     font=ctk.CTkFont(size=13)).pack(side="left", padx=(0, 4))
        self.min_n_entry = ctk.CTkEntry(thr_frame, width=52, height=28)
        self.min_n_entry.insert(0, "0")
        self.min_n_entry.pack(side="left")
        ctk.CTkLabel(thr_frame, text="(0 = без порога)",
                     font=ctk.CTkFont(size=11),
                     text_color=("gray45", "gray65")).pack(side="left", padx=(4, 0))

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
            # фоновое определение клиник и антибиотиков — чтобы конструктор и
            # режим «Сравнение клиник» сразу знали списки
            self._antibiotics_cache = []
            self._clinics_cache = []
            self._populate_clinics([])
            self._scan_token += 1
            token = self._scan_token
            threading.Thread(target=self._scan_antibiotics, args=(filename, token),
                             daemon=True).start()

    def _scan_antibiotics(self, path, token=None):
        try:
            clinics, antibiotics = get_file_info(path, log_func=None)
            if token is not None and token != self._scan_token:
                return  # пользователь уже выбрал другой файл
            self._antibiotics_cache = list(antibiotics)
            self._clinics_cache = list(clinics)
            counts = {}
            try:
                _, _, data, _, _ = load_data(path, log_func=None)
                counts = count_rows_by_clinic(data)
            except Exception:
                pass
            self._log_queue.put(f"Найдено клиник: {len(clinics)} "
                                f"({', '.join(clinics) if clinics else '—'}).")
            self._log_queue.put(f"Список антибиотиков обновлён: {len(antibiotics)} шт. "
                                f"(для конструктора листов и сравнения клиник)")
            # обновление виджетов — в главном потоке
            self.after(0, lambda: self._populate_clinics(clinics, counts))
        except Exception as e:
            self._antibiotics_cache = []
            self._clinics_cache = []
            self._log_queue.put(f"Не удалось заранее прочитать файл: {e}")

    # ------------------------------------------------------------------
    # Режим «Сравнение клиник» (отдельный режим работы)
    # ------------------------------------------------------------------
    def _is_compare_mode(self):
        return self.clinic_menu.get() == COMPARE_CLINICS_LABEL

    def _on_clinic_mode_changed(self, choice):
        """Переключение видимости панели сравнения + подсказки о режиме."""
        compare = choice == COMPARE_CLINICS_LABEL
        if compare:
            self.compare_frame.grid()
            self.builder_btn.configure(state="disabled")
            self.builder_info.configure(
                text="режим «Сравнение клиник»: базовые листы не создаются")
        else:
            self.compare_frame.grid_remove()
            self.builder_btn.configure(state="normal")
            n = len(self._sheet_specs)
            self.builder_info.configure(text=f"доп. листов: {n}")
        self._update_cmp_info()

    def _populate_clinics(self, clinics, counts=None):
        """Меню режимов + чекбоксы клиник из исходного файла (все отмечены)."""
        # обновляем выпадающий список режимов: Все клиники | клиники | Сравнение
        menu_values = [ALL_CLINICS_LABEL] + list(clinics) + [COMPARE_CLINICS_LABEL]
        self.clinic_menu.configure(values=menu_values)
        cur = self.clinic_menu.get()
        if cur not in menu_values:  # клиника из прежнего файла исчезла — сброс
            self.clinic_menu.set(ALL_CLINICS_LABEL)
            self._on_clinic_mode_changed(ALL_CLINICS_LABEL)

        for cb in self._clinic_checkboxes.values():
            cb.destroy()
        self._clinic_vars.clear()
        self._clinic_checkboxes.clear()
        if not clinics:
            self._clinics_placeholder.grid(row=0, column=0, sticky="w")
            self._update_cmp_info()
            return
        self._clinics_placeholder.grid_forget()
        counts = counts or {}
        import tkinter as tk
        for i, cl in enumerate(clinics):
            var = tk.BooleanVar(master=self, value=True)
            n = counts.get(cl)
            label = f"{cl} ({n} посевов)" if n else cl
            cb = ctk.CTkCheckBox(self.cmp_clinic_box, text=label, variable=var,
                                 font=ctk.CTkFont(size=13),
                                 command=self._update_cmp_info)
            cb.grid(row=0, column=i, padx=(0, 18), pady=2)
            self._clinic_vars[cl] = var
            self._clinic_checkboxes[cl] = cb
        self._update_cmp_info()

    def selected_clinics(self):
        return [cl for cl, v in self._clinic_vars.items() if v.get()]

    def selected_compare_antibiotics(self):
        sel = []
        for i in self.cmp_ab_lb.curselection():
            try:
                sel.append(self.cmp_ab_lb.get(i))
            except Exception:
                pass
        return sel

    def _refilter_cmp_antibiotics(self):
        """Фильтрация списка антибиотиков по строке поиска (выбор сохраняется)."""
        chosen = set(self.selected_compare_antibiotics())
        query = self.cmp_search_var.get().strip().lower()
        self.cmp_ab_lb.delete(0, "end")
        shown = 0
        for ab in self._antibiotics_cache:
            if query and query not in ab.lower():
                continue
            self.cmp_ab_lb.insert("end", ab)
            if ab in chosen:
                self.cmp_ab_lb.selection_set(shown)
            shown += 1
        self._update_cmp_info()

    def _cmp_select_all_abs(self):
        query = self.cmp_search_var.get().strip().lower()
        self.cmp_ab_lb.select_clear(0, "end")
        for i, ab in enumerate(self.cmp_ab_lb.get(0, "end")):
            if not query or query in ab.lower():
                self.cmp_ab_lb.selection_set(i)
        self._update_cmp_info()

    def _cmp_clear_abs(self):
        self.cmp_ab_lb.select_clear(0, "end")
        self._update_cmp_info()

    def _update_cmp_info(self):
        if not self._is_compare_mode():
            self.cmp_info.configure(text="")
            return
        n_abs = len(self.selected_compare_antibiotics())
        total_abs = len(self._antibiotics_cache)
        abs_txt = (f"антибиотики: все ({total_abs})" if n_abs == 0
                   else f"антибиотики: выбрано {n_abs} из {total_abs}")
        cl = self.selected_clinics()
        cl_txt = f"клиники: {', '.join(cl) if cl else 'не отмечены'}"
        warn = ""
        if self._clinics_cache and len(cl) < 2:
            warn = "  ⚠️ отметьте минимум две клиники"
        self.cmp_info.configure(text=f"ℹ️ {abs_txt};  {cl_txt}{warn}")

    def current_min_n(self):
        """Порог значимости со главного окна (0 = без порога)."""
        from app_core import parse_min_significant
        return parse_min_significant(self.min_n_entry.get())

    def _build_compare_spec(self):
        """Спецификация листа для отдельного режима «Сравнение клиник»."""
        return {
            "name": "Сравнение клиник",
            "mode": COMPARE_CLINICS_KEY,
            "clinics": self.selected_clinics(),
            "antibiotics": self.selected_compare_antibiotics(),
            "compare_by": ["organism", "material_raw"],
            "filter": "positive_only",
            "min_n": self.current_min_n(),
        }

    def _save_compare_settings(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить настройки сравнения",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            save_compare_preset(self._build_compare_spec(), path)
            self.log(f"Настройки сравнения сохранены: {path}")
        except Exception as e:
            self._error(f"Не удалось сохранить настройки: {e}")

    def _load_compare_settings(self):
        path = filedialog.askopenfilename(
            title="Загрузить настройки сравнения",
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            spec = load_compare_preset(path)
        except Exception as e:
            self._error(f"Не удалось загрузить настройки: {e}")
            return
        # антибиотики: отметить в списке только сохранённые (если они есть в файле)
        wanted = set(spec.get("antibiotics") or [])
        self.cmp_ab_lb.select_clear(0, "end")
        if wanted:
            for i, ab in enumerate(self.cmp_ab_lb.get(0, "end")):
                if ab in wanted:
                    self.cmp_ab_lb.selection_set(i)
        # клиники: отметить сохранённые (если они есть в текущем файле)
        saved_clinics = spec.get("clinics") or []
        known = [c for c in saved_clinics if c in self._clinic_vars]
        if known:
            for cl, var in self._clinic_vars.items():
                var.set(cl in set(known))
        self._update_cmp_info()
        self.log(f"Настройки сравнения загружены: {os.path.basename(path)}"
                 + ("" if known else " (клиники из пресета не найдены в текущем файле — "
                                     "будут использованы все клиники файла)"))

    # ------------------------------------------------------------------
    # Конструктор дополнительных листов
    # ------------------------------------------------------------------
    def open_sheet_builder(self):
        dlg = SheetBuilderDialog(self,
                                 antibiotics_getter=lambda: self._antibiotics_cache,
                                 clinics_getter=lambda: self._clinics_cache)
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
        min_n = self.current_min_n()
        compare_mode = self._is_compare_mode()
        specs = [] if compare_mode else list(self._sheet_specs)

        if compare_mode:
            # отдельный режим «Сравнение клиник»: клиники чекбоксами,
            # антибиотики списком (пусто = все); базовые листы не создаются
            clinics_sel = self.selected_clinics()
            if len(clinics_sel) < 2:
                self._error("В режиме «Сравнение клиник» отметьте минимум две клиники "
                            "(чекбоксы под выбором файла).")
                return
            spec = self._build_compare_spec()
            mode_txt = COMPARE_CLINICS_LABEL
        else:
            spec = None
            mode_txt = self.clinic_menu.get()

        self._processing = True
        self.run_btn.configure(state="disabled", text="Обработка…", fg_color="gray45")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.progress.start()
        self.log("Начало обработки...")

        if compare_mode:
            self.log(f"Режим: {COMPARE_CLINICS_LABEL}. "
                     f"Клиники: {', '.join(spec['clinics'])}. "
                     f"Антибиотики: "
                     + (f"все ({len(self._antibiotics_cache)})" if not spec['antibiotics']
                        else ", ".join(spec['antibiotics']))
                     + (f". Порог значимости: n≥{min_n}." if min_n else
                        ". Порог значимости не применяется.")
                     + " Базовые листы 1–3 не создаются.")
        else:
            self.log(f"Клиника: {mode_txt}"
                     + (f". Порог значимости: n≥{min_n}." if min_n else ""))
            if specs:
                self.log(f"План: 3 базовых листа + {len(specs)} из конструктора: "
                         + ", ".join(s['name'] for s in specs))

        thread = threading.Thread(target=self._task,
                                  args=(input_path, output_path, styled, specs,
                                        mode_txt, compare_mode, spec, min_n),
                                  daemon=True)
        thread.start()

    def _task(self, input_path, output_path, styled, specs, mode_txt,
              compare_mode=False, compare_spec=None, min_n=0):
        # Из фонового потока не трогаем Tk напрямую — только очередь и флаг.
        try:
            if compare_mode:
                built = run_clinic_comparison(input_path, output_path, compare_spec,
                                              styled=styled, log_func=self.log)
                names = ", ".join(built.keys())
                self._log_queue.put(f"Результат сохранён: {output_path} "
                                    f"(листы: {names})")
            elif mode_txt == ALL_CLINICS_LABEL:
                results = run_pipeline_multi_clinic(
                    input_path, output_path, styled=styled, log_func=self.log,
                    extra_sheets=specs, global_min_n=min_n)
                listing = "; ".join(f"{cl}: {os.path.basename(p)}"
                                    for cl, p in results.items())
                self._log_queue.put(f"Создано файлов: {len(results)} — {listing}")
            else:
                # выбранная клиника из меню; если её нет в файле (например,
                # «Общая выборка») — обычная обработка всего файла
                clinic = None if mode_txt in (ALL_CLINICS_LABEL, DEFAULT_CLINIC_LABEL) \
                    else mode_txt
                run_pipeline(input_path, output_path, styled=styled,
                             log_func=self.log, extra_sheets=specs,
                             clinic=clinic, global_min_n=min_n)
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
