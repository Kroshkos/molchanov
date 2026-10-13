# -*- coding: utf-8 -*-
"""
sheet_builder.py — конструктор дополнительных листов отчёта.

Первые 3 листа результирующего файла всегда базовые (анализ посевов,
эпидемиология, антибиотикорезистентность). Этот модуль позволяет добавлять
неограниченное количество дополнительных листов, в которых пользователь сам
задаёт группировку по нужным параметрам (микроорганизм, материал, категория,
год/месяц/квартал…), набор метрик и фильтр данных.

Используется как модальное окно из основного приложения (1.py):
    dlg = SheetBuilderDialog(parent, antibiotics_getter=lambda: [...])
    dlg.wait_window()
    specs = dlg.specs          # список спецификаций для run_pipeline(extra_sheets=...)
"""

import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app_core import (
    GROUP_FIELDS, METRIC_OPTIONS, FILTER_OPTIONS, validate_spec,
    save_sheet_specs, load_sheet_specs, SHEET_TEMPLATES, TOTAL_GROUP_KEY,
    AB_ALL_KEY, AB_ALL_LABEL, COMPARE_CLINICS_KEY, parse_min_significant,
    MATRIX_AB_ORG_KEY, MATRIX_AB_ONLY_KEY, MATRIX_AB_ORG_LABEL,
    MATRIX_AB_ONLY_LABEL,
)

ACCENT = "#2CC98F"

#: Пункт режима карточки: обычный лист или сравнение клиник
MODE_NORMAL_LABEL = "Обычный лист (группировка и метрики)"
MODE_COMPARE_LABEL = "⚖️ Сравнение клиник (АБГ по разным клиникам)"
#: Матричные режимы: колонки — антибиотики, строки — микроорганизмы;
#: значение — доля «не устойчивых» (S+I)/(S+I+R)×100
_MODE_LABELS = {
    MODE_NORMAL_LABEL: None,
    MODE_COMPARE_LABEL: COMPARE_CLINICS_KEY,
    MATRIX_AB_ORG_LABEL: MATRIX_AB_ORG_KEY,
    MATRIX_AB_ONLY_LABEL: MATRIX_AB_ONLY_KEY,
}
_MODE_BY_KEY = {v: k for k, v in _MODE_LABELS.items() if v}


def _mode_is_matrix(mode_key):
    return mode_key in (MATRIX_AB_ORG_KEY, MATRIX_AB_ONLY_KEY)


class MultiSelectFrame(ctk.CTkFrame):
    """Поиск + многосписочный выбор значений (антибиотики для сравнения)."""

    def __init__(self, master, values=None, height=140, on_change=None):
        super().__init__(master, fg_color="transparent")
        self._all = list(values or [])
        self.on_change = on_change
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", pady=(0, 4))
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refilter())
        ctk.CTkEntry(top, textvariable=self.search_var, height=28,
                     placeholder_text="Поиск по названию…").pack(side="left", fill="x", expand=True)
        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.pack(side="right", padx=(6, 0))
        ctk.CTkButton(btns, text="Все", width=52, height=26,
                      fg_color="transparent", border_width=1,
                      command=self.select_all).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btns, text="Снять", width=52, height=26,
                      fg_color="transparent", border_width=1,
                      command=self.clear_selection).pack(side="left")
        self.listbox = tk.Listbox(self, selectmode="extended", height=6,
                                  exportselection=False, activestyle="none",
                                  font=("Segoe UI", 11),
                                  bg=("#ebeef3" if ctk.get_appearance_mode() == "Light" else "#2b2b2b"),
                                  fgcolor=("#111318" if ctk.get_appearance_mode() == "Light" else "#e5e5e5"),
                                  selectbackground="#3B82F6", selectforeground="white",
                                  highlightthickness=0, relief="flat")
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._notify())
        self.set_values(self._all)

    def set_values(self, values):
        self._all = list(values)
        self._refilter()

    def _refilter(self):
        q = self.search_var.get().strip().lower()
        shown = [v for v in self._all if not q or q in str(v).lower()]
        sel = set(self.get_selected())
        keep_first = bool(q) and not sel  # при поиске выделяем первые совпадения
        self.listbox.delete(0, "end")
        for i, v in enumerate(shown):
            self.listbox.insert("end", v)
            if v in sel or (keep_first and i < 5):
                self.listbox.selection_set(i)

    def get_selected(self):
        return [self.listbox.get(i) for i in self.listbox.curselection()]

    def select_all(self):
        self.listbox.selection_set(0, "end")
        self._notify()

    def clear_selection(self):
        self.listbox.selection_clear(0, "end")
        self._notify()

    def set_selected(self, items):
        wanted = {str(i) for i in (items or [])}
        self.listbox.selection_clear(0, "end")
        for i in range(self.listbox.size()):
            if self.listbox.get(i) in wanted:
                self.listbox.selection_set(i)

    def _notify(self):
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass


class SheetCard(ctk.CTkFrame):
    """Карточка одного листа конструктора."""

    def __init__(self, master, index, spec=None, antibiotics_getter=None,
                 clinics_getter=None, on_delete=None):
        super().__init__(master, fg_color=("gray84", "gray22"), corner_radius=10)
        self.index = index
        self.on_delete = on_delete
        self._antibiotics_getter = antibiotics_getter or (lambda: [])
        self._clinics_getter = clinics_getter or (lambda: [])
        self._suppress_cb_logic = False  # защита от рекурсии при программной установке

        spec = spec or {}
        self._compare_mode = spec.get("mode") == COMPARE_CLINICS_KEY
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(10, 2))
        ctk.CTkLabel(header, text=f"🧩  Лист {index + 1}",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="✕", width=30, height=24, corner_radius=8,
                      fg_color="#E35D4A", hover_color="#c14a3a",
                      command=self._delete).pack(side="right")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="x", padx=12, pady=(2, 12))
        body.columnconfigure(1, weight=1)

        # Режим листа: обычный, сравнение клиник или матрица
        self.mode_menu = ctk.CTkOptionMenu(
            body, values=list(_MODE_LABELS.keys()), width=340,
            command=self._on_mode_changed)
        self.mode_menu.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))
        _mode_key_init = spec.get("mode") if spec else None
        self._compare_mode = (_mode_key_init == COMPARE_CLINICS_KEY)
        self._matrix_mode = _mode_is_matrix(_mode_key_init)
        self.mode_menu.set(_MODE_BY_KEY.get(_mode_key_init, MODE_NORMAL_LABEL)
                           if _mode_key_init else
                           (MODE_COMPARE_LABEL if self._compare_mode else MODE_NORMAL_LABEL))

        # Название листа
        ctk.CTkLabel(body, text="Название:", font=ctk.CTkFont(size=12)).grid(
            row=1, column=0, sticky="w", pady=4, padx=(0, 8))
        self.name_var = ctk.StringVar(value=spec.get("name", f"Лист {index + 1}"))
        ctk.CTkEntry(body, textvariable=self.name_var, width=220).grid(
            row=1, column=1, columnspan=3, sticky="ew", pady=4)

        # Группировка (несколько чекбоксов; «Всё вместе» — эксклюзивный режим)
        self.grp_label = ctk.CTkLabel(body, text="",
                                      font=ctk.CTkFont(size=12), justify="left")
        self.grp_label.grid(row=2, column=0, sticky="nw", pady=(8, 4), padx=(0, 8))
        grp_frame = ctk.CTkFrame(body, fg_color="transparent")
        grp_frame.grid(row=2, column=1, columnspan=3, sticky="w", pady=(8, 2))
        self.group_vars = {}
        chosen_groups = set(spec.get("compare_by" if self._compare_mode else "groups", []))
        if not spec:
            chosen_groups = {"organism"}
        elif (self._compare_mode or self._matrix_mode) and not chosen_groups:
            chosen_groups = {"organism"}
        for i, (key, label) in enumerate(GROUP_FIELDS):
            var = ctk.BooleanVar(value=key in chosen_groups)
            cb = ctk.CTkCheckBox(grp_frame, text=label, variable=var,
                                 font=ctk.CTkFont(size=12))
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 16), pady=2)
            self.group_vars[key] = var
            var.trace_add("write", lambda *_: self._on_group_toggled())
        self.group_hint = ctk.CTkLabel(grp_frame, text="", font=ctk.CTkFont(size=11),
                                       text_color=("gray45", "gray65"))
        self.group_hint.grid(row=(len(GROUP_FIELDS) + 1) // 2, column=0,
                             columnspan=2, sticky="w", pady=(2, 0))
        self._update_group_hint()

        # Метрики
        self.met_label = ctk.CTkLabel(body, text="Метрики:", font=ctk.CTkFont(size=12))
        self.met_label.grid(row=3, column=0, sticky="nw", pady=(8, 4), padx=(0, 8))
        met_frame = ctk.CTkFrame(body, fg_color="transparent")
        met_frame.grid(row=3, column=1, columnspan=3, sticky="w", pady=(8, 2))
        self.met_frame = met_frame
        self.metric_vars = {}
        chosen_metrics = set(spec.get("metrics", []))
        if not spec:
            chosen_metrics = {"total", "positive", "pct_positive"}
        for i, (key, label) in enumerate(METRIC_OPTIONS):
            var = ctk.BooleanVar(value=key in chosen_metrics)
            ctk.CTkCheckBox(met_frame, text=label, variable=var,
                            font=ctk.CTkFont(size=12)).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 16), pady=2)
            self.metric_vars[key] = var
        self.ab_var = ctk.BooleanVar(value="ab" in chosen_metrics or self._compare_mode)
        ab_row = (len(METRIC_OPTIONS) // 2)
        self.ab_cb = ctk.CTkCheckBox(met_frame, text="Антибиотикограмма (S/I/R/НД)",
                                     variable=self.ab_var, font=ctk.CTkFont(size=12))
        self.ab_cb.grid(row=ab_row, column=0, columnspan=2, sticky="w", pady=2)

        # Антибиотик + фильтр (одна строка, элементы растягиваются по ширине)
        row4 = ctk.CTkFrame(body, fg_color="transparent")
        row4.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        row4.columnconfigure(1, weight=1)
        row4.columnconfigure(3, weight=1)
        self.ab_title = ctk.CTkLabel(row4, text="Антибиотик:", font=ctk.CTkFont(size=12))
        self.ab_title.grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.antibiotic_menu = ctk.CTkOptionMenu(row4, values=["—"])
        self.antibiotic_menu.grid(row=0, column=1, sticky="ew", padx=(0, 18))
        ctk.CTkLabel(row4, text="Фильтр:", font=ctk.CTkFont(size=12)).grid(
            row=0, column=2, sticky="w", padx=(0, 6))
        self.filter_menu = ctk.CTkOptionMenu(
            row4, values=[label for _, label in FILTER_OPTIONS])
        self.filter_menu.grid(row=0, column=3, sticky="ew")

        # --- дополнительные поля режима «Сравнение клиник» ---
        # Несколько антибиотиков для сравнения (пустой список = все из файла)
        cmp_ab_frame = ctk.CTkFrame(body, fg_color="transparent")
        cmp_ab_frame.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        cmp_ab_frame.columnconfigure(1, weight=1)
        self.cmp_ab_title = ctk.CTkLabel(cmp_ab_frame,
                                         text="Антибиотики для сравнения\n(можно несколько):",
                                         font=ctk.CTkFont(size=12), justify="left")
        self.cmp_ab_title.grid(row=0, column=0, sticky="nw", padx=(0, 6))
        self.cmp_ab_listbox = tk.Listbox(cmp_ab_frame, selectmode="extended",
                                         height=5, exportselection=False,
                                         activestyle="none", font=("Segoe UI", 11))
        self.cmp_ab_listbox.grid(row=0, column=1, sticky="ew")
        self.cmp_ab_hint = ctk.CTkLabel(cmp_ab_frame, text="",
                                       font=ctk.CTkFont(size=11),
                                       text_color=("gray45", "gray65"))
        self.cmp_ab_hint.grid(row=1, column=1, sticky="w", pady=(2, 0))
        self._cmp_abs = []
        self.cmp_ab_listbox.bind("<<ListboxSelect>>", lambda *_: self._update_cmp_ab_hint())

        # Клиники для сравнения (чекбоксы из исходного файла)
        self.cmp_clinic_label = ctk.CTkLabel(body,
                                             text="Клиники для сравнения\n(отметьте ≥ 2):",
                                             font=ctk.CTkFont(size=12), justify="left")
        self.cmp_clinic_frame = ctk.CTkFrame(body, fg_color="transparent")
        self.cmp_clinic_frame.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        self._clinic_vars = {}      # код клиники -> BooleanVar
        self._clinic_checkboxes = {}  # код клиники -> CTkCheckBox

        # Порог значимости (общий для всех режимов конструктора): строки/ячейки,
        # где число определённых ответов S+I+R меньше порога, в лист не попадают.
        minn_row = ctk.CTkFrame(body, fg_color="transparent")
        minn_row.grid(row=7, column=0, columnspan=4, sticky="w", pady=(8, 0))
        self.min_n_title = ctk.CTkLabel(minn_row, text="Порог значимости n≥:",
                                        font=ctk.CTkFont(size=12))
        self.min_n_title.grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.min_n_var = ctk.StringVar(value=str(spec.get("min_n", "") or ""))
        self.min_n_entry = ctk.CTkEntry(minn_row, textvariable=self.min_n_var,
                                        width=70, height=26)
        self.min_n_entry.grid(row=0, column=1, sticky="w")
        self.min_n_hint = ctk.CTkLabel(
            minn_row, text="(пусто/0 = без порога; n — число определений S+I+R)",
            font=ctk.CTkFont(size=11), text_color=("gray45", "gray65"))
        self.min_n_hint.grid(row=0, column=2, sticky="w", padx=(6, 0))

        # восстановление выбранных антибиотика/фильтра
        if spec.get("antibiotic"):
            ab_val = spec["antibiotic"]
            # обратная совместимость: в старых пресетах имя могло храниться прямо
            self._pending_ab = AB_ALL_KEY if ab_val == AB_ALL_LABEL else ab_val
        if spec.get("filter"):
            fkey = spec["filter"]
            for key, label in FILTER_OPTIONS:
                if key == fkey:
                    self.filter_menu.set(label)
                    break
        self._refresh_antibiotics()
        self.set_clinics(spec.get("clinics") or [])
        if spec.get("antibiotics"):
            self._set_cmp_selected(spec["antibiotics"])
        # В Tk режим трассировки называется "write" (а не "changed") — иначе
        # создание карточки падало с TclError и кнопка «Добавить лист» не работала.
        self.ab_var.trace_add("write", lambda *_: self._refresh_antibiotics())
        self._apply_mode_layout()
        self._update_cmp_ab_hint()

    # ------------------------------------------------------------------
    def _is_compare(self):
        return bool(getattr(self, "_compare_mode", False))

    def _is_matrix(self):
        return bool(getattr(self, "_matrix_mode", False))

    def _on_mode_changed(self, choice):
        """Переключение режима карточки: обычный / сравнение / матрица."""
        key = _MODE_LABELS.get(choice)
        self._compare_mode = (key == COMPARE_CLINICS_KEY)
        self._matrix_mode = _mode_is_matrix(key)
        if self._compare_mode or self._matrix_mode:
            # в этих режимах АБГ/матрица обязательна, прочие метрики скрыты
            self.ab_var.set(True)
        self._apply_mode_layout()
        self._update_group_hint()

    def _apply_mode_layout(self):
        """Скрывает/показывает элементы в зависимости от режима карточки."""
        try:
            compare = self._is_compare()
            matrix = self._is_matrix()
            special = compare or matrix
            self.grp_label.configure(
                text=("Сравнивать по:\n(микроорганизм, материал…)" if compare else
                      "Строки матрицы должны включать:\n(микроорганизм — обязательно)"
                      if matrix else
                      "Группировать по:\n(можно выбирать несколько)"))
            self.ab_title.configure(
                text="Антибиотик для сравнения:" if compare else "Антибиотик:")
            # поля сравнения: список антибиотиков — и для сравнения, и для
            # матрицы; чекбоксы клиник — только для сравнения клиник
            if special:
                self.cmp_ab_title.configure(
                    text=("Колонки-антибиотики\n(пусто = все из файла):" if matrix
                          else "Антибиотики для сравнения\n(можно несколько):"))
                self.cmp_ab_title.grid()
                self.cmp_ab_listbox.grid()
                self.cmp_ab_hint.grid()
            else:
                self.cmp_ab_title.grid_remove()
                self.cmp_ab_listbox.grid_remove()
                self.cmp_ab_hint.grid_remove()
            if compare:
                self.cmp_clinic_label.grid()
                self.cmp_clinic_frame.grid()
            else:
                self.cmp_clinic_label.grid_remove()
                self.cmp_clinic_frame.grid_remove()
            if special:
                self.met_label.grid_remove()
                self.ab_cb.grid_remove()
                for v in self.metric_vars.values():
                    if v.get():
                        v.set(False)  # скрываем только чекбоксы метрик посева
                for w in self.met_frame.winfo_children():
                    if w is not self.ab_cb:
                        w.grid_remove()
            else:
                self.met_label.grid()
                self.ab_cb.grid()
                for w in self.met_frame.winfo_children():
                    w.grid()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Поля режима «Сравнение клиник»
    # ------------------------------------------------------------------
    def set_clinics(self, clinics, checked=None):
        """Перестраивает чекбоксы клиник (список кодов из исходного файла).

        checked — какие клиники отмечены (по умолчанию все, если ранее
        выбор уже был сделан карточкой — сохраняем пересечение).
        """
        clinics = list(clinics or [])
        prev_checked = {k for k, v in self._clinic_vars.items() if v.get()}
        for cb in self._clinic_checkboxes.values():
            try:
                cb.destroy()
            except Exception:
                pass
        self._clinic_checkboxes.clear()
        self._clinic_vars.clear()
        if checked is None:
            checked_set = set(clinics) if not prev_checked else (prev_checked & set(clinics)) \
                or set(clinics)
        else:
            checked_set = set(checked)
        for i, cl in enumerate(clinics):
            var = ctk.BooleanVar(value=cl in checked_set)
            cb = ctk.CTkCheckBox(self.cmp_clinic_frame, text=str(cl), variable=var,
                                 font=ctk.CTkFont(size=12))
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 18), pady=2)
            self._clinic_vars[cl] = var
            self._clinic_checkboxes[cl] = cb
        if not clinics:
            hint = ctk.CTkLabel(self.cmp_clinic_frame,
                                text="Клиники не определены — выберите исходный файл "
                                     "и обновите список.",
                                font=ctk.CTkFont(size=11),
                                text_color=("#B45309", "#F59E0B"))
            hint.grid(row=0, column=0, sticky="w")
            self._clinic_checkboxes["__hint__"] = hint

    def selected_clinics(self):
        return [k for k, v in self._clinic_vars.items() if v.get()]

    def refresh_antibiotics_multilist(self):
        """Обновляет список антибиотиков для многократного выбора."""
        try:
            abs_list = [a for a in self._antibiotics_getter() if a and a != AB_ALL_KEY]
        except Exception:
            abs_list = []
        sel = self._get_cmp_selected()
        self._cmp_abs = abs_list
        lb = self.cmp_ab_listbox
        lb.delete(0, "end")
        for a in abs_list:
            lb.insert("end", a)
        wanted = {s for s in sel if s in set(abs_list)}
        if wanted:
            for i, a in enumerate(abs_list):
                if a in wanted:
                    lb.selection_set(i)
        self._update_cmp_ab_hint()

    def _get_cmp_selected(self):
        try:
            return [self.cmp_ab_listbox.get(i) for i in self.cmp_ab_listbox.curselection()]
        except Exception:
            return []

    def _set_cmp_selected(self, items):
        wanted = set(items or [])
        lb = self.cmp_ab_listbox
        lb.selection_clear(0, "end")
        for i in range(lb.size()):
            if lb.get(i) in wanted:
                lb.selection_set(i)
        self._update_cmp_ab_hint()

    def _update_cmp_ab_hint(self):
        n_sel = len(self._get_cmp_selected())
        n_all = len(self._cmp_abs)
        if n_sel:
            text = f"Выбрано антибиотиков для сравнения: {n_sel} из {n_all}."
        elif n_all:
            text = f"Ничего не отмечено — будут сравниваться ВСЕ антибиотики файла ({n_all})."
        else:
            text = "Список антибиотиков пуст — выберите исходный файл и обновите список."
        try:
            self.cmp_ab_hint.configure(text=text)
        except Exception:
            pass

    def _on_group_toggled(self):
        """Эксклюзивность «Всё вместе»: при его выборе снимаем прочие поля,
        при выборе прочего — снимаем «Всё вместе»."""
        if getattr(self, "_suppress_cb_logic", False):
            return
        total_var = self.group_vars.get(TOTAL_GROUP_KEY)
        if total_var is None:
            self._update_group_hint()
            return
        others_on = [k for k, v in self.group_vars.items()
                     if k != TOTAL_GROUP_KEY and v.get()]
        if total_var.get():
            if others_on:
                self._suppress_cb_logic = True
                try:
                    for k in others_on:
                        self.group_vars[k].set(False)
                finally:
                    self._suppress_cb_logic = False
        elif others_on:
            pass  # обычная комбинированная группировка
        self._update_group_hint()

    def _update_group_hint(self):
        """Показывает итоговую схему группировки прямо под чекбоксами."""
        labels = dict(GROUP_FIELDS)
        chosen = [k for k, v in self.group_vars.items() if v.get()]
        compare = self._is_compare()
        if not chosen:
            text = ("ℹ️ Поля не выбраны — сравнение будет одной итоговой строкой "
                    "по всем выделениям." if compare else
                    "ℹ️ Группировка не выбрана — лист будет построен как «всё вместе» (итог).")
        elif TOTAL_GROUP_KEY in chosen:
            text = "📊 Режим «всё вместе»: одна итоговая строка по всей выборке."
        elif compare:
            text = ("⚖️ Сравнение клиник: " + " × ".join(labels[k] for k in chosen) +
                    " — столбцы S%/I%/R% для каждой клиники в одном листе.")
        else:
            text = "➡️ Схема: " + " × ".join(labels[k] for k in chosen)
        try:
            self.group_hint.configure(text=text)
        except Exception:
            pass

    def _refresh_antibiotics(self):
        """Заполняет список антибиотиков из исходного файла.
        Первым пунктом всегда идёт «Все антибиотики» (суммарная АБГ).

        Меню хранит только отображаемые имена; сопоставление имени с ключом
        антибиотика ведёт словарь self._ab_key (ранее использовался хак со
        скрытыми символами в значении, который ломал карточку при открытии).
        """
        self.refresh_antibiotics_multilist()  # синхронизируем мультивыбор сравнения
        try:
            abs_list = list(self._antibiotics_getter())
        except Exception:
            abs_list = []
        all_disp = f"⭐ {AB_ALL_LABEL}"
        # имя -> ключ (без коллизий: реальное название не перезапишет спецключ)
        self._ab_key = {all_disp: AB_ALL_KEY}
        display_values = [all_disp]

        def _free_name(name):
            """Возвращает отображаемое имя, не занятое другим пунктом меню."""
            if name not in self._ab_key:
                return name
            k = 2
            while f"{name} ({k})" in self._ab_key:
                k += 1
            return f"{name} ({k})"

        for a in abs_list:
            if not a or a == AB_ALL_KEY:
                continue
            # реальный антибиотик с именем «Все антибиотики» не должен
            # сливаться со служебным пунктом — даём ему уникальную метку
            disp = _free_name(f"{a} ⭐") if a == AB_ALL_LABEL else _free_name(a)
            self._ab_key[disp] = a
            display_values.append(disp)
        self.antibiotic_menu.configure(values=display_values)
        # восстановление выбранного значения
        want_key = getattr(self, "_pending_ab", None) or \
            self._ab_key.get(self.antibiotic_menu.get(), "")
        if want_key and want_key in self._ab_key.values():
            disp = next((d for d, k in self._ab_key.items() if k == want_key),
                        all_disp)
            self.antibiotic_menu.set(disp)
            self._pending_ab = None
        elif abs_list:
            first = next((a for a in abs_list
                          if any(k == a for k in self._ab_key.values())), None)
            if first is not None:
                disp = next(d for d, k in self._ab_key.items() if k == first)
                self.antibiotic_menu.set(disp)
            else:
                self.antibiotic_menu.set(all_disp)
        else:
            self.antibiotic_menu.set(all_disp)

    def _selected_antibiotic(self):
        """Возвращает ключ выбранного антибиотика (AB_ALL_KEY для «всех»,
        '' — если выбор некорректен)."""
        sel = self.antibiotic_menu.get()
        return getattr(self, "_ab_key", {}).get(sel, "")

    def _delete(self):
        if self.on_delete:
            self.on_delete(self)

    def apply_spec(self, spec):
        """Программно применяет спецификацию к карточке (для шаблонов)."""
        self._suppress_cb_logic = True
        try:
            mode_key = spec.get("mode")
            compare = mode_key == COMPARE_CLINICS_KEY
            matrix = _mode_is_matrix(mode_key)
            self._compare_mode = compare
            self._matrix_mode = matrix
            self.mode_menu.set(_MODE_BY_KEY.get(mode_key, MODE_NORMAL_LABEL))
            self.name_var.set(spec.get("name", self.name_var.get()))
            chosen_groups = set(spec.get("compare_by" if compare else "groups", []))
            for k, v in self.group_vars.items():
                v.set(k in chosen_groups)
            chosen_metrics = set(spec.get("metrics", []))
            for k, v in self.metric_vars.items():
                v.set(k in chosen_metrics and not (compare or matrix))
            self.ab_var.set("ab" in chosen_metrics or compare or matrix)
            if "min_n" in spec:
                self.min_n_var.set(str(spec.get("min_n") or ""))
            ab_val = spec.get("antibiotic", "")
            if ab_val:
                # обратная совместимость: в старых пресетах «все антибиотики»
                # могли храниться как текстовая метка, а не служебный ключ
                self._pending_ab = AB_ALL_KEY if ab_val == AB_ALL_LABEL else ab_val
                self._refresh_antibiotics()
            if compare and spec.get("antibiotics"):
                self._set_cmp_selected(spec["antibiotics"])
            if "clinics" in spec:
                self.set_clinics(self._clinics_getter(), checked=spec.get("clinics"))
            fkey = spec.get("filter", "none")
            for key, label in FILTER_OPTIONS:
                if key == fkey:
                    self.filter_menu.set(label)
                    break
            self._apply_mode_layout()
        finally:
            self._suppress_cb_logic = False
        self._update_group_hint()

    def get_spec(self):
        groups = [k for k, v in self.group_vars.items() if v.get()]
        metrics = [k for k, v in self.metric_vars.items() if v.get()]
        if self.ab_var.get():
            metrics.append("ab")
        filter_key = "none"
        sel_label = self.filter_menu.get()
        for key, label in FILTER_OPTIONS:
            if label == sel_label:
                filter_key = key
                break
        min_n = parse_min_significant(self.min_n_var.get())
        if self._is_compare():
            return {
                "name": self.name_var.get().strip(),
                "groups": [],
                "compare_by": groups,
                "metrics": ["ab"],
                "antibiotic": self._selected_antibiotic(),
                "antibiotics": self._get_cmp_selected(),
                "clinics": self.selected_clinics(),
                "filter": filter_key,
                "mode": COMPARE_CLINICS_KEY,
                "min_n": min_n,
            }
        if self._is_matrix():
            mode_key = (MATRIX_AB_ONLY_KEY
                        if self.mode_menu.get() == MATRIX_AB_ONLY_LABEL
                        else MATRIX_AB_ORG_KEY)
            return {
                "name": self.name_var.get().strip(),
                "groups": groups,       # подсказка билдера; реально строки
                                        # матрицы задаются режимом
                "matrix_mode": mode_key,
                "mode": mode_key,
                "metrics": ["ab"],
                "antibiotics": self._get_cmp_selected(),
                "filter": filter_key,
                "min_n": min_n,
            }
        return {
            "name": self.name_var.get().strip(),
            "groups": groups,
            "metrics": metrics,
            "antibiotic": self._selected_antibiotic() if self.ab_var.get() else "",
            "filter": filter_key,
            "min_n": min_n,
        }


class SheetBuilderDialog(ctk.CTkToplevel):
    """Модальное окно конструктора дополнительных листов."""

    def __init__(self, master, antibiotics_getter=None, clinics_getter=None):
        super().__init__(master)
        self.title("🧩 Конструктор листов отчёта")
        self.geometry("1100x680")
        self.minsize(940, 540)
        self.transient(master)
        self.grab_set()

        self.specs = []
        self._antibiotics_getter = antibiotics_getter or (lambda: [])
        self._clinics_getter = clinics_getter or (lambda: [])
        self._cards = []

        # Шапка
        head = ctk.CTkFrame(self, fg_color=("gray86", "gray17"), corner_radius=0)
        head.pack(fill="x")
        box = ctk.CTkFrame(head, fg_color="transparent")
        box.pack(fill="x", padx=18, pady=10)
        ctk.CTkLabel(box, text="Конструктор дополнительных листов",
                     font=ctk.CTkFont(size=17, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(box, text="Первые 3 листа формируются всегда. Здесь можно добавить "
                              "листы с любой группировкой по параметрам.",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray40", "gray65")).pack(anchor="w")

        # Пресеты и шаблоны (две строки, чтобы всё помещалось без обрезания)
        preset_bar = ctk.CTkFrame(self, fg_color="transparent")
        preset_bar.pack(fill="x", padx=18, pady=(8, 0))
        row_a = ctk.CTkFrame(preset_bar, fg_color="transparent")
        row_a.pack(fill="x")
        row_b = ctk.CTkFrame(preset_bar, fg_color="transparent")
        row_b.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(row_a, text="＋ Добавить лист", width=140, height=30,
                      fg_color=ACCENT, hover_color="#25a878",
                      command=lambda: self._add_card()).pack(side="left", padx=(0, 8))
        # Быстрый шаблон — применяет типовую сценарную настройку к последней карточке
        ctk.CTkLabel(row_a, text="Шаблон →", font=ctk.CTkFont(size=12),
                     text_color=("gray40", "gray65")).pack(side="left", padx=(0, 4))
        self.template_menu = ctk.CTkOptionMenu(
            row_a, values=[name for name, _ in SHEET_TEMPLATES], width=320,
            command=self._apply_template)
        self.template_menu.pack(side="left", padx=(0, 8))
        ctk.CTkButton(row_b, text="💾 Сохранить пресет…", width=160, height=30,
                      fg_color="transparent", border_width=1,
                      command=self._save_preset).pack(side="left", padx=(0, 8))
        ctk.CTkButton(row_b, text="📂 Загрузить пресет…", width=160, height=30,
                      fg_color="transparent", border_width=1,
                      command=self._load_preset).pack(side="left", padx=(0, 8))
        ctk.CTkButton(row_b, text="🔄 Обновить список антибиотиков", width=240, height=30,
                      fg_color="transparent", border_width=1,
                      command=self._refresh_all_antibiotics).pack(side="left")

        # Скроллируемая область карточек
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True, padx=14, pady=8)

        # Нижняя панель
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=18, pady=(0, 12))
        ctk.CTkLabel(bottom, text="ℹ️  Листы без ошибок будут добавлены после 3 базовых листов.",
                     font=ctk.CTkFont(size=11),
                     text_color=("gray45", "gray60")).pack(side="left")
        ctk.CTkButton(bottom, text="Отмена", width=100, height=34,
                      fg_color="transparent", border_width=1,
                      command=self.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(bottom, text="✔ Применить", width=130, height=34,
                      fg_color=ACCENT, hover_color="#25a878",
                      font=ctk.CTkFont(size=13, weight="bold"),
                      command=self._apply).pack(side="right")

        # стартовая карточка
        self._add_card()

    # ------------------------------------------------------------------
    def _add_card(self, spec=None):
        try:
            card = SheetCard(self.scroll, len(self._cards), spec=spec,
                             antibiotics_getter=self._antibiotics_getter,
                             clinics_getter=self._clinics_getter,
                             on_delete=self._remove_card)
            card.pack(fill="x", pady=6)
            self._cards.append(card)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать карточку листа: {e}",
                                 parent=self)

    def _refresh_all_antibiotics(self):
        """Обновляет список антибиотиков (и клиник) во всех карточках."""
        n = 0
        clinics = []
        try:
            clinics = list(self._clinics_getter())
        except Exception:
            pass
        for c in self._cards:
            try:
                c._refresh_antibiotics()
                if clinics:
                    c.set_clinics(clinics)
                n += 1
            except Exception:
                pass
        abs_list = []
        try:
            abs_list = list(self._antibiotics_getter())
        except Exception:
            pass
        if abs_list:
            messagebox.showinfo("Антибиотики",
                                f"Загружено антибиотиков: {len(abs_list)} "
                                f"(обновлено карточек: {n}).", parent=self)
        else:
            messagebox.showwarning(
                "Антибиотики",
                "Список пуст. Убедитесь, что исходный файл выбран на главном окне.",
                parent=self)

    def _remove_card(self, card):
        if card in self._cards:
            self._cards.remove(card)
        try:
            card.destroy()
        except Exception:
            pass
        # переиндексация оставшихся карточек не требуется — имена по умолчанию
        # фиксируются при создании

    def _apply_template(self, template_name):
        """Применяет выбранный шаблон к последней карточке (или создаёт новую)."""
        spec = None
        for name, s in SHEET_TEMPLATES:
            if name == template_name and s is not None:
                spec = dict(s)  # копия, чтобы не мутировать исходный шаблон
                break
        if spec is None:
            self.template_menu.set(SHEET_TEMPLATES[0][0])
            return
        try:
            if not self._cards:
                self._add_card()
            card = self._cards[-1]
            card.apply_spec(spec)
            self.template_menu.set(SHEET_TEMPLATES[0][0])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось применить шаблон: {e}",
                                 parent=self)

    def _apply(self):
        specs = [c.get_spec() for c in self._cards]
        abs_list = []
        try:
            abs_list = list(self._antibiotics_getter())
        except Exception:
            pass
        errors = []
        for i, s in enumerate(specs, start=1):
            for e in validate_spec(s, available_antibiotics=abs_list or None):
                errors.append(f"Лист {i} «{s['name'] or 'без имени'}»: {e}")
        if errors:
            messagebox.showwarning("Проверьте настройки", "\n".join(errors[:8]), parent=self)
            return
        # проверка дубликатов имён
        seen = set()
        for s in specs:
            nm = s["name"].lower()
            if nm in seen:
                messagebox.showwarning("Проверьте настройки",
                                       f"Дублируется название листа «{s['name']}».",
                                       parent=self)
                return
            seen.add(nm)
        self.specs = specs
        self.destroy()

    # ------------------------------------------------------------------
    def _save_preset(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Сохранить пресет листов",
            defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        specs = [c.get_spec() for c in self._cards]
        try:
            save_sheet_specs(specs, path)
            messagebox.showinfo("Пресет сохранён", f"Файл: {os.path.basename(path)}",
                                parent=self)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=self)

    def _load_preset(self):
        path = filedialog.askopenfilename(
            parent=self, title="Загрузить пресет листов",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            specs = load_sheet_specs(path)
        except (json.JSONDecodeError, OSError) as e:
            messagebox.showerror("Ошибка загрузки пресета", str(e), parent=self)
            return
        for c in list(self._cards):
            c.destroy()
        self._cards.clear()
        for s in specs:
            self._add_card(spec=s)
        if not self._cards:
            self._add_card()


if __name__ == "__main__":
    root = ctk.CTk()
    root.withdraw()
    dlg = SheetBuilderDialog(root, antibiotics_getter=lambda: ["Ампициллин", "Цефтриаксон"])
    dlg.wait_window()
    print(json.dumps(dlg.specs, ensure_ascii=False, indent=2))
