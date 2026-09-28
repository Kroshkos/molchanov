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
    AB_ALL_KEY, AB_ALL_LABEL,
)

ACCENT = "#2CC98F"


class SheetCard(ctk.CTkFrame):
    """Карточка одного листа конструктора."""

    def __init__(self, master, index, spec=None, antibiotics_getter=None, on_delete=None):
        super().__init__(master, fg_color=("gray84", "gray22"), corner_radius=10)
        self.index = index
        self.on_delete = on_delete
        self._antibiotics_getter = antibiotics_getter or (lambda: [])
        self._suppress_cb_logic = False  # защита от рекурсии при программной установке

        spec = spec or {}
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

        # Название листа
        ctk.CTkLabel(body, text="Название:", font=ctk.CTkFont(size=12)).grid(
            row=0, column=0, sticky="w", pady=4, padx=(0, 8))
        self.name_var = ctk.StringVar(value=spec.get("name", f"Лист {index + 1}"))
        ctk.CTkEntry(body, textvariable=self.name_var, width=220).grid(
            row=0, column=1, columnspan=3, sticky="ew", pady=4)

        # Группировка (несколько чекбоксов; «Всё вместе» — эксклюзивный режим)
        ctk.CTkLabel(body, text="Группировать по:\n(можно выбирать несколько)",
                     font=ctk.CTkFont(size=12), justify="left").grid(
            row=1, column=0, sticky="nw", pady=(8, 4), padx=(0, 8))
        grp_frame = ctk.CTkFrame(body, fg_color="transparent")
        grp_frame.grid(row=1, column=1, columnspan=3, sticky="w", pady=(8, 2))
        self.group_vars = {}
        chosen_groups = set(spec.get("groups", []))
        if not spec:
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
        ctk.CTkLabel(body, text="Метрики:", font=ctk.CTkFont(size=12)).grid(
            row=2, column=0, sticky="nw", pady=(8, 4), padx=(0, 8))
        met_frame = ctk.CTkFrame(body, fg_color="transparent")
        met_frame.grid(row=2, column=1, columnspan=3, sticky="w", pady=(8, 2))
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
        self.ab_var = ctk.BooleanVar(value="ab" in chosen_metrics)
        ab_row = (len(METRIC_OPTIONS) // 2)
        ctk.CTkCheckBox(met_frame, text="Антибиотикограмма (S/I/R/НД)",
                        variable=self.ab_var, font=ctk.CTkFont(size=12)).grid(
            row=ab_row, column=0, columnspan=2, sticky="w", pady=2)

        # Антибиотик + фильтр (одна строка, элементы растягиваются по ширине)
        row3 = ctk.CTkFrame(body, fg_color="transparent")
        row3.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        row3.columnconfigure(1, weight=1)
        row3.columnconfigure(3, weight=1)
        ctk.CTkLabel(row3, text="Антибиотик:", font=ctk.CTkFont(size=12)).grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        self.antibiotic_menu = ctk.CTkOptionMenu(row3, values=["—"])
        self.antibiotic_menu.grid(row=0, column=1, sticky="ew", padx=(0, 18))
        ctk.CTkLabel(row3, text="Фильтр:", font=ctk.CTkFont(size=12)).grid(
            row=0, column=2, sticky="w", padx=(0, 6))
        self.filter_menu = ctk.CTkOptionMenu(
            row3, values=[label for _, label in FILTER_OPTIONS])
        self.filter_menu.grid(row=0, column=3, sticky="ew")

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
        # В Tk режим трассировки называется "write" (а не "changed") — иначе
        # создание карточки падало с TclError и кнопка «Добавить лист» не работала.
        self.ab_var.trace_add("write", lambda *_: self._refresh_antibiotics())

    # ------------------------------------------------------------------
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
        if not chosen:
            text = "ℹ️ Группировка не выбрана — лист будет построен как «всё вместе» (итог)."
        elif TOTAL_GROUP_KEY in chosen:
            text = "📊 Режим «всё вместе»: одна итоговая строка по всей выборке."
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
            self.name_var.set(spec.get("name", self.name_var.get()))
            chosen_groups = set(spec.get("groups", []))
            for k, v in self.group_vars.items():
                v.set(k in chosen_groups)
            chosen_metrics = set(spec.get("metrics", []))
            for k, v in self.metric_vars.items():
                v.set(k in chosen_metrics)
            self.ab_var.set("ab" in chosen_metrics)
            ab_val = spec.get("antibiotic", "")
            if ab_val:
                # обратная совместимость: в старых пресетах «все антибиотики»
                # могли храниться как текстовая метка, а не служебный ключ
                self._pending_ab = AB_ALL_KEY if ab_val == AB_ALL_LABEL else ab_val
                self._refresh_antibiotics()
            fkey = spec.get("filter", "none")
            for key, label in FILTER_OPTIONS:
                if key == fkey:
                    self.filter_menu.set(label)
                    break
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
        return {
            "name": self.name_var.get().strip(),
            "groups": groups,
            "metrics": metrics,
            "antibiotic": self._selected_antibiotic() if self.ab_var.get() else "",
            "filter": filter_key,
        }


class SheetBuilderDialog(ctk.CTkToplevel):
    """Модальное окно конструктора дополнительных листов."""

    def __init__(self, master, antibiotics_getter=None):
        super().__init__(master)
        self.title("🧩 Конструктор листов отчёта")
        self.geometry("1100x680")
        self.minsize(940, 540)
        self.transient(master)
        self.grab_set()

        self.specs = []
        self._antibiotics_getter = antibiotics_getter or (lambda: [])
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
                             on_delete=self._remove_card)
            card.pack(fill="x", pady=6)
            self._cards.append(card)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать карточку листа: {e}",
                                 parent=self)

    def _refresh_all_antibiotics(self):
        """Обновляет список антибиотиков во всех карточках."""
        n = 0
        for c in self._cards:
            try:
                c._refresh_antibiotics()
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
