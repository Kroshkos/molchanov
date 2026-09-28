import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
import threading
import re
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 1. Загрузка и предобработка данных (расширенная)
# ----------------------------------------------------------------------

def load_data(filepath, log_func=None):
    if log_func:
        log_func("Чтение файла...")
    df_raw = pd.read_excel(filepath, sheet_name='TDSheet', header=None)
    
    # Находим строку с "Заказ"
    header_row_idx = None
    for i, row in df_raw.iterrows():
        if pd.notna(row[0]) and str(row[0]).strip() == 'Заказ':
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Не найдена строка с заголовком 'Заказ'")
    
    # Основные заголовки берём из строки с "Заказ"
    headers = df_raw.iloc[header_row_idx].values
    data = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = headers
    
    # Переименовываем основные колонки
    rename_map = {}
    for col in data.columns:
        col_str = str(col).strip()
        if col_str in ('История болезни', 'patient_id'):
            rename_map[col] = 'patient_id'
        elif col_str in ('Дата', 'date'):
            rename_map[col] = 'date'
        elif col_str in ('Материал', 'material_raw'):
            rename_map[col] = 'material_raw'
        elif col_str in ('Микроорганизм', 'organism'):
            rename_map[col] = 'organism'
        elif col_str in ('ФИО', 'fio'):
            rename_map[col] = 'fio'
    data = data.rename(columns=rename_map)
    
    # Поиск столбца с микроорганизмом
    if 'organism' not in data.columns:
        for col in data.columns:
            if 'микро' in str(col).lower() or 'organism' in str(col).lower():
                data = data.rename(columns={col: 'organism'})
                break
        else:
            raise ValueError("Не найден столбец с микроорганизмом")
    
    # Поиск столбца с датой
    if 'date' not in data.columns:
        for col in data.columns:
            if 'дат' in str(col).lower():
                data = data.rename(columns={col: 'date'})
                break
        else:
            raise ValueError("Не найден столбец с датой")
    
    # Поиск столбца с материалом
    if 'material_raw' not in data.columns:
        for col in data.columns:
            if 'матер' in str(col).lower():
                data = data.rename(columns={col: 'material_raw'})
                break
        else:
            raise ValueError("Не найден столбец с материалом")
    
    data['date'] = pd.to_datetime(data['date'], errors='coerce', dayfirst=True)
    
    # Находим позицию столбца 'organism'
    organism_pos = None
    for i, col in enumerate(data.columns):
        if col == 'organism':
            organism_pos = i
            break
    if organism_pos is None:
        raise ValueError("Не удалось определить позицию столбца 'organism'")
    
    # Строка с названиями антибиотиков (индекс 6, как в исходном коде)
    ab_name_row = 6
    if log_func:
        log_func(f"Используем строку {ab_name_row} для названий антибиотиков")
        sample_cells = []
        for j in range(organism_pos+1, min(organism_pos+11, len(df_raw.columns))):
            val = df_raw.iloc[ab_name_row, j] if ab_name_row < len(df_raw) else ''
            sample_cells.append(str(val) if pd.notna(val) else 'NaN')
        log_func(f"Столбцы {organism_pos+2}-{organism_pos+11} на строке {ab_name_row}: {sample_cells}")
    
    antibiotics = []
    ab_indices = []
    for j in range(organism_pos + 1, len(data.columns), 2):
        if j+1 >= len(data.columns):
            break
        ab_name_raw = df_raw.iloc[ab_name_row, j] if ab_name_row < len(df_raw) else ''
        ab_name = str(ab_name_raw).strip()
        if not ab_name or ab_name.lower() in ('nan', 'none', ''):
            ab_name_raw = df_raw.iloc[ab_name_row, j+1] if ab_name_row < len(df_raw) else ''
            ab_name = str(ab_name_raw).strip()
        if ab_name and ab_name.lower() not in ('nan', 'none', ''):
            ab_name_clean = ab_name
            if (ab_name_clean and not ab_name_clean.isdigit() 
                and ab_name_clean not in ('RSI', 'S', 'R', 'I', 'НД', 'Unnamed', 'ФИО', 'material_raw')):
                antibiotics.append(ab_name_clean)
                ab_indices.append(j+1)  # индекс RSI (правый столбец)
    
    if log_func:
        log_func(f"Найдено антибиотиков: {len(antibiotics)}. Примеры: {antibiotics[:10]}")
        if len(antibiotics) == 0:
            log_func("ВНИМАНИЕ: Не найдено ни одного антибиотика. Проверьте правильность строки с названиями.")
    
    # Функция диагностичности
    def is_diagnostic(row):
        if pd.isna(row['organism']) or str(row['organism']).strip().lower() in ('', 'роста нет', 'нет роста'):
            return False
        for idx in ab_indices:
            val = row.iloc[idx]
            if pd.notna(val) and isinstance(val, str):
                if val.strip().upper() in ('S', 'R', 'I'):
                    return True
        return False
    
    data['diagnostic'] = data.apply(is_diagnostic, axis=1)
    data['is_positive'] = data['organism'].apply(
        lambda x: False if pd.isna(x) or str(x).strip().lower() in ('', 'роста нет', 'нет роста') else True
    )
    
    data = data.dropna(subset=['date', 'material_raw'])
    data['material_raw'] = data['material_raw'].astype(str).str.strip()
    data['patient_id'] = data['patient_id'].astype(str).str.strip()
    
    df_main = data[['patient_id', 'date', 'material_raw', 'organism', 'is_positive', 'diagnostic']].copy()
    
    # Сбор данных по антибиотикам для всех образцов (группировка только по микроорганизму)
    ab_data = compute_ab_data(data, antibiotics, ab_indices, group_cols=['organism'], log_func=log_func)
    
    if log_func:
        log_func(f"Загружено {len(df_main)} строк, антибиотиков с данными: {len(ab_data)}")
    
    # Возвращаем также полный data, список антибиотиков и индексы для дальнейшего использования
    return df_main, ab_data, data, antibiotics, ab_indices


def compute_ab_data(df, antibiotics, ab_indices, group_cols=None, log_func=None):
    """
    Вычисляет статистику по антибиотикам для заданного DataFrame.
    df должен содержать столбец 'organism' и столбцы с RSI-значениями по индексам ab_indices.
    Параметр group_cols задаёт список столбцов для группировки (по умолчанию ['organism']).
    Возвращает словарь ab_data {антибиотик: DataFrame с подсчётами}.
    """
    if group_cols is None:
        group_cols = ['organism']
    
    ab_records = {ab: [] for ab in antibiotics}
    for idx, row in df[df['is_positive']].iterrows():
        org = row['organism']
        if pd.isna(org):
            continue
        # Формируем ключ группировки
        group_key = tuple(row[col] for col in group_cols)
        for i, ab in enumerate(antibiotics):
            rsi_val = row.iloc[ab_indices[i]]
            if pd.notna(rsi_val) and isinstance(rsi_val, str):
                rsi = rsi_val.strip().upper()
                if rsi in ('S', 'I', 'R', 'НД'):
                    # Сохраняем словарь с группирующими колонками и результатом
                    record = {col: row[col] for col in group_cols}
                    record['result'] = rsi
                    ab_records[ab].append(record)
    
    ab_data = {}
    for ab, records in ab_records.items():
        if records:
            df_ab = pd.DataFrame(records)
            # Группируем по group_cols и 'result'
            counts = df_ab.groupby(group_cols + ['result']).size().unstack(fill_value=0)
            # Убедимся, что все столбцы присутствуют
            for col in ['S', 'I', 'R', 'НД']:
                if col not in counts.columns:
                    counts[col] = 0
            counts = counts[['S', 'I', 'R', 'НД']]
            counts['Total'] = counts.sum(axis=1)
            for col in ['S', 'I', 'R', 'НД']:
                counts[f'{col}_%'] = counts[col] / counts['Total'] * 100
            ab_data[ab] = counts
        else:
            ab_data[ab] = pd.DataFrame()
    return ab_data


# ----------------------------------------------------------------------
# 2. Категоризация материалов для листов 2 и 3
# ----------------------------------------------------------------------

def categorize_material(material):
    material_lower = material.lower()
    if 'моча' in material_lower:
        return 'моча', 'моча'
    elif 'мокрота' in material_lower or 'аспират' in material_lower or 'бронхо' in material_lower:
        return 'мокрота и т.д.', 'мокрота'
    elif 'кровь' in material_lower:
        return 'кровь', 'катетер и кровь'
    elif 'катетер' in material_lower:
        return 'катетер', 'катетер и кровь'
    elif 'ликвор' in material_lower or 'цереброспинальная' in material_lower:
        return 'жидкость цереброспинальная', 'церебр'
    elif any(x in material_lower for x in ['рана', 'язва', 'соскоб', 'гной', 'абсцесс', 'отделяемое', 'эрозий', 'раневое', 'содержимое', 'пунктат']):
        return 'РАНЫ', 'раны'
    else:
        return 'другие', 'другие'

# ----------------------------------------------------------------------
# 3. Присвоение порядкового номера забора для каждого пациента и материала
# ----------------------------------------------------------------------

def assign_order(df):
    df = df.sort_values(['patient_id', 'material_raw', 'date']).reset_index(drop=True)
    order = []
    for (pid, mat), group in df.groupby(['patient_id', 'material_raw']):
        group = group.sort_values('date')
        orders = range(1, len(group)+1)
        order.extend(orders)
    df['order'] = order
    return df

# ----------------------------------------------------------------------
# 4. Вспомогательная функция подсчёта статистики
# ----------------------------------------------------------------------

def compute_stats(subdf):
    if subdf.empty:
        return (0,0,0,0,0.0,0,0.0)
    patients = subdf['patient_id'].nunique()
    total = len(subdf)
    no_growth = (~subdf['is_positive']).sum()
    pos = subdf['is_positive'].sum()
    pct_pos = (pos / total * 100) if total > 0 else 0.0
    diag = subdf[subdf['is_positive'] & subdf['diagnostic']].shape[0]
    pct_diag = (diag / pos * 100) if pos > 0 else 0.0
    return patients, total, no_growth, pos, pct_pos, diag, pct_diag

# ----------------------------------------------------------------------
# 5. Формирование листа 2 (Анализ результатов АБЧ общий)
# ----------------------------------------------------------------------

def generate_sheet2(df, log_func=None):
    df['cat2'] = df['material_raw'].apply(lambda x: categorize_material(x)[0])
    categories = ['всего', 'моча', 'мокрота и т.д.', 'кровь', 'катетер',
                  'жидкость цереброспинальная', 'РАНЫ', 'другие']
    rows = []
    for cat in categories:
        if log_func:
            log_func(f"Обработка категории: {cat}")
        if cat == 'всего':
            subdf = df
        else:
            subdf = df[df['cat2'] == cat]
        pat1, tot1, nog1, pos1, pct1, diag1, pctd1 = compute_stats(subdf)
        temp = subdf.copy()
        temp = assign_order(temp)
        subdf_ge2 = temp[temp['order'] >= 2]
        pat2, tot2, nog2, pos2, pct2, diag2, pctd2 = compute_stats(subdf_ge2)
        subdf_ge3 = temp[temp['order'] >= 3]
        pat3, tot3, nog3, pos3, pct3, diag3, pctd3 = compute_stats(subdf_ge3)
        rows.append({
            'Категория': cat,
            'кол-во больных': pat1, 'всего проб': tot1, 'роста нет': nog1, '+': pos1,
            '%': pct1, 'в диагностическом титре': diag1, '%_диаг': pctd1,
            'кол-во больных_2': pat2, 'всего проб_2': tot2, 'роста нет_2': nog2, '+_2': pos2,
            '%_2': pct2, 'в диагностическом титре_2': diag2, '%_диаг_2': pctd2,
            'кол-во больных_3': pat3, 'всего проб_3': tot3, 'роста нет_3': nog3, '+_3': pos3,
            '%_3': pct3, 'в диагностическом титре_3': diag3, '%_диаг_3': pctd3,
        })
    result_df = pd.DataFrame(rows)
    result_df = result_df.rename(columns={
        '%_диаг': '% (в диагн. титре)',
        '%_диаг_2': '% (в диагн. титре)',
        '%_диаг_3': '% (в диагн. титре)',
        '%': '% (+)', '%_2': '% (+)', '%_3': '% (+)'
    })
    return result_df

# ----------------------------------------------------------------------
# 6. Формирование листа 3 (Эпидемиология по локализациям)
# ----------------------------------------------------------------------

def generate_sheet3(df, log_func=None):
    df['cat3'] = df['material_raw'].apply(lambda x: categorize_material(x)[1])
    pos_df = df[df['is_positive']].copy()
    if pos_df.empty:
        if log_func:
            log_func("Нет положительных проб для листа 3.")
        return {}
    pos_df = assign_order(pos_df)
    organisms = pos_df['organism'].dropna().unique()
    categories3 = ['моча', 'мокрота', 'другие', 'церебр', 'катетер и кровь', 'раны']
    
    def build_table_for_category(cat):
        if cat == 'ВСЕ':
            subdf = pos_df
        else:
            subdf = pos_df[pos_df['cat3'] == cat]
        if subdf.empty:
            return None
        data = []
        for org in sorted(organisms):
            org_df = subdf[subdf['organism'] == org]
            if org_df.empty:
                continue
            row = {'Микроорганизм': org}
            for order in [1,2,3]:
                ord_df = org_df[org_df['order'] == order]
                total = len(ord_df)
                if total == 0:
                    row[f'diag_{order}'] = 0
                    row[f'nediag_{order}'] = 0
                    row[f'total_{order}'] = 0
                    row[f'pct_{order}'] = 0.0
                else:
                    diag = ord_df['diagnostic'].sum()
                    non_diag = total - diag
                    pct = (diag / total * 100) if total > 0 else 0.0
                    row[f'diag_{order}'] = diag
                    row[f'nediag_{order}'] = non_diag
                    row[f'total_{order}'] = total
                    row[f'pct_{order}'] = pct
            data.append(row)
        total_row = {'Микроорганизм': 'ВСЕ'}
        for order in [1,2,3]:
            total_diag = subdf[subdf['order'] == order]['diagnostic'].sum()
            total_all = subdf[subdf['order'] == order].shape[0]
            total_non = total_all - total_diag
            pct = (total_diag / total_all * 100) if total_all > 0 else 0.0
            total_row[f'diag_{order}'] = total_diag
            total_row[f'nediag_{order}'] = total_non
            total_row[f'total_{order}'] = total_all
            total_row[f'pct_{order}'] = pct
        data.append(total_row)
        df_cat = pd.DataFrame(data)
        col_map = {}
        for order in [1,2,3]:
            col_map[f'diag_{order}'] = f'диаг_{order}'
            col_map[f'nediag_{order}'] = f'недиаг_{order}'
            col_map[f'total_{order}'] = f'итого_{order}'
            col_map[f'pct_{order}'] = f'% диагн_{order}'
        df_cat = df_cat.rename(columns=col_map)
        return df_cat
    
    result = {}
    for cat in categories3:
        if log_func:
            log_func(f"Обработка категории листа 3: {cat}")
        tbl = build_table_for_category(cat)
        if tbl is not None:
            result[cat] = tbl
    total_tbl = build_table_for_category('ВСЕ')
    if total_tbl is not None:
        result['ВСЕ'] = total_tbl
    return result

# ----------------------------------------------------------------------
# 7. Формирование листа 4 (Антибиотикорезистентность) – общая
#    и листа 5 (раневые антибиотики с локализацией)
# ----------------------------------------------------------------------

def generate_antibiotic_sheet(ab_data, log_func=None):
    """
    Создаёт DataFrame для листа с антибиотикорезистентностью.
    Для каждого антибиотика выводит таблицу с результатами.
    Если индекс DataFrame содержит несколько уровней (например, organism и material_raw),
    то будет добавлена колонка "Локализация".
    """
    if not ab_data:
        if log_func:
            log_func("Нет данных по антибиотикам для листа")
        return pd.DataFrame()
    
    all_rows = []
    for antibiotic, df_ab in ab_data.items():
        if df_ab.empty:
            continue
        # Определяем, является ли индекс MultiIndex
        is_multi = isinstance(df_ab.index, pd.MultiIndex)
        # Заголовок антибиотика
        all_rows.append([f"Антибиотик: {antibiotic}"])
        # Заголовки столбцов
        if is_multi:
            header = ['Микроорганизм', 'Локализация', 'S', 'I', 'R', 'НД', 'Total', 'S%', 'I%', 'R%', 'НД%']
        else:
            header = ['Микроорганизм', 'S', 'I', 'R', 'НД', 'Total', 'S%', 'I%', 'R%', 'НД%']
        all_rows.append(header)
        
        # Обходим строки
        for idx, row in df_ab.iterrows():
            if is_multi:
                # idx - кортеж (organism, material_raw)
                organism, location = idx
                row_data = [
                    organism,
                    location,
                    row['S'], row['I'], row['R'], row['НД'],
                    row['Total'],
                    round(row['S_%'], 1), round(row['I_%'], 1),
                    round(row['R_%'], 1), round(row['НД_%'], 1)
                ]
            else:
                organism = idx
                row_data = [
                    organism,
                    row['S'], row['I'], row['R'], row['НД'],
                    row['Total'],
                    round(row['S_%'], 1), round(row['I_%'], 1),
                    round(row['R_%'], 1), round(row['НД_%'], 1)
                ]
            all_rows.append(row_data)
        all_rows.append([])  # пустая строка после таблицы
    # Убираем последнюю пустую строку
    if all_rows and all_rows[-1] == []:
        all_rows.pop()
    # Преобразуем в DataFrame
    df_result = pd.DataFrame(all_rows)
    return df_result

# (для обратной совместимости оставляем псевдоним)
generate_sheet4 = generate_antibiotic_sheet

# ----------------------------------------------------------------------
# 8. Сохранение в Excel (листы 2,3,4,5) с автонастройкой ширины столбцов
# ----------------------------------------------------------------------

def save_to_excel(sheet2_df, sheet3_dict, sheet4_df, sheet5_df, output_file, log_func=None):
    if log_func:
        log_func(f"Сохранение в {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sheet2_df.to_excel(writer, sheet_name='2', index=False)
        
        # Лист 3
        sheet3_rows = []
        for section_name, section_df in sheet3_dict.items():
            sheet3_rows.append([section_name] + [''] * (len(section_df.columns)-1))
            sheet3_rows.append(section_df.columns.tolist())
            for _, row in section_df.iterrows():
                sheet3_rows.append(row.tolist())
            sheet3_rows.append([''] * len(section_df.columns))
        if sheet3_rows:
            max_cols = max(len(row) for row in sheet3_rows)
            for row in sheet3_rows:
                if len(row) < max_cols:
                    row.extend([''] * (max_cols - len(row)))
            sheet3_combined = pd.DataFrame(sheet3_rows)
            sheet3_combined.to_excel(writer, sheet_name='3', index=False, header=False)
        
        # Лист 4
        if not sheet4_df.empty:
            sheet4_df.to_excel(writer, sheet_name='4', index=False, header=False)
        
        # Лист 5 (раневые антибиотики с локализацией)
        if not sheet5_df.empty:
            sheet5_df.to_excel(writer, sheet_name='5', index=False, header=False)
        
        # Автонастройка ширины столбцов
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
    if log_func:
        log_func("Сохранение завершено.")

# ----------------------------------------------------------------------
# 9. Графический интерфейс
# ----------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("Обработка микробиологических данных")
        root.geometry("700x500")
        
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        frame_input = tk.LabelFrame(root, text="Исходный файл", padx=5, pady=5)
        frame_input.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_input, textvariable=self.input_file, width=60).pack(side="left", padx=5)
        tk.Button(frame_input, text="Обзор...", command=self.select_input).pack(side="left")
        
        frame_output = tk.LabelFrame(root, text="Результирующий файл", padx=5, pady=5)
        frame_output.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_output, textvariable=self.output_file, width=60).pack(side="left", padx=5)
        tk.Button(frame_output, text="Сохранить как...", command=self.select_output).pack(side="left")
        
        self.run_btn = tk.Button(root, text="Запустить обработку", command=self.run_processing, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.run_btn.pack(pady=10)
        
        self.log_area = scrolledtext.ScrolledText(root, height=15, width=80, state="normal")
        self.log_area.pack(padx=10, pady=5, fill="both", expand=True)
        
        tk.Button(root, text="Выход", command=root.quit, bg="#f44336", fg="white").pack(pady=5)
    
    def select_input(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            if not self.output_file.get():
                import os
                dirname = os.path.dirname(filename)
                basename = os.path.basename(filename)
                name, ext = os.path.splitext(basename)
                default_out = os.path.join(dirname, f"{name}_обработанный.xlsx")
                self.output_file.set(default_out)
    
    def select_output(self):
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def log(self, message):
        self.log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()
    
    def run_processing(self):
        input_path = self.input_file.get().strip()
        output_path = self.output_file.get().strip()
        if not input_path:
            messagebox.showerror("Ошибка", "Выберите исходный файл.")
            return
        if not output_path:
            messagebox.showerror("Ошибка", "Укажите имя результирующего файла.")
            return
        
        self.run_btn.config(state="disabled")
        self.log_area.delete(1.0, tk.END)
        self.log("Начало обработки...")
        
        def task():
            try:
                # Загружаем данные, получаем всё необходимое
                df_main, ab_data, data, antibiotics, ab_indices = load_data(input_path, log_func=self.log)
                self.log("Формирование листа 2...")
                sheet2 = generate_sheet2(df_main, log_func=self.log)
                self.log("Формирование листа 3...")
                sheet3 = generate_sheet3(df_main, log_func=self.log)
                self.log("Формирование листа 4 (общая антибиотикорезистентность)...")
                sheet4 = generate_antibiotic_sheet(ab_data, log_func=self.log)
                
                # --- Дополнительно: лист 5 для раневых антибиотиков с локализацией ---
                self.log("Формирование листа 5 (раневые антибиотики с локализацией)...")
                data['cat_wound'] = data['material_raw'].apply(lambda x: categorize_material(x)[0])
                wound_data = data[data['cat_wound'] == 'РАНЫ'].copy()
                if wound_data.empty:
                    self.log("Нет раневых материалов, лист 5 будет пуст.")
                    sheet5 = pd.DataFrame()
                else:
                    # Группируем по микроорганизму и материалу
                    ab_data_wound = compute_ab_data(
                        wound_data, antibiotics, ab_indices,
                        group_cols=['organism', 'material_raw'],
                        log_func=self.log
                    )
                    sheet5 = generate_antibiotic_sheet(ab_data_wound, log_func=self.log)
                
                save_to_excel(sheet2, sheet3, sheet4, sheet5, output_path, log_func=self.log)
                self.log("Обработка успешно завершена!")
                messagebox.showinfo("Готово", f"Результат сохранён в:\n{output_path}")
            except Exception as e:
                self.log(f"ОШИБКА: {str(e)}")
                messagebox.showerror("Ошибка", f"При обработке произошла ошибка:\n{str(e)}")
            finally:
                self.root.after(0, lambda: self.run_btn.config(state="normal"))
        
        threading.Thread(target=task, daemon=True).start()

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()