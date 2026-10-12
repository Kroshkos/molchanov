# -*- coding: utf-8 -*-
"""
app_core.py — ядро обработки микробиологических данных.

Модуль не зависит от графического интерфейса: здесь загрузка и предобработка
данных из Excel, расчёт статистики по антибиотикограммам, формирование
аналитических листов и сохранение результата (в т.ч. с современным
оформлением через openpyxl).
"""
import os
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

#: Специальное значение выбора клиники: обработать все клиники, сохранив
#: результат для каждой клиники в отдельном файле.
ALL_CLINICS_LABEL = 'Все клиники'

#: Метка для строк исходного файла, в котором клиники не различаются
#: (одна выборка / файл без разделения на клиники).
DEFAULT_CLINIC_LABEL = 'Общая выборка'


def detect_clinics(filepath, log_func=None):
    """Определяет список клиник в исходном файле.

    Возвращает отсортированный список кодов/названий клиник. Если данных
    несколько и признаков клиник нет — возвращается пустой список (файл
    трактуется как одна выборка). Значение «Все клиники» в этот список не
    включается — оно является специальным пунктом выбора в интерфейсе.
    """
    _, _, data, _, _ = load_data(filepath, log_func=log_func)
    if 'clinic' not in data.columns:
        return []
    found = sorted({str(c).strip() for c in data['clinic'].dropna()
                    if str(c).strip() and str(c).strip() != DEFAULT_CLINIC_LABEL})
    return found

# ----------------------------------------------------------------------
# 1. Загрузка и предобработка данных (расширенная)
# ----------------------------------------------------------------------

# Ключевые слова, по которым распознаём служебные столбцы (без жёстких названий)
_HEADER_KEYWORDS = ('заказ', 'история болезни', 'дата', 'материал', 'фио',
                    'patient', 'date', 'material', 'fio')

# Варианты названий ключевых столбцов (синонимы, регистронезависимо)
_ORGANISM_KEYS = ('микро', 'organism', 'возбудитель', 'идентификац', 'microbiolog',
                  'species', 'специес')
_DATE_KEYS = ('дат', 'date', 'число', 'дата посева', 'поступил')
_MATERIAL_KEYS = ('матер', 'биоматеріал', 'локализ', 'object', 'bid', 'вид материала',
                  'материал')
_PATIENT_KEYS = ('история', 'patient', 'номер', 'карта', 'закз', 'id', '№', 'п/п', 'фио', 'fio')


def _matches(col_norm, keys):
    return any(k in col_norm for k in keys)


def _norm(s):
    """Нормализация строки для нечувствительного сравнения."""
    return str(s).strip().lower().replace('\xa0', ' ')


def pick_sheet(xl, log_func=None):
    """Выбирает лист: сначала «TDSheet», затем первый лист с распознанными данными."""
    names = xl.sheet_names
    candidates = []
    for name in names:
        if _norm(name) == 'tdsheet':
            candidates.insert(0, name)
        else:
            candidates.append(name)

    for name in candidates:
        df_try = xl.parse(sheet_name=name, header=None)
        hdr = find_header_row(df_try)
        if hdr is not None:
            if log_func and name != candidates[0]:
                log_func(f"Лист '{name}': используется как источник данных "
                         f"(строка заголовков — {hdr + 1}).")
            return df_try, hdr
        if log_func:
            log_func(f"Лист '{name}' просмотрен — знакомых заголовков не найдено, пробуем следующий…")
    raise ValueError(
        "Не удалось распознать данные ни на одном листе. "
        "Убедитесь, что в файле есть таблица с колонками вроде «Дата», «Материал», «Микроорганизм»."
    )


def find_header_row(df_raw):
    """Ищет строку-заголовок без жёстких требований к оформлению.

    Подходит строка, первая ячейка которой содержит «Заказ»/«№»/«ID», либо
    строка с двумя и более известными заголовками («дата», «материал», «ФИО»…).
    Дополнительно строка признаётся заголовком, если в ней есть столбец про
    микроорганизм и ещё один распознанный столбец — это позволяет открывать
    файлы с нестандартными названиями колонок.
    """
    best_idx, best_score = None, 0
    limit = min(len(df_raw), 60)
    for i in range(limit):
        vals = [_norm(v) for v in df_raw.iloc[i].tolist() if pd.notna(v)]
        if len(vals) < 2:
            continue
        first = vals[0]
        score = 0
        if first.startswith('заказ') or first in ('№', 'no', 'no.', 'id'):
            score += 2
        matched = set()
        has_organism = False
        other_hits = 0
        for v in vals:
            if _matches(v, _ORGANISM_KEYS):
                has_organism = True
            for kw in _HEADER_KEYWORDS:
                if kw in v and kw not in matched:
                    matched.add(kw)
                    score += 1
                    break
            else:
                if (_matches(v, _DATE_KEYS) or _matches(v, _MATERIAL_KEYS)
                        or _matches(v, _PATIENT_KEYS)):
                    other_hits += 1
        # Комбинированный признак: микроорганизм + любой другой служебный столбец
        if has_organism and (other_hits >= 1 or len(matched) >= 1):
            score += 3
        if score >= 2 and score > best_score:
            best_idx, best_score = i, score
    return best_idx


_AB_SERVICE_TOKENS = {'nan', 'none', 'rsi', 's', 'i', 'r', 'нд', 'nd', 's/i/r', 'si r',
                      's i r', 's/i/r/нд', 'чувствительность', 'резистентность',
                      'антибиотик', 'antibiotic', 'микроорганизм', 'микробиология',
                      'заказ', 'история', 'история болезни', 'дата', 'материал',
                      'фио', 'fio', 'patient', 'material', 'date', 'organism',
                      'номер', '№', 'id', 'прочее', 'диагностика', 'выделено',
                      'рост', 'посев', 'лаборатория', 'врач', 'отделение', 'клиника',
                      'пол', 'возраст', 'адрес', 'направление', 'примечание'}


def detect_ab_name_row(df_raw, organism_pos, header_row_idx, log_func=None):
    """Автоопределение строки с названиями антибиотиков.

    Штатное расположение — на одну строку ВЫШЕ заголовков таблицы
    (например, заголовки на 8-й строке Excel — названия антибиотиков на 7-й).
    Поэтому сначала проверяется именно строка над заголовками; если названий
    в ней не найдено, перебираются соседние строки выше и ниже заголовков.
    Служебные слова («RSI», «Дата», «ФИО» и т.п.) названиями не считаются.
    """
    ncols = df_raw.shape[1]
    start = organism_pos + 1
    if start >= ncols:
        return max(header_row_idx - 1, 0)

    def count_names(r):
        if r < 0 or r >= len(df_raw):
            return 0, []
        cnt, cols = 0, []
        for j in range(start, ncols):
            v = df_raw.iat[r, j]
            if pd.notna(v):
                s = str(v).strip()
                n = _norm(s)
                if (s and not s.isdigit() and n not in _AB_SERVICE_TOKENS
                        and '/' not in s and '\n' not in s):
                    cnt += 1
                    cols.append(j)
        return cnt, cols

    # 1) Основной кандидат — строка на 1 выше заголовков
    above = header_row_idx - 1
    cnt_above, cols_above = count_names(above)
    if cnt_above >= 2:
        if log_func:
            log_func(f"Названия антибиотиков взяты из строки №{above + 1} "
                     f"(на 1 строку выше заголовков, найдено {cnt_above} названий)")
        return above

    # 2) Запасной вариант — поиск среди соседних строк (выше приоритетнее)
    best_r, best_c = None, 0
    offsets = [-2, -3, 1, 2, 3, 4, 5]
    for off in offsets:
        r = header_row_idx + off
        c, _ = count_names(r)
        if c > best_c:
            best_r, best_c = r, c
    if best_c >= 2 and best_r is not None:
        if log_func:
            log_func(f"Строка с названиями антибиотиков определена автоматически: "
                     f"№{best_r + 1} (найдено {best_c} названий)")
        return best_r

    if log_func:
        log_func("Не удалось уверенно найти строку с названиями антибиотиков — "
                 "используется строка на 1 выше заголовков.")
    return max(above, 0)


def load_data(filepath, log_func=None):
    if log_func:
        log_func("Чтение файла…")
    try:
        xl = pd.ExcelFile(filepath)
    except ValueError as e:
        # Старый формат .xls под Windows может определяться как ZIP (xlsx)
        msg = str(e)
        if 'zip' in msg.lower():
            try:
                xl = pd.ExcelFile(filepath, engine='openpyxl')
            except Exception:
                xl = pd.ExcelFile(filepath, engine='xlrd')
        else:
            raise

    # Лист и строка заголовков определяются автоматически —
    # требования к имени листа и расположению таблицы смягчены
    df_raw, header_row_idx = pick_sheet(xl, log_func=log_func)
    if header_row_idx is None:
        raise ValueError(
            "Не найдена строка с заголовками таблицы. Ожидаются колонки вида "
            "«Заказ / История болезни», «Дата», «Материал», «Микроорганизм»."
        )
    if log_func:
        log_func(f"Строка заголовков найдена: №{header_row_idx + 1}")

    # Основные заголовки берём из найденной строки
    headers = df_raw.iloc[header_row_idx].values
    data = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = headers

    # Переименовываем основные колонки (в т.ч. по синонимам — без точного совпадения)
    rename_map = {}
    taken = set()
    for col in data.columns:
        col_str = str(col).strip()
        c_low = _norm(col_str)
        if col_str in ('История болезни', 'patient_id'):
            rename_map[col] = 'patient_id'
            taken.add('patient_id')
        elif col_str in ('Дата', 'date'):
            rename_map[col] = 'date'
            taken.add('date')
        elif col_str in ('Материал', 'material_raw'):
            rename_map[col] = 'material_raw'
            taken.add('material_raw')
        elif col_str in ('Микроорганизм', 'organism'):
            rename_map[col] = 'organism'
            taken.add('organism')
        elif col_str in ('ФИО', 'fio'):
            rename_map[col] = 'fio'
        else:
            # мягкое сопоставление по ключевым словам (первое подходящее имя выигрывает)
            if 'organism' not in taken and _matches(c_low, _ORGANISM_KEYS):
                rename_map[col] = 'organism'
                taken.add('organism')
            elif 'date' not in taken and _matches(c_low, _DATE_KEYS):
                rename_map[col] = 'date'
                taken.add('date')
            elif 'material_raw' not in taken and _matches(c_low, _MATERIAL_KEYS):
                rename_map[col] = 'material_raw'
                taken.add('material_raw')
            elif 'patient_id' not in taken and _matches(c_low, _PATIENT_KEYS):
                rename_map[col] = 'patient_id'
                taken.add('patient_id')
    data = data.rename(columns=rename_map)

    # Поиск столбца с микроорганизмом (расширенный, без жёстких названий)
    if 'organism' not in data.columns:
        for col in data.columns:
            c_low = _norm(col)
            if ('микро' in c_low or 'organism' in c_low or 'возбудитель' in c_low
                    or 'идентификац' in c_low or 'вид' == c_low):
                data = data.rename(columns={col: 'organism'})
                break
        else:
            raise ValueError(
                "Не найден столбец с микроорганизмом. Добавьте колонку с названием "
                "вида (например, «Микроорганизм» / «Возбудитель»)."
            )

    # Поиск столбца с датой (необязателен — при отсутствии заполняется пропуском)
    if 'date' not in data.columns:
        for col in data.columns:
            if 'дат' in _norm(col):
                data = data.rename(columns={col: 'date'})
                break
        else:
            if log_func:
                log_func("Столбец с датой не найден — статистика по датам будет недоступна.")
            data['date'] = pd.NaT

    # Поиск столбца с материалом (необязателен — при отсутствии используется «прочее»)
    if 'material_raw' not in data.columns:
        for col in data.columns:
            c_low = _norm(col)
            if 'матер' in c_low or 'биоматеріал' in c_low or 'локализ' in c_low or 'object' in c_low:
                data = data.rename(columns={col: 'material_raw'})
                break
        else:
            if log_func:
                log_func("Столбец с материалом не найден — все пробы будут отнесены к категории «другие».")
            data['material_raw'] = 'прочее'

    # Поиск столбца с пациентом (добавлена отказоустойчивость)
    if 'patient_id' not in data.columns:
        for col in data.columns:
            col_str = str(col).lower()
            if 'история' in col_str or 'patient' in col_str or 'номер' in col_str:
                data = data.rename(columns={col: 'patient_id'})
                break
        else:
            data['patient_id'] = data.index.astype(str)

    data['date'] = pd.to_datetime(data['date'], errors='coerce', dayfirst=True)

    # ---------------- Определение клиник ----------------
    # В исходном файле могут быть данные нескольких клиник. Клиника может
    # задаваться отдельным столбцом (например, «Клиника») либо указываться
    # в первой колонке вместо служебного слова «Заказ» (встречается в
    # выгрузках лабораторной системы: в строках блока клиники пишется её
    # код — «ВМХ», «ВПХ», «ГХ» и т.п.). Колонка 'clinic' добавляется всегда;
    # если клиник не найдено — все строки получают значение «Все клиники».
    clinic_col = None
    for col in data.columns:
        c_low = _norm(col)
        if c_low == 'клиника' or ('клиник' in c_low and c_low not in taken):
            clinic_col = col
            break

    def _clean_clinic(v):
        s = str(v).strip() if pd.notna(v) else ''
        if not s or s.lower() in ('нан', 'none', 'nan') or _norm(s) in _AB_SERVICE_TOKENS:
            return None
        return s

    if clinic_col is not None:
        data['clinic'] = data[clinic_col].apply(_clean_clinic)
        if log_func:
            log_func("Клиники определены по столбцу «Клиника».")
    elif len(data):
        first_vals = data.iloc[:, 0].map(_clean_clinic)
        # значениями клиники считаем только короткие строки без пробелов
        # (коды/аббревиатуры вроде «ВМХ»); длинные фразы — посторонний текст
        looks_like_clinic = first_vals.notna() & first_vals.map(
            lambda s: bool(s) and ' ' not in str(s) and len(str(s)) <= 24)
        clinic_series = pd.Series(None, index=data.index, dtype='object')
        clinic_series[looks_like_clinic] = first_vals[looks_like_clinic]
        data['clinic'] = clinic_series
        if log_func and data['clinic'].notna().any():
            log_func("Клиники определены по значениям первой колонки таблицы.")
    else:
        data['clinic'] = pd.Series(dtype='object')

    found_clinics = sorted({str(c).strip() for c in data.get('clinic', pd.Series(dtype='object')).dropna()})
    if not found_clinics:
        data['clinic'] = DEFAULT_CLINIC_LABEL
        if log_func:
            log_func("Признаков разделения на клиники не найдено — "
                     "все данные обработаны как одна выборка.")
    else:
        # строки без явного указания клиники относятся к последней указанной
        # выше по таблице (блоки строк в выгрузке идут под кодом клиники)
        data['clinic'] = data['clinic'].ffill()
        data['clinic'] = data['clinic'].fillna(found_clinics[0])
        counts = data[data['is_positive']].groupby('clinic').size() \
            if 'is_positive' in data.columns else data.groupby('clinic').size()
        if log_func:
            summary = ", ".join(f"{c}: {int(n)}" for c, n in counts.items())
            log_func(f"Обнаружены клиники ({len(found_clinics)}): {summary}")

    # Находим позицию столбца 'organism'
    organism_pos = None
    for i, col in enumerate(data.columns):
        if col == 'organism':
            organism_pos = i
            break
    if organism_pos is None:
        raise ValueError("Не удалось определить позицию столбца с микроорганизмом")

    # Строка с названиями антибиотиков определяется автоматически —
    # больше не требуется фиксированное расположение (7-я строка)
    ab_name_row = detect_ab_name_row(df_raw, organism_pos, header_row_idx, log_func=log_func)
    if log_func:
        log_func(f"Используем строку {ab_name_row + 1} для названий антибиотиков")
        sample_cells = []
        for j in range(organism_pos+1, min(organism_pos+11, len(df_raw.columns))):
            val = df_raw.iloc[ab_name_row, j] if ab_name_row < len(df_raw) else ''
            sample_cells.append(str(val) if pd.notna(val) else 'NaN')
        log_func(f"Столбцы {organism_pos+2}-{organism_pos+11} на строке {ab_name_row}: {sample_cells}")

    antibiotics = []
    ab_indices = []
    taken_rsi = set()
    for j in range(organism_pos + 1, len(data.columns) - 1):
        # Название антибиотика ищем в строке НАД заголовками таблицы
        # (например, заголовки на 8-й строке — названия на 7-й).
        ab_name_raw = df_raw.iloc[ab_name_row, j] if 0 <= ab_name_row < len(df_raw) else ''
        ab_name = str(ab_name_raw).strip()
        if not ab_name or _norm(ab_name) in _AB_SERVICE_TOKENS or ab_name.isdigit():
            continue
        # Столбец с результатами RSI — ближайшая колонка справа от названия,
        # содержащая значения S/I/R/НД среди данных (каждый столбец
        # используется только один раз)
        rsi_idx = None
        for k in range(j + 1, min(j + 3, len(data.columns))):
            if k in taken_rsi:
                continue
            vals = {str(v).strip().upper() for v in data.iloc[:60, k].dropna().tolist()
                    if isinstance(v, str)}
            if vals & {'S', 'I', 'R', 'НД'}:
                rsi_idx = k
                break
        if rsi_idx is None:
            continue
        antibiotics.append(ab_name)
        ab_indices.append(rsi_idx)
        taken_rsi.add(rsi_idx)

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

    data = data.dropna(subset=['material_raw'])
    # Пустые строки (без микроорганизма и без антибиотикограммы) отбрасываем,
    # но наличие даты больше не является обязательным
    data = data[~(data['organism'].isna() & ~data['diagnostic'])]
    data['material_raw'] = data['material_raw'].astype(str).str.strip()
    data['patient_id'] = data['patient_id'].astype(str).str.strip()

    df_main = data[['patient_id', 'date', 'material_raw', 'organism',
                    'is_positive', 'diagnostic', 'clinic']].copy()

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
    df = df.copy()
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
    df = df.copy()
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


def build_wound_sheet5(data, antibiotics, ab_indices, log_func=None):
    """Формирует лист 5: раневые антибиотики с разбивкой по локализации."""
    cat_wound = data['material_raw'].apply(lambda x: categorize_material(x)[0])
    wound_data = data[cat_wound == 'РАНЫ'].copy()
    if wound_data.empty:
        if log_func:
            log_func("Нет раневых материалов, лист 5 будет пуст.")
        return pd.DataFrame(), wound_data
    if log_func:
        log_func(f"Раневых образцов: {len(wound_data)}")
    ab_data_wound = compute_ab_data(
        wound_data, antibiotics, ab_indices,
        group_cols=['organism', 'material_raw'],
        log_func=log_func
    )
    sheet5 = generate_antibiotic_sheet(ab_data_wound, log_func=log_func)
    return sheet5, wound_data


# ----------------------------------------------------------------------
# 7b. Конструктор дополнительных листов
#     Первые 3 листа отчёта всегда базовые; дополнительные листы
#     описываются словарями-спецификациями и строятся функциями ниже.
# ----------------------------------------------------------------------

#: Доступные поля для группировки в конструкторе: (ключ, отображаемое имя)
#: Любые поля можно комбинировать между собой — данные группируются по
#: всем выбранным полям одновременно. Поле 'total' — режим «Всё вместе»:
#: одна итоговая строка по всей выборке без разбивки.
GROUP_FIELDS = [
    ('total', 'Всё вместе (итог без разбивки)'),
    ('clinic', 'Клиника'),
    ('organism', 'Микроорганизм'),
    ('material_raw', 'Материал / локализация'),
    ('material_organism', 'Материал + микроорганизм'),
    ('category', 'Категория материала'),
    ('year', 'Год'),
    ('year_quarter', 'Год + квартал'),
    ('quarter', 'Квартал'),
    ('month', 'Месяц'),
    ('year_month', 'Год + месяц'),
    ('week', 'Неделя года'),
    ('weekday', 'День недели'),
    ('patient_id', 'Пациент (история болезни)'),
]

#: Специальные ключи группировки
TOTAL_GROUP_KEY = 'total'

#: Специальный режим конструктора: сравнение клиник — антибиотикограмма на
#: одни и те же микроорганизмы/материал в условиях разных клиник.
#: Лист повторяет структуру листа 3 (блоки «Антибиотик: …», строки —
#: микроорганизм × локализация, колонки S/I/R/НД/Total/S%/I%/R%/НД%), но
#: вместо общей выборки даёт отдельные блоки для каждой клиники; вместе с
#: таким листом автоматически создаётся лист с графиком сравнения.
COMPARE_CLINICS_KEY = '__compare_clinics__'

#: Максимальное число блоков-клиник на одном графике (листья Excel)
COMPARE_CHART_MAX_BLOCKS = 8
#: Максимальное число столбцов (диапазон АБ × микроорганизм) на графике
COMPARE_CHART_MAX_POINTS = 60

#: Специальное значение в выборе антибиотика: «Все антибиотики» — суммарная
#: антибиотикограмма по всем АБ-колонкам исходного файла сразу.
AB_ALL_KEY = '__ALL__'
AB_ALL_LABEL = 'Все антибиотики'

#: Готовые шаблоны (сценарии) для кнопки «Шаблон» в конструкторе:
#: (имя шаблона, спецификация листа)
SHEET_TEMPLATES = [
    ('— выбрать шаблон —', None),
    ('Резистентность по микроорганизмам', {
        'name': 'Резистентность по микробам', 'groups': ['organism'],
        'metrics': ['total', 'positive', 'ab'], 'filter': 'none'}),
    ('Динамика по месяцам', {
        'name': 'Динамика по месяцам', 'groups': ['year_month'],
        'metrics': ['total', 'positive', 'pct_positive', 'diagnostic'],
        'filter': 'none'}),
    ('Динамика по кварталам', {
        'name': 'Динамика по кварталам', 'groups': ['year_quarter'],
        'metrics': ['total', 'positive', 'pct_positive'], 'filter': 'none'}),
    ('Материал × микроорганизм', {
        'name': 'Материал и микробы', 'groups': ['material_organism'],
        'metrics': ['total', 'positive', 'pct_positive'], 'filter': 'none'}),
    ('Эпидемиология: категории материалов', {
        'name': 'Категории материалов', 'groups': ['category'],
        'metrics': ['total', 'positive', 'pct_positive', 'patients'],
        'filter': 'none'}),
    ('Полный свод по году', {
        'name': 'Свод по годам', 'groups': ['year'],
        'metrics': ['total', 'positive', 'pct_positive', 'diagnostic',
                    'pct_diagnostic', 'patients'], 'filter': 'none'}),
    ('Год × микроорганизм (устойчивые)', {
        'name': 'Год и микробы', 'groups': ['year', 'organism'],
        'metrics': ['total', 'positive'], 'filter': 'positive_only'}),
    ('Антибиотикограмма по материалу', {
        'name': 'АБГ по материалу', 'groups': ['material_raw'],
        'metrics': ['ab'], 'filter': 'none'}),
    ('Суммарная АБГ (все антибиотики) по микробам', {
        'name': 'АБГ все АБ по микробам', 'groups': ['organism'],
        'metrics': ['total', 'positive', 'ab'], 'antibiotic': AB_ALL_KEY,
        'filter': 'none'}),
    ('Итог по всем данным («всё вместе»)', {
        'name': 'Итоговый свод', 'groups': ['total'],
        'metrics': ['total', 'positive', 'pct_positive', 'diagnostic',
                    'pct_diagnostic', 'patients'], 'filter': 'none'}),
]

#: Доступные метрики: (ключ, отображаемое имя)
METRIC_OPTIONS = [
    ('total', 'Всего проб'),
    ('positive', 'Проб с ростом'),
    ('pct_positive', '% положительных'),
    ('diagnostic', 'Диагностичных'),
    ('pct_diagnostic', '% диагностичных'),
    ('patients', 'Уникальных пациентов'),
]

#: Метрики антибиотикограммы (считаются отдельно, по выбранному антибиотику)
AB_METRICS = ['S', 'I', 'R', 'НД']

#: Фильтры: (ключ, отображаемое имя). Специальные префиксы: 'organism=',
#: 'category=' — фильтрация по конкретному значению.
FILTER_OPTIONS = [
    ('none', 'Без фильтра'),
    ('positive_only', 'Только пробы с ростом'),
    ('no_growth_only', 'Только «роста нет»'),
    ('diagnostic_only', 'Только диагностичные'),
    ('non_diagnostic_only', 'Только недиагностичные'),
    ('wounds_only', 'Только раны'),
    ('urine_only', 'Только моча'),
    ('sputum_only', 'Только мокрота и т.д.'),
    ('blood_cat_only', 'Только кровь и катетеры'),
    ('csf_only', 'Только ликвор'),
    ('other_materials', 'Только прочие материалы'),
]

_MONTH_NAMES = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
_WEEKDAY_NAMES = ['Понедельник', 'Вторник', 'Среда', 'Четверг',
                  'Пятница', 'Суббота', 'Воскресенье']

#: Порядок сортировки колонок-метрик в итоговом листе
METRIC_COLUMN_ORDER = ['Всего проб', 'Проб с ростом', '% положительных',
                       'Диагностичных', '% диагностичных', 'Уникальных пациентов',
                       'S', 'I', 'R', 'НД', 'Total', 'S%', 'I%', 'R%', 'НД%']


def _has_date(d):
    return pd.notna(d) and not pd.isna(d)


def prepare_base_table(df_main, data=None):
    """Возвращает копию данных с доп. полями для конструктора:
    category (категория материала), year/month/quarter/weekday/week (из даты),
    составные ключи material_organism / year_month / year_quarter."""
    base = df_main.copy()
    base['category'] = base['material_raw'].apply(lambda x: categorize_material(str(x))[0])
    # подстраховка: если в df_main нет колонки 'organism', берём её из полных данных
    if 'organism' not in base.columns and data is not None and 'organism' in data.columns:
        base['organism'] = data['organism'].values[:len(base)]
    if 'patient_id' not in base.columns and data is not None and 'patient_id' in data.columns:
        base['patient_id'] = data['patient_id'].values[:len(base)]

    def _year(d):
        return d.year if _has_date(d) else None

    base['year'] = base['date'].apply(_year)
    base['month'] = base['date'].apply(
        lambda d: _MONTH_NAMES[d.month - 1] if _has_date(d) else 'без даты')
    base['quarter'] = base['date'].apply(
        lambda d: f'Q{(d.month - 1) // 3 + 1}' if _has_date(d) else 'без даты')
    base['year_quarter'] = base['date'].apply(
        lambda d: f'{d.year} Q{(d.month - 1) // 3 + 1}' if _has_date(d) else 'без даты')
    base['year_month'] = base['date'].apply(
        lambda d: f'{d.year}-{d.month:02d}' if _has_date(d) else 'без даты')
    base['week'] = base['date'].apply(
        lambda d: f'{d.year}-W{d.isocalendar()[1]:02d}' if _has_date(d) else 'без даты')
    base['weekday'] = base['date'].apply(
        lambda d: _WEEKDAY_NAMES[d.weekday()] if _has_date(d) else 'без даты')
    # составной ключ «материал + микроорганизм»
    base['material_organism'] = (base['material_raw'].astype(str).str.strip() + ' | ' +
                                 base['organism'].fillna('роста нет').astype(str).str.strip())
    # сортировка месяцев по календарю
    base['_month_order'] = base['date'].apply(
        lambda d: d.month if _has_date(d) else 99)
    return base


def apply_spec_filter(base_df, spec):
    """Применяет фильтр спецификации листа к таблице посевов."""
    ftype = str(spec.get('filter', 'none'))
    df = base_df
    if ftype == 'positive_only':
        df = df[df['is_positive']]
    elif ftype == 'no_growth_only':
        df = df[~df['is_positive']]
    elif ftype == 'diagnostic_only':
        df = df[df['diagnostic']]
    elif ftype == 'non_diagnostic_only':
        df = df[~df['diagnostic']]
    elif ftype == 'wounds_only':
        df = df[df['category'] == 'РАНЫ']
    elif ftype == 'urine_only':
        df = df[df['category'] == 'моча']
    elif ftype == 'sputum_only':
        df = df[df['category'] == 'мокрота и т.д.']
    elif ftype == 'blood_cat_only':
        df = df[df['category'].isin(['кровь', 'катетер'])]
    elif ftype == 'csf_only':
        df = df[df['category'] == 'жидкость цереброспинальная']
    elif ftype == 'other_materials':
        df = df[df['category'] == 'другие']
    elif ftype.startswith('organism='):
        org = ftype.split('=', 1)[1].strip().lower()
        df = df[df['organism'].astype(str).str.strip().str.lower() == org]
    elif ftype.startswith('category='):
        cat = ftype.split('=', 1)[1].strip().lower()
        df = df[df['category'].astype(str).str.strip().str.lower() == cat]
    return df.copy()


def resolve_group_fields(groups):
    """Раскрывает составные ключи группировки в реальные колонки таблицы.
    Возвращает (cols, labels): cols — список колонок base_df,
    labels — {колонка: отображаемое имя}. Ключ 'total' раскрывается в []"""
    labels = dict(GROUP_FIELDS)
    COMPOSITE = {
        'material_organism': (['material_raw', 'organism'],
                              {'material_raw': 'Материал / локализация',
                               'organism': 'Микроорганизм'}),
        'year_month': (['year', 'month'], {'year': 'Год', 'month': 'Месяц'}),
        'year_quarter': (['year', 'quarter'], {'year': 'Год', 'quarter': 'Квартал'}),
    }
    cols, col_labels = [], {}
    for g in groups:
        if g == TOTAL_GROUP_KEY:
            continue
        if g in COMPOSITE:
            sub, sub_labels = COMPOSITE[g]
            for c in sub:
                if c not in cols:
                    cols.append(c)
                    col_labels[c] = sub_labels[c]
        elif g in labels:
            if g not in cols:
                cols.append(g)
                col_labels[g] = labels[g]
    return cols, col_labels


def _grouped_series(gb, func):
    """Применяет func к каждой группе DataFrame и возвращает Series с индексом
    групп. Обходит сбои pandas-а при агрегировании пустых/экзотических срезов."""
    keys = []
    vals = []
    for name, g in gb:
        keys.append(name)
        vals.append(func(g))
    if not keys:
        return pd.Series(dtype="float64")
    try:
        idx = pd.Index(keys)
    except Exception:
        idx = pd.Index([str(k) for k in keys])
    return pd.Series(vals, index=idx, dtype="float64")


def _group_keys_frame(gb, group_cols):
    """Кадр с ключами групп (по одной колонке на поле группировки).
    Устойчив к случаям, когда pandas не выносит колонки групп из индекса."""
    keys = list(gb.groups.keys())
    if not group_cols:
        return pd.DataFrame({c: [] for c in group_cols})
    if len(group_cols) == 1:
        vals = [k[0] if isinstance(k, tuple) else k for k in keys]
        return pd.DataFrame({group_cols[0]: vals})
    return pd.DataFrame(keys, columns=group_cols)


def compute_group_stats(df, group_cols):
    """Сводная статистика по выбранным полям группировки (метрики посева).
    Пустой group_cols означает режим «Всё вместе» — одна итоговая строка."""
    import numpy as np
    if group_cols:
        gb = df.groupby(group_cols, dropna=False)
        out = pd.concat(
            [_group_keys_frame(gb, group_cols).reset_index(drop=True),
             pd.DataFrame({
                 'Всего проб': _grouped_series(gb, len),
                 'Проб с ростом': _grouped_series(gb, lambda g: int(g['is_positive'].sum())),
                 'Диагностичных': _grouped_series(
                     gb, lambda g: int((g['is_positive'] & g['diagnostic']).sum())),
                 'Уникальных пациентов': _grouped_series(
                     gb, lambda g: int(g['patient_id'].nunique())),
             }).reset_index(drop=True)],
            axis=1)
    else:
        total = len(df)
        pos = int(df['is_positive'].sum()) if total else 0
        diag = int((df['is_positive'] & df['diagnostic']).sum()) if total else 0
        pats = int(df['patient_id'].nunique()) if total else 0
        out = pd.DataFrame([{
            'Всего проб': total, 'Проб с ростом': pos,
            'Диагностичных': diag, 'Уникальных пациентов': pats}])
    t = out['Всего проб'].astype(float)
    p = out['Проб с ростом'].astype(float)
    d = out['Диагностичных'].astype(float)
    out['% положительных'] = (p / t.replace(0, np.nan) * 100).round(1).fillna(0)
    out['% диагностичных'] = (d / p.replace(0, np.nan) * 100).round(1).fillna(0)
    return out


def compute_ab_group_stats(data, antibiotics, ab_indices, group_cols, antibiotic,
                           only_positive=True, base_df=None):
    """Статистика S/I/R/НД по выбранному антибиотику для заданных полей группировки.

    group_cols — реальные колонки (после resolve_group_fields); пустой список
    означает режим «Всё вместе» (одна итоговая строка). Производные поля
    (category, year, month, quarter, week, weekday) подтягиваются из base_df.
    Счёт идёт по всем образцам с ростом: пустое значение RSI считается как «НД».

    Особый случай: antibiotic == AB_ALL_KEY («Все антибиотики») — считается
    суммарная антибиотикограмма по всем АБ-колонкам сразу (каждая ячейка
    «проба × антибиотик» даёт один зачёт; непустые MIC-значения не учитываются).
    """
    import numpy as np

    def _rsi(v):
        if isinstance(v, str):
            s = v.strip().upper()
            return s if s in AB_METRICS else 'НД'
        return 'НД'  # пусто / NaN / число MIC — определяемость не установлена

    df = data[data['is_positive']].copy() if only_positive else data.copy()
    # добавляем производные поля из базовой таблицы (позиционно совпадают строки)
    derived = [g for g in group_cols
               if g in ('category', 'year', 'month', 'quarter', 'year_month',
                        'year_quarter', 'week', 'weekday', 'material_organism')]
    if derived and base_df is not None:
        base_aligned = base_df.reset_index(drop=True)
        src = df.reset_index(drop=True)
        for g in derived:
            if g not in src.columns and g in base_aligned.columns:
                df[g] = base_aligned[g].values[:len(df)]

    if antibiotic == AB_ALL_KEY:
        # суммарная АБГ: «разворачиваем» все АБ-колонки в длинные данные
        if not antibiotics or not ab_indices:
            raise ValueError("В исходном файле не найдено колонок антибиотиков.")
        ab_cols = list(ab_indices)[:len(antibiotics)]
        long_parts = []
        for k, (ab_name, col_idx) in enumerate(zip(antibiotics, ab_cols)):
            part = df[group_cols].copy() if group_cols else pd.DataFrame(index=df.index)
            part['_rsi'] = df.iloc[:, col_idx].apply(_rsi)
            long_parts.append(part)
        long_df = pd.concat(long_parts, ignore_index=True)
        if group_cols:
            counts = (long_df.groupby(group_cols + ['_rsi'], dropna=False).size()
                      .unstack('_rsi', fill_value=0))
            # unstack может поместить колонки группировки в MultiIndex — снимаем
            if isinstance(counts.index, pd.MultiIndex):
                counts = counts.reset_index()
        else:
            vc = long_df['_rsi'].value_counts()
            counts = pd.DataFrame({m: [int(vc.get(m, 0))] for m in AB_METRICS})
    else:
        if antibiotic not in antibiotics:
            raise ValueError(f"Антибиотик «{antibiotic}» не найден в исходном файле.")
        i = antibiotics.index(antibiotic)
        col_idx = ab_indices[i]
        df['_rsi'] = df.iloc[:, col_idx].apply(_rsi)
        if group_cols:
            counts = (df.groupby(group_cols + ['_rsi'], dropna=False).size()
                      .unstack('_rsi', fill_value=0))
        else:
            vc = df['_rsi'].value_counts()
            counts = pd.DataFrame({m: [int(vc.get(m, 0))] for m in AB_METRICS})
    for m in AB_METRICS:
        if m not in counts.columns:
            counts[m] = 0
    counts = counts[AB_METRICS]
    # после unstack колонки группировки могут остаться в индексе (в т.ч. с
    # именем '_rsi') — гарантированно выносим их в обычные столбцы
    if any(c in counts.index.names for c in group_cols) or \
            counts.index.name is not None:
        counts = counts.reset_index()
    if isinstance(counts.columns, pd.MultiIndex):
        counts.columns = [str(c) for c in counts.columns]
    counts = counts.drop(columns=['_rsi'], errors='ignore')
    counts['Total'] = counts[AB_METRICS].sum(axis=1)
    total_f = counts['Total'].astype(float).replace(0, np.nan)
    for m in AB_METRICS:
        counts[f'{m}%'] = (counts[m].astype(float) / total_f * 100).round(1).fillna(0)
    return counts.reset_index()


def sort_by_month(df, month_col='Месяц'):
    """Сортирует строки по календарному порядку месяцев (для группировки «Месяц»)."""
    order = {name: k for k, name in enumerate(_MONTH_NAMES, start=1)}
    tmp = df.copy()
    tmp['_mo'] = tmp[month_col].map(order).fillna(99)
    return tmp.sort_values('_mo').drop(columns='_mo').reset_index(drop=True)


def _sort_key_for_group(col_label):
    """Возвращает порядковый ключ для человекочитаемых значений группировки."""
    if col_label == 'Месяц':
        order = {name: k for k, name in enumerate(_MONTH_NAMES, start=1)}
        return lambda v: (order.get(str(v), 99),)
    if col_label == 'Квартал':
        return lambda v: (str(v),)
    if col_label == 'День недели':
        order = {name: k for k, name in enumerate(_WEEKDAY_NAMES, start=1)}
        return lambda v: (order.get(str(v), 99),)
    return None


def build_sheet_from_spec(spec, base_df, data, antibiotics, ab_indices, log_func=None):
    """Строит один дополнительный лист по спецификации конструктора.

    Спецификация — словарь:
      name        — имя листа Excel
      groups      — список полей группировки (ключи GROUP_FIELDS); любые
                    комбинации поддерживаются, включая 'total' («всё вместе»)
      metrics     — список метрик; если содержит 'ab', считается антибиотикограмма
      antibiotic  — название антибиотика (когда metrics включает 'ab')
      filter      — ключ фильтра (FILTER_OPTIONS)
    Возвращает DataFrame или None, если лист построить не удалось.
    """
    valid_groups = dict(GROUP_FIELDS)
    groups = [g for g in spec.get('groups', []) if g in valid_groups]
    if not groups:
        # пустая группировка трактуется как режим «Всё вместе» (итог без разбивки)
        if log_func:
            log_func(f"Лист «{spec.get('name')}»: поля группировки не выбраны — "
                     f"будет построена одна итоговая строка («всё вместе»).")

    # «Всё вместе» имеет приоритет: если выбран — игнорируем прочие поля
    is_total = TOTAL_GROUP_KEY in groups or not groups
    groups = [] if is_total else groups
    group_cols, col_labels = resolve_group_fields(groups)

    filtered = apply_spec_filter(base_df, spec)
    if filtered.empty and log_func:
        log_func(f"Лист «{spec.get('name')}»: нет данных после применения фильтра.")

    _metric_keys = {k for k, _ in METRIC_OPTIONS}
    metrics = [m for m in spec.get('metrics', []) if m == 'ab' or m in _metric_keys]
    if not metrics:
        # метрики не выбраны (или выбраны недопустимо) — берём базовый набор
        metrics = ['total', 'positive', 'pct_positive']
        if log_func:
            log_func(f"Лист «{spec.get('name')}»: метрики не выбраны — применён "
                     f"набор по умолчанию (всего проб / с ростом / % положительных).")
    frames = []

    if any(m in _metric_keys for m in metrics):
        stats = compute_group_stats(filtered, group_cols)
        keep = group_cols + [label for key, label in METRIC_OPTIONS if key in metrics]
        frames.append(stats[[c for c in keep if c in stats.columns]])

    if 'ab' in metrics:
        ab_name = spec.get('antibiotic', '')
        try:
            ab_stats = compute_ab_group_stats(data, antibiotics, ab_indices, group_cols,
                                              ab_name, base_df=base_df)
            # нормализуем структуру АБГ-таблицы для слияния: оставляем только
            # колонки группировки и метрики S/I/R/НД (служебные index/_rsi — в мусор)
            ab_keep = [c for c in group_cols if c in ab_stats.columns]
            ab_mcols = [c for c in AB_METRICS + ['Total'] + [f'{m}%' for m in AB_METRICS]
                        if c in ab_stats.columns]
            missing = [c for c in group_cols if c not in ab_stats.columns]
            ab_stats = ab_stats[ab_keep + ab_mcols].copy()
            if missing and log_func:
                log_func(f"Лист «{spec.get('name')}»: в антибиотикограмме не найдены "
                         f"поля группировки {missing} — лист построен без АБГ.")
                ab_stats = None
            if ab_stats is not None:
                # приводим колонки группировки к одному типу с посевной статистикой
                # (иначе merge не найдёт ключей из-за различий dtype)
                for c in group_cols:
                    if c in filtered.columns:
                        ab_stats[c] = ab_stats[c].astype(filtered[c].dtype)
                frames.append(ab_stats)
        except ValueError as e:
            if log_func:
                log_func(f"Лист «{spec.get('name')}»: {e}")
    if not frames:
        if log_func:
            log_func(f"Лист «{spec.get('name')}»: не выбрано ни одной метрики — лист пропущен.")
        return None

    result = frames[0]
    for f in frames[1:]:
        if group_cols:
            result = result.merge(f, on=group_cols, how='outer')
        else:
            # режим «всё вместе»: склеиваем единственные строки без ключа
            result = pd.concat([result.reset_index(drop=True),
                                f.drop(columns=[c for c in f.columns
                                                if c in result.columns],
                                       errors='ignore').reset_index(drop=True)],
                               axis=1)
        # строки без антибиотикограммы (группы только в посевной статистике)
        for c in ('S', 'I', 'R', 'НД', 'Total'):
            if c in result.columns:
                result[c] = result[c].fillna(0).astype(int)
        for c in ('S%', 'I%', 'R%', 'НД%'):
            if c in result.columns:
                result[c] = result[c].fillna(0)

    # переименуем технические имена группировок в человекочитаемые
    result = result.rename(columns=col_labels)
    display_cols = [col_labels[c] for c in group_cols]
    if is_total:
        display_cols = ['Итог']
        if result.columns.size and display_cols[0] not in result.columns:
            result.insert(0, 'Итог', 'Все данные')

    # упорядочиваем столбцы: сначала группы, затем метрики в фиксированном порядке
    other = [c for c in result.columns if c not in display_cols]
    ordered = [c for c in METRIC_COLUMN_ORDER if c in other]
    rest = [c for c in other if c not in ordered]
    result = result[display_cols + ordered + rest]

    # ---- сортировка строк -------------------------------------------------
    if is_total:
        pass  # одна строка
    else:
        sort_specs = []          # [(колонка, ascending, ключ-функция)]
        if 'year' in group_cols:
            sort_specs.append(('Год', True, None))
        if 'quarter' in group_cols:
            sort_specs.append(('Квартал', True, _sort_key_for_group('Квартал')))
        if 'month' in group_cols:
            sort_specs.append(('Месяц', True, _sort_key_for_group('Месяц')))
        if 'week' in group_cols:
            sort_specs.append(('Неделя года', True, None))
        if 'weekday' in group_cols:
            sort_specs.append(('День недели', True, _sort_key_for_group('День недели')))
        remaining = [col_labels[c] for c in group_cols
                     if col_labels[c] not in {s[0] for s in sort_specs}]
        if sort_specs:
            # сначала временна́я шкала, затем прочие выбранные группы,
            # внутри группы — по убыванию объёма проб
            extra = [('Всего проб', False, None)] if 'Всего проб' in result.columns else []
            keys = sort_specs + [(c, True, None) for c in remaining] + extra
            tmp = result.copy()
            by, asc = [], []
            for col, a, kf in keys:
                if col not in tmp.columns:
                    continue
                if kf:
                    tmp['_k_' + col] = tmp[col].map(lambda v: kf(v)[0])
                    by.append('_k_' + col)
                else:
                    by.append(col)
                asc.append(a)
            if by:
                result = tmp.sort_values(by, ascending=asc).drop(
                    columns=[c for c in tmp.columns if str(c).startswith('_k_')])
        elif 'material_raw' in group_cols or 'category' in group_cols:
            # эпидемиологические срезы: категории/материалы по убыванию числа проб
            if 'Всего проб' in result.columns:
                primary = 'Материал / локализация' if 'material_raw' in group_cols else 'Категория материала'
                result = result.sort_values([primary, 'Всего проб'],
                                            ascending=[True, False])
            if 'category' in group_cols:
                pass
        elif remaining:
            first = col_labels[group_cols[0]]
            if 'Всего проб' in result.columns and first in result.columns:
                result = result.sort_values([first, 'Всего проб'], ascending=[True, False])
            elif 'Всего проб' in result.columns:
                result = result.sort_values('Всего проб', ascending=False)
        result = result.reset_index(drop=True)

    if log_func:
        desc = 'всё вместе (итог)' if is_total else \
               ', '.join(col_labels[c] for c in group_cols)
        log_func(f"Лист «{spec.get('name')}»: {len(result)} строк (группировка: {desc}).")
    return result


def sanitize_sheet_name(name):
    """Делает имя листа допустимым для Excel: убирает запрещённые символы,
    ограничивает длину 31 символом. Пустое имя заменяется на «Лист»."""
    bad = set(':\\/?*[]')
    cleaned = ''.join(ch for ch in str(name) if ch not in bad).strip()
    return cleaned[:31] if cleaned else 'Лист'


def validate_spec(spec, available_antibiotics=None):
    """Проверяет спецификацию листа; возвращает список сообщений об ошибках.

    Название автокорректируется (санитизируется), отсутствие метрик не является
    ошибкой — применится набор по умолчанию. Пустая группировка допустима и
    трактуется как режим «Всё вместе» (итог без разбивки).
    """
    errors = []
    name = sanitize_sheet_name(str(spec.get('name', '')).strip())
    if len(name) > 31:
        errors.append("Название листа не должно превышать 31 символ.")
    is_compare = spec.get('mode') == COMPARE_CLINICS_KEY
    if is_compare:
        # отдельный режим «Сравнение клиник»: список антибиотиков + клиники
        abs_list = [a for a in (spec.get('antibiotics') or
                                ([spec['antibiotic']] if spec.get('antibiotic') else []))
                    if a and a != AB_ALL_KEY]
        if not abs_list:
            errors.append("Для сравнения клиников выберите хотя бы один антибиотик "
                          "(пустой список = все антибиотики файла).")
        if available_antibiotics is not None:
            missing = [a for a in abs_list if a not in available_antibiotics]
            if missing:
                errors.append("Антибиотики отсутствуют в текущем исходном файле: "
                              + ", ".join(missing[:5]))
        clinics = [c for c in (spec.get('clinics') or []) if c]
        if len(clinics) < 2:
            errors.append("Для сравнения отметьте минимум две клиники.")
        return errors
    groups = [g for g in spec.get('groups', []) if g in dict(GROUP_FIELDS)]
    if not groups:
        # не ошибка: будет построен итог «всё вместе»
        pass
    metrics = [m for m in spec.get('metrics', [])
               if m == 'ab' or m in {k for k, _ in METRIC_OPTIONS}]
    if 'ab' in metrics:
        ab = spec.get('antibiotic', '')
        if not ab:
            errors.append("Для метрики «антибиотикограмма» выберите антибиотик.")
        elif ab != AB_ALL_KEY and available_antibiotics is not None \
                and ab not in available_antibiotics:
            errors.append(f"Антибиотик «{ab}» отсутствует в текущем исходном файле.")
    return errors


# --- Пресеты конструктора (сохранение/загрузка настроек листов) -----------

def save_sheet_specs(specs, path):
    """Сохраняет список спецификаций листов в JSON-файл."""
    import json
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(list(specs), f, ensure_ascii=False, indent=2)


def load_sheet_specs(path):
    """Загружает список спецификаций листов из JSON-файла."""
    import json
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, dict):
        data = data.get('sheets', [])
    specs = []
    for item in data:
        if not isinstance(item, dict):
            continue
        mode = str(item.get('mode', ''))
        if mode and mode != COMPARE_CLINICS_KEY:
            mode = ''
        specs.append({
            'name': str(item.get('name', '')),
            'groups': list(item.get('groups', [])),
            'metrics': list(item.get('metrics', [])),
            'antibiotic': str(item.get('antibiotic', '')),
            'filter': str(item.get('filter', 'none')),
            **({'mode': COMPARE_CLINICS_KEY,
                'compare_by': list(item.get('compare_by', ['organism'])),
                'antibiotics': list(item.get('antibiotics', [])),
                'clinics': list(item.get('clinics', []))} if mode else {}),
        })
    return specs


def filter_data_by_clinic(data, antibiotics, ab_indices, clinic):
    """Оставляет в данных строки только выбранной клиники.

    Возвращает новую тройку (data, antibiotics, ab_indices); список
    антибиотиков остаётся общим для всего файла — так антибиотикограммы
    разных клиник сопоставимы по составу.
    """
    if not clinic or 'clinic' not in data.columns:
        return data, antibiotics, ab_indices
    mask = data['clinic'].astype(str).str.strip() == str(clinic).strip()
    return data[mask].copy(), antibiotics, ab_indices


# --- Отдельный режим работы: сравнение клиник ---------------------------

#: Пункт выбора клиники на главном окне: отдельный режим «Сравнение клиник»
COMPARE_CLINICS_LABEL = '⚖️ Сравнение клиник'


def get_compare_sheet_names(spec_name):
    """Имена листов режима сравнения: таблица + лист с графиком."""
    base = sanitize_sheet_name(str(spec_name).strip())[:25] or 'Сравнение'
    return base, sanitize_sheet_name(f'{base} (график)')


def _clinic_codes(data):
    """Клиники в порядке встречаемости (без служебной «общей выборки»)."""
    clinics = []
    if data is not None and 'clinic' in getattr(data, 'columns', []):
        for c in data['clinic'].dropna().astype(str).str.strip():
            if c and c != DEFAULT_CLINIC_LABEL and c not in clinics:
                clinics.append(c)
    return clinics


def get_file_info(filepath, log_func=None):
    """Возвращает (список клиник, список антибиотиков) исходного файла.

    Используется интерфейсом: клиники — для чекбоксов режима «Сравнение
    клиник» и выбора клиники на главном окне; антибиотики — для конструктора.
    """
    _, _, data, antibiotics, _ = load_data(filepath, log_func=log_func)
    return _clinic_codes(data), list(antibiotics)


def count_rows_by_clinic(data):
    """dict {клиника: число строк посевов} — для подписей чекбоксов."""
    counts = {}
    if data is not None and 'clinic' in getattr(data, 'columns', []):
        vc = data['clinic'].dropna().astype(str).str.strip().value_counts()
        for cl, n in vc.items():
            if cl and cl != DEFAULT_CLINIC_LABEL:
                counts[cl] = int(n)
    return counts


def _cmp_block_stats(sub, data, col_lab, row_fields, combos):
    """Считает строки блока (клиника × антибиотик) в формате листа 3.

    Возвращает (rows_out, total_row, s_list, labels_out):
      rows_out  — строки [ключ группы..., S, I, R, НД, Total, S%, I%, R%, НД%];
      total_row — итоговая строка «ВСЕГО» в том же формате;
      s_list/labels_out — значения S% и подписи строк для графика.
    Строки сортируются по общему списку комбинаций (combos), поэтому блоки
    разных клиник содержат одинаковые наборы «микроорганизм × материал»
    в одном порядке — сравнение визуально сопоставимо.
    """
    def _rsi(v):
        if isinstance(v, str):
            s = v.strip().upper()
            return s if s in AB_METRICS else 'НД'
        return 'НД'  # пусто / NaN / число MIC — определяемость не установлена

    if col_lab is not None and col_lab in data.columns:
        rsi_src = data.loc[sub.index, col_lab]
    else:
        rsi_src = pd.Series(index=sub.index, dtype=object)
    tmp = pd.DataFrame({'_rsi': rsi_src.apply(_rsi)})
    if row_fields:
        for f in row_fields:
            # берём значения по индексам (а не .values): pos может быть вырезан
            # из base_df, индексы строк которого не совпадают с data
            tmp[f] = sub[f].reindex(tmp.index).fillna('').astype(str).str.strip().values
        counts = tmp.groupby(row_fields + ['_rsi']).size().unstack('_rsi', fill_value=0)
        if isinstance(counts.index, pd.MultiIndex):
            counts = counts.reset_index()
        counts = counts.reset_index(drop=True)
    else:
        vc = tmp['_rsi'].value_counts()
        counts = pd.DataFrame({m: [int(vc.get(m, 0))] for m in AB_METRICS})
    for m in AB_METRICS:
        if m not in counts.columns:
            counts[m] = 0

    stats = {}
    for _, r in counts.iterrows():
        key = tuple(str(r[f]) for f in row_fields) if row_fields else ('Все данные',)
        s, i_, rr, nd = (int(r.get('S', 0)), int(r.get('I', 0)),
                        int(r.get('R', 0)), int(r.get('НД', 0)))
        tot = s + i_ + rr + nd
        tf = float(tot) if tot else np.nan
        pcts = [round(x / tf * 100, 1) if tf == tf else 0.0 for x in (s, i_, rr, nd)]
        stats[key] = [s, i_, rr, nd, tot] + pcts

    rows_out, s_list, labels_out = [], [], []
    for key in combos:
        if key not in stats:
            continue
        vals = stats[key]
        rows_out.append(list(key) + vals)
        s_list.append(vals[5])
        lbl = ' | '.join(key) if row_fields else 'Все данные'
        labels_out.append(lbl)
    if not rows_out:
        return rows_out, None, s_list, labels_out

    arr = np.array([stats[k] for k in combos if k in stats], dtype=float)
    sums = arr.sum(axis=0)
    stot = sums[4]
    total_row = (['ВСЕГО'] * max(len(row_fields), 1) +
                 [int(sums[0]), int(sums[1]), int(sums[2]), int(sums[3]), int(stot),
                  round(sums[0] / stot * 100, 1) if stot else 0.0,
                  round(sums[1] / stot * 100, 1) if stot else 0.0,
                  round(sums[2] / stot * 100, 1) if stot else 0.0,
                  round(sums[3] / stot * 100, 1) if stot else 0.0])
    return rows_out, total_row, s_list, labels_out


def build_clinic_comparison(spec, base_df, data, antibiotics, ab_indices, log_func=None):
    """Лист(ы) отдельного режима «Сравнение клиник».

    Возвращает dict {имя листа: DataFrame}. Табличный лист повторяет
    структуру листа 3 (антибиотикорезистентность): блоки заголовков
    «Антибиотик: …», строки — микроорганизм × локализация, колонки
    S / I / R / НД / Total / S% / I% / R% / НД%. Отличие — каждый блок
    сравнивает одни и те же антибиотики, микроорганизмы и материал
    в условиях разных клиник: под общим заголовком идёт подзаголовок
    клиники, затем строки по этой клинике и итоговая строка «ВСЕГО».
    Вместе с таблицей создаётся лист с графиками S% по клиникам.

    Ключи спецификации:
      antibiotics — список антибиотиков (пусто/AB_ALL_KEY — все найденные);
      compare_by  — поля строк: 'organism' и/или 'material_raw';
      clinics     — список клиник для сравнения (пусто — все из файла);
      filter      — обычный фильтр конструктора (применяется ко всем блокам).
    """
    # --- разбор спецификации ---
    compare_by = [g for g in spec.get('compare_by', ['organism'])
                  if g in ('organism', 'material_raw')]
    if not compare_by:
        compare_by = ['organism']

    wanted_abs = [a for a in (spec.get('antibiotics') or
                              ([spec['antibiotic']] if spec.get('antibiotic') else []))
                  if a and a != AB_ALL_KEY]
    if wanted_abs:
        selected_abs = [a for a in antibiotics if a in set(wanted_abs)]
    else:
        selected_abs = list(antibiotics)
    if not selected_abs:
        raise ValueError("Для сравнения клиник выберите хотя бы один антибиотик "
                         "(иначе лист будет пуст).")

    all_clinics = _clinic_codes(data)
    if len(all_clinics) < 2:
        raise ValueError("В исходном файле не найдено минимум двух клиник — "
                         "сравнивать нечего.")
    wanted_clinics = [c for c in (spec.get('clinics') or []) if c in all_clinics]
    clinics = wanted_clinics if len(wanted_clinics) >= 2 else all_clinics
    if len(clinics) < 2:
        raise ValueError("Для сравнения отметьте минимум две клиники.")

    # --- строки таблицы: посевы (фильтр) + антибиотикограмма по АБ-колонкам ---
    filtered = apply_spec_filter(base_df, spec)
    pos = filtered[filtered['is_positive']].copy() if 'is_positive' in filtered.columns else filtered.copy()
    if pos.empty:
        raise ValueError("После фильтра не осталось проб с ростом — сравнение пустое.")

    row_fields = []
    if 'organism' in compare_by:
        row_fields.append('organism')
    if 'material_raw' in compare_by:
        row_fields.append('material_raw')
    for f in row_fields:
        if f not in pos.columns:
            pos[f] = ''
        pos[f] = pos[f].fillna('').astype(str).str.strip()
    if row_fields:
        combos = sorted(set(map(tuple, pos[row_fields].drop_duplicates().values.tolist())))
    else:
        combos = [('Все данные',)]

    # индексы колонок данных: position -> column label
    col_labels_map = {}
    for i, c in enumerate(data.columns):
        col_labels_map[i] = c

    blocks = []          # (ab, clinic, rows(list), total_row(list))
    chart_data = []      # (ab, clinic, [(label, S%)])
    for ab in selected_abs:
        idx = antibiotics.index(ab)
        col_i = ab_indices[idx]
        col_lab = col_labels_map.get(col_i)
        for cl in clinics:
            sub = pos[pos['clinic'].astype(str).str.strip() == cl] if 'clinic' in pos.columns else pos
            if sub.empty:
                continue
            rows_out, total_row, s_list, labels_out = _cmp_block_stats(
                sub, data, col_lab, row_fields, combos)
            if not rows_out:
                continue
            blocks.append((ab, cl, rows_out, total_row))
            chart_data.append((ab, cl, list(zip(labels_out, s_list))))

    if not blocks:
        raise ValueError("По выбранным антибиотикам и клиникам нет определений "
                         "чувствительности — измените настройки сравнения.")

    # --- сборка табличного DataFrame в стиле листа 3 ---
    header = (['Микроорганизм', 'Локализация'] if len(row_fields) == 2 else
              ['Микроорганизм'] if row_fields == ['organism'] else
              ['Локализация'] if row_fields == ['material_raw'] else ['Группа'])
    header += AB_METRICS + ['Total'] + [f'{m}%' for m in AB_METRICS]
    ncol = len(header)
    all_rows = []
    cur_ab = None
    for ab, cl, rows_out, total_row in blocks:
        if ab != cur_ab:
            if cur_ab is not None:
                all_rows.append([''] * ncol)
            all_rows.append([f"Антибиотик: {ab}"] + [''] * (ncol - 1))
            cur_ab = ab
        all_rows.append([f"Клиника: {cl}"] + [''] * (ncol - 1))
        all_rows.append(header)
        all_rows.extend(rows_out)
        all_rows.append(total_row)
        all_rows.append([''] * ncol)
    while all_rows and all_rows[-1] == [''] * ncol:
        all_rows.pop()
    table_df = pd.DataFrame(all_rows)

    # --- лист с графиками: S% по клиникам для одних и тех же наборов ---
    groups = []  # (ab, [labels], {clinic: [S%]}, used_labels)
    by_ab = {}
    for ab, cl, pairs in chart_data:
        by_ab.setdefault(ab, []).append((cl, pairs))
    for ab in selected_abs:
        if ab not in by_ab:
            continue
        dmaps = {cl: dict(pairs) for cl, pairs in by_ab[ab]}
        cl_list = [cl for cl in clinics if cl in dmaps]
        labels = []
        for cl in cl_list:
            for lbl in dmaps[cl]:
                if lbl not in labels:
                    labels.append(lbl)
        labels.sort()
        # приоритет: диапазоны, присутствующие во всех клиниках
        common = [l for l in labels if all(l in dmaps[cl] for cl in cl_list)]
        ordered = (common + [l for l in labels if l not in set(common)])[:COMPARE_CHART_MAX_POINTS]
        series = {cl: [dmaps[cl].get(l, 0.0) for l in ordered] for cl in cl_list}
        if ordered:
            groups.append((ab, ordered, series, cl_list))
    group_rows = []
    for ab, labels, series, cl_list in groups:
        group_rows.append([ab] + [''] * len(labels))
        group_rows.append(['Клиника'] + labels)
        for cl in cl_list:
            group_rows.append([cl] + [round(v, 1) for v in series[cl]])
        group_rows.append([''] * (len(labels) + 1))
    while group_rows and group_rows[-1] == ['']:
        group_rows.pop()
    chart_df = pd.DataFrame(group_rows) if group_rows else pd.DataFrame()

    tab_name, chart_name = get_compare_sheet_names(spec.get('name', 'Сравнение клиник'))
    result = {tab_name: table_df}
    if not chart_df.empty:
        result[chart_name] = chart_df

    if log_func:
        log_func(f"Сравнение клиник «{spec.get('name')}»: {len(selected_abs)} антибиотик(ов), "
                 f"{len(clinics)} клиник ({', '.join(clinics)}), "
                 f"{len(blocks)} блоков, лист графика: «{chart_name}».")
    return result


def run_pipeline(input_path, output_path, styled=True, log_func=None, extra_sheets=None,
                 clinic=None, preloaded=None, base_sheets=True):
    """Полный цикл обработки одной выборки: загрузка -> базовые листы ->
    конструктор -> Excel.

    clinic — код клиники из исходного файла; если задан, обрабатываются
    только её данные (базовые листы 1–3 строятся по этой клинике). Листы с
    mode == COMPARE_CLINICS_KEY всегда строятся по ПОЛНОМУ файлу (все
    клиники) — даже когда выбрана одна клиника, иначе сравнивать нечего.
    preloaded — кортеж из load_data (df_main, ab_data, data, antibiotics,
    ab_indices), чтобы не читать файл повторно в многотомной обработке.
    extra_sheets — список спецификаций конструктора (дополнительные листы
    после трёх базовых). Каждое описание проверяется validate_spec().
    """
    if preloaded is not None:
        df_main, ab_data, data, antibiotics, ab_indices = preloaded
    else:
        df_main, ab_data, data, antibiotics, ab_indices = load_data(input_path, log_func=log_func)

    # полные (межклинические) данные для листов сравнения
    cmp_state = dict(data=data, antibiotics=antibiotics, ab_indices=ab_indices,
                     base=prepare_base_table(df_main, data))

    if clinic:
        before = len(data)
        # работаем с копией, чтобы не «сжигать» строки в preloaded-данных
        data, antibiotics, ab_indices = filter_data_by_clinic(
            data.copy(), antibiotics, ab_indices, clinic)
        df_main = df_main[df_main['clinic'].astype(str).str.strip() == str(clinic).strip()] \
            if 'clinic' in df_main.columns else df_main
        ab_data = compute_ab_data(data, antibiotics, ab_indices, group_cols=['organism'],
                                  log_func=None)
        if log_func:
            log_func(f"Фильтр по клинике «{clinic}»: {before} строк → {len(data)}.")
        if data.empty:
            raise ValueError(f"Клиника «{clinic}» не найдена в исходном файле.")

    if base_sheets:
        if log_func:
            log_func("Формирование листа 1 (анализ посевов по категориям)...")
        sheet2 = generate_sheet2(df_main, log_func=log_func)

        if log_func:
            log_func("Формирование листа 2 (эпидемиология по локализациям)...")
        sheet3 = generate_sheet3(df_main, log_func=log_func)

        if log_func:
            log_func("Формирование листа 3 (общая антибиотикорезистентность)...")
        sheet4 = generate_antibiotic_sheet(ab_data, log_func=log_func)
    else:
        # отдельный режим (например, сравнение клиник): базовые листы не создаются
        sheet2 = pd.DataFrame()
        sheet3 = {}
        sheet4 = pd.DataFrame()

    # Листы конструктора
    built_sheets = {}
    if extra_sheets:
        base = prepare_base_table(df_main, data)
        used_names = {'1', '2', '3'} if base_sheets else set()
        for n, spec in enumerate(extra_sheets, start=1):
            errs = validate_spec(spec, available_antibiotics=antibiotics)
            if errs:
                if log_func:
                    log_func(f"Лист №{n}: {errs[0]} — лист пропущен.")
                continue
            try:
                if spec.get('mode') == COMPARE_CLINICS_KEY:
                    res = build_clinic_comparison(
                        spec, cmp_state['base'], cmp_state['data'],
                        cmp_state['antibiotics'], cmp_state['ab_indices'],
                        log_func=log_func)
                    # режим сравнения возвращает dict {имя: DataFrame}
                    dfs = res if isinstance(res, dict) else {spec['name']: res}
                else:
                    df_sheet = build_sheet_from_spec(spec, base, data, antibiotics, ab_indices,
                                                     log_func=log_func)
                    dfs = {str(spec['name']).strip(): df_sheet}
            except ValueError as e:
                if log_func:
                    log_func(f"Лист №{n}: {e} — лист пропущен.")
                continue
            for raw_name, df_sheet in dfs.items():
                if df_sheet is None or df_sheet.empty:
                    continue
                nm = sanitize_sheet_name(str(raw_name).strip())
                k = 1
                while nm in used_names:
                    k += 1
                    nm = f"{sanitize_sheet_name(str(raw_name).strip())[:28]}_{k}"
                used_names.add(nm)
                built_sheets[nm] = df_sheet

    save_to_excel(sheet2, sheet3, sheet4, output_path,
                  log_func=log_func, styled=styled,
                  raw_data=data,
                  user_sheets=built_sheets, base_sheets=base_sheets,
                  chart_sheets={k: v for k, v in built_sheets.items()
                                if '(график)' in k})

    if log_func:
        log_func("Обработка успешно завершена!")
    return dict(df_main=df_main, ab_data=ab_data, sheet2=sheet2, sheet3=sheet3,
                sheet4=sheet4, user_sheets=built_sheets)


def make_per_clinic_output_path(output_path, clinic):
    """«/путь/отчёт.xlsx» + клиника «ВМХ» → «/путь/отчёт — ВМХ.xlsx»."""
    root, ext = os.path.splitext(output_path)
    safe = sanitize_sheet_name(str(clinic)) or 'Клиника'
    return f"{root} — {safe}{ext or '.xlsx'}"


def run_pipeline_multi_clinic(input_path, output_path, styled=True, log_func=None,
                              extra_sheets=None, clinics=None):
    """Обработка «Все клиники»: каждый набор данных сохраняется в отдельный файл.

    Возвращает dict {код клиники: путь к созданному файлу}. Если клиник в
    файле не обнаружено — создаётся один общий файл (как в run_pipeline).
    """
    if clinics is None:
        clinics = detect_clinics(input_path, log_func=None)
    results = {}
    if not clinics:
        if log_func:
            log_func("Клиники в файле не обнаружены — обработка одной общей выборкой.")
        run_pipeline(input_path, output_path, styled=styled, log_func=log_func,
                     extra_sheets=extra_sheets)
        results[DEFAULT_CLINIC_LABEL] = output_path
        return results
    # читаем исходный файл один раз для всех клиник
    preloaded = load_data(input_path, log_func=log_func)
    if log_func:
        log_func(f"Режим «{ALL_CLINICS_LABEL}»: будет создано отдельных файлов — "
                 f"{len(clinics)} ({', '.join(clinics)}).")
    for cl in clinics:
        path = make_per_clinic_output_path(output_path, cl)
        if log_func:
            log_func(f"— Клиника {cl}: сохранение в {os.path.basename(path)}")
        run_pipeline(input_path, path, styled=styled, log_func=log_func,
                     extra_sheets=extra_sheets, clinic=cl, preloaded=preloaded)
        results[cl] = path
    if log_func:
        log_func(f"Готово: создано {len(results)} файлов (по одному на клинику).")
    return results


def run_clinic_comparison(input_path, output_path, spec, styled=True, log_func=None):
    """Отдельный режим работы «Сравнение клиник».

    Строит только листы сравнения (таблица в стиле листа 3 + лист с
    графиками); базовые листы 1–3 НЕ создаются. Спецификация:
      spec.clinics    — отмеченные чекбоксами клиники из исходного файла;
      spec.antibiotics — антибиотики для сравнения (пусто — все);
      spec.compare_by — 'organism' и/или 'material_raw' (строки таблиц);
      spec.filter     — обычный фильтр конструктора.
    Возвращает dict {имя листа: DataFrame}.
    """
    df_main, ab_data, data, antibiotics, ab_indices = load_data(
        input_path, log_func=log_func)
    # клиники из спецификации должны существовать в текущем файле;
    # если ни одна не совпала — сравниваем все найденные клиники
    known = _clinic_codes(data)
    valid_clinics = [c for c in (spec.get('clinics') or []) if c in known]
    spec = dict(spec)
    spec['clinics'] = valid_clinics
    errs = validate_spec(spec, available_antibiotics=antibiotics)
    if errs:
        raise ValueError(errs[0])
    base = prepare_base_table(df_main, data)
    built = build_clinic_comparison(spec, base, data, antibiotics, ab_indices,
                                    log_func=log_func)
    save_to_excel(pd.DataFrame(), {}, pd.DataFrame(), output_path,
                  log_func=log_func, styled=styled, raw_data=data,
                  user_sheets=built, base_sheets=False,
                  chart_sheets={k: v for k, v in built.items()
                                if '(график)' in k})
    if log_func:
        log_func(f"Сравнение клиник сохранено: {output_path}")
    return built


def save_compare_preset(spec, path):
    """Сохраняет настройки режима «Сравнение клиник» в JSON-файл."""
    import json
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(dict(spec), f, ensure_ascii=False, indent=2)


def load_compare_preset(path):
    """Загружает настройки режима «Сравнение клиник» из JSON-файла."""
    import json
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("Файл пресета не содержит объект настроек.")
    return {
        'name': str(data.get('name', 'Сравнение клиник')),
        'mode': COMPARE_CLINICS_KEY,
        'clinics': [str(c) for c in data.get('clinics', [])],
        'antibiotics': [str(a) for a in data.get('antibiotics', [])],
        'compare_by': [g for g in data.get('compare_by', ['organism'])
                       if g in ('organism', 'material_raw')] or ['organism'],
        'filter': str(data.get('filter', 'none')),
        'metrics': ['ab'],
    }

# ----------------------------------------------------------------------
# 8. Сохранение в Excel
# ----------------------------------------------------------------------

def save_to_excel(sheet2_df, sheet3_dict, sheet4_df, output_file,
                  log_func=None, styled=False, raw_data=None, wound_data=None,
                  user_sheets=None, base_sheets=True, chart_sheets=None):
    """
    Сохраняет аналитические листы в Excel.
    При base_sheets=True первые три листа всегда базовые: «1» (анализ посевов),
    «2» (эпидемиология), «3» (резистентность). Остальные листы добавляются
    конструктором (user_sheets — dict {имя: DataFrame}); при styled=True
    дополнительно создаются справочные листы «Сводка» и «Данные».
    base_sheets=False используется отдельным режимом «Сравнение клиник» —
    базовые листы не создаются.
    chart_sheets — dict {имя листа: DataFrame} для листов с графиками: после
    записи таблицы на такой лист добавляются столбчатые диаграммы openpyxl.
    styled=True — современное оформление (шапки, зебра, заморозка панелей,
    условное форматирование % чувствительности).
    """
    if log_func:
        log_func(f"Сохранение в {output_file}...")
    dirname = os.path.dirname(os.path.abspath(output_file))
    if dirname:
        os.makedirs(dirname, exist_ok=True)

    user_sheets = user_sheets or {}
    chart_sheets = chart_sheets or {}

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if base_sheets:
            sheet2_df.to_excel(writer, sheet_name='1', index=False)

            # Лист 2
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
                sheet3_combined.to_excel(writer, sheet_name='2', index=False, header=False)

            # Лист 3
            if not sheet4_df.empty:
                sheet4_df.to_excel(writer, sheet_name='3', index=False, header=False)

        # Листы конструктора
        for name, df_sheet in user_sheets.items():
            df_sheet.to_excel(writer, sheet_name=name[:31], index=False,
                              header=False if name in chart_sheets else True)

        extra_sheets = list(user_sheets.keys())
        if styled:
            summary_df = _build_summary(raw_data, sheet2_df) if base_sheets else None
            if summary_df is not None:
                summary_df.to_excel(writer, sheet_name='Сводка', index=False)
                extra_sheets.append('Сводка')
            if raw_data is not None and not raw_data.empty:
                _write_raw_sheet(raw_data, writer)
                extra_sheets.append('Данные')
            _apply_styling(writer, sheet2_df, sheet3_dict, sheet4_df, extra_sheets,
                           base_sheets=base_sheets)

        else:
            # Автонастройка ширины столбцов (как в базовой версии)
            for sheetname in writer.sheets:
                worksheet = writer.sheets[sheetname]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except Exception:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # Графики сравнения клиников (листы «… (график)»)
        for name in chart_sheets:
            if name in writer.sheets:
                _add_comparison_charts(writer.sheets[name], log_func=log_func)

    if log_func:
        log_func("Сохранение завершено.")


def _build_summary(raw_data, sheet2_df):
    """Небольшая сводная таблица ключевых показателей для листа «Сводка»."""
    if raw_data is None or getattr(raw_data, 'empty', True):
        return None
    total = len(raw_data)
    pos = int(raw_data['is_positive'].sum())
    diag = int((raw_data['is_positive'] & raw_data['diagnostic']).sum())
    patients = raw_data['patient_id'].nunique()
    dates = raw_data['date'].dropna()
    period = ''
    if not dates.empty:
        period = f"{dates.min():%d.%m.%Y} — {dates.max():%d.%m.%Y}"
    top_org = ''
    orgs = raw_data.loc[raw_data['is_positive'], 'organism'].dropna()
    if not orgs.empty:
        top = orgs.value_counts().head(3)
        top_org = ', '.join(f"{name} ({cnt})" for name, cnt in top.items())
    rows = [
        ('Период', period),
        ('Пациентов', patients),
        ('Всего проб', total),
        ('Положительных проб', pos),
        ('Доля положительных, %', round(pos / total * 100, 1) if total else 0),
        ('В диагностическом титре', diag),
        ('Доля диагностичных, %', round(diag / pos * 100, 1) if pos else 0),
        ('Лидеры по выделению', top_org),
    ]
    return pd.DataFrame(rows, columns=['Показатель', 'Значение'])


def _write_raw_sheet(raw_data, writer):
    """Лист «Данные»: очищенные строки посевов с категорией материала."""
    df = raw_data.copy()
    cats = df['material_raw'].apply(lambda x: categorize_material(x)[0])
    out = pd.DataFrame({
        'Клиника': df['clinic'] if 'clinic' in df.columns else DEFAULT_CLINIC_LABEL,
        'Пациент': df['patient_id'],
        'Дата': df['date'].dt.strftime('%d.%m.%Y'),
        'Материал': df['material_raw'],
        'Категория': cats,
        'Микроорганизм': df['organism'],
        'Рост': df['is_positive'].map({True: '+', False: 'нет'}),
        'Диагностичный': df['diagnostic'].map({True: 'да', False: 'нет'}),
    })
    out.to_excel(writer, sheet_name='Данные', index=False)


def _sheet_has_blocks(ws, max_scan_rows=5):
    """Лист построен блоками («Антибиотик: …») — оформляется как лист 3."""
    try:
        for r in range(1, min(max_scan_rows, ws.max_row) + 1):
            v = ws.cell(row=r, column=1).value
            if v and str(v).startswith(('Антибиотик:', 'Клиника:')):
                return True
    except Exception:
        pass
    return False


def _add_comparison_charts(ws, log_func=None):
    """Добавляет столбчатые диаграммы на лист графика сравнения клиник.

    Структура листа (см. build_clinic_comparison): блок из трёх строк —
    заголовок антибиотика, строка «Клиника | <диапазоны…>», затем по строке
    на клинику со значениями S%. На каждый блок создаётся своя диаграмма
    BarChart с осями категорий (микроорганизм/материал) и сериями-клиниками.
    """
    from openpyxl.chart import BarChart, Reference

    n_charts = 0
    r = 1
    max_row = ws.max_row
    while r <= max_row:
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r + 1, column=1).value if r + 1 <= max_row else None
        if a and b == 'Клиника':
            header_row = r + 1
            rr = r + 2
            last_data = rr - 1
            while last_data + 1 <= max_row:
                v = ws.cell(row=last_data + 1, column=1).value
                if v in (None, ''):
                    break
                last_data += 1
            ncols = ws.max_column
            if last_data >= header_row + 1 and ncols >= 2:
                chart = BarChart()
                chart.type = "col"
                chart.grouping = "clustered"
                chart.title = str(a)
                chart.y_axis.title = 'Чувствительность, %'
                chart.x_axis.title = 'Микроорганизм / материал'
                data = Reference(ws, min_col=2, max_col=ncols,
                                 min_row=header_row, max_row=last_data)
                cats = Reference(ws, min_col=2, max_col=ncols,
                                 min_row=header_row, max_row=header_row)
                chart.add_data(data, titles_from_data=False, from_rows=True)
                chart.set_categories(cats)
                # имена серий — из первого столбца блока
                series_names = []
                for sr in range(header_row + 1, last_data + 1):
                    v = ws.cell(row=sr, column=1).value
                    series_names.append(str(v) if v not in (None, '') else f"Ряд {len(series_names)+1}")
                for s, nm in zip(chart.series, series_names):
                    from openpyxl.chart.series import SeriesLabel
                    from openpyxl.chart.data_source import StrRef
                    s.tx = SeriesLabel(v=nm)
                chart.width = 32
                chart.height = 12
                anchor_col = 1
                anchor_row = last_data + 2 + n_charts * 1
                ws.add_chart(chart, f"{get_column_letter(anchor_col)}{anchor_row}")
                n_charts += 1
                if n_charts >= COMPARE_CHART_MAX_BLOCKS:
                    break
                r = last_data + 2
                continue
            r = last_data + 2
        else:
            r += 1
    if log_func and n_charts:
        log_func(f"Лист «{ws.title}»: добавлено графиков сравнения — {n_charts}.")


def _apply_styling(writer, sheet2_df, sheet3_dict, sheet4_df, extra_sheets,
                   base_sheets=True):
    """Современное оформление книги: шапки, зебра, числовые форматы, панели.

    extra_sheets — имена дополнительных листов (конструктор, «Сводка», «Данные»),
    оформляемых как простые таблицы с шапкой. base_sheets=False — режим
    сравнения клиник: базовые листы 1–3 отсутствуют, но листы-блоки сравнения
    («Антибиотик: …» / «Клиника: …») оформляются по тому же шаблону."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.utils import get_column_letter as gcl

    ACCENT = "1F4E79"      # тёмно-синий (шапки)
    ACCENT_SOFT = "DCE6F1"  # светло-голубой (зебра/секции)
    ZEBRA = "F5F8FC"
    GOOD = "C6EFCE"
    BAD = "FFC7CE"

    thin = Side(style="thin", color="B8C4D4")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=ACCENT)
    zebra_fill = PatternFill("solid", fgColor=ZEBRA)
    section_fill = PatternFill("solid", fgColor=ACCENT_SOFT)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_al = Alignment(horizontal="left", vertical="center")

    def style_header_row(ws, row_idx, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border_all

    def autofit(ws, max_width=50):
        widths = {}
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None:
                    l = len(str(cell.value))
                    widths[cell.column] = max(widths.get(cell.column, 0), l)
        for col, length in widths.items():
            ws.column_dimensions[gcl(col)].width = min(length + 2.5, max_width)

    # --- Сводка / Данные / листы конструктора / лист 1: простые таблицы ------
    block_sheet_names = {n for n in extra_sheets
                         if n in writer.sheets and _sheet_has_blocks(writer.sheets[n])}
    simple_with_header = [s for s in (['Сводка', 'Данные'] + list(extra_sheets) +
                                      (['1'] if base_sheets else []))
                          if s in writer.sheets and s not in block_sheet_names]
    for name in simple_with_header:
        ws = writer.sheets[name]
        ncols = ws.max_column
        style_header_row(ws, 1, ncols)
        for r in range(2, ws.max_row + 1):
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border_all
                if r % 2 == 0:
                    cell.fill = zebra_fill
        # подсветка столбцов процентов (листы конструктора)
        for c in range(1, ncols + 1):
            head = ws.cell(row=1, column=c).value
            if head and str(head).startswith('%'):
                rng = f"{gcl(c)}2:{gcl(c)}{ws.max_row}"
                ws.conditional_formatting.add(rng, CellIsRule(
                    operator='greaterThanOrEqual', formula=['80'],
                    fill=PatternFill(start_color=GOOD, end_color=GOOD, fill_type='solid')))
                ws.conditional_formatting.add(rng, CellIsRule(
                    operator='lessThan', formula=['50'],
                    fill=PatternFill(start_color=BAD, end_color=BAD, fill_type='solid')))
        ws.freeze_panes = "A2"
        autofit(ws)

    # --- Лист 2: секции (заголовок + шапка + строки) ---------------------
    if '2' in writer.sheets and sheet3_dict:
        ws = writer.sheets['2']
        ncols = ws.max_column
        r = 1
        max_row = ws.max_row
        while r <= max_row:
            a = ws.cell(row=r, column=1).value
            if a in (None, ''):
                r += 1
                continue
            b = ws.cell(row=r + 1, column=1).value if r + 1 <= max_row else None
            if b == 'Микроорганизм':
                style_header_row(ws, r + 1, ncols)
                rr = r + 2
                k = 0
                while rr <= max_row:
                    val = ws.cell(row=rr, column=1).value
                    if val in (None, ''):
                        break
                    for c in range(1, ncols + 1):
                        cell = ws.cell(row=rr, column=c)
                        cell.border = border_all
                        if k % 2 == 1:
                            cell.fill = zebra_fill
                    if str(val) == 'ВСЕ':
                        for c in range(1, ncols + 1):
                            ws.cell(row=rr, column=c).font = Font(bold=True)
                    k += 1
                    rr += 1
                # подсветка процентов диагностичности
                for c in range(1, ncols + 1):
                    head = ws.cell(row=r + 1, column=c).value
                    if head and str(head).startswith('%'):
                        rng = f"{gcl(c)}{r + 2}:{gcl(c)}{r + 1 + k}"
                        ws.conditional_formatting.add(rng, CellIsRule(
                            operator='greaterThanOrEqual', formula=['80'],
                            fill=PatternFill(start_color=GOOD, end_color=GOOD, fill_type='solid')))
                        ws.conditional_formatting.add(rng, CellIsRule(
                            operator='lessThan', formula=['50'],
                            fill=PatternFill(start_color=BAD, end_color=BAD, fill_type='solid')))
                ws.cell(row=r, column=1).fill = section_fill
                ws.cell(row=r, column=1).font = Font(bold=True, color=ACCENT, size=12)
                r = rr + 1
            else:
                r += 1
        autofit(ws)

    # --- Лист «3» и листы сравнения клиник: блоки «Антибиотик: ...» -----
    block_names = ['3'] + sorted(block_sheet_names)
    for name in block_names:
        if name not in writer.sheets:
            continue
        if name == '3' and (sheet4_df is None or sheet4_df.empty):
            continue
        ws = writer.sheets[name]
        ncols = ws.max_column
        r = 1
        max_row = ws.max_row
        while r <= max_row:
            val = ws.cell(row=r, column=1).value
            if val and str(val).startswith('Антибиотик:'):
                ws.cell(row=r, column=1).font = Font(bold=True, color=ACCENT, size=12)
                ws.cell(row=r, column=1).fill = section_fill
                if r + 1 <= max_row:
                    style_header_row(ws, r + 1, ncols)
                rr = r + 2
                k = 0
                s_col = None
                while rr <= max_row:
                    v = ws.cell(row=rr, column=1).value
                    if v in (None, '') or str(v).startswith('Антибиотик:'):
                        break
                    for c in range(1, ncols + 1):
                        cell = ws.cell(row=rr, column=c)
                        cell.border = border_all
                        if k % 2 == 1:
                            cell.fill = zebra_fill
                    if s_col is None:
                        for c in range(1, ncols + 1):
                            if ws.cell(row=r + 1, column=c).value == 'S%':
                                s_col = c
                                break
                    k += 1
                    rr += 1
                if s_col:
                    rng = f"{gcl(s_col)}{r + 2}:{gcl(s_col)}{r + 1 + k}"
                    ws.conditional_formatting.add(rng, CellIsRule(
                        operator='greaterThanOrEqual', formula=['80'],
                        fill=PatternFill(start_color=GOOD, end_color=GOOD, fill_type='solid')))
                    ws.conditional_formatting.add(rng, CellIsRule(
                        operator='lessThan', formula=['50'],
                        fill=PatternFill(start_color=BAD, end_color=BAD, fill_type='solid')))
                r = rr
            else:
                r += 1
        autofit(ws)

    # Активный лист — «Сводка», если он есть
    if 'Сводка' in writer.sheets:
        writer.book.active = writer.sheets['Сводка']
