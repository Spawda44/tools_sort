"""
Сбор уникальных инструментов из всех листов Excel-файла в одну таблицу.

Логика в двух словах:
1. На каждом листе ищем таблицы, у которых в колонке A стоит заголовок "Название".
2. Читаем строки таблицы (A..I) до первой пустой строки / стоп-фразы.
3. Дубликаты (совпадение по всем характеристикам) схлопываем в одну строку
   и считаем, сколько раз инструмент встретился.
4. Результат сортируем, пишем на отдельный лист и красиво форматируем.
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

STOP_PHRASE = ["Список калибров и мерительного инструмента", "Список калибров"]
HEADER_NAME = "Название"
DATA_COLS = 9  # A..I: Название + 7 характеристик + Комментарий

COLUMNS = [
    "Название", "Вид Фрезы", "Диаметр раб.ч", "Длина обнижения",
    "Диаметр хвостовика", "Длина рабочая", "Длина общая",
    "Кол-во зубьев", "Комментарий", "Количество",
]
# по этим полям сравниваем и сортируем строки (без "Комментарий" и "Количество")
KEY_COLS = COLUMNS[:8]

# --- стили оформления итоговой таблицы ---
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY_FONT = Font(name="Calibri", size=11)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
STRIPE_FILL = PatternFill("solid", fgColor="EAF0FB")
BORDER = Border(*[Side(style="thin", color="B7C6E5")] * 4)


def norm_text(x):
    """Приводит значение ячейки к сравнимому виду: без NBSP, лишних пробелов, в нижнем регистре."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).replace("\u00A0", " ")
    return re.sub(r"\s+", " ", s.strip()).lower()


def clear_output_sheet(out_path, out_sheet):
    """Готовит пустой лист под результат (создаёт файл/лист, если их ещё нет)."""
    file_exists = os.path.exists(out_path)
    wb = load_workbook(out_path) if file_exists else Workbook()

    if out_sheet in wb.sheetnames:
        wb[out_sheet].delete_rows(1, wb[out_sheet].max_row)
    elif not file_exists:
        # новый файл: переименовываем автосозданный лист по умолчанию
        wb.active.title = out_sheet
    else:
        wb.create_sheet(out_sheet)

    wb.save(out_path)


def row_has_stop_phrase(ws, row_idx):
    """Проверяет, не начинается ли с этой строки список калибров (конец нужных данных)."""
    stop_norm = {norm_text(p) for p in STOP_PHRASE}
    return any(norm_text(ws.cell(row_idx, c).value) in stop_norm for c in range(1, 5))


def get_merged_rows(ws):
    """Строки, попадающие в объединённые ячейки A..I — их пропускаем (это заголовки/шапки таблиц)."""
    merged = set()
    for rng in ws.merged_cells.ranges:
        min_r, min_c, max_r, max_c = rng.bounds
        if min_c <= DATA_COLS:
            merged.update(range(min_r, max_r + 1))
    return merged


def collect_tables_from_sheet(ws):
    """Находит на листе все таблицы под заголовком 'Название' и возвращает их строки."""
    rows = []
    merged_rows = get_merged_rows(ws)

    r = 1
    while r <= ws.max_row:
        if r in merged_rows:
            r += 1
            continue
        if row_has_stop_phrase(ws, r):
            break

        # нашли заголовок таблицы — читаем строки под ним
        if norm_text(ws.cell(r, 1).value) == norm_text(HEADER_NAME):
            i = r + 1
            while i <= ws.max_row:
                if i in merged_rows:
                    i += 1
                    continue
                if row_has_stop_phrase(ws, i):
                    i = ws.max_row + 1
                    break

                a = ws.cell(i, 1).value
                if norm_text(a) == norm_text(HEADER_NAME):  # вложенный заголовок — пропускаем
                    i += 1
                    continue
                if norm_text(a) == "":  # пустая строка — конец таблицы
                    break

                rows.append(tuple(ws.cell(i, c).value for c in range(1, DATA_COLS + 1)))
                i += 1
            r = i
        else:
            r += 1

    return rows


def collect_unique_tools(in_path, out_sheet):
    """Проходит по всем листам книги и схлопывает повторяющиеся инструменты, считая их количество."""
    wb = load_workbook(in_path, data_only=True)
    first_seen = {}  # ключ -> исходная строка
    counts = {}       # ключ -> число повторов

    for sheet_name in wb.sheetnames:
        if sheet_name == out_sheet:
            continue
        for row in collect_tables_from_sheet(wb[sheet_name]):
            name, *specs, comment = row
            key = "|".join(norm_text(v) for v in (name, *specs))
            if key not in first_seen:
                first_seen[key] = row
                counts[key] = 1
            else:
                counts[key] += 1

    return [(*first_seen[k], counts[k]) for k in first_seen]


def build_dataframe(data):
    """Собирает и сортирует итоговую таблицу."""
    df = pd.DataFrame(data, columns=COLUMNS)
    if df.empty:
        return df

    # сортируем по нормализованному названию, но сам текст в ячейках не трогаем
    sort_key = df["Название"].apply(norm_text)
    return (
        df.assign(_sort=sort_key)
        .sort_values(["_sort", *KEY_COLS[1:]], kind="mergesort")
        .drop(columns="_sort")
    )


def style_worksheet(ws, n_cols, count_col_idx):
    """Оформляет шапку, тело таблицы (рамки + зебра) и закрепляет верхнюю строку."""
    for c in range(1, n_cols + 1):
        cell = ws.cell(1, c)
        cell.fill, cell.font, cell.alignment, cell.border = HEADER_FILL, HEADER_FONT, HEADER_ALIGN, BORDER

    for r in range(2, ws.max_row + 1):
        row_fill = STRIPE_FILL if r % 2 == 0 else None
        for c in range(1, n_cols + 1):
            cell = ws.cell(r, c)
            cell.font, cell.border = BODY_FONT, BORDER
            cell.alignment = CENTER_ALIGN if c == count_col_idx else BODY_ALIGN
            if row_fill:
                cell.fill = row_fill

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def autofit_columns(ws, n_cols, min_width=10, max_width=60, padding=4):
    """Подгоняет ширину каждой колонки под самую длинную строку текста в ней."""
    for c in range(1, n_cols + 1):
        longest = max(
            (len(line) for r in range(1, ws.max_row + 1)
             for line in str(ws.cell(r, c).value or "").split("\n")),
            default=0,
        )
        ws.column_dimensions[get_column_letter(c)].width = min(max(longest + padding, min_width), max_width)


def main(in_path, out_path, out_sheet="Уникальные_инструменты"):
    # 1. подготавливаем пустой лист под результат
    clear_output_sheet(out_path, out_sheet)
    print('Подготовка пустой таблицы для сортировки...\n')

    # 2. собираем уникальные инструменты со всех листов исходной книги
    data = collect_unique_tools(in_path, out_sheet)
    df = build_dataframe(data)
    print('Чтение и сбор уникального инструмента...\n')

    # 3. записываем таблицу
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=out_sheet, index=False)
    print('Запись инструмента в таблицу...\n')
    
    # 4. оформляем: рамки, зебра, автоподбор ширины колонок
    wb = load_workbook(out_path)
    ws = wb[out_sheet]
    count_col_idx = COLUMNS.index("Количество") + 1
    style_worksheet(ws, len(COLUMNS), count_col_idx)
    autofit_columns(ws, len(COLUMNS))
    wb.save(out_path)
    print('Оформление готовой таблицы...\n')
    print('Готово!')

if __name__ == "__main__":
    main("Название_книги_со_списком_инструмента.xlsx", "Уникальный инструмент.xlsx")