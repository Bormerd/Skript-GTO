from pathlib import Path
from openpyxl import load_workbook, Workbook

# ==========================
# НАСТРОЙКИ
# ==========================

MAIN_FILE = r"г. Кострома Политех.xlsx"
COLLEGES_FOLDER = r"Колледжи"

OUTPUT_FILE = r"г. Кострома Политех_с_колледжами.xlsx"
NOT_FOUND_FILE = r"г. Кострома Политех_НЕ_НАЙДЕНО.xlsx"


# ==========================
# НОРМАЛИЗАЦИЯ
# ==========================

def normalize_uin(value):
    if value is None:
        return ""

    value = str(value).strip()
    value = value.replace(" ", "").replace("-", "")
    value = value.replace("\n", "").replace("\r", "")

    if value.endswith(".0"):
        value = value[:-2]

    return value


def normalize_fio(value):
    if value is None:
        return ""

    value = str(value).strip().lower()
    value = " ".join(value.split())
    return value


# ==========================
# ПОИСК СТОЛБЦОВ
# ==========================

def find_columns(ws):
    uin_col = None
    fio_col = None
    header_row = None

    for row in ws.iter_rows(max_row=60):
        for cell in row:

            if cell.value is None:
                continue

            text = str(cell.value).lower()

            if "уин" in text:
                uin_col = cell.column
                header_row = cell.row

            if "фио" in text:
                fio_col = cell.column
                header_row = cell.row

        if uin_col:
            break

    return uin_col, fio_col, header_row


# ==========================
# ИНДЕКСЫ ПРОТОКОЛОВ
# ==========================

uin_map = {}     # uin -> college
fio_map = {}     # fio -> list(uin)

print("Считывание протоколов...\n")

files_count = 0

for folder in Path(COLLEGES_FOLDER).iterdir():

    if not folder.is_dir():
        continue

    college = folder.name
    print(f"Колледж: {college}")

    for file in folder.glob("*.xlsx"):

        try:
            wb = load_workbook(file, data_only=True)
            files_count += 1

            for ws in wb.worksheets:

                uin_col, fio_col, header_row = find_columns(ws)

                if not uin_col:
                    continue

                for r in range(header_row + 1, ws.max_row + 1):

                    uin = normalize_uin(ws.cell(r, uin_col).value)
                    fio = normalize_fio(ws.cell(r, fio_col).value if fio_col else None)

                    if uin:
                        uin_map[uin] = college

                    if fio:
                        if fio not in fio_map:
                            fio_map[fio] = []
                        fio_map[fio].append(uin)

        except:
            continue


print("\n==============================")
print(f"Файлов обработано: {files_count}")
print(f"УИН в базе: {len(uin_map)}")
print("==============================\n")


# ==========================
# ОСНОВНАЯ КНИГА
# ==========================

wb = load_workbook(MAIN_FILE)

total = 0
found_uin = 0
found_fio = 0
not_found = 0

not_found_rows = []   # для отдельного файла

for sheet in ["Золото", "Серебро", "Бронза"]:

    ws = wb[sheet]

    uin_col, fio_col, header_row = find_columns(ws)

    college_col = ws.max_column + 1
    ws.cell(header_row, college_col, "Колледж")

    for r in range(header_row + 1, ws.max_row + 1):

        total += 1

        uin = normalize_uin(ws.cell(r, uin_col).value)
        fio = normalize_fio(ws.cell(r, fio_col).value if fio_col else None)

        college = None
        status = "НЕ НАЙДЕН"

        # 1. УИН
        if uin in uin_map:
            college = uin_map[uin]
            status = "УИН"
            found_uin += 1

        # 2. ФИО
        elif fio in fio_map:

            uins = fio_map[fio]

            if len(uins) == 1:
                college = uin_map.get(uins[0])
                status = "ФИО"
                found_fio += 1

        if not college:
            college = "НЕ НАЙДЕН"
            not_found += 1

            # сохраняем полную строку в отчёт
            row_data = [
                sheet,
                fio,
                uin
            ]
            not_found_rows.append(row_data)

        ws.cell(r, college_col, f"{college} ({status})")


# ==========================
# СОХРАНЕНИЕ ОСНОВНОГО ФАЙЛА
# ==========================

wb.save(OUTPUT_FILE)


# ==========================
# ОТДЕЛЬНЫЙ ФАЙЛ НЕ НАЙДЕНО
# ==========================

report = Workbook()
ws_rep = report.active
ws_rep.title = "НЕ НАЙДЕНО"

ws_rep.append(["Лист", "ФИО", "УИН"])

for row in not_found_rows:
    ws_rep.append(row)

report.save(NOT_FOUND_FILE)


# ==========================
# КОНСОЛЬНЫЙ ОТЧЁТ
# ==========================

print("\n==============================")
print("ГОТОВО")
print("==============================")

print(f"Всего записей: {total}")
print(f"По УИН найдено: {found_uin}")
print(f"По ФИО найдено: {found_fio}")
print(f"Не найдено: {not_found}")

print("\nПервые 30 НЕ НАЙДЕННЫХ:")

for r in not_found_rows[:30]:
    print(r)

print("\nФайлы сохранены:")
print("Основной:", OUTPUT_FILE)
print("Не найдено:", NOT_FOUND_FILE)