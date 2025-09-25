# -*- coding: utf-8 -*-
"""
Приложение "Импорт анкет поставщиков в реестр"
- Окно с полем для перетаскивания файлов (или выбора через кнопку).
- Запоминает путь к реестру (можно изменить в настройках).
- На каждую анкету создаёт одну или несколько строк (по каждой выбранной ТГ).
- Поле "Наличие складка (да/нет)" определяется по диапазону B5:E29 на листе "Пр.№ 2 Производств. мощности ".
- Есть настройка fallback ТГ (если "галочки" в анкете не читаются).
Сборка EXE: см. инструкцию внизу файла.
"""
import json
import os
from pathlib import Path
import re
import traceback
import PySimpleGUI as sg
import pandas as pd

# ------------------ Константы и настройки ------------------
APP_NAME = "Импорт анкет поставщиков"
CONFIG_DIR = Path.home() / ".supplier_import_app"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = CONFIG_DIR / "config.json"

SHEET_FORM = "Анкета поставщика"
SHEET_TG = "Пр.№ 1 Товарн. категория "
SHEET_CAP = "Пр.№ 2 Производств. мощности "

TRUTHY_MARKS = {"x", "х", "✓", "да", "true", "1", "+", "y", "yes"}

FIELD_KEYS = {
    "Полное наименование организации": ["полное наименование", "наименование организации"],
    "Сокращенное наименование": ["сокращенное наименование", "торговая марка"],
    "ИНН": ["инн"],
    "Контактное лицо": ["контактное лицо"],
    "Должность": ["должност", "должность"],
    "Телефон": ["телефон"],
    "E-mail": ["e-mail", "email", "электронной"],
    "Регион(поле)": ["регион (астрахань/екатеринбург)"],
    "Город": ["город"],
    "Система налогообложения": ["система налогооблож", "с ндс", "без ндс"],
    "ФИО": ["ф.и.о"],
}

# ------------------ Работа с конфигом ------------------
def load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "registry_path": "",
        "fallback_tgs": []
    }

def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

# ------------------ Утилиты парсинга ------------------
def normalize_tax(value: str) -> str:
    s = (value or "").strip().lower()
    if not s:
        return ""
    if "без" in s and "ндс" in s:
        return "Без НДС"
    if "с ндс" in s or ("ндс" in s and ("с " in s or "включ" in s or "вкл" in s)):
        return "С НДС"
    return value.strip()

def read_excel_sheet(path: Path, sheet: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet, header=None)

def extract_form_fields(xlsx_path: Path) -> dict:
    df = read_excel_sheet(xlsx_path, SHEET_FORM).fillna("")
    rows = df.values.tolist()
    pairs = []
    for r in rows:
        label = ""
        if len(r) > 1 and str(r[1]).strip():
            label = str(r[1]).strip()
        elif len(r) > 0 and str(r[0]).strip():
            label = str(r[0]).strip()
        value = str(r[2]).strip() if len(r) > 2 else ""
        if label or value:
            pairs.append((label.lower(), value))

    def find_val(keys):
        for lab, val in pairs:
            if any(k in lab for k in keys):
                if val and val.lower() != "nan":
                    return str(val).strip()
        return ""

    data = {}
    for field, keys in FIELD_KEYS.items():
        data[field] = find_val([k.lower() for k in keys])

    contact = data.get("ФИО") or data.get("Контактное лицо") or ""
    data["Контактное лицо"] = contact

    if not data.get("Телефон"):
        for lab, val in pairs:
            if "телефон" in lab:
                data["Телефон"] = val
                break

    data["Система налогообложения"] = normalize_tax(data.get("Система налогообложения", ""))

    region = data.get("Регион(поле)") or ""
    if not region:
        city = (data.get("Город") or "").lower()
        if "астрахан" in city:
            region = "Астрахань"
        elif "екатерин" in city:
            region = "Екатеринбург"
    data["Регион"] = region

    org_name = data.get("Сокращенное наименование") or data.get("Полное наименование организации") or ""
    data["ОргНазваниеДляРеестра"] = org_name
    return data

def extract_tk_tg(xlsx_path: Path, fallback_tgs=None):
    result = []
    try:
        df = read_excel_sheet(xlsx_path, SHEET_TG).fillna("")
    except Exception:
        return result

    mark_col = None
    header_row = None
    for i in range(min(20, len(df))):
        row = [str(x) for x in df.iloc[i].tolist()]
        if any("отметить галочкой" in c.lower() for c in row):
            header_row = i
            for j, c in enumerate(row):
                if "отметить галочкой" in c.lower():
                    mark_col = j
                    break
            break

    tk_col = None
    tg_col = None
    if header_row is not None:
        hdr = [str(x) for x in df.iloc[header_row].tolist()]
        for j, c in enumerate(hdr):
            cl = c.lower()
            if "категор" in cl and tk_col is None:
                tk_col = j
            if ("группа" in cl or "тг" in cl) and tg_col is None:
                tg_col = j

    if mark_col is not None and tk_col is not None and tg_col is not None:
        for i in range(header_row+1, len(df)):
            tk = str(df.iat[i, tk_col]).strip()
            tg = str(df.iat[i, tg_col]).strip()
            mark = str(df.iat[i, mark_col]).strip().lower()
            if not tk and not tg:
                continue
            if mark in TRUTHY_MARKS:
                result.append((tk, tg))

    if not result and fallback_tgs:
        for i in range(len(df)):
            row = [str(x).strip() for x in df.iloc[i].tolist()]
            for j, cell in enumerate(row):
                if any(cell.lower() == tg.lower() for tg in fallback_tgs):
                    tk_guess = row[j-1] if j-1 >= 0 else ""
                    result.append((tk_guess, cell))

    result = [(tk, tg) for tk, tg in result if tg]
    seen = set(); uniq = []
    for pair in result:
        if pair not in seen:
            seen.add(pair); uniq.append(pair)
    return uniq

def has_warehouse(xlsx_path: Path) -> bool:
    try:
        cap = read_excel_sheet(xlsx_path, SHEET_CAP)
    except Exception:
        return False
    rng = cap.iloc[4:29, 1:5]  # B5:E29
    return bool(rng.notna().any().any())

def find_registry_header_row(df: pd.DataFrame) -> int:
    for i in range(min(15, len(df))):
        row = df.iloc[i].astype(str).tolist()
        if "№" in row and any("ИНН" in c for c in row):
            return i
    raise RuntimeError("Не найден заголовок в реестре.")

def next_number(data_rows: pd.DataFrame) -> int:
    nums = []
    for v in data_rows.iloc[:, 1]:
        try:
            nums.append(int(float(str(v).replace(",", "."))))
        except:
            pass
    return (max(nums) + 1) if nums else 1

def process_files(files, registry_path: Path, fallback_tgs):
    logs = []
    # Загружаем реестр
    reg_raw = pd.read_excel(registry_path, sheet_name="Реестр поставщиков", header=None)
    hdr_idx = find_registry_header_row(reg_raw)
    headers = reg_raw.iloc[hdr_idx].tolist()
    idx = {str(h): i for i, h in enumerate(headers)}

    data_rows = reg_raw.iloc[hdr_idx+1:]
    def valid_row(row):
        v = row.iloc[1]
        try:
            float(str(v).replace(",", ".")); return True
        except: return False
    existing = data_rows[data_rows.apply(valid_row, axis=1)]
    seq = next_number(existing)

    append_rows = []

    for file in files:
        x = Path(file)
        if not x.suffix.lower().endswith("xlsx"):
            logs.append(f"[SKIP] {x.name}: не .xlsx")
            continue
        if x.name.lower().startswith("реестр"):
            logs.append(f"[SKIP] {x.name}: похоже на реестр, пропускаю")
            continue
        try:
            fields = extract_form_fields(x)
        except Exception as e:
            logs.append(f"[WARN] {x.name}: ошибка чтения анкеты: {e}")
            continue
        try:
            tk_tg = extract_tk_tg(x, fallback_tgs=fallback_tgs)
        except Exception as e:
            logs.append(f"[WARN] {x.name}: ошибка чтения ТК/ТГ: {e}")
            tk_tg = []

        if not tk_tg:
            tk_tg = [("", "")]

        wh = "ДА" if has_warehouse(x) else "НЕТ"

        for tk, tg in tk_tg:
            row = [None] * len(headers)
            def setv(col_name, value):
                i = idx.get(col_name)
                if i is not None and i < len(row):
                    row[i] = value

            setv("№", seq); seq += 1
            setv("Полное наименование организации:", fields.get("ОргНазваниеДляРеестра"))
            setv("ИНН:", fields.get("ИНН") or "")
            setv("Контактное лицо:", fields.get("Контактное лицо") or "")
            setv("Должность:", fields.get("Должность") or "")
            setv("Телефон", fields.get("Телефон") or "")
            setv("E-mail:", fields.get("E-mail") or "")
            setv("Система налогообложения: указать с НДС или без НДС", fields.get("Система налогообложения") or "")
            setv("Регион (Астрахань/Екатеринбург)", fields.get("Регион") or "")
            setv("ТОВАРНАЯ КАТЕГОРИЯ (ТК)", tk or "")
            setv("ТОВАРНАЯ ГРУППА (ТГ)", tg or "")
            setv("Наличие складка (да/нет)", wh)
            append_rows.append(row)

        logs.append(f"[OK] {x.name}: строк добавлено — {len(tk_tg)}")

    if not append_rows:
        logs.append("Новые строки не добавлены.")
        return logs, None

    out_df = pd.concat([
        reg_raw.iloc[:hdr_idx+1],
        pd.DataFrame(append_rows),
        reg_raw.iloc[hdr_idx+1:]
    ], ignore_index=True)

    # Сохраняем прямо в исходный реестр
    with pd.ExcelWriter(registry_path, engine="openpyxl") as w:
        pd.DataFrame(out_df).to_excel(w, index=False, header=False, sheet_name="Реестр поставщиков")

    logs.append(f"Готово: обновлён файл {registry_path}")
    return logs, registry_path

# ------------------ GUI ------------------
def make_window(cfg):
    sg.theme("SystemDefault")
    layout = [
        [sg.Text(APP_NAME, font=("Segoe UI", 16, "bold"))],
        [sg.Text("Путь к реестру:"), 
         sg.Input(cfg.get("registry_path",""), key="-REG-", expand_x=True, enable_events=True),
         sg.FileBrowse("Выбрать...", file_types=(("Excel", "*.xlsx"),))],
        [sg.Text("Fallback ТГ (через ; ):"), 
         sg.Input("; ".join(cfg.get("fallback_tgs", [])), key="-FB-", expand_x=True)],
        [sg.Frame("Перетащите сюда анкеты (*.xlsx) или выберите кнопкой ниже", [
            [sg.Multiline(size=(80,10), key="-DROP-", disabled=True, autoscroll=True, no_scrollbar=False)],
        ], expand_x=True)],
        [sg.FilesBrowse("Выбрать файлы...", key="-SELECT-", file_types=(("Excel", "*.xlsx"),), enable_events=True),
         sg.Push(), sg.Button("Импортировать", key="-IMPORT-", size=(16,1), button_color=("white","#0078D7"))],
        [sg.Output(size=(100,8), key="-LOG-")],
        [sg.Button("Сохранить настройки"), sg.Button("Выход")]
    ]
    return sg.Window(APP_NAME, layout, finalize=True, resizable=True)

def main_gui():
    cfg = load_config()
    win = make_window(cfg)

    dropped_files = []

    while True:
        event, values = win.read()
        if event in (sg.WINDOW_CLOSED, "Выход"):
            break

        if event == "-SELECT-":
            files = values["-SELECT-"]
            if files:
                # Может прийти строка с ; в качестве разделителя
                file_list = [f.strip() for f in files.split(";") if f.strip()]
                dropped_files.extend(file_list)
                win["-DROP-"].update("\n".join(dropped_files))

        if event == "-IMPORT-":
            reg = values.get("-REG-", "").strip()
            fb_raw = values.get("-FB-", "").strip()
            fallback = [s.strip() for s in fb_raw.split(";") if s.strip()] if fb_raw else []
            if not reg:
                print("Укажите путь к реестру.")
                continue
            if not Path(reg).exists():
                print("Файл реестра не найден:", reg)
                continue
            if not dropped_files:
                print("Не выбраны файлы анкет.")
                continue
            try:
                logs, _ = process_files(dropped_files, Path(reg), fallback)
                for line in logs:
                    print(line)
                # Очистим список после успешного импорта
                dropped_files = []
                win["-DROP-"].update("")
            except Exception as e:
                print("Ошибка импорта:", e)
                traceback.print_exc()

        if event == "Сохранить настройки":
            cfg["registry_path"] = values.get("-REG-", "").strip()
            fb_raw = values.get("-FB-", "").strip()
            cfg["fallback_tgs"] = [s.strip() for s in fb_raw.split(";") if s.strip()] if fb_raw else []
            save_config(cfg)
            print("Настройки сохранены:", CONFIG_PATH)

    win.close()

if __name__ == "__main__":
    main_gui()

"""
Установка и запуск
1) Установите Python 3.10+ и пакеты:
   pip install pandas openpyxl PySimpleGUI

2) Запуск приложения:
   python supplier_import_app.py

3) Сборка .EXE (Windows) с PyInstaller (опционально):
   pip install pyinstaller
   pyinstaller --noconsole --onefile --name SupplierImportApp supplier_import_app.py

   EXE появится в папке dist\SupplierImportApp.exe
"""
