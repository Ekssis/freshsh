
“””
Fresh Auto — Генератор документов с искусственным ПВ
Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках
“””

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, re, subprocess, sys
from pathlib import Path
from datetime import datetime

try:
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import num2words as nw
except ImportError:
subprocess.check_call([sys.executable, “-m”, “pip”, “install”,
“python-docx”, “num2words”, “-q”])
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import num2words as nw

CONFIG_FILE = Path(**file**).parent / “config.json”

# ══════════════════════════════════════════════════════════════════

# УТИЛИТЫ ДЛЯ РАБОТЫ С СУММАМИ

# ══════════════════════════════════════════════════════════════════

def _rub_word(n: int) -> str:
“”“Правильное склонение слова ‘рубль’”””
n100, n10 = n % 100, n % 10
if 11 <= n100 <= 19:
return “рублей”
if n10 == 1:
return “рубль”
if 2 <= n10 <= 4:
return “рубля”
return “рублей”

def format_amount(amount: float) -> str:
“”“2138000 → ‘2 138 000,00’”””
rub = int(amount)
kop = round((amount - rub) * 100)
return f”{rub:,}”.replace(”,”, “ “) + f”,{kop:02d}”

def amount_to_words(amount: float) -> str:
“”“2138000 → ‘Два миллиона сто тридцать восемь тысяч рублей 00 копеек’”””
rub = int(amount)
kop = round((amount - rub) * 100)
words = nw.num2words(rub, lang=“ru”)
words = words[0].upper() + words[1:]
return f”{words} {_rub_word(rub)} {kop:02d} копеек”

def parse_amount(text: str) -> float:
“””‘2 000 000,00’ или ‘2000000’ → 2000000.0”””
cleaned = re.sub(r”[^\d,.]”, “”, text.strip())
cleaned = cleaned.replace(”,”, “.”)
# Если несколько точек — берём последнюю как десятичную
parts = cleaned.split(”.”)
if len(parts) > 2:
cleaned = “”.join(parts[:-1]) + “.” + parts[-1]
try:
return float(cleaned)
except ValueError:
return 0.0

# ══════════════════════════════════════════════════════════════════

# РАБОТА С DOCX — НИЗКОУРОВНЕВЫЕ УТИЛИТЫ

# ══════════════════════════════════════════════════════════════════

def _iter_paragraphs(doc: Document):
“”“Все параграфы: тело + таблицы + колонтитулы”””
yield from doc.paragraphs
for tbl in doc.tables:
for row in tbl.rows:
for cell in row.cells:
yield from cell.paragraphs
for sec in doc.sections:
for hf in (sec.header, sec.footer):
if hf:
yield from hf.paragraphs
for tbl in hf.tables:
for row in tbl.rows:
for cell in row.cells:
yield from cell.paragraphs

def _para_text(para) -> str:
“”“Полный текст параграфа из всех runs”””
return “”.join(r.text for r in para.runs)

def _replace_in_para(para, old: str, new: str) -> bool:
“””
Заменяет old→new в параграфе, сохраняя форматирование.
Стратегия: сначала пробуем в каждом run по отдельности,
если не получилось — объединяем текст всех runs в первый.
“””
# Быстрый путь: текст целиком в одном run
for run in para.runs:
if old in run.text:
run.text = run.text.replace(old, new)
return True

```
# Медленный путь: текст разбит по runs
full = _para_text(para)
if old not in full:
    return False
new_full = full.replace(old, new)
if para.runs:
    para.runs[0].text = new_full
    for r in para.runs[1:]:
        r.text = ""
return True
```

def replace_all(doc: Document, replacements: dict):
“”“Заменяет все ключи словаря на значения во всём документе”””
for para in _iter_paragraphs(doc):
full = _para_text(para)
for old, new in replacements.items():
if old in full:
_replace_in_para(para, old, new)
full = _para_text(para)   # обновляем после замены

def regex_replace_all(doc: Document, pattern: str, repl, flags=0):
“”“Regex-замена во всём документе”””
for para in _iter_paragraphs(doc):
full = _para_text(para)
new_full = re.sub(pattern, repl, full, flags=flags)
if new_full != full and para.runs:
para.runs[0].text = new_full
for r in para.runs[1:]:
r.text = “”

# ══════════════════════════════════════════════════════════════════

# ПАРСИНГ СЧЁТА №1

# ══════════════════════════════════════════════════════════════════

def extract_invoice_data(doc: Document) -> dict:
“”“Извлекает данные из Счёта №1 для авто-заполнения полей”””
data = {}

```
for para in _iter_paragraphs(doc):
    t = para.text.strip()

    # Номер ДКП и дата из строки "По договору: Продажа т/с № ФКП-20/06/26-8 ... от 20.06.26"
    m = re.search(
        r"[Пп]родажа\s+[тТ][/\\][сС]\s+№\s*([\w\-/]+).*?от\s+(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))",
        t)
    if m and "dkp_number" not in data:
        data["dkp_number"] = m.group(1).strip()
        d = m.group(2)
        # 20.06.26 → 20.06.2026
        parts = d.split(".")
        if len(parts[2]) == 2:
            parts[2] = "20" + parts[2]
        data["date"] = ".".join(parts)

    # ФИО покупателя: "Покупатель: ИНН ..., Камбаров Ханлар Владимирович"
    m = re.search(
        r"[Пп]окупатель[:\s]+(?:ИНН\s+\d+[,\s]+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)",
        t)
    if m and "buyer_name" not in data:
        data["buyer_name"] = m.group(1).strip()

    # Данные авто из строки таблицы: "HYUNDAI ELANTRA серый № C176OO193 VIN LBECNAFD6MZ128269"
    m = re.search(
        r"([A-Z]{2,}\s+[A-Z]{2,})\s+([а-яё]+)\s+№\s*([A-ZА-ЯЁa-zA-Zа-яё0-9]+)\s+VIN\s+([A-Z0-9]{17})",
        t)
    if m and "car_vin" not in data:
        data["car_brand"] = m.group(1).strip()
        data["car_color"] = m.group(2).strip()
        data["car_reg"]   = m.group(3).strip()
        data["car_vin"]   = m.group(4).strip()

# Дополнительный поиск по таблицам
for tbl in doc.tables:
    for row in tbl.rows:
        row_text = " ".join(c.text.strip() for c in row.cells)

        if "car_vin" not in data:
            m = re.search(r"VIN\s+([A-Z0-9]{17})", row_text)
            if m:
                data["car_vin"] = m.group(1)

        if "car_brand" not in data:
            m = re.search(r"([A-Z]{2,}\s+[A-Z]{2,})", row_text)
            if m:
                data["car_brand"] = m.group(1)

        if "car_color" not in data:
            m = re.search(
                r"\b(серый|белый|чёрный|черный|синий|красный|серебристый"
                r"|золотистый|коричневый|зелёный|зеленый|бежевый|жёлтый|желтый)\b",
                row_text, re.IGNORECASE)
            if m:
                data["car_color"] = m.group(1).lower()

        if "car_reg" not in data:
            m = re.search(r"№\s*([A-ZА-ЯЁ]{1,2}\d{3}[A-ZА-ЯЁ]{2}\d{2,3})", row_text)
            if m:
                data["car_reg"] = m.group(1)

return data
```

# ══════════════════════════════════════════════════════════════════

# ОБРАБОТКА ДОКУМЕНТОВ

# ══════════════════════════════════════════════════════════════════

# Паттерн для поиска суммы цифрами в тексте (2 000 000,00 или 2000000,00)

_AMOUNT_PAT = r”\d[\d\s]*\d[,.]\d{2}”

# Паттерн для суммы прописью

_WORDS_PAT  = (r”[А-ЯЁ][а-яёА-ЯЁ\s]+”
r”(?:тысяч|миллион|миллиард)[а-яё\s]*”
r”(?:рубл[а-яё]+)\s+\d{2}\s+копеек”)

def _remove_nds(text: str) -> str:
“”“Убирает все упоминания НДС из строки”””
text = re.sub(r”,?\s*в\s+т.?\s*ч.?\s*НДС[^.;\n]*”, “”, text)
text = re.sub(r”(22/122%?)\s*[\d\s,]+руб.?”, “”, text)
text = re.sub(r”НДС\s*(22/122%?)\s*[\d\s,]+руб.?”, “без НДС”, text)
return text

def process_dkp(doc: Document, p: dict) -> Document:
“””
ДКП:
- Меняет цену ТС на doc_price (цифрами и прописью)
- Убирает НДС везде
- Добавляет/обновляет пункт про ПВ
- В Акте приёма-передачи тоже меняет стоимость
“””
new_price  = p[“new_price”]
pv_amount  = p[“pv_amount”]
new_str    = format_amount(new_price)
new_words  = amount_to_words(new_price)
pv_str     = format_amount(pv_amount)
pv_words   = amount_to_words(pv_amount)

```
pv_para_found = False

for para in _iter_paragraphs(doc):
    full = _para_text(para)

    # Заменяем суммы цифрами
    if re.search(_AMOUNT_PAT, full):
        new_full = re.sub(_AMOUNT_PAT, new_str, full)
        new_full = _remove_nds(new_full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""

    # Заменяем суммы прописью
    elif re.search(_WORDS_PAT, full):
        new_full = re.sub(_WORDS_PAT, new_words, full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""

    # Убираем НДС в строках без суммы
    elif re.search(r"НДС|22/122", full):
        new_full = _remove_nds(full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""

    # Проверяем наличие пункта про ПВ
    if "первоначальный взнос" in full.lower():
        pv_para_found = True
        # Обновляем сумму в существующем пункте
        new_full = re.sub(_AMOUNT_PAT, pv_str, full)
        new_full = re.sub(_WORDS_PAT, pv_words, new_full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""

# Если пункта про ПВ нет — добавляем после параграфа с ценой
if not pv_para_found:
    pv_text = (f"Первоначальный взнос по оплате цены Договора составляет "
               f"{pv_str} руб ({pv_words}).")
    for para in doc.paragraphs:
        if re.search(r"размере\s+" + _AMOUNT_PAT, _para_text(para)):
            # Вставляем новый параграф следующим
            new_p = OxmlElement("w:p")
            new_r = OxmlElement("w:r")
            # Копируем стиль из текущего параграфа
            if para.runs:
                rpr = para.runs[0]._r.find(qn("w:rPr"))
                if rpr is not None:
                    from copy import deepcopy
                    new_r.append(deepcopy(rpr))
            new_t = OxmlElement("w:t")
            new_t.text = pv_text
            new_r.append(new_t)
            new_p.append(new_r)
            para._element.addnext(new_p)
            break

return doc
```

def process_pko(doc: Document, p: dict) -> Document:
“””
ПКО:
- Дата (все вхождения ДД.ММ.ГГГГ)
- ФИО покупателя
- Сумма = ПВ (цифрами и прописью)
- Основание из данных авто
- Убирает НДС
“””
pv_str   = format_amount(p[“pv_amount”])
pv_words = amount_to_words(p[“pv_amount”])
date     = p[“date”]
buyer    = p[“buyer_name”]
osnov    = (f”По ДКП №{p[‘dkp_number’]} от {date} “
f”за а/м {p[‘car_brand’]} {p[‘car_color’]} “
f”№{p[‘car_reg’]} VIN {p[‘car_vin’]}”)

```
for para in _iter_paragraphs(doc):
    full = _para_text(para)
    if not full.strip():
        continue

    new_full = full

    # Основание — целиком заменяем строку
    if re.search(r"[Пп]о\s+ДКП|за\s+а/м|VIN\s+[A-Z0-9]{17}", full):
        new_full = osnov
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # Сумма прописью — целиком заменяем строку
    if re.search(_WORDS_PAT, full):
        new_full = re.sub(_WORDS_PAT, pv_words, full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # НДС — строку целиком заменяем на "Без НДС" (приоритет выше суммы)
    if re.search(r"НДС|22/122", full):
        new_full = re.sub(r"НДС\s*\(22/122%?\)[^\n]*", "Без НДС", full)
        new_full = re.sub(r"В том числе НДС[^\n]*", "Без НДС", new_full)
        new_full = _remove_nds(new_full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # Сумма цифрами (строка только с суммой, без даты)
    if re.search(_AMOUNT_PAT, full) and not re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
        new_full = re.sub(_AMOUNT_PAT, pv_str, full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # Дата — только если строка выглядит как дата или содержит дату без суммы
    if re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
        new_full = re.sub(r"\b\d{2}\.\d{2}\.\d{4}\b", date, full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # ФИО
    m = re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+", full)
    if m and m.group(0) != buyer:
        new_full = full.replace(m.group(0), buyer)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""
        continue

    # НДС — строка целиком заменяется на "Без НДС"
    if re.search(r"НДС|22/122", full):
        new_full = re.sub(r"НДС\s*\(22/122%?\)[^\n]*", "Без НДС", full)
        new_full = re.sub(r"В том числе НДС[^\n]*", "Без НДС", new_full)
        new_full = _remove_nds(new_full)
        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]: r.text = ""

return doc
```

def process_invoice(doc: Document, p: dict, amount: float) -> Document:
“””
Счёт (№1 или №2):
- Сумма цифрами и прописью
- Дата
- НДС убираем везде: ставка→’Без НДС’, сумма НДС→0,00
- Итого = amount
“””
amt_str   = format_amount(amount)
amt_words = amount_to_words(amount)
date      = p[“date”]

```
# Сначала обрабатываем таблицы точечно
for tbl in doc.tables:
    # Определяем индекс колонки "Сумма НДС" по заголовку первой строки
    nds_sum_col = -1
    if len(tbl.rows) > 0:
        for ci, hcell in enumerate(tbl.rows[0].cells):
            if re.search(r"[Сс]умма\s*НДС|НДС\s*[Сс]умма", hcell.text):
                nds_sum_col = ci
                break
        # Если заголовка нет — ищем ячейку со ставкой НДС (22/122) в первой строке данных
        if nds_sum_col == -1 and len(tbl.rows) > 1:
            data_row = tbl.rows[1].cells
            nds_rate_col = next((ci for ci, c in enumerate(data_row)
                                 if re.search(r"22/122", c.text)), -1)
            if nds_rate_col >= 0 and nds_rate_col + 1 < len(data_row):
                nds_sum_col = nds_rate_col + 1

    for ri, row in enumerate(tbl.rows):
        cells = row.cells
        for j, cell in enumerate(cells):
            ct = cell.text.strip()

            # Ставка НДС
            if re.search(r"22/122", ct):
                for para in cell.paragraphs:
                    new_f = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                    if new_f != _para_text(para) and para.runs:
                        para.runs[0].text = new_f
                        for r in para.runs[1:]: r.text = ""

            elif re.search(_AMOUNT_PAT, ct):
                # Колонка суммы НДС → зануляем (не трогаем заголовок ri=0)
                if j == nds_sum_col and ri > 0:
                    for para in cell.paragraphs:
                        new_f = re.sub(_AMOUNT_PAT, "0,00", _para_text(para))
                        if new_f != _para_text(para) and para.runs:
                            para.runs[0].text = new_f
                            for r in para.runs[1:]: r.text = ""
                else:
                    # Цена / итого → новая сумма
                    for para in cell.paragraphs:
                        new_f = re.sub(_AMOUNT_PAT, amt_str, _para_text(para))
                        if new_f != _para_text(para) and para.runs:
                            para.runs[0].text = new_f
                            for r in para.runs[1:]: r.text = ""

# Затем параграфы вне таблиц
for para in doc.paragraphs:
    full = _para_text(para)
    new_full = full

    # Дата
    new_full = re.sub(r"\b\d{2}\.\d{2}\.(?:\d{2}|\d{4})\b", date, new_full)

    # Сумма прописью (строка "Всего наименований...")
    if re.search(_WORDS_PAT, new_full):
        new_full = re.sub(_WORDS_PAT, amt_words, new_full)
        # Убираем НДС из этой строки
        new_full = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", new_full)
        new_full = re.sub(r"\s{2,}", " ", new_full)

    # НДС
    new_full = _remove_nds(new_full)

    if new_full != full and para.runs:
        para.runs[0].text = new_full
        for r in para.runs[1:]:
            r.text = ""

return doc
```

# ══════════════════════════════════════════════════════════════════

# КОНФИГ

# ══════════════════════════════════════════════════════════════════

def load_config() -> dict:
if CONFIG_FILE.exists():
try:
return json.loads(CONFIG_FILE.read_text(encoding=“utf-8”))
except Exception:
pass
return {}

def save_config(data: dict):
CONFIG_FILE.write_text(
json.dumps(data, ensure_ascii=False, indent=2), encoding=“utf-8”)

# ══════════════════════════════════════════════════════════════════

# GUI

# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
def **init**(self):
super().**init**()
self.title(“Fresh Auto — Генератор документов с ПВ”)
self.geometry(“720x800”)
self.resizable(False, False)
self.configure(bg=”#1a2540”)

```
    self.cfg = load_config()

    # Пути к файлам (внутренние переменные)
    self._paths = {}

    # StringVar для отображения имён файлов
    self._file_labels = {
        "dkp": tk.StringVar(value="файл не выбран"),
        "pko": tk.StringVar(value="файл не выбран"),
        "inv1": tk.StringVar(value="файл не выбран"),
        "inv2": tk.StringVar(value="файл не выбран"),
    }

    # Поля данных
    self.v_buyer   = tk.StringVar()
    self.v_dkp_num = tk.StringVar()
    self.v_date    = tk.StringVar(value=datetime.today().strftime("%d.%m.%Y"))
    self.v_brand   = tk.StringVar()
    self.v_vin     = tk.StringVar()
    self.v_color   = tk.StringVar()
    self.v_reg     = tk.StringVar()

    # Суммы
    self.v_real    = tk.StringVar()
    self.v_pv      = tk.StringVar()
    self.v_total   = tk.StringVar(value="—")
    self.v_inv2    = tk.StringVar(value="—")

    self.v_real.trace("w", self._recalc)
    self.v_pv.trace("w", self._recalc)

    self._build()

# ── построение интерфейса ──────────────────────────────────

def _build(self):
    # Шапка
    hdr = tk.Frame(self, bg="#1e3a6e", height=56)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(hdr, text="🚗  Fresh Auto  ·  Генератор документов с ПВ",
             font=("Arial", 13, "bold"), bg="#1e3a6e", fg="white").pack(expand=True)

    # Прокручиваемый контейнер
    outer = tk.Frame(self, bg="#1a2540")
    outer.pack(fill="both", expand=True)

    cv = tk.Canvas(outer, bg="#1a2540", highlightthickness=0)
    sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
    self._sf = tk.Frame(cv, bg="#1a2540")

    self._sf.bind("<Configure>",
                  lambda e: cv.configure(scrollregion=cv.bbox("all")))
    cv.create_window((0, 0), window=self._sf, anchor="nw")
    cv.configure(yscrollcommand=sb.set)

    cv.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    cv.bind_all("<MouseWheel>",
                lambda e: cv.yview_scroll(-1 * (e.delta // 120), "units"))

    # Блоки
    self._section("📁  Шаблоны документов")
    self._file_row("ДКП",      "dkp")
    self._file_row("ПКО",      "pko")
    self._file_row("Счёт №1",  "inv1",
                   hint="← загружается авто-заполнение")
    self._file_row("Счёт №2",  "inv2")

    self._section("📋  Данные сделки")
    self._field("ФИО покупателя",       self.v_buyer,   edit=True)
    self._field("Номер ДКП",            self.v_dkp_num, edit=False)
    self._field("Дата документов",      self.v_date,    edit=True)
    self._field("Марка / Модель",       self.v_brand,   edit=False)
    self._field("VIN",                  self.v_vin,     edit=False)
    self._field("Цвет",                 self.v_color,   edit=False)
    self._field("Гос. номер",           self.v_reg,     edit=False)

    self._section("💰  Суммы")
    self._field("Реальная цена авто (₽)",      self.v_real, edit=True)
    self._field("Сумма искусственного ПВ (₽)", self.v_pv,   edit=True)

    calc = tk.Frame(self._sf, bg="#1e2f56", padx=12, pady=8)
    calc.pack(fill="x", padx=14, pady=4)
    for row_i, (lbl, var) in enumerate([
        ("Цена по документам (ДКП + Счёт №1)", self.v_total),
        ("Счёт №2  (к доплате покупателем)",   self.v_inv2),
    ]):
        tk.Label(calc, text=lbl, bg="#1e2f56", fg="#8899bb",
                 font=("Arial", 10), anchor="w", width=36).grid(
                     row=row_i, column=0, sticky="w")
        tk.Label(calc, textvariable=var, bg="#1e2f56", fg="#2ecc71",
                 font=("Arial", 11, "bold")).grid(
                     row=row_i, column=1, padx=16, sticky="w")

    self._section("💾  Сохранение")
    dir_f = tk.Frame(self._sf, bg="#1a2540")
    dir_f.pack(fill="x", padx=14, pady=3)
    self.v_dir = tk.StringVar(
        value=self.cfg.get("save_dir", str(Path.home() / "Desktop")))
    tk.Entry(dir_f, textvariable=self.v_dir, bg="#243055", fg="white",
             font=("Arial", 10), relief="flat", width=50,
             insertbackground="white").pack(side="left", padx=(0, 8))
    tk.Button(dir_f, text="Обзор", command=self._choose_dir,
              bg="#4a90d9", fg="white", relief="flat",
              padx=10).pack(side="left")

    # Прогресс
    tk.Frame(self._sf, bg="#1a2540", height=10).pack()
    self.progress = ttk.Progressbar(self._sf, mode="determinate",
                                    length=680, maximum=4)
    self.progress.pack(padx=14, pady=2)

    self.status = tk.Label(self._sf, text="", bg="#1a2540",
                           fg="#8899bb", font=("Arial", 9))
    self.status.pack()

    # Кнопки
    btn_f = tk.Frame(self._sf, bg="#1a2540")
    btn_f.pack(pady=14)
    tk.Button(btn_f, text="✅  Сгенерировать документы",
              command=self._generate,
              bg="#2ecc71", fg="#0d1a30",
              font=("Arial", 13, "bold"),
              relief="flat", padx=22, pady=10).pack(side="left", padx=8)

    self._open_btn = tk.Button(btn_f, text="📂  Открыть папку",
              command=self._open_folder,
              bg="#4a90d9", fg="white",
              font=("Arial", 11), relief="flat",
              padx=16, pady=10, state="disabled")
    self._open_btn.pack(side="left", padx=8)

    tk.Frame(self._sf, bg="#1a2540", height=24).pack()

def _section(self, title: str):
    f = tk.Frame(self._sf, bg="#1a2540")
    f.pack(fill="x", padx=14, pady=(14, 3))
    tk.Label(f, text=title, bg="#1a2540", fg="#4a90d9",
             font=("Arial", 11, "bold")).pack(side="left")
    tk.Frame(f, bg="#4a90d9", height=1).pack(
        side="left", fill="x", expand=True, padx=8, pady=6)

def _file_row(self, label: str, key: str, hint: str = ""):
    f = tk.Frame(self._sf, bg="#1a2540")
    f.pack(fill="x", padx=14, pady=2)
    tk.Label(f, text=label, bg="#1a2540", fg="#8899bb",
             font=("Arial", 10), width=10, anchor="w").pack(side="left")
    tk.Label(f, textvariable=self._file_labels[key],
             bg="#243055", fg="#aaddff", font=("Arial", 9),
             width=40, anchor="w", padx=6).pack(side="left", padx=4)
    tk.Button(f, text="Выбрать",
              command=lambda k=key: self._choose_file(k),
              bg="#4a90d9", fg="white", relief="flat",
              padx=8, font=("Arial", 9)).pack(side="left")
    if hint:
        tk.Label(f, text=hint, bg="#1a2540", fg="#556688",
                 font=("Arial", 8)).pack(side="left", padx=6)

def _field(self, label: str, var: tk.StringVar, edit: bool):
    f = tk.Frame(self._sf, bg="#1a2540")
    f.pack(fill="x", padx=14, pady=2)
    tk.Label(f, text=label, bg="#1a2540", fg="#8899bb",
             font=("Arial", 10), width=30, anchor="w").pack(side="left")
    state = "normal" if edit else "readonly"
    bg    = "#243055" if edit else "#1e2840"
    fg    = "white"   if edit else "#aaddff"
    tk.Entry(f, textvariable=var, bg=bg, fg=fg,
             font=("Arial", 10), relief="flat", width=34,
             state=state, readonlybackground="#1e2840",
             insertbackground="white").pack(side="left", padx=4)

# ── логика ────────────────────────────────────────────────

def _choose_file(self, key: str):
    path = filedialog.askopenfilename(
        filetypes=[("Word документы", "*.docx"), ("Все файлы", "*.*")])
    if not path:
        return
    self._paths[key] = path
    self._file_labels[key].set(Path(path).name)

    if key == "inv1":
        self._parse_inv1(path)

def _parse_inv1(self, path: str):
    try:
        doc  = Document(path)
        data = extract_invoice_data(doc)
        mapping = {
            "dkp_number": self.v_dkp_num,
            "date":       self.v_date,
            "buyer_name": self.v_buyer,
            "car_brand":  self.v_brand,
            "car_vin":    self.v_vin,
            "car_color":  self.v_color,
            "car_reg":    self.v_reg,
        }
        filled = 0
        for k, var in mapping.items():
            if data.get(k):
                var.set(data[k])
                filled += 1
        self.status.config(
            text=f"✓ Счёт №1 загружен, заполнено {filled} полей",
            fg="#2ecc71")
    except Exception as e:
        self.status.config(text=f"Ошибка парсинга: {e}", fg="#e74c3c")

def _choose_dir(self):
    path = filedialog.askdirectory()
    if path:
        self.v_dir.set(path)
        cfg = load_config()
        cfg["save_dir"] = path
        save_config(cfg)

def _recalc(self, *_):
    try:
        real = parse_amount(self.v_real.get())
        pv   = parse_amount(self.v_pv.get())
        if real > 0 and pv > 0:
            total = real + pv
            self.v_total.set(f"{format_amount(total)} ₽")
            self.v_inv2.set(f"{format_amount(real)} ₽")
        else:
            self.v_total.set("—")
            self.v_inv2.set("—")
    except Exception:
        pass

def _validate(self) -> bool:
    errors = []
    for k, lbl in [("dkp","ДКП"),("pko","ПКО"),
                    ("inv1","Счёт №1"),("inv2","Счёт №2")]:
        if k not in self._paths:
            errors.append(f"Не выбран файл: {lbl}")
    if not self.v_buyer.get().strip():
        errors.append("Не заполнено ФИО покупателя")
    if parse_amount(self.v_real.get()) <= 0:
        errors.append("Укажите реальную цену авто")
    if parse_amount(self.v_pv.get()) <= 0:
        errors.append("Укажите сумму ПВ")
    if errors:
        messagebox.showerror("Заполните все поля", "\n".join(errors))
        return False
    return True

def _generate(self):
    if not self._validate():
        return

    real  = parse_amount(self.v_real.get())
    pv    = parse_amount(self.v_pv.get())
    total = real + pv

    params = {
        "buyer_name": self.v_buyer.get().strip(),
        "dkp_number": self.v_dkp_num.get().strip(),
        "date":       self.v_date.get().strip(),
        "car_brand":  self.v_brand.get().strip(),
        "car_vin":    self.v_vin.get().strip(),
        "car_color":  self.v_color.get().strip(),
        "car_reg":    self.v_reg.get().strip(),
        "real_price": real,
        "pv_amount":  pv,
        "new_price":  total,
    }

    surname   = (self.v_buyer.get().strip().split() or ["Клиент"])[0]
    date_safe = self.v_date.get().replace(".", "-")
    save_dir  = self.v_dir.get()
    os.makedirs(save_dir, exist_ok=True)

    tasks = [
        ("dkp",  f"ДКП_{surname}_{date_safe}.docx",    process_dkp,
         lambda d, p: process_dkp(d, p)),
        ("pko",  f"ПКО_{surname}_{date_safe}.docx",    process_pko,
         lambda d, p: process_pko(d, p)),
        ("inv1", f"Счёт1_{surname}_{date_safe}.docx",  None,
         lambda d, p: process_invoice(d, p, p["new_price"])),
        ("inv2", f"Счёт2_{surname}_{date_safe}.docx",  None,
         lambda d, p: process_invoice(d, p, p["real_price"])),
    ]

    self.progress["value"] = 0
    errors = []

    for key, fname, _, processor in tasks:
        try:
            self.status.config(text=f"Обрабатываю {fname}…", fg="#4a90d9")
            self.update()
            doc = Document(self._paths[key])
            doc = processor(doc, params)
            doc.save(os.path.join(save_dir, fname))
            self.progress["value"] += 1
            self.update()
        except Exception as e:
            errors.append(f"{fname}: {e}")

    if errors:
        self.status.config(text="⚠ Завершено с ошибками", fg="#e67e22")
        messagebox.showwarning("Ошибки при генерации", "\n".join(errors))
    else:
        self.status.config(
            text=f"✅ 4 документа сохранены → {save_dir}", fg="#2ecc71")
        self._open_btn.config(state="normal")
        messagebox.showinfo("Готово!",
            f"Документы сохранены в папку:\n{save_dir}")

def _open_folder(self):
    path = self.v_dir.get()
    if os.path.exists(path):
        if os.name == "nt":
            os.startfile(path)
        else:
            subprocess.Popen(["open", path])
```

# ══════════════════════════════════════════════════════════════════

if **name** == “**main**”:
app = App()
app.mainloop()
