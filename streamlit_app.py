"""
Fresh Auto - Генератор документов с искусственным ПВ
Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках
"""

import streamlit as st
import json, os, re, subprocess, sys
from pathlib import Path
from datetime import datetime
import base64
from io import BytesIO

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import num2words as nw
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                          "python-docx", "num2words", "-q"])
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import num2words as nw

# ==============================================================================
# УТИЛИТЫ ДЛЯ РАБОТЫ С СУММАМИ
# ==============================================================================

def _rub_word(n: int) -> str:
    """Правильное склонение слова 'рубль'"""
    n100, n10 = n % 100, n % 10
    if 11 <= n100 <= 19:
        return "рублей"
    if n10 == 1:
        return "рубль"
    if 2 <= n10 <= 4:
        return "рубля"
    return "рублей"

def format_amount(amount: float) -> str:
    """2138000 -> '2 138 000,00'"""
    rub = int(amount)
    kop = round((amount - rub) * 100)
    return f"{rub:,}".replace(",", " ") + f",{kop:02d}"

def amount_to_words(amount: float) -> str:
    """2138000 -> 'Два миллиона сто тридцать восемь тысяч рублей 00 копеек'"""
    rub = int(amount)
    kop = round((amount - rub) * 100)
    words = nw.num2words(rub, lang="ru")
    words = words[0].upper() + words[1:]
    return f"{words} {_rub_word(rub)} {kop:02d} копеек"

def parse_amount(text: str) -> float:
    """'2 000 000,00' или '2000000' -> 2000000.0"""
    if not text:
        return 0.0
    cleaned = re.sub(r"[^\d,.]", "", text.strip())
    cleaned = cleaned.replace(",", ".")
    parts = cleaned.split(".")
    if len(parts) > 2:
        cleaned = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

# ==============================================================================
# РАБОТА С DOCX - НИЗКОУРОВНЕВЫЕ УТИЛИТЫ
# ==============================================================================

def _iter_paragraphs(doc: Document):
    """Все параграфы: тело + таблицы + колонтитулы"""
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
    """Полный текст параграфа из всех runs"""
    return "".join(r.text for r in para.runs)

def _replace_in_para(para, old: str, new: str) -> bool:
    """Заменяет old->new в параграфе, сохраняя форматирование"""
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            return True

    full = _para_text(para)
    if old not in full:
        return False
    new_full = full.replace(old, new)
    if para.runs:
        para.runs[0].text = new_full
        for r in para.runs[1:]:
            r.text = ""
    return True

# Паттерны для сумм
_AMOUNT_PAT = r"\d[\d\s]*\d[,.]\d{2}"
_WORDS_PAT = (r"[А-ЯЁ][а-яёА-ЯЁ\s]+"
              r"(?:тысяч|миллион|миллиард)[а-яё\s]*"
              r"(?:рубл[а-яё]+)\s+\d{2}\s+копеек")

def _remove_nds(text: str) -> str:
    """Убирает все упоминания НДС из строки"""
    text = re.sub(r",?\s*в\s+т.?\s*ч.?\s*НДС[^.;\n]*", "", text)
    text = re.sub(r"(22/122%?)\s*[\d\s,]+руб.?", "", text)
    text = re.sub(r"НДС\s*(22/122%?)\s*[\d\s,]+руб.?", "без НДС", text)
    return text

# ==============================================================================
# ПАРСИНГ СЧЁТА №1
# ==============================================================================

def extract_invoice_data(doc: Document) -> dict:
    """Извлекает данные из Счёта №1 для авто-заполнения полей"""
    data = {}

    for para in _iter_paragraphs(doc):
        t = para.text.strip()

        m = re.search(
            r"[Пп]родажа\s+[тТ][/\\][сС]\s+№\s*([\w\-/]+).*?от\s+(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))",
            t)
        if m and "dkp_number" not in data:
            data["dkp_number"] = m.group(1).strip()
            d = m.group(2)
            parts = d.split(".")
            if len(parts[2]) == 2:
                parts[2] = "20" + parts[2]
            data["date"] = ".".join(parts)

        m = re.search(
            r"[Пп]окупатель[:\s]+(?:ИНН\s+\d+[,\s]+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)",
            t)
        if m and "buyer_name" not in data:
            data["buyer_name"] = m.group(1).strip()

        m = re.search(
            r"([A-Z]{2,}\s+[A-Z]{2,})\s+([а-яё]+)\s+№\s*([A-ZА-ЯЁa-zA-Zа-яё0-9]+)\s+VIN\s+([A-Z0-9]{17})",
            t)
        if m and "car_vin" not in data:
            data["car_brand"] = m.group(1).strip()
            data["car_color"] = m.group(2).strip()
            data["car_reg"] = m.group(3).strip()
            data["car_vin"] = m.group(4).strip()

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

# ==============================================================================
# ОБРАБОТКА ДОКУМЕНТОВ
# ==============================================================================

def process_dkp(doc: Document, p: dict) -> Document:
    """Обработка ДКП"""
    new_price = p["new_price"]
    pv_amount = p["pv_amount"]
    new_str = format_amount(new_price)
    new_words = amount_to_words(new_price)
    pv_str = format_amount(pv_amount)
    pv_words = amount_to_words(pv_amount)

    pv_para_found = False

    for para in _iter_paragraphs(doc):
        full = _para_text(para)

        if re.search(_AMOUNT_PAT, full):
            new_full = re.sub(_AMOUNT_PAT, new_str, full)
            new_full = _remove_nds(new_full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]:
                    r.text = ""

        elif re.search(_WORDS_PAT, full):
            new_full = re.sub(_WORDS_PAT, new_words, full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]:
                    r.text = ""

        elif re.search(r"НДС|22/122", full):
            new_full = _remove_nds(full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]:
                    r.text = ""

        if "первоначальный взнос" in full.lower():
            pv_para_found = True
            new_full = re.sub(_AMOUNT_PAT, pv_str, full)
            new_full = re.sub(_WORDS_PAT, pv_words, new_full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]:
                    r.text = ""

    if not pv_para_found:
        pv_text = (f"Первоначальный взнос по оплате цены Договора составляет "
                   f"{pv_str} руб ({pv_words}).")
        for para in doc.paragraphs:
            if re.search(r"размере\s+" + _AMOUNT_PAT, _para_text(para)):
                new_p = OxmlElement("w:p")
                new_r = OxmlElement("w:r")
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

def process_pko(doc: Document, p: dict) -> Document:
    """Обработка ПКО"""
    pv_str = format_amount(p["pv_amount"])
    pv_words = amount_to_words(p["pv_amount"])
    date = p["date"]
    buyer = p["buyer_name"]
    osnov = (f"По ДКП №{p['dkp_number']} от {date} "
             f"за а/м {p['car_brand']} {p['car_color']} "
             f"№{p['car_reg']} VIN {p['car_vin']}")

    for para in _iter_paragraphs(doc):
        full = _para_text(para)
        if not full.strip():
            continue

        if re.search(r"[Пп]о\s+ДКП|за\s+а/м|VIN\s+[A-Z0-9]{17}", full):
            if para.runs:
                para.runs[0].text = osnov
                for r in para.runs[1:]: r.text = ""
            continue

        if re.search(_WORDS_PAT, full):
            new_full = re.sub(_WORDS_PAT, pv_words, full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""
            continue

        if re.search(r"НДС|22/122", full):
            new_full = re.sub(r"НДС\s*\(22/122%?\)[^\n]*", "Без НДС", full)
            new_full = re.sub(r"В том числе НДС[^\n]*", "Без НДС", new_full)
            new_full = _remove_nds(new_full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""
            continue

        if re.search(_AMOUNT_PAT, full) and not re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
            new_full = re.sub(_AMOUNT_PAT, pv_str, full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""
            continue

        if re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
            new_full = re.sub(r"\b\d{2}\.\d{2}\.\d{4}\b", date, full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""
            continue

        m = re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+", full)
        if m and m.group(0) != buyer:
            new_full = full.replace(m.group(0), buyer)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""

    return doc

def process_invoice(doc: Document, p: dict, amount: float) -> Document:
    """Обработка счёта"""
    amt_str = format_amount(amount)
    amt_words = amount_to_words(amount)
    date = p["date"]

    for tbl in doc.tables:
        nds_sum_col = -1
        if len(tbl.rows) > 0:
            for ci, hcell in enumerate(tbl.rows[0].cells):
                if re.search(r"[Сс]умма\s*НДС|НДС\s*[Сс]умма", hcell.text):
                    nds_sum_col = ci
                    break
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

                if re.search(r"22/122", ct):
                    for para in cell.paragraphs:
                        new_f = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                        if new_f != _para_text(para) and para.runs:
                            para.runs[0].text = new_f
                            for r in para.runs[1:]: r.text = ""

                elif re.search(_AMOUNT_PAT, ct):
                    if j == nds_sum_col and ri > 0:
                        for para in cell.paragraphs:
                            new_f = re.sub(_AMOUNT_PAT, "0,00", _para_text(para))
                            if new_f != _para_text(para) and para.runs:
                                para.runs[0].text = new_f
                                for r in para.runs[1:]: r.text = ""
                    else:
                        for para in cell.paragraphs:
                            new_f = re.sub(_AMOUNT_PAT, amt_str, _para_text(para))
                            if new_f != _para_text(para) and para.runs:
                                para.runs[0].text = new_f
                                for r in para.runs[1:]: r.text = ""

    for para in doc.paragraphs:
        full = _para_text(para)
        new_full = full

        new_full = re.sub(r"\b\d{2}\.\d{2}\.(?:\d{2}|\d{4})\b", date, new_full)

        if re.search(_WORDS_PAT, new_full):
            new_full = re.sub(_WORDS_PAT, amt_words, new_full)
            new_full = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", new_full)
            new_full = re.sub(r"\s{2,}", " ", new_full)

        new_full = _remove_nds(new_full)

        if new_full != full and para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""

    return doc

# ==============================================================================
# STREAMLIT APP
# ==============================================================================

def get_binary_file_downloader_html(bin_data, file_label='File'):
    """Генерация ссылки для скачивания"""
    b64 = base64.b64encode(bin_data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{file_label}">Скачать {file_label}</a>'

def main():
    st.set_page_config(
        page_title="Fresh Auto - Генератор документов с ПВ",
        page_icon=":car:",
        layout="wide"
    )

    st.title(":car: Fresh Auto - Генератор документов с ПВ")
    st.markdown("Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках")

    # Инициализация session_state для хранения файлов
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = {}
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {
            'buyer_name': '',
            'dkp_number': '',
            'date': datetime.today().strftime("%d.%m.%Y"),
            'car_brand': '',
            'car_vin': '',
            'car_color': '',
            'car_reg': '',
            'real_price': '',
            'pv_amount': ''
        }

    # Боковая панель для загрузки файлов
    with st.sidebar:
        st.header(":file_folder: Шаблоны документов")

        dkp_file = st.file_uploader("ДКП (шаблон)", type=['docx'], key='dkp')
        pko_file = st.file_uploader("ПКО (шаблон)", type=['docx'], key='pko')
        inv1_file = st.file_uploader("Счёт №1 (шаблон) :arrow_left: авто-заполнение", type=['docx'], key='inv1')
        inv2_file = st.file_uploader("Счёт №2 (шаблон)", type=['docx'], key='inv2')

        # Сохраняем файлы в session_state
        if dkp_file:
            st.session_state.uploaded_files['dkp'] = dkp_file.getvalue()
        if pko_file:
            st.session_state.uploaded_files['pko'] = pko_file.getvalue()
        if inv1_file:
            st.session_state.uploaded_files['inv1'] = inv1_file.getvalue()
        if inv2_file:
            st.session_state.uploaded_files['inv2'] = inv2_file.getvalue()

        st.markdown("---")
        st.markdown("### Статус загрузки:")
        for name, label in [('dkp', 'ДКП'), ('pko', 'ПКО'), ('inv1', 'Счёт №1'), ('inv2', 'Счёт №2')]:
            if name in st.session_state.uploaded_files:
                st.success(f"✅ {label} загружен")
            else:
                st.warning(f"⚠️ {label} не загружен")

    # Основная область
    col1, col2 = st.columns(2)

    with col1:
        st.header(":clipboard: Данные сделки")

        buyer_name = st.text_input("ФИО покупателя",
                                  value=st.session_state.form_data['buyer_name'],
                                  key='buyer_name_input')
        dkp_number = st.text_input("Номер ДКП",
                                   value=st.session_state.form_data['dkp_number'],
                                   key='dkp_number_input',
                                   disabled=True)
        date = st.text_input("Дата документов",
                            value=st.session_state.form_data['date'],
                            key='date_input')
        car_brand = st.text_input("Марка / Модель",
                                 value=st.session_state.form_data['car_brand'],
                                 key='car_brand_input',
                                 disabled=True)
        car_vin = st.text_input("VIN",
                               value=st.session_state.form_data['car_vin'],
                               key='car_vin_input',
                               disabled=True)
        car_color = st.text_input("Цвет",
                                 value=st.session_state.form_data['car_color'],
                                 key='car_color_input',
                                 disabled=True)
        car_reg = st.text_input("Гос. номер",
                               value=st.session_state.form_data['car_reg'],
                               key='car_reg_input',
                               disabled=True)

    with col2:
        st.header(":moneybag: Суммы")

        real_price_str = st.text_input("Реальная цена авто (руб.)",
                                      value=st.session_state.form_data['real_price'],
                                      key='real_price_input')
        pv_amount_str = st.text_input("Сумма искусственного ПВ (руб.)",
                                     value=st.session_state.form_data['pv_amount'],
                                     key='pv_amount_input')

        try:
            real_price = parse_amount(real_price_str) if real_price_str else 0
            pv_amount = parse_amount(pv_amount_str) if pv_amount_str else 0
            if real_price > 0 and pv_amount > 0:
                total_price = real_price + pv_amount
                st.info(f"**Цена по документам (ДКП + Счёт №1):** {format_amount(total_price)} руб.")
                st.info(f"**Счёт №2 (к доплате покупателем):** {format_amount(real_price)} руб.")
        except:
            pass

    # Авто-заполнение из Счёта №1
    if 'inv1' in st.session_state.uploaded_files:
        if st.button(":arrows_counterclockwise: Авто-заполнить из Счёта №1", use_container_width=True):
            try:
                doc = Document(BytesIO(st.session_state.uploaded_files['inv1']))
                data = extract_invoice_data(doc)

                updates = []
                if data.get('buyer_name'):
                    st.session_state.form_data['buyer_name'] = data['buyer_name']
                    updates.append("ФИО")
                if data.get('dkp_number'):
                    st.session_state.form_data['dkp_number'] = data['dkp_number']
                    updates.append("Номер ДКП")
                if data.get('date'):
                    st.session_state.form_data['date'] = data['date']
                    updates.append("Дата")
                if data.get('car_brand'):
                    st.session_state.form_data['car_brand'] = data['car_brand']
                    updates.append("Марка")
                if data.get('car_vin'):
                    st.session_state.form_data['car_vin'] = data['car_vin']
                    updates.append("VIN")
                if data.get('car_color'):
                    st.session_state.form_data['car_color'] = data['car_color']
                    updates.append("Цвет")
                if data.get('car_reg'):
                    st.session_state.form_data['car_reg'] = data['car_reg']
                    updates.append("Гос.номер")

                if updates:
                    st.success(f"✅ Заполнено полей: {', '.join(updates)}")
                    st.rerun()
                else:
                    st.warning("Не удалось извлечь данные из счёта")
            except Exception as e:
                st.error(f"Ошибка парсинга: {e}")

    # Генерация документов
    st.markdown("---")
    if st.button(":white_check_mark: Сгенерировать документы", type="primary", use_container_width=True):
        # Валидация
        errors = []
        if 'dkp' not in st.session_state.uploaded_files:
            errors.append("Не загружен файл: ДКП")
        if 'pko' not in st.session_state.uploaded_files:
            errors.append("Не загружен файл: ПКО")
        if 'inv1' not in st.session_state.uploaded_files:
            errors.append("Не загружен файл: Счёт №1")
        if 'inv2' not in st.session_state.uploaded_files:
            errors.append("Не загружен файл: Счёт №2")
        if not buyer_name.strip():
            errors.append("Не заполнено ФИО покупателя")
        if not date.strip():
            errors.append("Не заполнена дата")
        if real_price <= 0:
            errors.append("Укажите реальную цену авто")
        if pv_amount <= 0:
            errors.append("Укажите сумму ПВ")

        if errors:
            for error in errors:
                st.error(error)
        else:
            total = real_price + pv_amount
            params = {
                "buyer_name": buyer_name.strip(),
                "dkp_number": dkp_number.strip(),
                "date": date.strip(),
                "car_brand": car_brand.strip(),
                "car_vin": car_vin.strip(),
                "car_color": car_color.strip(),
                "car_reg": car_reg.strip(),
                "real_price": real_price,
                "pv_amount": pv_amount,
                "new_price": total,
            }

            surname = (buyer_name.strip().split() or ["Клиент"])[0]
            date_safe = date.replace(".", "-")

            tasks = [
                ("dkp", f"ДКП_{surname}_{date_safe}.docx",
                 lambda d, p: process_dkp(d, p)),
                ("pko", f"ПКО_{surname}_{date_safe}.docx",
                 lambda d, p: process_pko(d, p)),
                ("inv1", f"Счёт1_{surname}_{date_safe}.docx",
                 lambda d, p: process_invoice(d, p, p["new_price"])),
                ("inv2", f"Счёт2_{surname}_{date_safe}.docx",
                 lambda d, p: process_invoice(d, p, p["real_price"])),
            ]

            progress_bar = st.progress(0)
            status_text = st.empty()

            generated_files = []

            for i, (key, fname, processor) in enumerate(tasks):
                try:
                    status_text.text(f"Обрабатываю {fname}...")
                    doc = Document(BytesIO(st.session_state.uploaded_files[key]))
                    doc = processor(doc, params)
                    
                    # Сохраняем в BytesIO
                    output = BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    generated_files.append((fname, output.getvalue()))
                    progress_bar.progress((i + 1) / len(tasks))
                except Exception as e:
                    st.error(f"Ошибка при обработке {fname}: {e}")

            progress_bar.empty()
            status_text.empty()

            if len(generated_files) == 4:
                st.success(f"✅ Успешно сгенерировано {len(generated_files)} документа!")
                
                # Показываем ссылки для скачивания
                st.subheader("Скачать документы:")
                for fname, data in generated_files:
                    st.markdown(
                        get_binary_file_downloader_html(data, fname),
                        unsafe_allow_html=True
                    )
                    st.markdown("---")
            else:
                st.warning(f"Сгенерировано только {len(generated_files)} из 4 документов")

if __name__ == "__main__":
    main()
