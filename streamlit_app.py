"""
FRESH - Генератор документов с искусственным ПВ
Автоматическое заполнение ДКП, ПКО и счетов при кредищих сделках
"""

import streamlit as st
import re, sys, subprocess
from datetime import datetime
from io import BytesIO

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt
    import num2words as nw
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "python-docx", "num2words", "-q"])
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt
    import num2words as nw

# ==============================================================================
# УТИЛИТЫ ДЛЯ РАБОТЫ С СУММАМИ
# ==============================================================================

def _rub_word(n: int) -> str:
    n100, n10 = n % 100, n % 10
    if 11 <= n100 <= 19:
        return "рублей"
    if n10 == 1:
        return "рубль"
    if 2 <= n10 <= 4:
        return "рубля"
    return "рублей"

def format_amount(amount: float) -> str:
    rub = int(amount)
    kop = round((amount - rub) * 100)
    return f"{rub:,}".replace(",", " ") + f",{kop:02d}"

def amount_to_words(amount: float) -> str:
    rub = int(amount)
    kop = round((amount - rub) * 100)
    words = nw.num2words(rub, lang="ru")
    words = words[0].upper() + words[1:]
    return f"{words} {_rub_word(rub)} {kop:02d} копеек"

def parse_amount(text: str) -> float:
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
# РАБОТА С DOCX
# ==============================================================================

def _iter_paragraphs(doc: Document):
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
    return "".join(r.text for r in para.runs)

def _replace_para_text(para, new_text):
    if not para.runs:
        para.add_run(new_text)
        return
    if _para_text(para) == new_text:
        return
    
    first_run = para.runs[0]
    style = first_run._element.find(qn("w:rPr"))
    
    for r in para.runs[1:]:
        r.text = ""
    
    first_run.text = new_text
    if style is not None:
        first_run._element.append(style)

def insert_paragraph_after(paragraph, text, style_source_para=None):
    """Вставка абзаца строго после текущего с принудительным размером шрифта 8pt"""
    new_p = OxmlElement('w:p')
    paragraph._p.getparent().insert(paragraph._p.getparent().index(paragraph._p) + 1, new_p)
    new_para = paragraph.__class__(new_p, paragraph._parent)
    
    # Создаем раны с текстом и жестко ставим 8pt
    run = new_para.add_run(text)
    run.font.size = Pt(8)
    
    # Копируем базовые стили (шрифт Arial/Times), если они есть
    if style_source_para and style_source_para.runs:
        src_run = style_source_para.runs[0]
        run.font.name = src_run.font.name
        
    # Настройки абзаца (чтобы не было огромных отступов)
    new_para.paragraph_format.space_before = Pt(0)
    new_para.paragraph_format.space_after = Pt(4)
    new_para.paragraph_format.line_spacing = 1.0
            
    return new_para

_AMOUNT_PAT = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
_WORDS_PAT = (r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+"
              r"(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*"
              r"\d{2}\s+копеек")
_DATE_PAT = r"\b\d{2}\.\d{2}\.\d{4}\b"

def _remove_nds(text: str) -> str:
    # Агрессивное удаление НДС с учетом возможных переносов строк и пробелов
    text = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.;\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\(?22/122%?\)?\s*[\d\s,.]+руб\.?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"НДС\s*\(?22/122%?\)?\s*[\d\s,.]+руб\.?", "", text, flags=re.IGNORECASE)
    # Чистим двойные пробелы в конце
    text = re.sub(r"\s+", " ", text).strip()
    return text

def extract_invoice_data(doc: Document) -> dict:
    data = {}
    
    for para in _iter_paragraphs(doc):
        t = para.text.strip()
        
        m = re.search(
            r"[Пп]родажа\s+[тТ][/\\][сС]\s+№\s*([\w\-/]+).*?от\s+(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))", t)
        if m and "dkp_number" not in data:
            data["dkp_number"] = m.group(1).strip()
            d = m.group(2)
            parts = d.split(".")
            if len(parts[2]) == 2:
                parts[2] = "20" + parts[2]
            data["date"] = ".".join(parts)
        
        m = re.search(
            r"[Пп]окупатель[:\s]+(?:ИНН\s+\d+[,\s]+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)", t)
        if m and "buyer_name" not in data:
            data["buyer_name"] = m.group(1).strip()
            
        m = re.search(
            r"([A-ZА-ЯЁ]{2,}\s+[A-ZА-ЯЁ0-9\s\-]+?)\s+([а-яё]+)\s+№\s*([A-ZА-ЯЁ0-9]+)\s+VIN\s+([A-Z0-9]{17})", t, re.IGNORECASE)
        if m and "car_vin" not in data:
            data["car_brand"] = m.group(1).strip()
            data["car_color"] = m.group(2).strip()
            data["car_reg"] = m.group(3).strip()
            data["car_vin"] = m.group(4).strip()
    
    for tbl in doc.tables:
        for row in tbl.rows:
            row_text = " ".join(c.text.strip() for c in row.cells)
            
            if "car_vin" not in data:
                m = re.search(r"VIN\s*[:\s]*([A-Z0-9]{17})", row_text, re.IGNORECASE)
                if m: data["car_vin"] = m.group(1)
            
            if "car_brand" not in data:
                m = re.search(r"Марка\s*[:\s]*([A-ZА-ЯЁ]{2,}\s+[A-ZА-ЯЁ0-9\s\-]+)", row_text, re.IGNORECASE)
                if m: data["car_brand"] = m.group(1).strip()
            
            if "car_color" not in data:
                m = re.search(
                    r"\b(серый|белый|чёрный|черный|синий|красный|серебристый"
                    r"|золотистый|коричневый|зелёный|зеленый|бежевый|жёлтый|желтый)\b",
                    row_text, re.IGNORECASE)
                if m: data["car_color"] = m.group(1).lower()
            
            if "car_reg" not in data:
                m = re.search(r"(?:№|Гос\s*знак)\s*([A-ZА-ЯЁ]{1,2}\d{3}[A-ZА-ЯЁ]{2}\d{2,3})", row_text, re.IGNORECASE)
                if m: data["car_reg"] = m.group(1)
    
    return data

# ==============================================================================
# ПРОЦЕССОРЫ ДОКУМЕНТОВ
# ==============================================================================

def process_dkp(doc: Document, p: dict) -> Document:
    new_price = p["new_price"]
    pv_amount = p["pv_amount"]
    new_str = format_amount(new_price)
    new_words = amount_to_words(new_price)
    pv_str = format_amount(pv_amount)
    pv_words = amount_to_words(pv_amount)
    
    pv_para_found = False
    target_payment_para = None
    
    A_PAT = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
    W_PAT = r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*\d{2}\s+копеек\s*\)?"

    for para in doc.paragraphs:
        full = "".join(r.text for r in para.runs)
        full_normalized = re.sub(r"\s+", " ", full).strip()
        
        if not full_normalized:
            continue
        
        if full_normalized.startswith("Цена ТС оплачивается Покупателем в течение"):
            target_payment_para = para
        
        if any(word in full_normalized.lower() for word in ["гаранти", "техническая защита", "лимит ответственности", "пробег"]):
            continue

        # 1. Корректируем пункт с Первоначальным Взносом
        if "первоначальный" in full_normalized.lower() or "взнос" in full_normalized.lower():
            clean_full = full_normalized
            if re.search(A_PAT, clean_full):
                clean_full = re.sub(A_PAT, pv_str, clean_full)
            if re.search(W_PAT, clean_full):
                clean_full = re.sub(W_PAT, pv_words, clean_full)
            
            para.text = ""
            run = para.add_run(clean_full)
            run.font.size = Pt(8)  # Строго 8pt
            pv_para_found = True
        
        # 2. Корректируем пункт цены договора / стоимости ТС
        elif any(marker in full_normalized.lower() for marker in ["цена договора", "стоимость тс", "цена тс", "цену за тс", "стоимость автомобиля", "уплачивает покупатель"]):
            clean_full = _remove_nds(full_normalized)
            
            has_amt = re.search(A_PAT, clean_full)
            has_wrd = re.search(W_PAT, clean_full)
            
            if has_amt and has_wrd:
                clean_full = re.sub(A_PAT, new_str, clean_full)
                clean_full = re.sub(W_PAT, new_words + " )" if ")" in has_wrd.group(0) else new_words, clean_full)
            elif has_amt:
                clean_full = re.sub(A_PAT, new_str, clean_full)
            
            clean_full = clean_full.replace(" ) )", " )").replace("))", ")")
            clean_full = re.sub(r",\s*\.", ".", clean_full)
            clean_full = re.sub(r"\s+", " ", clean_full).strip()
            
            if not clean_full.endswith("."):
                clean_full += "."

            para.text = ""
            run = para.add_run(clean_full)
            run.font.size = Pt(8)  # Строго 8pt для измененной цены договора

    # Если пункта ПВ не было, создаем его с размером 8pt
    if not pv_para_found and target_payment_para is not None:
        pv_text = f"Первоначальный взнос по оплате цены Договора составляет {pv_str} руб ({pv_words})."
        insert_paragraph_after(target_payment_para, pv_text, style_source_para=target_payment_para)

    return doc
def process_pko(doc: Document, p: dict) -> Document:
    pv_str = format_amount(p["pv_amount"])
    pv_words = amount_to_words(p["pv_amount"])
    date = p["date"]
    
    # Формируем чистое однострочное основание
    osnov = (f"По ДКП №{p['dkp_number']} от {date} "
             f"за а/м {p['car_brand']} {p['car_color']} "
             f"№{p['car_reg']} VIN {p['car_vin']}")

    A_PAT = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
    W_PAT = r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*\d{2}\s+копеек\s*\)?"

    # Проходим по абсолютно всем таблицам и их ячейкам (так как весь ПКО — это таблица)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                # Читаем текст из ячейки и нормализуем пробелы
                cell_text = "".join(r.text for para in cell.paragraphs for r in para.runs)
                cell_text_clean = re.sub(r"\s+", " ", cell_text).strip()
                
                if not cell_text_clean:
                    continue

                # 1. Полностью сносим строки с НДС, где бы они ни находились
                if "в том числе" in cell_text_clean.lower() and "ндс" in cell_text_clean.lower():
                    for para in cell.paragraphs:
                        para.text = ""
                    continue

                # 2. Обрабатываем блок "Основание:" (в левой части)
                if cell_text_clean.startswith("Основание:"):
                    # Очищаем ячейку и пишем строго один раз
                    for para in cell.paragraphs:
                        para.text = ""
                    run = cell.paragraphs[0].add_run(f"Основание: {osnov}")
                    run.font.size = Pt(8)
                    continue

                # 3. Обрабатываем правое основание (где нет слова "Основание:", но есть ДКП и VIN)
                if ("дкп" in cell_text_clean.lower() or "vin" in cell_text_clean.lower()) and "кассир" not in cell_text_clean.lower():
                    # Защита от повторного срабатывания на уже измененном поле
                    if cell_text_clean == osnov:
                        continue
                    for para in cell.paragraphs:
                        para.text = ""
                    run = cell.paragraphs[0].add_run(osnov)
                    run.font.size = Pt(8)
                    continue

                # 4. Меняем пропись суммы (Принято от... / Сумма прописью)
                if re.search(W_PAT, cell_text_clean) and any(m in cell_text_clean.lower() for m in ["принято", "сумма", "руб"]):
                    new_text = re.sub(W_PAT, pv_words, cell_text_clean)
                    for para in cell.paragraphs:
                        para.text = ""
                    run = cell.paragraphs[0].add_run(new_text)
                    run.font.size = Pt(8)
                    continue

                # 5. Меняем цифры суммы в тексте (например, "Сумма 600 000,00" или "Сумма цифрами")
                if re.search(A_PAT, cell_text_clean) and "сумма" in cell_text_clean.lower():
                    new_text = re.sub(A_PAT, pv_str, cell_text_clean)
                    for para in cell.paragraphs:
                        para.text = ""
                    run = cell.paragraphs[0].add_run(new_text)
                    run.font.size = Pt(8)
                    continue

                # 6. Меняем чисто цифровые ячейки (Дебет/Кредит/Сумма посередине бланка)
                if re.match(r"^[\d\s.,]+$", cell_text_clean) and re.search(A_PAT, cell_text_clean):
                    for para in cell.paragraphs:
                        para.text = ""
                    run = cell.paragraphs[0].add_run(pv_str)
                    run.font.size = Pt(8)
                    continue

    return doc

def process_invoice(doc: Document, p: dict, amount: float) -> Document:
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

        for ri, row in enumerate(tbl.rows):
            for j, cell in enumerate(row.cells):
                ct = cell.text.strip()

                if re.search(r"22/122", ct):
                    for para in cell.paragraphs:
                        new_f = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                        _replace_para_text(para, new_f)

                elif re.search(_AMOUNT_PAT, ct):
                    if j == nds_sum_col and ri > 0:
                        for para in cell.paragraphs:
                            if re.search(_AMOUNT_PAT, _para_text(para)):
                                _replace_para_text(para, "0,00")
                    else:
                        if ri > 0 and (j >= len(row.cells) - 2): 
                            for para in cell.paragraphs:
                                if re.search(_AMOUNT_PAT, _para_text(para)):
                                    _replace_para_text(para, amt_str)

    for para in doc.paragraphs:
        full = _para_text(para)
        new_full = full

        if re.search(_DATE_PAT, full) and "счет" in full.lower():
            new_full = re.sub(_DATE_PAT, date, new_full)

        if re.search(_WORDS_PAT, new_full):
            new_full = re.sub(_WORDS_PAT, amt_words, new_full)
            new_full = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", new_full)
            new_full = re.sub(r"\s{2,}", " ", new_full)

        new_full = _remove_nds(new_full)

        if new_full != full:
            _replace_para_text(para, new_full)

    return doc

# ==============================================================================
# STREAMLIT APP
# ==============================================================================

def main():
    st.set_page_config(
        page_title="FRESH - Генератор документов с ПВ",
        page_icon="🚗",
        layout="wide"
    )

    st.title("🚗 FRESH - Генератор документов с ПВ")
    st.markdown("Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках")

    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = {}
    
    def auto_fill():
        if 'inv1' in st.session_state.uploaded_files:
            try:
                doc = Document(BytesIO(st.session_state.uploaded_files['inv1']))
                data = extract_invoice_data(doc)
                
                updates = []
                if data.get('buyer_name'):
                    st.session_state.buyer_name = data['buyer_name']
                    updates.append("ФИО")
                if data.get('dkp_number'):
                    st.session_state.dkp_number = data['dkp_number']
                    updates.append("Номер ДКП")
                if data.get('date'):
                    st.session_state.date = data['date']
                    updates.append("Дата")
                if data.get('car_brand'):
                    st.session_state.car_brand = data['car_brand']
                    updates.append("Марка")
                if data.get('car_vin'):
                    st.session_state.car_vin = data['car_vin']
                    updates.append("VIN")
                if data.get('car_color'):
                    st.session_state.car_color = data['car_color']
                    updates.append("Цвет")
                if data.get('car_reg'):
                    st.session_state.car_reg = data['car_reg']
                    updates.append("Гос.номер")
                
                if updates:
                    st.success(f"✅ Успешно импортировано: {', '.join(updates)}")
                else:
                    st.warning("Не удалось извлечь данные. Проверьте формат Счёта №1.")
            except Exception as e:
                st.error(f"Ошибка парсинга: {e}")

    with st.sidebar:
        st.header("📁 Шаблоны документов")

        dkp_file = st.file_uploader("ДКП (шаблон)", type=['docx'], key='dkp_file')
        pko_file = st.file_uploader("ПКО (шаблон)", type=['docx'], key='pko_file')
        inv1_file = st.file_uploader("Счёт №1 (шаблон) ⬅️ авто-заполнение", type=['docx'], key='inv1_file')
        inv2_file = st.file_uploader("Счёт №2 (шаблон)", type=['docx'], key='inv2_file')

        if dkp_file: st.session_state.uploaded_files['dkp'] = dkp_file.getvalue()
        if pko_file: st.session_state.uploaded_files['pko'] = pko_file.getvalue()
        if inv1_file: st.session_state.uploaded_files['inv1'] = inv1_file.getvalue()
        if inv2_file: st.session_state.uploaded_files['inv2'] = inv2_file.getvalue()

        st.markdown("---")
        st.markdown("### Статус загрузки:")
        for name, label in [('dkp', 'ДКП'), ('pko', 'ПКО'), ('inv1', 'Счёт №1'), ('inv2', 'Счёт №2')]:
            if name in st.session_state.uploaded_files:
                st.success(f"✅ {label} загружен")
            else:
                st.warning(f"⚠️ {label} не загружен")

    col1, col2 = st.columns(2)

    with col1:
        st.header("📋 Данные сделки")

        if 'buyer_name' not in st.session_state: st.session_state.buyer_name = ''
        if 'dkp_number' not in st.session_state: st.session_state.dkp_number = ''
        if 'date' not in st.session_state: st.session_state.date = datetime.today().strftime("%d.%m.%Y")
        if 'car_brand' not in st.session_state: st.session_state.car_brand = ''
        if 'car_vin' not in st.session_state: st.session_state.car_vin = ''
        if 'car_color' not in st.session_state: st.session_state.car_color = ''
        if 'car_reg' not in st.session_state: st.session_state.car_reg = ''
        if 'real_price' not in st.session_state: st.session_state.real_price = ''
        if 'pv_amount' not in st.session_state: st.session_state.pv_amount = ''

        st.text_input("ФИО покупателя", key='buyer_name')
        st.text_input("Номер ДКП", key='dkp_number')
        st.text_input("Дата документов", key='date')
        st.text_input("Марка / Модель", key='car_brand')
        st.text_input("VIN", key='car_vin')
        st.text_input("Цвет", key='car_color')
        st.text_input("Гос. номер", key='car_reg')

    with col2:
        st.header("💰 Суммы")

        real_price_str = st.text_input("Реальная цена авто (руб.)", key='real_price')
        pv_amount_str = st.text_input("Сумма искусственного ПВ (руб.)", key='pv_amount')

        try:
            real_price = parse_amount(real_price_str) if real_price_str else 0
            pv_amount = parse_amount(pv_amount_str) if pv_amount_str else 0
            if real_price > 0 and pv_amount > 0:
                total_price = real_price + pv_amount
                st.info(f"**Цена по документам (ДКП + Счёт №1):** {format_amount(total_price)} руб.")
                st.info(f"**Счёт №2 (к доплате покупателем):** {format_amount(real_price)} руб.")
        except:
            pass

    if 'inv1' in st.session_state.uploaded_files:
        st.button("🔄 Авто-заполнить из Счёта №1", on_click=auto_fill, use_container_width=True)

    st.markdown("---")
    if st.button("✅ Сгенерировать документы", type="primary", use_container_width=True):
        errors = []
        if not any(k in st.session_state.uploaded_files for k in ['dkp', 'pko', 'inv1', 'inv2']):
            errors.append("Загрузите все необходимые шаблоны в боковой панели")
        if not st.session_state.buyer_name.strip():
            errors.append("Не заполнено ФИО покупателя")
        
        real_price = parse_amount(st.session_state.real_price)
        pv_amount = parse_amount(st.session_state.pv_amount)
        
        if real_price <= 0: errors.append("Укажите корректную реальную цену авто")
        if pv_amount <= 0: errors.append("Укажите корректную сумму ПВ")

        if errors:
            for error in errors: st.error(error)
        else:
            total = real_price + pv_amount
            params = {
                "buyer_name": st.session_state.buyer_name.strip(),
                "dkp_number": st.session_state.dkp_number.strip(),
                "date": st.session_state.date.strip(),
                "car_brand": st.session_state.car_brand.strip(),
                "car_vin": st.session_state.car_vin.strip(),
                "car_color": st.session_state.car_color.strip(),
                "car_reg": st.session_state.car_reg.strip(),
                "real_price": real_price,
                "pv_amount": pv_amount,
                "new_price": total,
            }

            surname = (st.session_state.buyer_name.strip().split() or ["Клиент"])[0]
            date_safe = st.session_state.date.replace(".", "-")

            tasks = [
                ("dkp", f"ДКП_{surname}_{date_safe}.docx", process_dkp),
                ("pko", f"ПКО_{surname}_{date_safe}.docx", process_pko),
                ("inv1", f"Счёт1_{surname}_{date_safe}.docx", lambda d, p: process_invoice(d, p, p["new_price"])),
                ("inv2", f"Счёт2_{surname}_{date_safe}.docx", lambda d, p: process_invoice(d, p, p["real_price"])),
            ]

            progress_bar = st.progress(0)
            generated_files = []

            for i, (key, fname, processor) in enumerate(tasks):
                if key in st.session_state.uploaded_files:
                    try:
                        doc = Document(BytesIO(st.session_state.uploaded_files[key]))
                        doc = processor(doc, params)
                        output = BytesIO()
                        doc.save(output)
                        generated_files.append((fname, output.getvalue()))
                    except Exception as e:
                        st.error(f"Ошибка при обработке {fname}: {e}")
                progress_bar.progress((i + 1) / len(tasks))

            progress_bar.empty()

            if generated_files:
                st.subheader("📥 Скачать готовые документы:")
                for fname, data in generated_files:
                    st.download_button(label=f"Скачать {fname}", data=data, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    main()
