"""
FRESH - Генератор документов с искусственным ПВ
Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках
v3 - исправлен process_pko (замены в таблицах и параграфах)
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
# РАБОТА С DOCX — БАЗОВЫЕ УТИЛИТЫ
# ==============================================================================

A_PAT    = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
W_PAT    = (r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+"
            r"(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*\d{2}\s+копеек")
DATE_PAT = r"\b\d{2}\.\d{2}\.\d{4}\b"

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

def _set_para(para, text: str):
    """Записывает текст, сохраняя форматирование первого run"""
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(text)

def _remove_nds(text: str) -> str:
    text = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.;\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\(?22/122%?\)?\s*[\d\s,.]+руб\.?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"НДС\s*\(?22/122%?\)?\s*[\d\s,.]+руб\.?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def insert_paragraph_after(paragraph, text, style_source_para=None):
    """Вставка абзаца после текущего"""
    new_p = OxmlElement('w:p')
    paragraph._p.getparent().insert(
        paragraph._p.getparent().index(paragraph._p) + 1, new_p)
    new_para = paragraph.__class__(new_p, paragraph._parent)
    run = new_para.add_run(text)
    run.font.size = Pt(8)
    if style_source_para and style_source_para.runs:
        run.font.name = style_source_para.runs[0].font.name
    new_para.paragraph_format.space_before = Pt(0)
    new_para.paragraph_format.space_after = Pt(4)
    new_para.paragraph_format.line_spacing = 1.0
    return new_para

# ==============================================================================
# ПАРСИНГ СЧЁТА №1
# ==============================================================================

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
            r"([A-ZА-ЯЁ]{2,}\s+[A-ZА-ЯЁ0-9\s\-]+?)\s+([а-яё]+)\s+№\s*([A-ZА-ЯЁ0-9]+)\s+VIN\s+([A-Z0-9]{17})",
            t, re.IGNORECASE)
        if m and "car_vin" not in data:
            data["car_brand"] = m.group(1).strip()
            data["car_color"] = m.group(2).strip()
            data["car_reg"]   = m.group(3).strip()
            data["car_vin"]   = m.group(4).strip()

    for tbl in doc.tables:
        for row in tbl.rows:
            row_text = " ".join(c.text.strip() for c in row.cells)

            if "car_vin" not in data:
                m = re.search(r"VIN\s*[:\s]*([A-Z0-9]{17})", row_text, re.IGNORECASE)
                if m: data["car_vin"] = m.group(1)

            if "car_brand" not in data:
                m = re.search(r"([A-ZА-ЯЁ]{2,}\s+[A-ZА-ЯЁ0-9\s\-]+)", row_text)
                if m: data["car_brand"] = m.group(1).strip()

            if "car_color" not in data:
                m = re.search(
                    r"\b(серый|белый|чёрный|черный|синий|красный|серебристый"
                    r"|золотистый|коричневый|зелёный|зеленый|бежевый|жёлтый|желтый)\b",
                    row_text, re.IGNORECASE)
                if m: data["car_color"] = m.group(1).lower()

            if "car_reg" not in data:
                m = re.search(
                    r"(?:№|Гос\s*знак)\s*([A-ZА-ЯЁ]{1,2}\d{3}[A-ZА-ЯЁ]{2}\d{2,3})",
                    row_text, re.IGNORECASE)
                if m: data["car_reg"] = m.group(1)

    return data

# ==============================================================================
# ПРОЦЕССОР ДКП — БЕЗ ИЗМЕНЕНИЙ (работает отлично)
# ==============================================================================

def process_dkp(doc: Document, p: dict) -> Document:
    new_price = p["new_price"]
    pv_amount = p["pv_amount"]
    new_str   = format_amount(new_price)
    new_words = amount_to_words(new_price)
    pv_str    = format_amount(pv_amount)
    pv_words  = amount_to_words(pv_amount)

    pv_para_found      = False
    target_payment_para = None

    for para in doc.paragraphs:
        full            = "".join(r.text for r in para.runs)
        full_normalized = re.sub(r"\s+", " ", full).strip()

        if not full_normalized:
            continue

        if full_normalized.startswith("Цена ТС оплачивается Покупателем в течение"):
            target_payment_para = para

        if any(word in full_normalized.lower() for word in
               ["гаранти", "техническая защита", "лимит ответственности", "пробег"]):
            continue

        if "первоначальный" in full_normalized.lower() or "взнос" in full_normalized.lower():
            clean_full = full_normalized
            if re.search(A_PAT, clean_full):
                clean_full = re.sub(A_PAT, pv_str, clean_full)
            if re.search(W_PAT, clean_full):
                clean_full = re.sub(W_PAT, pv_words, clean_full)
            para.text = ""
            run = para.add_run(clean_full)
            run.font.size = Pt(8)
            pv_para_found = True

        elif any(marker in full_normalized.lower() for marker in
                 ["цена договора", "стоимость тс", "цена тс", "цену за тс",
                  "стоимость автомобиля", "уплачивает покупатель"]):
            clean_full = _remove_nds(full_normalized)
            has_amt = re.search(A_PAT, clean_full)
            has_wrd = re.search(W_PAT, clean_full)
            if has_amt and has_wrd:
                clean_full = re.sub(A_PAT, new_str, clean_full)
                suffix = " )" if ")" in has_wrd.group(0) else ""
                clean_full = re.sub(W_PAT, new_words + suffix, clean_full)
            elif has_amt:
                clean_full = re.sub(A_PAT, new_str, clean_full)
            clean_full = clean_full.replace(" ) )", " )").replace("))", ")")
            clean_full = re.sub(r",\s*\.", ".", clean_full)
            clean_full = re.sub(r"\s+", " ", clean_full).strip()
            if not clean_full.endswith("."):
                clean_full += "."
            para.text = ""
            run = para.add_run(clean_full)
            run.font.size = Pt(8)

    if not pv_para_found and target_payment_para is not None:
        pv_text = (f"Первоначальный взнос по оплате цены Договора составляет "
                   f"{pv_str} руб ({pv_words}).")
        insert_paragraph_after(target_payment_para, pv_text,
                                style_source_para=target_payment_para)

    return doc

# ==============================================================================
# ПРОЦЕССОР ПКО — v4 (написан под реальную структуру файла)
# ==============================================================================

def process_pko(doc: Document, p: dict) -> Document:
    """
    ПКО — структура: почти всё в параграфах, таблицы минимальны.
    Замены делаются точечно по содержимому каждого параграфа.
    """
    pv_str   = format_amount(p["pv_amount"])
    pv_words = amount_to_words(p["pv_amount"])
    date     = p["date"]
    buyer    = p["buyer_name"]
    osnov1   = f"По ДКП №{p['dkp_number']} от {date}"
    osnov2   = f"за а/м {p['car_brand']} {p['car_color']} № {p['car_reg']} VIN {p['car_vin']}"
    osnov_full = f"{osnov1} {osnov2}"

    DATE_FULL = r"\b\d{2}\.\d{2}\.\d{4}\b"
    A_FULL    = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
    W_FULL    = r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*\d{2}\s+копеек\s*"

    for para in doc.paragraphs:
        full = _para_text(para)
        t    = full.strip()
        if not t:
            continue

        # "от ДД.ММ.ГГГГ" — дата в шапке
        if re.match(r"^от\s+\d{2}\.\d{2}\.\d{4}$", t):
            _set_para(para, f"от {date}")
            continue

        # "подразделение\tФИО" — ФИО в правой части шапки
        if t.startswith("подразделение") and "\t" in full:
            _set_para(para, full[: full.index("\t") + 1] + buyer)
            continue

        # "По ДКП №... от дата" — первая строка основания (без "за а/м")
        if re.match(r"^По ДКП №", t) and "за а/м" not in t:
            _set_para(para, osnov1)
            continue

        # "за а/м ..." — вторая строка основания
        if t.startswith("за а/м "):
            _set_para(para, osnov2)
            continue

        # "По ДКП №... за а/м..." — полное основание одной строкой
        if re.match(r"^По ДКП №", t) and "за а/м" in t:
            _set_para(para, osnov_full)
            continue

        # "Сумма\t428 000,00" — сумма цифрами в квитанции
        if re.match(r"^Сумма", t) and re.search(A_FULL, t):
            _set_para(para, re.sub(A_FULL, pv_str, full))
            continue

        # "Принято от: ФИО\tпропись" — ФИО + пропись через таб
        if t.startswith("Принято от:") and "\t" in full:
            _set_para(para, f"Принято от: {buyer}\t{pv_words}")
            continue

        # Только прописью (квитанция — правая часть)
        if re.match(W_FULL, t) and not re.search(r"Принято|Основание|Сумма", t):
            _set_para(para, pv_words)
            continue

        # "В том числе" + "НДС..." — убираем обе строки
        if t in ("В том числе", "В том числе:"):
            _set_para(para, "")
            continue

        if re.match(r"НДС\s*\(22/122\)", t):
            _set_para(para, "")
            continue

        # "В том числе: НДС...\tМ.П. (штампа)" — оставляем только М.П.
        if t.startswith("В том числе") and "НДС" in t and "\t" in full:
            _set_para(para, "\t" + full[full.index("\t") + 1:])
            continue

        # Голая дата "ДД.ММ.ГГГГ"
        if re.fullmatch(DATE_FULL, t):
            _set_para(para, date)
            continue

    # Таблицы (их мало, но на всякий случай — дата и ФИО)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    ct = para.text.strip()
                    if re.fullmatch(DATE_FULL, ct):
                        _set_para(para, date)
                    elif re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+", ct):
                        # ФИО — три слова с заглавной, но не "Главный бухгалтер" и т.п.
                        if not re.search(r"Главный|Кассир|Робот|Принято|КАССОВЫЙ", ct):
                            _set_para(para, buyer)
                    elif (re.search(A_FULL, ct) and
                          not re.search(r"18\.08|88|0310001|ОКУД|ОКПО|КО-1|ЦБ", ct)):
                        _set_para(para, re.sub(A_FULL, pv_str, _para_text(para)))

    return doc

# ==============================================================================
# ПРОЦЕССОР СЧЁТОВ — БЕЗ ИЗМЕНЕНИЙ (работает хорошо)
# ==============================================================================

def process_invoice(doc: Document, p: dict, amount: float) -> Document:
    amt_str   = format_amount(amount)
    amt_words = amount_to_words(amount)
    date      = p["date"]

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
                        nf = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                        _set_para(para, nf)

                elif re.search(A_PAT, ct):
                    if j == nds_sum_col and ri > 0:
                        for para in cell.paragraphs:
                            if re.search(A_PAT, _para_text(para)):
                                _set_para(para, "0,00")
                    else:
                        if ri > 0 and (j >= len(row.cells) - 2):
                            for para in cell.paragraphs:
                                if re.search(A_PAT, _para_text(para)):
                                    _set_para(para, amt_str)

    for para in doc.paragraphs:
        full     = _para_text(para)
        new_full = full

        if re.search(DATE_PAT, full) and "счет" in full.lower():
            new_full = re.sub(DATE_PAT, date, new_full)

        if re.search(W_PAT, new_full):
            new_full = re.sub(W_PAT, amt_words, new_full)
            new_full = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", new_full)
            new_full = re.sub(r"\s{2,}", " ", new_full)

        new_full = _remove_nds(new_full)

        if new_full != full:
            _set_para(para, new_full)

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
                doc  = Document(BytesIO(st.session_state.uploaded_files['inv1']))
                data = extract_invoice_data(doc)
                updates = []
                mapping = {
                    'buyer_name': 'ФИО',
                    'dkp_number': 'Номер ДКП',
                    'date':       'Дата',
                    'car_brand':  'Марка',
                    'car_vin':    'VIN',
                    'car_color':  'Цвет',
                    'car_reg':    'Гос.номер',
                }
                for key, label in mapping.items():
                    if data.get(key):
                        st.session_state[key] = data[key]
                        updates.append(label)
                if updates:
                    st.success(f"✅ Импортировано: {', '.join(updates)}")
                else:
                    st.warning("Не удалось извлечь данные. Проверьте формат Счёта №1.")
            except Exception as e:
                st.error(f"Ошибка парсинга: {e}")

    # ── БОКОВАЯ ПАНЕЛЬ ──────────────────────────────────────
    with st.sidebar:
        st.header("📁 Шаблоны документов")

        dkp_file  = st.file_uploader("ДКП",                            type=['docx'], key='dkp_file')
        pko_file  = st.file_uploader("ПКО",                            type=['docx'], key='pko_file')
        inv1_file = st.file_uploader("Счёт №1  ⬅️ авто-заполнение",   type=['docx'], key='inv1_file')
        inv2_file = st.file_uploader("Счёт №2",                        type=['docx'], key='inv2_file')

        if dkp_file:  st.session_state.uploaded_files['dkp']  = dkp_file.getvalue()
        if pko_file:  st.session_state.uploaded_files['pko']  = pko_file.getvalue()
        if inv1_file: st.session_state.uploaded_files['inv1'] = inv1_file.getvalue()
        if inv2_file: st.session_state.uploaded_files['inv2'] = inv2_file.getvalue()

        st.markdown("---")
        st.markdown("### Статус загрузки:")
        for name, label in [('dkp','ДКП'), ('pko','ПКО'),
                             ('inv1','Счёт №1'), ('inv2','Счёт №2')]:
            if name in st.session_state.uploaded_files:
                st.success(f"✅ {label} загружен")
            else:
                st.warning(f"⚠️ {label} не загружен")

    # ── ОСНОВНАЯ ЧАСТЬ ──────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.header("📋 Данные сделки")

        defaults = {
            'buyer_name': '', 'dkp_number': '',
            'date': datetime.today().strftime("%d.%m.%Y"),
            'car_brand': '', 'car_vin': '',
            'car_color': '', 'car_reg': '',
            'real_price': '', 'pv_amount': '',
        }
        for k, v in defaults.items():
            if k not in st.session_state:
                st.session_state[k] = v

        st.text_input("ФИО покупателя",  key='buyer_name')
        st.text_input("Номер ДКП",        key='dkp_number')
        st.text_input("Дата документов",  key='date',
                      help="Формат: ДД.ММ.ГГГГ")
        st.text_input("Марка / Модель",   key='car_brand')
        st.text_input("VIN",              key='car_vin')
        st.text_input("Цвет",             key='car_color')
        st.text_input("Гос. номер",       key='car_reg')

    with col2:
        st.header("💰 Суммы")

        st.text_input("Реальная цена авто (₽)",      key='real_price',
                      placeholder="3 250 000")
        st.text_input("Сумма искусственного ПВ (₽)", key='pv_amount',
                      placeholder="250 000")

        real_price = parse_amount(st.session_state.real_price)
        pv_amount  = parse_amount(st.session_state.pv_amount)

        if real_price > 0 and pv_amount > 0:
            total = real_price + pv_amount
            st.success(f"**Цена по документам (ДКП + Счёт №1):**  \n{format_amount(total)} ₽")
            st.info(f"**Счёт №2 (к доплате покупателем):**  \n{format_amount(real_price)} ₽")
            st.markdown(f"""
**Прописью (цена по документам):**  
_{amount_to_words(total)}_

**Прописью (ПВ):**  
_{amount_to_words(pv_amount)}_
""")

        # Авто-заполнение из счёта
        if 'inv1' in st.session_state.uploaded_files:
            st.button("🔄 Авто-заполнить из Счёта №1",
                      on_click=auto_fill, use_container_width=True)

    st.markdown("---")

    # ── ГЕНЕРАЦИЯ ───────────────────────────────────────────
    errors = []
    missing_files = [lbl for k, lbl in [('dkp','ДКП'),('pko','ПКО'),
                                         ('inv1','Счёт №1'),('inv2','Счёт №2')]
                     if k not in st.session_state.uploaded_files]
    if missing_files:
        errors.append(f"Не загружены файлы: {', '.join(missing_files)}")
    if not st.session_state.buyer_name.strip():
        errors.append("Не заполнено ФИО покупателя")
    if parse_amount(st.session_state.real_price) <= 0:
        errors.append("Укажите реальную цену авто")
    if parse_amount(st.session_state.pv_amount) <= 0:
        errors.append("Укажите сумму ПВ")

    if errors:
        for e in errors:
            st.warning(f"⚠️ {e}")

    btn = st.button(
        "✅  Сгенерировать документы",
        type="primary",
        use_container_width=True,
        disabled=bool(errors))

    if btn:
        real  = parse_amount(st.session_state.real_price)
        pv    = parse_amount(st.session_state.pv_amount)
        total = real + pv

        params = {
            "buyer_name": st.session_state.buyer_name.strip(),
            "dkp_number": st.session_state.dkp_number.strip(),
            "date":       st.session_state.date.strip(),
            "car_brand":  st.session_state.car_brand.strip(),
            "car_vin":    st.session_state.car_vin.strip(),
            "car_color":  st.session_state.car_color.strip(),
            "car_reg":    st.session_state.car_reg.strip(),
            "real_price": real,
            "pv_amount":  pv,
            "new_price":  total,
        }

        surname   = (st.session_state.buyer_name.strip().split() or ["Клиент"])[0]
        date_safe = st.session_state.date.replace(".", "-")

        tasks = [
            ("dkp",  f"ДКП_{surname}_{date_safe}.docx",
             lambda d, p: process_dkp(d, p)),
            ("pko",  f"ПКО_{surname}_{date_safe}.docx",
             lambda d, p: process_pko(d, p)),
            ("inv1", f"Счёт1_{surname}_{date_safe}.docx",
             lambda d, p: process_invoice(d, p, p["new_price"])),
            ("inv2", f"Счёт2_{surname}_{date_safe}.docx",
             lambda d, p: process_invoice(d, p, p["real_price"])),
        ]

        progress_bar    = st.progress(0, text="Начинаю обработку...")
        generated_files = []

        for i, (key, fname, processor) in enumerate(tasks):
            if key in st.session_state.uploaded_files:
                try:
                    progress_bar.progress(i / len(tasks), text=f"Обрабатываю: {fname}")
                    doc    = Document(BytesIO(st.session_state.uploaded_files[key]))
                    doc    = processor(doc, params)
                    output = BytesIO()
                    doc.save(output)
                    generated_files.append((fname, output.getvalue()))
                except Exception as e:
                    st.error(f"❌ Ошибка при обработке {fname}: {e}")
            progress_bar.progress((i + 1) / len(tasks))

        progress_bar.empty()

        if generated_files:
            st.success(f"✅ Сгенерировано {len(generated_files)} документа")
            st.subheader("📥 Скачать готовые документы:")
            cols = st.columns(len(generated_files))
            for col, (fname, data) in zip(cols, generated_files):
                with col:
                    st.download_button(
                        label=f"📄 {fname}",
                        data=data,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True)

if __name__ == "__main__":
    main()
