
"""
Fresh Auto — Генератор документов с искусственным ПВ
Streamlit web-интерфейс
"""

import streamlit as st
import re, io, json, zipfile
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
import num2words as nw

# ══════════════════════════════════════════════════════════════════
#  УТИЛИТЫ СУММ
# ══════════════════════════════════════════════════════════════════

def _rub_word(n: int) -> str:
    n100, n10 = n % 100, n % 10
    if 11 <= n100 <= 19: return "рублей"
    if n10 == 1: return "рубль"
    if 2 <= n10 <= 4: return "рубля"
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
    cleaned = re.sub(r"[^\d,.]", "", text.strip())
    cleaned = cleaned.replace(",", ".")
    parts = cleaned.split(".")
    if len(parts) > 2:
        cleaned = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

# ══════════════════════════════════════════════════════════════════
#  DOCX УТИЛИТЫ
# ══════════════════════════════════════════════════════════════════

_AMOUNT_PAT = r"\d[\d\s]*\d[,\.]\d{2}"
_WORDS_PAT  = (r"[А-ЯЁ][а-яёА-ЯЁ\s]+"
               r"(?:тысяч|миллион|миллиард)[а-яё\s]*"
               r"(?:рубл[а-яё]+)\s+\d{2}\s+копеек")

def _iter_paragraphs(doc):
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

def _set_para(para, new_text: str):
    if not para.runs:
        new_run = OxmlElement("w:r")
        new_t = OxmlElement("w:t")
        new_t.text = new_text
        new_t.set(qn('xml:space'), 'preserve')
        new_run.append(new_t)
        para._element.append(new_run)
        return
    
    para.runs[0].text = new_text
    for r in para.runs[1:]:
        r.text = ""

def _remove_nds(text: str) -> str:
    text = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.;\n]*", "", text)
    text = re.sub(r"\(22/122%?\)\s*[\d\s,]+руб\.?", "", text)
    text = re.sub(r"НДС\s*\(22/122%?\)\s*[\d\s,]+руб\.?", "без НДС", text)
    return text

# ══════════════════════════════════════════════════════════════════
#  ПАРСИНГ СЧЁТА №1
# ══════════════════════════════════════════════════════════════════

def extract_invoice_data(doc: Document) -> dict:
    data = {}
    for para in _iter_paragraphs(doc):
        t = para.text.strip()

        m = re.search(
            r"[Пп]родажа\s+[тТ][/\\][сС]\s+№\s*([\w\-/]+).*?от\s+(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))", t)
        if m and "dkp_number" not in data:
            data["dkp_number"] = m.group(1).strip()
            parts = m.group(2).split(".")
            if len(parts[2]) == 2: parts[2] = "20" + parts[2]
            data["date"] = ".".join(parts)

        m = re.search(
            r"[Пп]окупатель[:\s]+(?:ИНН\s+\d+[,\s]+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)", t)
        if m and "buyer_name" not in data:
            data["buyer_name"] = m.group(1).strip()

        m = re.search(
            r"([A-Z]{2,}\s+[A-Z]{2,})\s+([а-яё]+)\s+№\s*([A-ZА-ЯЁa-zA-Zа-яё0-9]+)\s+VIN\s+([A-Z0-9]{17})", t)
        if m and "car_vin" not in data:
            data["car_brand"] = m.group(1).strip()
            data["car_color"] = m.group(2).strip()
            data["car_reg"]   = m.group(3).strip()
            data["car_vin"]   = m.group(4).strip()

    for tbl in doc.tables:
        for row in tbl.rows:
            row_text = " ".join(c.text.strip() for c in row.cells)
            if "car_vin" not in data:
                m = re.search(r"VIN\s+([A-Z0-9]{17})", row_text)
                if m: data["car_vin"] = m.group(1)
            if "car_brand" not in data:
                m = re.search(r"([A-Z]{2,}\s+[A-Z]{2,})", row_text)
                if m: data["car_brand"] = m.group(1)
            if "car_color" not in data:
                m = re.search(
                    r"\b(серый|белый|чёрный|черный|синий|красный|серебристый"
                    r"|золотистый|коричневый|зелёный|зеленый|бежевый|жёлтый|желтый)\b",
                    row_text, re.IGNORECASE)
                if m: data["car_color"] = m.group(1).lower()
            if "car_reg" not in data:
                m = re.search(r"№\s*([A-ZА-ЯЁ]{1,2}\d{3}[A-ZА-ЯЁ]{2}\d{2,3})", row_text)
                if m: data["car_reg"] = m.group(1)
    return data

# ══════════════════════════════════════════════════════════════════
#  ОБРАБОТКА ДОКУМЕНТОВ
# ══════════════════════════════════════════════════════════════════

def process_dkp(doc: Document, p: dict) -> Document:
    new_str   = format_amount(p["new_price"])
    new_words = amount_to_words(p["new_price"])
    pv_str    = format_amount(p["pv_amount"])
    pv_words  = amount_to_words(p["pv_amount"])
    pv_found  = False

    for para in _iter_paragraphs(doc):
        full = _para_text(para)
        if not full.strip(): 
            continue

        if "первоначальный взнос" in full.lower():
            pv_found = True
            nf = re.sub(_AMOUNT_PAT, pv_str, full)
            nf = re.sub(_WORDS_PAT, pv_words, nf)
            if nf != full: 
                _set_para(para, nf)
            continue

        has_amount = bool(re.search(_AMOUNT_PAT, full))
        has_words = bool(re.search(_WORDS_PAT, full))
        has_nds = bool(re.search(r"НДС|22/122", full))
        
        if has_amount:
            nf = re.sub(_AMOUNT_PAT, new_str, full)
            nf = _remove_nds(nf)
            if nf != full: 
                _set_para(para, nf)
        elif has_words:
            nf = re.sub(_WORDS_PAT, new_words, full)
            if nf != full: 
                _set_para(para, nf)
        elif has_nds:
            nf = _remove_nds(full)
            if nf != full: 
                _set_para(para, nf)

    if not pv_found:
        pv_text = (f"Первоначальный взнос по оплате цены Договора составляет "
                   f"{pv_str} руб ({pv_words}).")
        
        inserted = False
        for para in doc.paragraphs:
            if re.search(r"размере\s+" + _AMOUNT_PAT, _para_text(para)):
                try:
                    new_p = deepcopy(para._element)
                    for r_elem in new_p.findall(qn('w:r')):
                        new_p.remove(r_elem)
                    
                    new_r = OxmlElement("w:r")
                    if para.runs and para.runs[0]._r.find(qn("w:rPr")) is not None:
                        rpr = deepcopy(para.runs[0]._r.find(qn("w:rPr")))
                        new_r.append(rpr)
                    
                    new_t = OxmlElement("w:t")
                    new_t.text = pv_text
                    new_t.set(qn('xml:space'), 'preserve')
                    new_r.append(new_t)
                    new_p.append(new_r)
                    
                    para._element.addnext(new_p)
                    inserted = True
                    break
                except Exception:
                    continue
        
        if not inserted and doc.paragraphs:
            try:
                last_para = doc.paragraphs[-1]
                new_p = OxmlElement("w:p")
                new_r = OxmlElement("w:r")
                new_t = OxmlElement("w:t")
                new_t.text = pv_text
                new_t.set(qn('xml:space'), 'preserve')
                new_r.append(new_t)
                new_p.append(new_r)
                last_para._element.addnext(new_p)
            except Exception:
                pass
    
    return doc


def process_pko(doc: Document, p: dict) -> Document:
    pv_str   = format_amount(p["pv_amount"])
    pv_words = amount_to_words(p["pv_amount"])
    date     = p["date"]
    buyer    = p["buyer_name"]
    osnov    = (f"По ДКП №{p['dkp_number']} от {date} "
                f"за а/м {p['car_brand']} {p['car_color']} "
                f"№{p['car_reg']} VIN {p['car_vin']}")

    for para in _iter_paragraphs(doc):
        full = _para_text(para)
        if not full.strip(): continue

        if re.search(r"[Пп]о\s+ДКП|за\s+а/м|VIN\s+[A-Z0-9]{17}", full):
            _set_para(para, osnov); continue

        if re.search(r"НДС|22/122", full):
            nf = re.sub(r"НДС\s*\(22/122%?\)[^\n]*", "Без НДС", full)
            nf = re.sub(r"В том числе НДС[^\n]*", "Без НДС", nf)
            nf = _remove_nds(nf)
            if nf != full: _set_para(para, nf)
            continue

        if re.search(_WORDS_PAT, full):
            nf = re.sub(_WORDS_PAT, pv_words, full)
            if nf != full: _set_para(para, nf)
            continue

        if re.search(_AMOUNT_PAT, full) and not re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
            nf = re.sub(_AMOUNT_PAT, pv_str, full)
            if nf != full: _set_para(para, nf)
            continue

        if re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", full):
            nf = re.sub(r"\b\d{2}\.\d{2}\.\d{4}\b", date, full)
            if nf != full: _set_para(para, nf)
            continue

        m = re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+", full)
        if m and m.group(0) != buyer:
            _set_para(para, full.replace(m.group(0), buyer))
    return doc


def process_invoice(doc: Document, p: dict, amount: float) -> Document:
    amt_str   = format_amount(amount)
    amt_words = amount_to_words(amount)
    date      = p["date"]

    for tbl in doc.tables:
        nds_sum_col = -1
        if tbl.rows:
            for ci, hc in enumerate(tbl.rows[0].cells):
                if re.search(r"[Сс]умма\s*НДС|НДС\s*[Сс]умма", hc.text):
                    nds_sum_col = ci; break
            if nds_sum_col == -1 and len(tbl.rows) > 1:
                rate_col = next((ci for ci, c in enumerate(tbl.rows[1].cells)
                                 if re.search(r"22/122", c.text)), -1)
                if rate_col >= 0 and rate_col + 1 < len(tbl.rows[1].cells):
                    nds_sum_col = rate_col + 1

        for ri, row in enumerate(tbl.rows):
            for j, cell in enumerate(row.cells):
                ct = cell.text.strip()
                if re.search(r"22/122", ct):
                    for para in cell.paragraphs:
                        nf = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                        if nf != _para_text(para): _set_para(para, nf)
                elif re.search(_AMOUNT_PAT, ct):
                    if j == nds_sum_col and ri > 0:
                        for para in cell.paragraphs:
                            nf = re.sub(_AMOUNT_PAT, "0,00", _para_text(para))
                            if nf != _para_text(para): _set_para(para, nf)
                    else:
                        for para in cell.paragraphs:
                            nf = re.sub(_AMOUNT_PAT, amt_str, _para_text(para))
                            if nf != _para_text(para): _set_para(para, nf)

    for para in doc.paragraphs:
        full = _para_text(para)
        nf = re.sub(r"\b\d{2}\.\d{2}\.(?:\d{2}|\d{4})\b", date, full)
        if re.search(_WORDS_PAT, nf):
            nf = re.sub(_WORDS_PAT, amt_words, nf)
            nf = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", nf)
            nf = re.sub(r"\s{2,}", " ", nf)
        nf = _remove_nds(nf)
        if nf != full: _set_para(para, nf)
    return doc

# ══════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Fresh Auto — ПВ",
    page_icon="🚗",
    layout="centered",
)

# Стили
st.markdown("""
<style>
    .stApp { background-color: #0f1729; }
    section[data-testid="stSidebar"] { display: none; }
    
    h1 { color: #4a90d9 !important; }
    h2, h3 { color: #7ab8f5 !important; }
    
    .block-container { max-width: 860px; padding-top: 2rem; }
    
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input {
        background: #1e2f56 !important;
        color: white !important;
        border: 1px solid #2a4080 !important;
        border-radius: 6px !important;
    }
    
    .calc-box {
        background: #1a2a4a;
        border: 1px solid #2a4a80;
        border-radius: 10px;
        padding: 18px 24px;
        margin: 12px 0;
    }
    .calc-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 6px 0;
    }
    .calc-label { color: #7899cc; font-size: 14px; }
    .calc-value { color: #2ecc71; font-size: 18px; font-weight: bold; }
    
    .result-box {
        background: #0d2a1a;
        border: 1px solid #1a6a3a;
        border-radius: 10px;
        padding: 16px 24px;
        margin: 12px 0;
    }
    .file-item {
        background: #1e2f56;
        border-radius: 6px;
        padding: 8px 14px;
        margin: 4px 0;
        color: #aaddff;
        font-size: 13px;
    }
    div[data-testid="stFileUploader"] {
        background: #1a2540 !important;
        border: 1px dashed #2a4080 !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# ── Заголовок ──────────────────────────────────────────────────
st.markdown("# 🚗 Fresh Auto")
st.markdown("### Генератор документов с искусственным первоначальным взносом")
st.divider()

# ══════════════════════════════════════════════════════════════════
# БЛОК 1 — ЗАГРУЗКА ФАЙЛОВ
# ══════════════════════════════════════════════════════════════════
st.markdown("## 📁 Шаблоны документов")
st.caption("Загрузите 4 файла .docx. Счёт №1 заполнит поля автоматически.")

col1, col2 = st.columns(2)
with col1:
    f_dkp  = st.file_uploader("ДКП",     type="docx", key="dkp")
    f_inv1 = st.file_uploader("Счёт №1", type="docx", key="inv1",
                               help="После загрузки поля заполнятся автоматически")
with col2:
    f_pko  = st.file_uploader("ПКО",     type="docx", key="pko")
    f_inv2 = st.file_uploader("Счёт №2", type="docx", key="inv2")

# ── Авто-парсинг при загрузке Счёта №1 ──
auto_data = {}
if f_inv1:
    try:
        doc_tmp = Document(io.BytesIO(f_inv1.read()))
        f_inv1.seek(0)
        auto_data = extract_invoice_data(doc_tmp)
        if auto_data:
            st.success(f"✓ Данные из Счёта №1 загружены: {', '.join(k for k in auto_data if auto_data[k])}")
    except Exception as e:
        st.warning(f"Не удалось распарсить Счёт №1: {e}")

st.divider()

# ══════════════════════════════════════════════════════════════════
# БЛОК 2 — ДАННЫЕ СДЕЛКИ
# ══════════════════════════════════════════════════════════════════
st.markdown("## 📋 Данные сделки")

col1, col2 = st.columns(2)
with col1:
    buyer_name = st.text_input(
        "ФИО покупателя *",
        value=auto_data.get("buyer_name", ""),
        placeholder="Иванов Иван Иванович")
    
    dkp_number = st.text_input(
        "Номер ДКП",
        value=auto_data.get("dkp_number", ""),
        placeholder="ФКП-20/06/26-8")
    
    date = st.text_input(
        "Дата документов *",
        value=auto_data.get("date", ""),
        placeholder="20.06.2026")
    
    car_brand = st.text_input(
        "Марка / Модель",
        value=auto_data.get("car_brand", ""),
        placeholder="HYUNDAI ELANTRA")

with col2:
    car_vin = st.text_input(
        "VIN",
        value=auto_data.get("car_vin", ""),
        placeholder="LBECNAFD6MZ128269")
    
    car_color = st.text_input(
        "Цвет",
        value=auto_data.get("car_color", ""),
        placeholder="серый")
    
    car_reg = st.text_input(
        "Гос. номер",
        value=auto_data.get("car_reg", ""),
        placeholder="C176OO193")

st.divider()

# ══════════════════════════════════════════════════════════════════
# БЛОК 3 — СУММЫ
# ══════════════════════════════════════════════════════════════════
st.markdown("## 💰 Суммы")

col1, col2 = st.columns(2)
with col1:
    real_price_str = st.text_input(
        "Реальная цена авто (₽) *",
        placeholder="2 000 000",
        help="Фактическая стоимость автомобиля")
with col2:
    pv_amount_str = st.text_input(
        "Сумма искусственного ПВ (₽) *",
        placeholder="138 000",
        help="Первоначальный взнос для банка")

# Расчёт
real_price = parse_amount(real_price_str) if real_price_str else 0.0
pv_amount  = parse_amount(pv_amount_str)  if pv_amount_str  else 0.0
doc_price  = real_price + pv_amount

if real_price > 0 and pv_amount > 0:
    st.markdown(f"""
    <div class="calc-box">
        <div class="calc-row">
            <span class="calc-label">📄 Цена по документам (ДКП + Счёт №1)</span>
            <span class="calc-value">{format_amount(doc_price)} ₽</span>
        </div>
        <div class="calc-row" style="border-top:1px solid #2a4a80; padding-top:10px; margin-top:6px;">
            <span class="calc-label">📄 Счёт №2 (к доплате покупателем)</span>
            <span class="calc-value">{format_amount(real_price)} ₽</span>
        </div>
        <div style="margin-top:12px; padding-top:10px; border-top:1px solid #1a3060;">
            <span class="calc-label">Прописью (цена по документам):</span><br>
            <span style="color:#aaddff; font-size:13px;">{amount_to_words(doc_price)}</span>
        </div>
        <div style="margin-top:8px;">
            <span class="calc-label">Прописью (ПВ):</span><br>
            <span style="color:#aaddff; font-size:13px;">{amount_to_words(pv_amount)}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.info("Введите реальную цену и сумму ПВ для расчёта", icon="💡")

st.divider()

# ══════════════════════════════════════════════════════════════════
# КНОПКА ГЕНЕРАЦИИ
# ══════════════════════════════════════════════════════════════════
st.markdown("## ✅ Генерация")

# Проверка готовности
missing = []
if not f_dkp:  missing.append("ДКП")
if not f_pko:  missing.append("ПКО")
if not f_inv1: missing.append("Счёт №1")
if not f_inv2: missing.append("Счёт №2")
if not buyer_name.strip(): missing.append("ФИО покупателя")
if real_price <= 0: missing.append("реальная цена")
if pv_amount <= 0:  missing.append("сумма ПВ")

if missing:
    st.warning(f"Не заполнено: {', '.join(missing)}", icon="⚠️")
else:
    st.success("Все данные заполнены — готово к генерации!", icon="✓")

btn_generate = st.button(
    "🚀 Сгенерировать документы и скачать ZIP",
    type="primary",
    use_container_width=True,
    disabled=bool(missing))

if btn_generate:
    params = {
        "buyer_name": buyer_name.strip(),
        "dkp_number": dkp_number.strip(),
        "date":       date.strip(),
        "car_brand":  car_brand.strip(),
        "car_vin":    car_vin.strip(),
        "car_color":  car_color.strip(),
        "car_reg":    car_reg.strip(),
        "pv_amount":  pv_amount,
        "new_price":  doc_price,
        "real_price": real_price,
    }

    surname   = (buyer_name.strip().split() or ["Клиент"])[0]
    date_safe = date.replace(".", "-")

    tasks = [
        (f_dkp,  f"ДКП_{surname}_{date_safe}.docx",   lambda d, p: process_dkp(d, p)),
        (f_pko,  f"ПКО_{surname}_{date_safe}.docx",   lambda d, p: process_pko(d, p)),
        (f_inv1, f"Счёт1_{surname}_{date_safe}.docx",  lambda d, p: process_invoice(d, p, p["new_price"])),
        (f_inv2, f"Счёт2_{surname}_{date_safe}.docx",  lambda d, p: process_invoice(d, p, p["real_price"])),
    ]

    zip_buf = io.BytesIO()
    errors  = []
    results = []

    progress = st.progress(0, text="Начинаю обработку...")

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (file_obj, fname, processor) in enumerate(tasks):
            try:
                progress.progress((i) / 4, text=f"Обрабатываю: {fname}")
                file_obj.seek(0)
                doc = Document(io.BytesIO(file_obj.read()))
                doc = processor(doc, params)

                doc_buf = io.BytesIO()
                doc.save(doc_buf)
                doc_buf.seek(0)
                zf.writestr(fname, doc_buf.read())
                results.append(fname)
            except Exception as e:
                errors.append(f"{fname}: {e}")

    progress.progress(1.0, text="✅ Готово!")

    if errors:
        for e in errors:
            st.error(e)
    
    if results:
        zip_buf.seek(0)
        
        st.markdown(f"""
        <div class="result-box">
            <b style="color:#2ecc71;">✅ Сгенерированы документы:</b><br>
            {"".join(f'<div class="file-item">📄 {f}</div>' for f in results)}
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label=f"📥 Скачать ZIP ({len(results)} документа)",
            data=zip_buf,
            file_name=f"ПВ_{surname}_{date_safe}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )

st.divider()
st.caption("Fresh Auto · Генератор документов с ПВ · v2.0")

