"""
FRESH - Генератор документов с искусственным ПВ
Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках
"""

import streamlit as st
import re, sys, subprocess
from datetime import datetime
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
    """Исправленная безопасная вставка нового абзаца строго после текущего"""
    new_p = OxmlElement('w:p')
    paragraph._p.getparent().insert(paragraph._p.getparent().index(paragraph._p) + 1, new_p)
    new_para = paragraph.__class__(new_p, paragraph._parent)
    new_para.text = text
    
    if style_source_para and style_source_para.runs:
        src_run = style_source_para.runs[0]
        style = src_run._element.find(qn("w:rPr"))
        if style is not None and new_para.runs:
            new_para.runs[0]._element.append(style)
            
    return new_para

_AMOUNT_PAT = r"\b\d[\d\s]{0,12}[,.]\d{2}\b"
_WORDS_PAT = (r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+"
              r"(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*"
              r"\d{2}\s+копеек")
_DATE_PAT = r"\b\d{2}\.\d{2}\.\d{4}\b"

def _remove_nds(text: str) -> str:
    text = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.;\n]*", "", text)
    text = re.sub(r"(22/122%?)\s*[\d\s,]+руб\.?", "", text)
    text = re.sub(r"НДС\s*(22/122%?)\s*[\d\s,]+руб\.?", "без НДС", text)
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
    
    for para in doc.paragraphs:
        full = _para_text(para)
        if not full.strip():
            continue
        
        if full.strip().startswith("Цена ТС оплачивается Покупателем в течение"):
            target_payment_para = para
        
        if any(word in full.lower() for
