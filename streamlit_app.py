"""
FRESH - Генератор документов с искусственным ПВ
v5 FINAL — process_pko переписан с нуля, без escape-багов
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
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "num2words", "-q"])
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt
    import num2words as nw

# ─────────────────────────────────────────────────────────────
# УТИЛИТЫ СУММ
# ─────────────────────────────────────────────────────────────

def _rub_word(n: int) -> str:
    n100, n10 = n % 100, n % 10
    if 11 <= n100 <= 19: return "рублей"
    if n10 == 1:         return "рубль"
    if 2 <= n10 <= 4:    return "рубля"
    return "рублей"

def format_amount(amount: float) -> str:
    rub = int(amount)
    kop = round((amount - rub) * 100)
    return f"{rub:,}".replace(",", " ") + f",{kop:02d}"

def amount_to_words(amount: float) -> str:
    rub = int(amount)
    kop = round((amount - rub) * 100)
    w = nw.num2words(rub, lang="ru")
    return w[0].upper() + w[1:] + f" {_rub_word(rub)} {kop:02d} копеек"

def parse_amount(text: str) -> float:
    if not text: return 0.0
    cleaned = re.sub(r"[^\d,.]", "", text.strip()).replace(",", ".")
    parts = cleaned.split(".")
    if len(parts) > 2:
        cleaned = "".join(parts[:-1]) + "." + parts[-1]
    try:    return float(cleaned)
    except: return 0.0

# ─────────────────────────────────────────────────────────────
# DOCX УТИЛИТЫ
# ─────────────────────────────────────────────────────────────

# Паттерны (одно место — без дублирования)
RE_DATE   = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
RE_AMOUNT = re.compile(r"\b\d[\d\s]{0,12}[,.]\d{2}\b")
RE_WORDS  = re.compile(
    r"[А-ЯЁа-яё][а-яёА-ЯЁ\s\-\,]+"
    r"(?:тысяч|миллион|миллиард|рубл)[а-яё\s\-\,]*"
    r"\d{2}\s+копеек\s*"
)

def _para_text(para) -> str:
    return "".join(r.text for r in para.runs)

def _set_para(para, text: str):
    """Пишет text в параграф, очищая все runs кроме первого."""
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(text)

def _remove_nds(text: str) -> str:
    text = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.;\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\(?22/122%?\)?\s*[\d\s,.]+руб\.?",   "", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()

def _iter_all_paragraphs(doc):
    """Все параграфы: тело + таблицы + колонтитулы."""
    yield from doc.paragraphs
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for sec in doc.sections:
        for hf in (sec.header, sec.footer):
            if hf:
                yield from hf.paragraphs

def insert_paragraph_after(paragraph, text, source_para=None):
    new_p = OxmlElement("w:p")
    paragraph._p.getparent().insert(
        paragraph._p.getparent().index(paragraph._p) + 1, new_p)
    new_para = paragraph.__class__(new_p, paragraph._parent)
    run = new_para.add_run(text)
    run.font.size = Pt(8)
    if source_para and source_para.runs:
        run.font.name = source_para.runs[0].font.name
    new_para.paragraph_format.space_before = Pt(0)
    new_para.paragraph_format.space_after  = Pt(4)
    new_para.paragraph_format.line_spacing = 1.0
    return new_para

# ─────────────────────────────────────────────────────────────
# ПАРСИНГ СЧЁТА №1
# ─────────────────────────────────────────────────────────────

def extract_invoice_data(doc: Document) -> dict:
    data = {}
    full_text = "\n".join([p.text for p in _iter_all_paragraphs(doc)])
    
    # 1. Поиск ДКП и даты
    m = re.search(r"[Пп]родажа\s+[тТ][/\\][сС]\s+№\s*([\w\-/]+).*?от\s+(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))", full_text)
    if m:
        data["dkp_number"] = m.group(1).strip()
        data["date"] = m.group(2) if len(m.group(2).split('.')[2]) == 4 else m.group(2)
        
    # 2. Поиск покупателя
    m = re.search(r"[Пп]окупатель[:\s]+(?:ИНН\s+\d+[,\s]+)?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)", full_text)
    if m: data["buyer_name"] = m.group(1).strip()
    
    # 3. Улучшенный парсинг авто (захватываем всё до цвета)
    # Ищем: Марка Модель [цвет] № [номер] VIN [вин]
    colors = ["серый", "белый", "черный", "чёрный", "синий", "красный", "серебристый", "золотистый", "коричневый", "зелёный", "зеленый", "бежевый", "жёлтый", "желтый"]
    color_pattern = "|".join(colors)
    
    # Ищем строку с товаром в таблицах
    for tbl in doc.tables:
        for row in tbl.rows:
            rt = " ".join(c.text.strip() for c in row.cells)
            m = re.search(f"(.*?)\s+({color_pattern})\s+№\s*([A-ZА-ЯЁ0-9]+)\s+VIN\s+([A-Z0-9]{17})", rt, re.IGNORECASE)
            if m:
                data["car_brand"] = m.group(1).strip()
                data["car_color"] = m.group(2).strip()
                data["car_reg"]   = m.group(3).strip()
                data["car_vin"]   = m.group(4).strip()
                break
    return data

# ─────────────────────────────────────────────────────────────
# ПРОЦЕССОР ДКП
# ─────────────────────────────────────────────────────────────

def process_dkp(doc: Document, p: dict) -> Document:
    new_str   = format_amount(p["new_price"])
    new_words = amount_to_words(p["new_price"])
    pv_str    = format_amount(p["pv_amount"])
    pv_words  = amount_to_words(p["pv_amount"])

    pv_found  = False
    target_para = None

    for para in doc.paragraphs:
        full = _para_text(para)
        t    = re.sub(r"\s+", " ", full).strip()
        if not t:
            continue

        if t.startswith("Цена ТС оплачивается Покупателем"):
            target_para = para

        if any(w in t.lower() for w in ["гаранти", "техническая защита", "лимит", "пробег"]):
            continue

        if "первоначальный" in t.lower() or "взнос" in t.lower():
            t2 = RE_AMOUNT.sub(pv_str, t)
            t2 = RE_WORDS.sub(pv_words, t2)
            para.text = ""
            run = para.add_run(t2); run.font.size = Pt(8)
            pv_found = True

        elif any(m in t.lower() for m in ["цена договора", "стоимость тс", "цена тс",
                                            "цену за тс", "стоимость автомобиля", "уплачивает покупатель"]):
            t2 = _remove_nds(t)
            if RE_AMOUNT.search(t2): t2 = RE_AMOUNT.sub(new_str, t2)
            if RE_WORDS.search(t2):  t2 = RE_WORDS.sub(new_words, t2)
            t2 = re.sub(r"\s+", " ", t2).strip()
            if not t2.endswith("."): t2 += "."
            para.text = ""
            run = para.add_run(t2); run.font.size = Pt(8)

    if not pv_found and target_para is not None:
        pv_text = (f"Первоначальный взнос по оплате цены Договора составляет "
                   f"{pv_str} руб ({pv_words}).")
        insert_paragraph_after(target_para, pv_text, source_para=target_para)
    return doc

# ─────────────────────────────────────────────────────────────
# ПРОЦЕССОР ПКО  (v5 — точечные замены, скомпилированные regex)
# ─────────────────────────────────────────────────────────────

def process_pko(doc: Document, p: dict) -> Document:
    pv_str    = format_amount(p["pv_amount"])
    pv_words  = amount_to_words(p["pv_amount"])
    date      = p["date"]
    buyer     = p["buyer_name"]
    
    osnov_full = f"По ДКП №{p['dkp_number']} от {date} за а/м {p['car_brand']} {p['car_color']} № {p['car_reg']} VIN {p['car_vin']}"

    for para in _iter_all_paragraphs(doc):
        t = para.text.strip()
        if not t: continue

        # 1. Замена даты (только если параграф короткий и выглядит как дата)
        if re.match(r"^\d{2}\.\d{2}\.\d{4}$", t):
            _set_para(para, date)
            continue
            
        # 2. Принято от
        if "Принято от" in t:
            # Оставляем маркер, меняем только имя
            new_text = re.sub(r"(Принято от[:\s]*).*?($)", rf"\g<1>{buyer}", t)
            _set_para(para, new_text)
            continue

        # 3. Основание
        if "По ДКП" in t or "за а/м" in t:
            _set_para(para, osnov_full)
            continue

        # 4. Замена суммы (ТОЛЬКО там, где есть слово "Сумма")
        if "Сумма" in t and RE_AMOUNT.search(t):
            # Заменяем только цифры, игнорируя слова
            new_text = RE_AMOUNT.sub(pv_str, t)
            # Если есть пропись суммы, меняем её
            if RE_WORDS.search(new_text):
                new_text = RE_WORDS.sub(pv_words, new_text)
            _set_para(para, new_text)
            continue
            
        # 5. Очистка НДС (точечно)
        if "НДС" in t or "22/122" in t:
            _set_para(para, "") # Просто удаляем строку с НДС, если она встречается отдельно
            
    return doc

# ─────────────────────────────────────────────────────────────
# ПРОЦЕССОР СЧЁТОВ
# ─────────────────────────────────────────────────────────────

def process_invoice(doc: Document, p: dict, amount: float) -> Document:
    amt_str   = format_amount(amount)
    amt_words = amount_to_words(amount)
    date      = p["date"]

    for tbl in doc.tables:
        # Находим колонку "Сумма НДС" по заголовку
        nds_sum_col = -1
        if tbl.rows:
            for ci, hc in enumerate(tbl.rows[0].cells):
                if re.search(r"[Сс]умма\s*НДС|НДС\s*[Сс]умма", hc.text):
                    nds_sum_col = ci; break

        for ri, row in enumerate(tbl.rows):
            for j, cell in enumerate(row.cells):
                ct = cell.text.strip()
                if re.search(r"22/122", ct):
                    for para in cell.paragraphs:
                        nf = re.sub(r"22/122%?", "Без НДС", _para_text(para))
                        _set_para(para, nf)
                elif RE_AMOUNT.search(ct):
                    if j == nds_sum_col and ri > 0:
                        for para in cell.paragraphs:
                            if RE_AMOUNT.search(_para_text(para)):
                                _set_para(para, "0,00")
                    else:
                        for para in cell.paragraphs:
                            nf = RE_AMOUNT.sub(amt_str, _para_text(para))
                            if nf != _para_text(para):
                                _set_para(para, nf)

    for para in doc.paragraphs:
        full = _para_text(para)
        nf   = full
        if RE_DATE.search(full) and "счет" in full.lower():
            nf = RE_DATE.sub(date, nf)
        if RE_WORDS.search(nf):
            nf = RE_WORDS.sub(amt_words, nf)
            nf = re.sub(r",?\s*в\s+т\.?\s*ч\.?\s*НДС[^.]*", ", без НДС", nf)
            nf = re.sub(r"\s{2,}", " ", nf)
        nf = _remove_nds(nf)
        if nf != full:
            _set_para(para, nf)
    return doc

# ─────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="FRESH — Генератор ПВ", page_icon="🚗", layout="wide")
    st.title("🚗 FRESH — Генератор документов с ПВ")
    st.markdown("Автоматическое заполнение ДКП, ПКО и счетов при кредитных сделках")

    if "uploaded_files" not in st.session_state:
        st.session_state.uploaded_files = {}

    def auto_fill():
        if "inv1" in st.session_state.uploaded_files:
            try:
                doc  = Document(BytesIO(st.session_state.uploaded_files["inv1"]))
                data = extract_invoice_data(doc)
                mapping = {
                    "buyer_name": "ФИО",      "dkp_number": "Номер ДКП",
                    "date":       "Дата",     "car_brand":  "Марка",
                    "car_vin":    "VIN",      "car_color":  "Цвет",
                    "car_reg":    "Гос.номер",
                }
                filled = [lbl for k, lbl in mapping.items()
                          if data.get(k) and not st.session_state.get(k)]
                for k in mapping:
                    if data.get(k):
                        st.session_state[k] = data[k]
                if filled:
                    st.success(f"✅ Импортировано: {', '.join(filled)}")
                else:
                    st.warning("Данные не найдены — заполните вручную")
            except Exception as e:
                st.error(f"Ошибка: {e}")

    # ── Боковая панель ───────────────────────────────────────
    with st.sidebar:
        st.header("📁 Шаблоны документов")
        dkp_f  = st.file_uploader("ДКП",                           type=["docx"])
        pko_f  = st.file_uploader("ПКО",                           type=["docx"])
        inv1_f = st.file_uploader("Счёт №1  ⬅️ авто-заполнение",  type=["docx"])
        inv2_f = st.file_uploader("Счёт №2",                       type=["docx"])

        if dkp_f:  st.session_state.uploaded_files["dkp"]  = dkp_f.getvalue()
        if pko_f:  st.session_state.uploaded_files["pko"]  = pko_f.getvalue()
        if inv1_f: st.session_state.uploaded_files["inv1"] = inv1_f.getvalue()
        if inv2_f: st.session_state.uploaded_files["inv2"] = inv2_f.getvalue()

        st.markdown("---")
        st.markdown("**Статус загрузки:**")
        for k, lbl in [("dkp","ДКП"),("pko","ПКО"),("inv1","Счёт №1"),("inv2","Счёт №2")]:
            if k in st.session_state.uploaded_files:
                st.success(f"✅ {lbl}")
            else:
                st.warning(f"⚠️ {lbl} не загружен")

    # ── Основная область ────────────────────────────────────
    col1, col2 = st.columns(2)

    defaults = {
        "buyer_name": "", "dkp_number": "",
        "date": datetime.today().strftime("%d.%m.%Y"),
        "car_brand": "", "car_vin": "",
        "car_color": "", "car_reg": "",
        "real_price": "", "pv_amount": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    with col1:
        st.header("📋 Данные сделки")
        st.text_input("ФИО покупателя *",    key="buyer_name")
        st.text_input("Номер ДКП",            key="dkp_number")
        st.text_input("Дата документов *",    key="date", help="ДД.ММ.ГГГГ")
        st.text_input("Марка / Модель",       key="car_brand")
        st.text_input("VIN",                  key="car_vin")
        st.text_input("Цвет",                 key="car_color")
        st.text_input("Гос. номер",           key="car_reg")
        if "inv1" in st.session_state.uploaded_files:
            st.button("🔄 Авто-заполнить из Счёта №1",
                      on_click=auto_fill, use_container_width=True)

    with col2:
        st.header("💰 Суммы")
        st.text_input("Реальная цена авто (₽) *",      key="real_price", placeholder="3 250 000")
        st.text_input("Сумма искусственного ПВ (₽) *", key="pv_amount",  placeholder="250 000")

        real = parse_amount(st.session_state.real_price)
        pv   = parse_amount(st.session_state.pv_amount)
        if real > 0 and pv > 0:
            total = real + pv
            st.success(f"**Цена по документам (ДКП + Счёт №1)**\n\n{format_amount(total)} ₽")
            st.info(f"**Счёт №2 (к доплате покупателем)**\n\n{format_amount(real)} ₽")
            with st.expander("Суммы прописью"):
                st.write(f"Цена по документам: *{amount_to_words(total)}*")
                st.write(f"ПВ: *{amount_to_words(pv)}*")

    st.markdown("---")

    # ── Валидация и кнопка ──────────────────────────────────
    missing = []
    for k, lbl in [("dkp","ДКП"),("pko","ПКО"),("inv1","Счёт №1"),("inv2","Счёт №2")]:
        if k not in st.session_state.uploaded_files:
            missing.append(f"файл {lbl}")
    if not st.session_state.buyer_name.strip(): missing.append("ФИО покупателя")
    if parse_amount(st.session_state.real_price) <= 0: missing.append("реальная цена")
    if parse_amount(st.session_state.pv_amount)  <= 0: missing.append("сумма ПВ")

    if missing:
        st.warning("⚠️ Не заполнено: " + ", ".join(missing))

    if st.button("✅  Сгенерировать документы", type="primary",
                 use_container_width=True, disabled=bool(missing)):
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
        surname   = (params["buyer_name"].split() or ["Клиент"])[0]
        date_safe = params["date"].replace(".", "-")

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

        bar   = st.progress(0, text="Обрабатываю...")
        files = []
        for i, (key, fname, proc) in enumerate(tasks):
            if key in st.session_state.uploaded_files:
                try:
                    bar.progress(i / len(tasks), text=f"Обрабатываю {fname}…")
                    doc = Document(BytesIO(st.session_state.uploaded_files[key]))
                    doc = proc(doc, params)
                    buf = BytesIO(); doc.save(buf)
                    files.append((fname, buf.getvalue()))
                except Exception as e:
                    st.error(f"❌ {fname}: {e}")
            bar.progress((i + 1) / len(tasks))
        bar.empty()

        if files:
            st.success(f"✅ Готово — {len(files)} документа")
            cols = st.columns(len(files))
            for col, (fname, data) in zip(cols, files):
                with col:
                    st.download_button(
                        label=f"📄 {fname}",
                        data=data, file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True)

if __name__ == "__main__":
    main()
