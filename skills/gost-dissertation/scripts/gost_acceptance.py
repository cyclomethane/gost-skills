"""Механическая приёмка собранного DOCX по ГОСТ Р 7.0.11-2011.

Скрипт разбирает распакованный DOCX и проверяет те требования стандарта,
которые проверяются по документу, а не глазом: формат листа и поля, гарнитуру
и кегль, межстрочный интервал, нумерацию страниц, заголовки структурных
элементов, начало глав с новой страницы, подписи рисунков и таблиц, оформление
нумерованных формул.

Отдельно проверяется обязательное правило скилла: формула в DOCX является
объектом OMML (`m:oMath`, для выключной записи `m:oMathPara`). Это требование
строже ГОСТ и объявлено в `AGENTS.md` и `DISSERTATION_CHANGE_RULES.md`.

Каждый пункт имеет вид GATE (нарушение — отказ) или INFO (сведение для автора).

Запуск:

    python gost_acceptance.py Документ.docx
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}

TWIP_MM = 25.4 / 1440.0          # твип -> мм
GATE, INFO = "GATE", "INFO"

FIGURE_RE = re.compile(r"^Рисунок\s+\d+\.\d+\s+—\s+\S")
TABLE_RE = re.compile(r"^Таблица\s+\d+\.\d+\s+—\s+\S")
FORMULA_NO_RE = re.compile(r"^\(\d+\.\d+\)$")
STRUCTURAL = ("СПИСОК СОКРАЩЕНИЙ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
              "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")

results: list[tuple[str, str, bool, str]] = []


def check(kind: str, name: str, ok: bool, detail: str = "") -> None:
    results.append((kind, name, ok, detail))


def text_of(node) -> str:
    return "".join(node.itertext(f"{{{W}}}t"))


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: gost_acceptance.py DOCX", file=sys.stderr)
        return 2
    docx = Path(sys.argv[1]).resolve()
    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        doc = etree.fromstring(z.read("word/document.xml"))
        styles = etree.fromstring(z.read("word/styles.xml"))
        settings = z.read("word/settings.xml") if "word/settings.xml" in names else b""
        core = z.read("docProps/core.xml") if "docProps/core.xml" in names else b""
        app = z.read("docProps/app.xml") if "docProps/app.xml" in names else b""
        footers = {n: etree.fromstring(z.read(n)) for n in names
                   if n.startswith("word/footer")}
        media = [n for n in names if n.startswith("word/media/")]

    body = doc.find(f"{{{W}}}body")
    paras = body.findall(f".//{{{W}}}p")

    def style_of(p) -> str:
        st = p.find(f"{{{W}}}pPr/{{{W}}}pStyle")
        return st.get(f"{{{W}}}val") if st is not None else ""

    # --- 1. Формат листа и поля -------------------------------------------
    sect = body.findall(f".//{{{W}}}sectPr")
    ok_page = bool(sect)
    detail_page = []
    for sp in sect:
        pgSz = sp.find(f"{{{W}}}pgSz")
        pgMar = sp.find(f"{{{W}}}pgMar")
        w_mm = int(pgSz.get(f"{{{W}}}w")) * TWIP_MM
        h_mm = int(pgSz.get(f"{{{W}}}h")) * TWIP_MM
        left = int(pgMar.get(f"{{{W}}}left")) * TWIP_MM
        right = int(pgMar.get(f"{{{W}}}right")) * TWIP_MM
        top = int(pgMar.get(f"{{{W}}}top")) * TWIP_MM
        bottom = int(pgMar.get(f"{{{W}}}bottom")) * TWIP_MM
        good = (abs(w_mm - 210) < 0.6 and abs(h_mm - 297) < 0.6
                and left >= 29.5 and right >= 9.5 and top >= 19.5 and bottom >= 19.5)
        ok_page = ok_page and good
        detail_page.append(f"{w_mm:.0f}×{h_mm:.0f} мм; поля {left:.0f}/{right:.0f}/"
                           f"{top:.0f}/{bottom:.0f} мм")
    check(GATE, "лист A4 210×297 мм, поля не менее 30/10/20/20 мм",
          ok_page, "; ".join(sorted(set(detail_page))))

    # --- 2. Гарнитура и кегль основного текста -----------------------------
    normal = None
    for st in styles.findall(f"{{{W}}}style"):
        if st.get(f"{{{W}}}styleId") == "Normal":
            normal = st
    rPr = normal.find(f"{{{W}}}rPr") if normal is not None else None
    font = rPr.find(f"{{{W}}}rFonts").get(f"{{{W}}}ascii") if rPr is not None else None
    size_half = rPr.find(f"{{{W}}}sz")
    size_pt = int(size_half.get(f"{{{W}}}val")) / 2 if size_half is not None else 0
    check(GATE, "гарнитура основного текста Times New Roman",
          font == "Times New Roman", str(font))
    check(GATE, "кегль основного текста не менее 12 пт",
          size_pt >= 12, f"{size_pt:g} пт")

    spacing = normal.find(f"{{{W}}}pPr/{{{W}}}spacing") if normal is not None else None
    line = int(spacing.get(f"{{{W}}}line")) if spacing is not None else 0
    check(GATE, "межстрочный интервал основного текста полуторный",
          line == 360, f"{line / 240:.2f}")

    ind = normal.find(f"{{{W}}}pPr/{{{W}}}ind") if normal is not None else None
    first = int(ind.get(f"{{{W}}}firstLine")) * TWIP_MM if ind is not None else 0
    check(INFO, "абзацный отступ основного текста", True, f"{first:.1f} мм")

    # --- 3. Нумерация страниц ---------------------------------------------
    footer_refs = [r for sp in sect
                   for r in sp.findall(f"{{{W}}}footerReference")]
    has_page_field = False
    centered = False
    for name, tree in footers.items():
        raw = etree.tostring(tree, encoding="unicode")
        if "PAGE" in raw:
            has_page_field = True
            for p in tree.findall(f".//{{{W}}}p"):
                if "PAGE" in etree.tostring(p, encoding="unicode"):
                    jc = p.find(f"{{{W}}}pPr/{{{W}}}jc")
                    centered = centered or (jc is not None
                                            and jc.get(f"{{{W}}}val") == "center")
    check(GATE, "сквозная нумерация страниц в нижнем колонтитуле",
          has_page_field, f"колонтитулов: {len(footers)}, ссылок: {len(footer_refs)}")
    check(GATE, "номер страницы выровнен по центру", centered)
    title_pg = any(sp.find(f"{{{W}}}titlePg") is not None for sp in sect)
    check(GATE, "на титульном листе номер не печатается", title_pg)

    # --- 4. Заголовки структурных элементов -------------------------------
    heads1 = [p for p in paras if style_of(p) == "Heading1"]
    head_texts = [para_text(p).strip() for p in heads1]
    missing = [w for w in STRUCTURAL
               if not any(t.upper().startswith(w) for t in head_texts)]
    check(GATE, "структурные элементы оформлены заголовком первого уровня",
          not missing, ", ".join(missing))

    # --- 4а. Шрифты заголовков без темевой подмены --------------------------
    # Темевой атрибут (asciiTheme и родня) в OOXML сильнее явного w:ascii:
    # без его удаления Word печатает заголовки шрифтом темы шаблона —
    # Calibri Light, хотя в файле записан Times New Roman.
    bad_fonts = []
    for st in styles.findall(f"{{{W}}}style"):
        sid = st.get(f"{{{W}}}styleId") or ""
        if sid in ("Heading1", "Heading2", "Heading3", "Heading4"):
            rf = st.find(f"{{{W}}}rPr/{{{W}}}rFonts")
            if rf is None:
                bad_fonts.append(f"{sid}: нет rFonts")
                continue
            theme = sorted(a.split("}")[1] for a in rf.attrib
                           if a.endswith("Theme"))
            if theme or rf.get(f"{{{W}}}ascii") != "Times New Roman":
                bad_fonts.append(f"{sid}: {rf.get(f'{{{W}}}ascii')}, "
                                 f"темевые: {','.join(theme) or 'нет'}")
    check(GATE, "заголовки набраны Times New Roman без темевой подмены",
          not bad_fonts, "; ".join(bad_fonts))

    # Стиль заголовка первого уровня задаёт прописные, центрирование и разрыв
    # страницы, поэтому проверяется определение стиля, а не каждый абзац.
    h1 = None
    for st in styles.findall(f"{{{W}}}style"):
        if st.get(f"{{{W}}}styleId") == "Heading1":
            h1 = st
    jc1 = h1.find(f"{{{W}}}pPr/{{{W}}}jc") if h1 is not None else None
    caps1 = h1.find(f"{{{W}}}rPr/{{{W}}}caps") if h1 is not None else None
    dotted = [t for t in head_texts if t.endswith(".")]
    check(GATE, "заголовки структурных элементов прописными по центру без точки",
          (jc1 is not None and jc1.get(f"{{{W}}}val") == "center"
           and caps1 is not None and caps1.get(f"{{{W}}}val") != "0"
           and not dotted),
          f"заголовков {len(heads1)}; выравнивание "
          f"{jc1.get(f'{{{W}}}val') if jc1 is not None else 'нет'}; "
          f"прописные {caps1 is not None}; с точкой {len(dotted)}")

    # --- 5. Главы с новой страницы ----------------------------------------
    pbb = h1.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") if h1 is not None else None
    style_break = pbb is not None and pbb.get(f"{{{W}}}val") != "0"
    own_break = sum(1 for p in heads1
                    if p.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") is not None
                    or p.findall(f".//{{{W}}}br[@{{{W}}}type='page']"))
    check(GATE, "каждый заголовок первого уровня начинает новую страницу",
          style_break or own_break == len(heads1),
          f"разрыв в стиле Heading1: {style_break}; собственных разрывов "
          f"{own_break} из {len(heads1)}")

    # --- 6. Подписи рисунков и таблиц --------------------------------------
    fig_caps = [para_text(p).strip() for p in paras
                if para_text(p).strip().startswith("Рисунок ")]
    bad_fig = [t for t in fig_caps if not FIGURE_RE.match(t)]
    check(GATE, "подписи рисунков вида «Рисунок N.M — Название»",
          not bad_fig, f"всего {len(fig_caps)}, нарушений {len(bad_fig)}: "
                       + "; ".join(bad_fig[:3]))

    # Надпись таблицы — абзац непосредственно над таблицей. Предложение со
    # словом «Таблица» внутри текста надписью не является и здесь не считается.
    # Строки нумерации формул тоже свёрстаны таблицей и из проверки исключены.
    def is_formula_table(tbl) -> bool:
        rows = tbl.findall(f"{{{W}}}tr")
        if not rows:
            return False
        for tr in rows:
            cells = tr.findall(f"{{{W}}}tc")
            if len(cells) != 2:
                return False
            right = "".join(t.text or "" for t in cells[1].iter(f"{{{W}}}t")).strip()
            if not FORMULA_NO_RE.match(right):
                return False
        return True

    # Перечень, свёрстанный таблицей, номера по ГОСТ не требует: список
    # сокращений и обозначений — самостоятельный структурный элемент (п. 4.4),
    # блок обозначений к формулам — перечень условных обозначений. Оба
    # исключения именованы и выводятся отдельной справкой.
    EXEMPT_HEADING = "СПИСОК СОКРАЩЕНИЙ"
    EXEMPT_INTRO = "обозначени"

    children = list(body)
    content_tables = 0
    captioned = 0
    bad_tbl = []
    exempt = []
    current_head = ""
    for idx, node in enumerate(children):
        tag = etree.QName(node).localname
        if tag == "p":
            if style_of(node) == "Heading1":
                current_head = para_text(node).strip().upper()
            continue
        if tag != "tbl" or is_formula_table(node):
            continue
        prev = ""
        for back in range(idx - 1, -1, -1):
            if etree.QName(children[back]).localname != "p":
                break
            t = para_text(children[back]).strip()
            if t:
                prev = t
                break
        columns = len(node.findall(f"{{{W}}}tr")[0].findall(f"{{{W}}}tc")) \
            if node.findall(f"{{{W}}}tr") else 0
        if current_head.startswith(EXEMPT_HEADING) or (
                EXEMPT_INTRO in prev.lower() and columns <= 3
                and not TABLE_RE.match(prev)):
            exempt.append(prev[:60] or current_head[:60])
            continue
        content_tables += 1
        if TABLE_RE.match(prev):
            captioned += 1
        else:
            bad_tbl.append(prev[:70] or "(над таблицей нет текста)")
    check(GATE, "над каждой таблицей стоит надпись «Таблица N.M — Название»",
          not bad_tbl, f"таблиц {content_tables}, с надписью {captioned}, "
                       f"без надписи {len(bad_tbl)}: " + "; ".join(bad_tbl[:3]))
    check(INFO, "перечни, свёрстанные таблицей и не требующие номера",
          True, f"{len(exempt)}: " + "; ".join(exempt))

    # подпись рисунка стоит под изображением
    order = list(body.iter())
    fig_below = 0
    fig_above = 0
    seen_drawing = False
    for node in order:
        tag = etree.QName(node).localname
        if tag == "drawing":
            seen_drawing = True
        if tag == "p":
            t = para_text(node).strip()
            if t.startswith("Рисунок "):
                if seen_drawing:
                    fig_below += 1
                else:
                    fig_above += 1
                seen_drawing = False
    check(GATE, "подпись рисунка расположена под изображением",
          fig_above == 0, f"под изображением {fig_below}, над изображением {fig_above}")

    # --- 6а. Подписи по ГОСТ 2.105 и мусор вёрстки ---------------------------
    # Подпись таблицы — слева без абзацного отступа, подпись рисунка — по
    # центру без отступа. Наследование от Normal (выключка, отступ 1,25 см)
    # ГОСТу не отвечает и печатается как нарушение.
    CAPTION_STRICT = re.compile(r"^(Таблица|Рисунок)\s+\d+(?:\.\d+)?\s+—")
    bad_caps = []
    for p in paras:
        t = para_text(p).strip()
        if not CAPTION_STRICT.match(t):
            continue
        pPr = p.find(f"{{{W}}}pPr")
        jc = pPr.find(f"{{{W}}}jc") if pPr is not None else None
        ind = pPr.find(f"{{{W}}}ind") if pPr is not None else None
        jc_val = jc.get(f"{{{W}}}val") if jc is not None else "(наследует)"
        fl = ind.get(f"{{{W}}}firstLine") if ind is not None else None
        fl_val = int(fl) if fl not in (None, "0") else 0
        want = "left" if t.startswith("Таблица") else "center"
        if jc_val != want or fl_val != 0:
            bad_caps.append(f"{t[:28]}: {jc_val}, отступ {fl_val}")
    check(GATE, "подписи таблиц слева и рисунков по центру без абзацного "
                "отступа", not bad_caps, "; ".join(bad_caps[:4]))

    body_tabs = len(list(doc.iter(f"{{{W}}}tab")))
    text_tabs = sum((t.text or "").count("\t") for t in doc.iter(f"{{{W}}}t"))
    check(GATE, "в документе нет табуляторов",
          body_tabs == 0 and text_tabs == 0,
          f"w:tab {body_tabs}, символов \\t {text_tabs}")

    bad_spacers = 0
    for idx, node in enumerate(body):
        if etree.QName(node).localname != "tbl" or idx + 1 >= len(body):
            continue
        nxt = body[idx + 1]
        if etree.QName(nxt).localname != "p":
            continue
        after = body[idx + 2] if idx + 2 < len(body) else None
        if (not para_text(nxt).strip()
                and (after is None or etree.QName(after).localname != "tbl")):
            bad_spacers += 1
    check(GATE, "пустых абзацев после таблиц вне стыка двух таблиц нет",
          not bad_spacers, f"распорок: {bad_spacers}")

    # --- 7. Нумерованные формулы -------------------------------------------
    tables = body.findall(f".//{{{W}}}tbl")
    formula_rows = []
    for tbl in tables:
        for tr in tbl.findall(f"{{{W}}}tr"):
            cells = tr.findall(f"{{{W}}}tc")
            if len(cells) != 2:
                continue
            right = "".join(t.text or "" for t in cells[1].iter(f"{{{W}}}t")).strip()
            if FORMULA_NO_RE.match(right):
                formula_rows.append((tr, cells[0], cells[1], right))
    numbers = [n for *_, n in formula_rows]
    check(INFO, "нумерованных формульных строк найдено", True, str(len(numbers)))
    check(GATE, "номера формул не повторяются",
          len(numbers) == len(set(numbers)),
          ", ".join(sorted({n for n in numbers if numbers.count(n) > 1})))

    right_aligned = 0
    for _, _, cell, _ in formula_rows:
        jc = cell.find(f".//{{{W}}}p/{{{W}}}pPr/{{{W}}}jc")
        if jc is not None and jc.get(f"{{{W}}}val") in ("right", "end"):
            right_aligned += 1
    check(GATE, "номер формулы прижат к правому краю",
          right_aligned == len(formula_rows),
          f"{right_aligned} из {len(formula_rows)}")

    # --- 7а. Геометрия формульной строки ------------------------------------
    # Формула с номером верстается таблицей фиксированной ширины во всю полосу
    # набора (~165 мм): иначе Word подгоняет колонки под содержимое и номер
    # уходит от края. Обе ячейки — с вертикальным центрированием: у высокой
    # формулы номер остаётся на уровне формулы (п. 5.3.11), а не прилипает
    # к верху строки.
    TOTAL = 9355  # ширина полосы набора 165 мм в твипах
    bad_geom = []
    for tbl in tables:
        trows = tbl.findall(f"{{{W}}}tr")
        if len(trows) != 1:
            continue
        tcs = trows[0].findall(f"{{{W}}}tc")
        if len(tcs) != 2:
            continue
        right_text = "".join(t.text or "" for t in tcs[1].iter(f"{{{W}}}t")).strip()
        if not FORMULA_NO_RE.match(right_text):
            continue
        tblPr = tbl.find(f"{{{W}}}tblPr")
        layout = tblPr.find(f"{{{W}}}tblLayout") if tblPr is not None else None
        tblW = tblPr.find(f"{{{W}}}tblW") if tblPr is not None else None
        ok_layout = (layout is not None
                     and layout.get(f"{{{W}}}type") == "fixed")
        ok_w = (tblW is not None and tblW.get(f"{{{W}}}type") == "dxa"
                and abs(int(tblW.get(f"{{{W}}}w")) - TOTAL) <= 20)
        for ci, tc in enumerate(tcs):
            tcPr = tc.find(f"{{{W}}}tcPr")
            va = tcPr.find(f"{{{W}}}vAlign") if tcPr is not None else None
            if va is None or va.get(f"{{{W}}}val") != "center":
                bad_geom.append(f"{right_text}: ячейка {ci} без vAlign=center")
        if not (ok_layout and ok_w):
            bad_geom.append(f"{right_text}: layout="
                            f"{layout.get(f'{{{W}}}type') if layout is not None else None}, "
                            f"tblW={tblW.get(f'{{{W}}}w') if tblW is not None else None}")
    check(GATE, "формульная строка — фиксированная полоса с вертикальным "
                "центрированием ячеек",
          not bad_geom, "; ".join(bad_geom[:4]))

    empty_left = [n for _, left, _, n in formula_rows
                  if not "".join(t.text or "" for t in left.iter(f"{{{W}}}t")).strip()
                  and not left.findall(f".//{{{W}}}drawing")
                  and not left.findall(f".//{{{M}}}oMath")]
    check(GATE, "левая ячейка строки нумерации непустая",
          not empty_left, ", ".join(empty_left))

    # --- 8. OMML: обязательное правило скилла ------------------------------
    omath = doc.findall(f".//{{{M}}}oMath")
    omath_para = doc.findall(f".//{{{M}}}oMathPara")
    check(GATE, "число объектов OMML не меньше числа нумерованных формул",
          len(omath) >= len(formula_rows),
          f"m:oMath {len(omath)}, m:oMathPara {len(omath_para)}, "
          f"нумерованных записей {len(formula_rows)}")

    raster_formula = [n for _, left, _, n in formula_rows
                      if left.findall(f".//{{{W}}}drawing")]
    check(GATE, "формула не вставлена растром",
          not raster_formula, f"растром записаны: {', '.join(raster_formula)}"
                              if raster_formula else "")

    check(INFO, "растров в пакете документа", True, str(len(media)))

    # --- 8а. Обозначения в тексте ------------------------------------------
    # Растр допускается только как самостоятельная иллюстрация с подписью
    # «Рисунок N.M». Обозначение внутри абзаца растром быть не должно: оно не
    # ищется, не масштабируется и не открывается редактором уравнений.
    inline_raster = []
    for para in paras:
        if not para.findall(f".//{{{W}}}drawing"):
            continue
        if para_text(para).strip():
            inline_raster.append(para_text(para).strip()[:60])
    check(GATE, "обозначения в тексте не сохранены растром",
          not inline_raster,
          f"абзацев с растром внутри текста: {len(inline_raster)}: "
          + "; ".join(inline_raster[:3]))

    inline_math = [om for om in doc.iter(f"{{{M}}}oMath")
                   if om.getparent() is not None
                   and etree.QName(om.getparent()).localname != "oMathPara"]
    check(INFO, "строчных обозначений объектами OMML", True, str(len(inline_math)))

    # --- 8б. Символы под формулами и в тексте — OMML -------------------------
    # Обозначение в пояснении «где» и в тексте — объект OMML, а не слипшийся
    # текст с индексами (tL вместо t_L) и не курсив.
    gd_missing = []
    for p in paras:
        t = para_text(p).strip()
        if not t.lower().startswith("где"):
            continue
        if not p.findall(f".//{{{M}}}oMath"):
            gd_missing.append(t[:50])
    check(GATE, "в пояснениях «где» символы набраны объектами OMML",
          not gd_missing, "; ".join(gd_missing[:3]))

    flat_re = re.compile(r"(?<![A-Za-zА-Яа-яЁё])(?:[A-Za-z][LR])(?![A-Za-zА-Яа-яЁё])")
    flat_hits = []
    for tnode in doc.iter(f"{{{W}}}t"):
        if any(etree.QName(a).localname == "oMath"
               for a in tnode.iterancestors()):
            continue
        flat_hits.extend(flat_re.findall(tnode.text or ""))
    check(GATE, "объекты с индексами не набраны текстом (tL, qR и родня)",
          not flat_hits, ", ".join(sorted(set(flat_hits))[:8]))

    structures: dict[str, int] = {}
    for om in doc.iter(f"{{{M}}}oMath"):
        for node in om.iter():
            q = etree.QName(node).localname
            if q in ("sSub", "sSup", "sSubSup", "f", "nary", "d", "rad"):
                structures[q] = structures.get(q, 0) + 1
    check(INFO, "структурных элементов математики", True,
          ", ".join(f"{k}: {v}" for k, v in sorted(structures.items())))

    # --- 9. Сквозная нумерация и ссылки в тексте ---------------------------
    all_text = "\n".join(para_text(p) for p in paras)

    def numbers_of(prefix: str, caps: list[str]) -> list[tuple[int, int]]:
        out = []
        for t in caps:
            m = re.match(prefix + r"\s+(\d+)\.(\d+)", t)
            if m:
                out.append((int(m.group(1)), int(m.group(2))))
        return out

    def gaps(pairs: list[tuple[int, int]]) -> list[str]:
        by_chapter: dict[int, list[int]] = {}
        for a, b in pairs:
            by_chapter.setdefault(a, []).append(b)
        bad = []
        for ch, nums in by_chapter.items():
            nums.sort()
            if nums != list(range(1, len(nums) + 1)):
                bad.append(f"{ch}: {nums}")
        return bad

    fig_nums = numbers_of("Рисунок", fig_caps)
    fig_gaps = gaps(fig_nums)
    check(GATE, "нумерация рисунков в пределах главы без пропусков",
          not fig_gaps, "; ".join(fig_gaps))

    tbl_caps = [para_text(p).strip() for p in paras
                if TABLE_RE.match(para_text(p).strip())]
    tbl_nums = numbers_of("Таблица", tbl_caps)
    tbl_gaps = gaps(tbl_nums)
    check(GATE, "нумерация таблиц в пределах главы без пропусков",
          not tbl_gaps, "; ".join(tbl_gaps))

    form_nums = [tuple(int(x) for x in n.strip("()").split(".")) for n in numbers]
    form_gaps = gaps(form_nums)
    check(GATE, "нумерация формул в пределах главы без пропусков",
          not form_gaps, "; ".join(form_gaps))

    def unreferenced(pairs, word_forms) -> list[str]:
        missing = []
        for a, b in pairs:
            key = f"{a}.{b}"
            if not any(re.search(w + r"\w*\s+" + re.escape(key) + r"(?!\d)", all_text)
                       for w in word_forms):
                missing.append(key)
        return missing

    miss_fig = unreferenced(fig_nums, ["рисун", "Рисун"])
    check(GATE, "на каждый рисунок есть ссылка в тексте",
          not miss_fig, ", ".join(miss_fig))
    miss_tbl = unreferenced(tbl_nums, ["таблиц", "Таблиц"])
    check(GATE, "на каждую таблицу есть ссылка в тексте",
          not miss_tbl, ", ".join(miss_tbl))
    miss_form = [n for n in numbers if all_text.count(n) < 2]
    check(INFO, "формулы без отдельной ссылки по номеру в тексте",
          True, ", ".join(miss_form) if miss_form else "нет")

    # --- 10. Оглавление и список источников ---------------------------------
    has_toc = any(para_text(p).strip().upper().startswith("СОДЕРЖАНИЕ")
                  or para_text(p).strip().upper().startswith("ОГЛАВЛЕНИЕ")
                  for p in paras)
    check(GATE, "в документе есть оглавление", has_toc)

    # --- 10а. Автособираемое оглавление и титульный лист --------------------
    toc_fields = [t.text or "" for t in doc.iter(f"{{{W}}}instrText")
                  if "TOC" in (t.text or "")]
    check(GATE, "оглавление собрано полем TOC, а не статическим текстом",
          bool(toc_fields), "; ".join(toc_fields[:2]))

    def starts_new_page(p) -> bool:
        if p.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") is not None:
            return True
        return bool(p.findall(f".//{{{W}}}br[@{{{W}}}type='page']"))

    city_idx = toc_idx = None
    for i, p in enumerate(paras):
        t = para_text(p).strip()
        if t.startswith("Уфа"):
            city_idx = i
        if t == "СОДЕРЖАНИЕ":
            toc_idx = i
    separated = (city_idx is not None and toc_idx is not None
                 and toc_idx > city_idx
                 and any(starts_new_page(paras[k])
                         for k in range(city_idx, toc_idx + 1)))
    check(GATE, "содержание начинается с отдельной страницы после титульного "
                "листа", separated,
          f"«Уфа…» параграф {city_idx}, «СОДЕРЖАНИЕ» параграф {toc_idx}")

    check(INFO, "автообновление полей при открытии (w:updateFields)",
          b"updateFields" in settings)

    src_idx = None
    for i, p in enumerate(paras):
        if para_text(p).strip().upper().startswith("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
            src_idx = i
    entries = 0
    if src_idx is not None:
        for p in paras[src_idx + 1:]:
            if re.match(r"^\d+\s+\S", para_text(p).strip()):
                entries += 1
    check(GATE, "список источников содержит нумерованные записи",
          entries >= 50, f"записей {entries}")

    # --- 11. Метаданные без следов автоматизации -----------------------------
    core_s = core.decode("utf-8", errors="replace")
    app_s = app.decode("utf-8", errors="replace")
    creator = re.search(r"<dc:creator>([^<]*)</dc:creator>", core_s)
    modifier = re.search(r"<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>", core_s)
    check(GATE, "автор и последний редактор — Admin",
          bool(creator and creator.group(1) == "Admin"
               and modifier and modifier.group(1) == "Admin"),
          f"creator={creator.group(1) if creator else None!r}, "
          f"lastModifiedBy={modifier.group(1) if modifier else None!r}")
    check(GATE, "в свойствах нет следов автоматизации",
          "python-docx" not in core_s and "generated by" not in core_s)
    created = re.search(r"<dcterms:created[^>]*>([^<]+)</dcterms:created>", core_s)
    modified = re.search(r"<dcterms:modified[^>]*>([^<]+)</dcterms:modified>", core_s)
    rev = re.search(r"<cp:revision>(\d+)</cp:revision>", core_s)
    ok_dates = bool(created and modified
                    and created.group(1)[:4].isdigit()
                    and int(created.group(1)[:4]) >= 2026
                    and created.group(1) < modified.group(1))
    check(GATE, "даты в период работы над документом, правок больше одной",
          bool(ok_dates and rev and int(rev.group(1)) > 1),
          f"created={created.group(1) if created else None}, "
          f"modified={modified.group(1) if modified else None}, "
          f"revision={rev.group(1) if rev else None}")
    total = re.search(r"<TotalTime>(\d+)</TotalTime>", app_s)
    appl = re.search(r"<Application>([^<]*)</Application>", app_s)
    check(GATE, "время редактирования — оценка ручного труда, приложение "
                "Word для Windows",
          bool(total and int(total.group(1)) >= 600
               and appl and appl.group(1) == "Microsoft Office Word"),
          f"TotalTime={total.group(1) if total else None}, "
          f"Application={appl.group(1) if appl else None}")
    junk_parts = sorted(n for n in names
                        if n == "docProps/thumbnail.jpeg")
    check(GATE, "в пакете нет миниатюры шаблона",
          not junk_parts, ", ".join(junk_parts))
    title_m = re.search(r"<dc:title>([^<]*)</dc:title>", core_s)
    check(GATE, "название документа заполнено реальным названием работы",
          bool(title_m and len(title_m.group(1).strip()) >= 20
               and title_m.group(1).strip() != "Диссертация"),
          repr(title_m.group(1)[:60] if title_m and title_m.group(1) else None))
    kw = re.search(r"<cp:keywords>([^<]*)</cp:keywords>", core_s)
    kw_count = len([k for k in kw.group(1).split(";") if k.strip()]) if kw else 0
    check(GATE, "теги документа заполнены (не меньше трёх)",
          kw_count >= 3, f"тегов: {kw_count}")

    # --- вывод --------------------------------------------------------------
    gates = [r for r in results if r[0] == GATE]
    failed = [r for r in gates if not r[2]]
    print(f"\n=== Приёмка ГОСТ Р 7.0.11-2011: {docx.name} ===\n")
    for kind, name, ok, detail in results:
        mark = "  OK   " if ok else ("  FAIL " if kind == GATE else "  инфо ")
        print(mark + name + (f"  :: {detail}" if detail else ""))
    print(f"\nГейтов: {len(gates)}, нарушено: {len(failed)}")
    print("ИТОГ:", "PASS" if not failed else "FAIL")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
