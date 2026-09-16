"""Механическая приёмка DOCX по ГОСТ: обязательные элементы, объём, вёрстка.

Скрипт разбирает пакет DOCX и проверяет то, что проверяется по файлу, а не
глазом. Отчёт разложен на те же разделы, что и отчёт скилла:

  1. ОБЯЗАТЕЛЬНЫЕ ЭЛЕМЕНТЫ — структурные разделы, элементы введения, объём;
  2. ВЁРСТКА — страница, текст, заголовки, нумерация, подписи, оглавление;
  3. ФОРМУЛЫ — OMML, нумерация, геометрия формульной строки;
  4. БИБЛИОГРАФИЯ И ССЫЛКИ — записи списка, ссылки на рисунки и таблицы;
  5. МЕТАДАННЫЕ — свойства пакета.

Профили оформления (`--profile`):

  A — ГОСТ Р 7.0.11-2011: поля 25/10/20/20 мм, номер страницы в верхнем поле;
  B — ГОСТ 7.32-2017: поля 30/15/20/20 мм, номер в центре нижней части листа.

Значения профиля A, совпадающие с профилем B (поля 30/15, номер внизу),
печатаются как DEVIATION — сознательное расхождение, а не отказ; см.
`references/layout-core.md`, раздел 15.

Виды документа (`--type`) задают набор обязательных структурных элементов и
элементов введения: dissertation, vkr, nir, generic.

Правило скилла, более строгое, чем ГОСТ: вся математика — объекты OMML
(`m:oMath`, для выключной записи `m:oMathPara`).

Каждый пункт имеет вид GATE (нарушение — отказ), INFO (сведение для автора)
или DEVIATION (учтённое расхождение).

Запуск:

    python gost_acceptance.py Документ.docx --profile B --type vkr
    python gost_acceptance.py Документ.docx --min-pages 60 --max-pages 90
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

# Консоль Windows по умолчанию открывает stdout в cp1251/cp866: символы вроде
# «≈» в отчёте валят печать UnicodeEncodeError. Отчёт всегда должен напечататься.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

TWIP_MM = 25.4 / 1440.0          # твип -> мм
CHARS_PER_PAGE = 1800            # страница А4, 14 пт, интервал 1,5

GATE, INFO, DEVIATION = "GATE", "INFO", "DEVIATION"

S_MUST, S_LAYOUT, S_MATH, S_BIB, S_META = (
    "ОБЯЗАТЕЛЬНЫЕ ЭЛЕМЕНТЫ И ОБЪЁМ",
    "ВЁРСТКА",
    "ФОРМУЛЫ",
    "БИБЛИОГРАФИЯ И ССЫЛКИ",
    "МЕТАДАННЫЕ ПАКЕТА",
)
SECTIONS = (S_MUST, S_LAYOUT, S_MATH, S_BIB, S_META)

# Подписи: «Рисунок 1 — …», «Рисунок 1.2 — …», «Рисунок А.1 — …»
NUM = r"(?:[А-ЯA-Z]|\d+)(?:\.\d+)*"
FIGURE_RE = re.compile(rf"^Рисунок\s+({NUM})\s+—\s+\S")
TABLE_RE = re.compile(rf"^Таблица\s+({NUM})\s+—\s+\S")
FORMULA_NO_RE = re.compile(rf"^\(({NUM})\)$")
CAPTION_STRICT = re.compile(rf"^(Таблица|Рисунок)\s+{NUM}\s+—")
CITY_YEAR_RE = re.compile(r"^[А-ЯЁ][А-Яа-яЁё\-\. ]{2,30}[\s,–—-]+(?:19|20)\d\d\.?$")
GENERATOR_RE = re.compile(r"python|docx|generated|script|bot|openai|claude|gpt",
                          re.IGNORECASE)

# Обязательные структурные элементы по виду документа: (метка, варианты заголовка)
MUST_HEADINGS = {
    "dissertation": [
        ("оглавление", ("ОГЛАВЛЕНИЕ", "СОДЕРЖАНИЕ")),
        ("введение", ("ВВЕДЕНИЕ",)),
        ("заключение", ("ЗАКЛЮЧЕНИЕ",)),
        ("список литературы", ("СПИСОК ЛИТЕРАТУРЫ",
                               "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                               "СПИСОК ИСТОЧНИКОВ")),
    ],
    "vkr": [
        ("содержание", ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ")),
        ("введение", ("ВВЕДЕНИЕ",)),
        ("заключение", ("ЗАКЛЮЧЕНИЕ",)),
        ("список источников", ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                               "СПИСОК ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ")),
    ],
    "nir": [
        ("реферат", ("РЕФЕРАТ",)),
        ("содержание", ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ")),
        ("введение", ("ВВЕДЕНИЕ",)),
        ("заключение", ("ЗАКЛЮЧЕНИЕ",)),
        ("список источников", ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                               "СПИСОК ИСТОЧНИКОВ")),
    ],
    "generic": [
        ("содержание", ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ")),
        ("введение", ("ВВЕДЕНИЕ",)),
        ("заключение", ("ЗАКЛЮЧЕНИЕ",)),
    ],
}

OPTIONAL_HEADINGS = (
    ("перечень сокращений", ("СПИСОК СОКРАЩЕНИЙ", "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ",
                             "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ")),
    ("словарь терминов", ("СЛОВАРЬ ТЕРМИНОВ", "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ")),
    ("список иллюстративного материала", ("СПИСОК ИЛЛЮСТРАТИВНОГО МАТЕРИАЛА",)),
    ("приложения", ("ПРИЛОЖЕНИЕ",)),
)

# Элементы введения: (метка, регулярное выражение по тексту введения)
INTRO_ELEMENTS = [
    ("актуальность", r"актуальност"),
    ("степень разработанности", r"разработанност|изученност"),
    ("цель и задачи", r"цель\w*\s+(?:работы|исследования|ВКР)|задачи\s+(?:работы|исследования)|поставлен\w*\s+задач"),
    ("научная новизна", r"новизн"),
    ("теоретическая и практическая значимость", r"значимост|практическая ценность"),
    ("методология и методы", r"методолог|методы\s+исследования"),
    ("положения, выносимые на защиту", r"выносим\w*\s+на\s+защиту|положения,\s*представляем|результаты,\s*выносим"),
    ("достоверность и апробация", r"достоверност|апробац"),
]
INTRO_VKR_EXTRA = [
    ("объект исследования", r"объект\w*\s+исследования"),
    ("предмет исследования", r"предмет\w*\s+исследования"),
]

PROFILES = {
    "A": {
        "name": "ГОСТ Р 7.0.11-2011",
        "margins": (25, 10, 20, 20),
        "font_min": 12, "font_max": 14,
        "page_number": "header",
        "caps_headings": False,
    },
    "B": {
        "name": "ГОСТ 7.32-2017",
        "margins": (30, 15, 20, 20),
        "font_min": 12, "font_max": None,
        "page_number": "footer",
        "caps_headings": True,
    },
}

results: list[tuple[str, str, str, bool, str]] = []


def check(section: str, kind: str, name: str, ok: bool, detail: str = "") -> None:
    results.append((section, kind, name, ok, detail))


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))


def parse_args(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Механическая приёмка DOCX по ГОСТ.")
    ap.add_argument("docx", help="файл документа")
    ap.add_argument("--profile", choices=("A", "B"), default="B",
                    help="профиль оформления: A — ГОСТ Р 7.0.11, B — ГОСТ 7.32 "
                         "(по умолчанию B)")
    ap.add_argument("--type", dest="doctype", default="generic",
                    choices=tuple(MUST_HEADINGS),
                    help="вид документа для набора обязательных элементов")
    ap.add_argument("--min-pages", type=int, default=None,
                    help="нижняя граница объёма основного текста в страницах")
    ap.add_argument("--max-pages", type=int, default=None,
                    help="верхняя граница объёма основного текста в страницах")
    ap.add_argument("--min-sources", type=int, default=None,
                    help="минимальное число записей в списке источников")
    ap.add_argument("--font", default="Times New Roman",
                    help="требуемая гарнитура основного текста")
    ap.add_argument("--author", default=None,
                    help="требуемое значение dc:creator и cp:lastModifiedBy")
    ap.add_argument("--strict-metadata", action="store_true",
                    help="проверять даты, число правок и TotalTime как GATE")
    return ap.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    prof = PROFILES[args.profile]
    alt = PROFILES["B" if args.profile == "A" else "A"]
    docx = Path(args.docx).resolve()

    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        doc = etree.fromstring(z.read("word/document.xml"))
        styles = etree.fromstring(z.read("word/styles.xml"))
        settings = z.read("word/settings.xml") if "word/settings.xml" in names else b""
        core = z.read("docProps/core.xml") if "docProps/core.xml" in names else b""
        app = z.read("docProps/app.xml") if "docProps/app.xml" in names else b""
        footers = {n: etree.fromstring(z.read(n)) for n in names
                   if n.startswith("word/footer")}
        headers = {n: etree.fromstring(z.read(n)) for n in names
                   if n.startswith("word/header")}
        media = [n for n in names if n.startswith("word/media/")]

    body = doc.find(f"{{{W}}}body")
    paras = body.findall(f".//{{{W}}}p")

    def style_of(p) -> str:
        st = p.find(f"{{{W}}}pPr/{{{W}}}pStyle")
        return st.get(f"{{{W}}}val") if st is not None else ""

    heads1 = [p for p in paras if style_of(p) == "Heading1"]
    heads_all = [p for p in paras
                 if style_of(p) in ("Heading1", "Heading2", "Heading3", "Heading4")]
    head_texts = [para_text(p).strip() for p in heads1]
    head_upper = [t.upper() for t in head_texts]
    all_text = "\n".join(para_text(p) for p in paras)

    # =====================================================================
    # 1. ОБЯЗАТЕЛЬНЫЕ ЭЛЕМЕНТЫ И ОБЪЁМ
    # =====================================================================
    def has_heading(variants: tuple[str, ...]) -> bool:
        return any(any(h.startswith(v) for v in variants) for h in head_upper)

    missing = []
    for label, variants in MUST_HEADINGS[args.doctype]:
        ok = has_heading(variants)
        if not ok:
            missing.append(label)
        check(S_MUST, GATE, f"структурный элемент «{label}»", ok,
              "" if ok else "заголовок первого уровня не найден: "
                            + " / ".join(variants))

    present_opt = [label for label, variants in OPTIONAL_HEADINGS
                   if has_heading(variants)]
    check(S_MUST, INFO, "факультативные элементы в документе", True,
          ", ".join(present_opt) if present_opt else "нет")

    # Элементы введения ищутся в тексте от «ВВЕДЕНИЕ» до следующего заголовка.
    intro_text = ""
    collecting = False
    for p in paras:
        t = para_text(p).strip()
        if style_of(p) == "Heading1":
            if t.upper().startswith("ВВЕДЕНИЕ"):
                collecting = True
                continue
            if collecting:
                break
        if collecting:
            intro_text += " " + t
    intro_text = intro_text.lower()

    wanted = list(INTRO_ELEMENTS)
    if args.doctype in ("vkr", "nir"):
        wanted += INTRO_VKR_EXTRA
    intro_kind = GATE if args.doctype == "dissertation" else INFO
    if not intro_text:
        check(S_MUST, intro_kind, "введение найдено для разбора", False,
              "раздел «ВВЕДЕНИЕ» не найден или пуст")
    else:
        for label, pattern in wanted:
            found = bool(re.search(pattern, intro_text))
            check(S_MUST, intro_kind, f"введение: {label}", found,
                  "" if found else "упоминание не найдено")

    # Объём и насыщенность
    fig_caps = [para_text(p).strip() for p in paras
                if FIGURE_RE.match(para_text(p).strip())]
    tbl_caps = [para_text(p).strip() for p in paras
                if TABLE_RE.match(para_text(p).strip())]
    appendices = sum(1 for h in head_upper if h.startswith("ПРИЛОЖЕНИЕ"))

    src_idx = None
    for i, p in enumerate(paras):
        if style_of(p) == "Heading1" and any(
                para_text(p).strip().upper().startswith(v)
                for v in ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСТОЧНИКОВ",
                          "СПИСОК ЛИТЕРАТУРЫ")):
            src_idx = i
    entries = 0
    if src_idx is not None:
        for p in paras[src_idx + 1:]:
            t = para_text(p).strip()
            if style_of(p) == "Heading1":
                break
            if re.match(r"^\d+[\.\)]?\s+\S", t):
                entries += 1

    chars = len(all_text)
    pages = chars / CHARS_PER_PAGE
    check(S_MUST, INFO, "оценка объёма основного текста", True,
          f"{chars} знаков ≈ {pages:.0f} страниц при {CHARS_PER_PAGE} знаках "
          f"на странице")
    check(S_MUST, INFO, "насыщенность работы", True,
          f"рисунков {len(fig_caps)}, таблиц {len(tbl_caps)}, "
          f"источников {entries}, приложений {appendices}")
    if args.min_pages or args.max_pages:
        lo = args.min_pages or 0
        hi = args.max_pages or 10 ** 6
        check(S_MUST, GATE, "объём в требуемых границах", lo <= pages <= hi,
              f"фактически ≈ {pages:.0f} страниц при требуемых "
              f"{args.min_pages or '—'}–{args.max_pages or '—'}")
    else:
        check(S_MUST, INFO, "границы объёма не заданы", True,
              "норму объёма возьми из локального акта и передай "
              "--min-pages/--max-pages; без неё объём не оценивается")

    if args.min_sources is not None:
        check(S_MUST, GATE, "число записей в списке источников",
              entries >= args.min_sources,
              f"записей {entries} при требуемых не менее {args.min_sources}")
    else:
        check(S_MUST, INFO, "записи в списке источников", entries > 0,
              f"записей {entries}")

    # =====================================================================
    # 2. ВЁРСТКА
    # =====================================================================
    sect = body.findall(f".//{{{W}}}sectPr")
    want = prof["margins"]
    alt_m = alt["margins"]
    ok_page, deviated_margins = bool(sect), False
    detail_page = []
    for sp in sect:
        pgSz = sp.find(f"{{{W}}}pgSz")
        pgMar = sp.find(f"{{{W}}}pgMar")
        if pgSz is None or pgMar is None:
            ok_page = False
            continue
        w_mm = int(pgSz.get(f"{{{W}}}w")) * TWIP_MM
        h_mm = int(pgSz.get(f"{{{W}}}h")) * TWIP_MM
        vals = tuple(int(pgMar.get(f"{{{W}}}{k}")) * TWIP_MM
                     for k in ("left", "right", "top", "bottom"))
        a4 = (abs(w_mm - 210) < 0.6 and abs(h_mm - 297) < 0.6) or \
             (abs(w_mm - 297) < 0.6 and abs(h_mm - 420) < 0.6)  # A4 или A3
        fits = all(abs(v - t) <= 1.0 for v, t in zip(vals, want))
        fits_alt = all(abs(v - t) <= 1.0 for v, t in zip(vals, alt_m))
        if not fits and fits_alt:
            deviated_margins = True
        ok_page = ok_page and a4 and (fits or fits_alt)
        detail_page.append(f"{w_mm:.0f}×{h_mm:.0f} мм; поля "
                           + "/".join(f"{v:.0f}" for v in vals) + " мм")
    detail_page = "; ".join(sorted(set(detail_page)))
    if deviated_margins and ok_page:
        check(S_LAYOUT, DEVIATION,
              f"поля профиля {'B' if args.profile == 'A' else 'A'} вместо "
              f"{'/'.join(str(x) for x in want)} мм", True, detail_page)
    else:
        check(S_LAYOUT, GATE,
              f"лист A4 и поля {'/'.join(str(x) for x in want)} мм "
              f"(профиль {args.profile}, {prof['name']})", ok_page, detail_page)

    normal = None
    for st in styles.findall(f"{{{W}}}style"):
        if st.get(f"{{{W}}}styleId") == "Normal":
            normal = st
    rPr = normal.find(f"{{{W}}}rPr") if normal is not None else None
    rf = rPr.find(f"{{{W}}}rFonts") if rPr is not None else None
    font = rf.get(f"{{{W}}}ascii") if rf is not None else None
    size_half = rPr.find(f"{{{W}}}sz") if rPr is not None else None
    size_pt = int(size_half.get(f"{{{W}}}val")) / 2 if size_half is not None else 0
    check(S_LAYOUT, GATE, f"гарнитура основного текста {args.font}",
          font == args.font, str(font))
    size_ok = size_pt >= prof["font_min"] and (
        prof["font_max"] is None or size_pt <= prof["font_max"])
    check(S_LAYOUT, GATE,
          f"кегль основного текста {prof['font_min']}"
          + (f"–{prof['font_max']} пт" if prof["font_max"] else " пт и больше"),
          size_ok, f"{size_pt:g} пт")

    spacing = normal.find(f"{{{W}}}pPr/{{{W}}}spacing") if normal is not None else None
    line = int(spacing.get(f"{{{W}}}line")) if spacing is not None else 0
    check(S_LAYOUT, GATE, "межстрочный интервал основного текста полуторный",
          line == 360, f"{line / 240:.2f}")

    ind = normal.find(f"{{{W}}}pPr/{{{W}}}ind") if normal is not None else None
    first = int(ind.get(f"{{{W}}}firstLine")) * TWIP_MM if ind is not None else 0
    check(S_LAYOUT, GATE, "абзацный отступ 1,25 см единый по тексту",
          abs(first - 12.5) <= 0.6, f"{first:.1f} мм")

    # Нумерация страниц
    def page_field(trees: dict) -> tuple[bool, bool]:
        found = centered = False
        for tree in trees.values():
            raw = etree.tostring(tree, encoding="unicode")
            if "PAGE" not in raw:
                continue
            found = True
            for p in tree.findall(f".//{{{W}}}p"):
                if "PAGE" in etree.tostring(p, encoding="unicode"):
                    jc = p.find(f"{{{W}}}pPr/{{{W}}}jc")
                    centered = centered or (jc is not None
                                            and jc.get(f"{{{W}}}val") == "center")
        return found, centered

    in_footer, foot_centered = page_field(footers)
    in_header, head_centered = page_field(headers)
    want_footer = prof["page_number"] == "footer"
    placed_right = in_footer if want_footer else in_header
    placed_alt = in_header if want_footer else in_footer
    where = "нижнем колонтитуле" if want_footer else "верхнем поле"
    if placed_right:
        check(S_LAYOUT, GATE, f"номер страницы в {where}", True,
              f"колонтитулов: верхних {len(headers)}, нижних {len(footers)}")
    elif placed_alt:
        check(S_LAYOUT, DEVIATION,
              f"номер страницы не в {where}, а в противоположном колонтитуле",
              True, "значение профиля " + ("A" if want_footer else "B"))
    else:
        check(S_LAYOUT, GATE, f"номер страницы в {where}", False,
              "поле PAGE в колонтитулах не найдено")
    check(S_LAYOUT, GATE, "номер страницы выровнен по центру",
          foot_centered or head_centered)
    check(S_LAYOUT, GATE, "на титульном листе номер не печатается",
          any(sp.find(f"{{{W}}}titlePg") is not None for sp in sect))

    # Заголовки
    bad_fonts = []
    for st in styles.findall(f"{{{W}}}style"):
        sid = st.get(f"{{{W}}}styleId") or ""
        if sid in ("Heading1", "Heading2", "Heading3", "Heading4"):
            hrf = st.find(f"{{{W}}}rPr/{{{W}}}rFonts")
            if hrf is None:
                bad_fonts.append(f"{sid}: нет rFonts")
                continue
            theme = sorted(a.split("}")[1] for a in hrf.attrib if a.endswith("Theme"))
            if theme or hrf.get(f"{{{W}}}ascii") != args.font:
                bad_fonts.append(f"{sid}: {hrf.get(f'{{{W}}}ascii')}, "
                                 f"темевые: {','.join(theme) or 'нет'}")
    check(S_LAYOUT, GATE,
          f"заголовки набраны {args.font} без темевой подмены",
          not bad_fonts, "; ".join(bad_fonts))

    # Стиль Heading может ссылаться на rFonts, но если у него нет собственной
    # записи, реальный шрифт берётся из w:docDefaults — там же живёт тема
    # Calibri, которую python-docx подставляет по умолчанию.
    doc_defaults_rf = styles.find(
        f"{{{W}}}docDefaults/{{{W}}}rPrDefault/{{{W}}}rPr/{{{W}}}rFonts")
    if doc_defaults_rf is not None:
        dd_theme = sorted(a.split("}")[1] for a in doc_defaults_rf.attrib
                          if a.endswith("Theme"))
        dd_ascii = doc_defaults_rf.get(f"{{{W}}}ascii")
        check(S_LAYOUT, GATE,
              f"шрифт по умолчанию пакета (w:docDefaults) — {args.font} без темы",
              not dd_theme and dd_ascii == args.font,
              f"ascii={dd_ascii}, темевые={','.join(dd_theme) or 'нет'}")
    else:
        check(S_LAYOUT, INFO, "w:docDefaults не переопределяет шрифт", True)

    # Прямое форматирование прогона всегда сильнее стиля: если сборщик (или
    # человек в Word) точечно поставил Calibri/темевой шрифт на текст самого
    # заголовка, проверка стиля Heading1–4 этого не увидит. Смотрим каждый
    # прогон каждого заголовка отдельно — это и есть тот случай, на котором
    # скилл уже дважды ошибался (заголовки Calibri в статье и в диссертации).
    bad_direct = []
    for p in heads_all:
        lvl = style_of(p)
        for r in p.findall(f"{{{W}}}r"):
            rPr = r.find(f"{{{W}}}rPr")
            rf = rPr.find(f"{{{W}}}rFonts") if rPr is not None else None
            if rf is None:
                continue
            theme = sorted(a.split("}")[1] for a in rf.attrib if a.endswith("Theme"))
            ascii_f = rf.get(f"{{{W}}}ascii")
            if theme or (ascii_f and ascii_f != args.font):
                txt = "".join(t.text or "" for t in r.iter(f"{{{W}}}t")).strip()
                bad_direct.append(f"{lvl} «{txt[:30]}»: ascii={ascii_f}, "
                                  f"темевые={','.join(theme) or 'нет'}")
    check(S_LAYOUT, GATE,
          "в тексте заголовков нет прямого форматирования шрифтом "
          f"мимо {args.font} (Calibri темой в т.ч.)",
          not bad_direct, "; ".join(bad_direct[:4]))

    h1 = None
    for st in styles.findall(f"{{{W}}}style"):
        if st.get(f"{{{W}}}styleId") == "Heading1":
            h1 = st
    jc1 = h1.find(f"{{{W}}}pPr/{{{W}}}jc") if h1 is not None else None
    caps1 = h1.find(f"{{{W}}}rPr/{{{W}}}caps") if h1 is not None else None
    dotted = [t for t in head_texts if t.endswith(".")]
    hyphened = [t for t in head_texts if "­" in t or "-\n" in t]
    centered1 = jc1 is not None and jc1.get(f"{{{W}}}val") == "center"
    check(S_LAYOUT, GATE,
          "заголовки структурных элементов по центру без точки и переносов",
          centered1 and not dotted and not hyphened,
          f"заголовков {len(heads1)}; выравнивание "
          f"{jc1.get(f'{{{W}}}val') if jc1 is not None else 'нет'}; "
          f"с точкой {len(dotted)}; с переносами {len(hyphened)}")
    has_caps = caps1 is not None and caps1.get(f"{{{W}}}val") != "0"
    check(S_LAYOUT, GATE if prof["caps_headings"] else INFO,
          "заголовки структурных элементов прописными буквами"
          + ("" if prof["caps_headings"] else " (профиль A не требует)"),
          has_caps if prof["caps_headings"] else True,
          f"прописные {has_caps}")

    pbb = h1.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") if h1 is not None else None
    style_break = pbb is not None and pbb.get(f"{{{W}}}val") != "0"
    own_break = sum(1 for p in heads1
                    if p.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") is not None
                    or p.findall(f".//{{{W}}}br[@{{{W}}}type='page']"))
    check(S_LAYOUT, GATE, "каждый заголовок первого уровня начинает новую страницу",
          style_break or (heads1 and own_break == len(heads1)),
          f"разрыв в стиле Heading1: {style_break}; собственных разрывов "
          f"{own_break} из {len(heads1)}")

    # Ручной разрыв страницы в пустом абзаце вместо pageBreakBefore в стиле —
    # типичный дефект вёрстки (layout-core.md, п. 14): создаёт пустые страницы
    # и ломается при правке текста выше. Пустой абзац с разрывом легален,
    # только если сразу за ним идёт заголовок первого уровня.
    stray_breaks = []
    for idx, p in enumerate(paras):
        if not p.findall(f".//{{{W}}}br[@{{{W}}}type='page']"):
            continue
        if style_of(p) == "Heading1":
            continue  # разрыв внутри самого заголовка уже учтён выше
        next_head = idx + 1 < len(paras) and style_of(paras[idx + 1]) == "Heading1"
        if para_text(p).strip() or not next_head:
            stray_breaks.append(f"параграф {idx}: «{para_text(p).strip()[:40]}»")
    check(S_LAYOUT, GATE,
          "нет ручных разрывов страницы вне пустого абзаца перед заголовком",
          not stray_breaks, "; ".join(stray_breaks[:4]))

    # Подписи рисунков и таблиц
    loose_fig = [para_text(p).strip() for p in paras
                 if para_text(p).strip().startswith("Рисунок ")]
    bad_fig = [t for t in loose_fig if not FIGURE_RE.match(t)]
    check(S_LAYOUT, GATE, "подписи рисунков вида «Рисунок N — Название»",
          not bad_fig, f"всего {len(loose_fig)}, нарушений {len(bad_fig)}: "
                       + "; ".join(bad_fig[:3]))

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
    # сокращений — самостоятельный структурный элемент, блок обозначений к
    # формулам — перечень условных обозначений. Оба исключения именованы.
    EXEMPT_HEADINGS = ("СПИСОК СОКРАЩЕНИЙ", "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ",
                       "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ")
    EXEMPT_INTRO = "обозначени"

    children = list(body)
    content_tables = captioned = 0
    bad_tbl, exempt = [], []
    TALL_TABLE_ROWS = 8  # эвристика: с такого числа строк таблица обычно переходит на следующую страницу
    no_repeat_header = []
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
        rows = node.findall(f"{{{W}}}tr")
        columns = len(rows[0].findall(f"{{{W}}}tc")) if rows else 0
        if current_head.startswith(EXEMPT_HEADINGS) or (
                EXEMPT_INTRO in prev.lower() and columns <= 3
                and not TABLE_RE.match(prev)):
            exempt.append(prev[:60] or current_head[:60])
            continue
        content_tables += 1
        if TABLE_RE.match(prev):
            captioned += 1
        else:
            bad_tbl.append(prev[:70] or "(над таблицей нет текста)")
        if len(rows) >= TALL_TABLE_ROWS:
            trPr = rows[0].find(f"{{{W}}}trPr")
            header = trPr.find(f"{{{W}}}tblHeader") if trPr is not None else None
            repeats = header is not None and header.get(f"{{{W}}}val") != "0"
            if not repeats:
                no_repeat_header.append(f"{prev[:60] or '(без надписи)'} ({len(rows)} строк)")
    check(S_LAYOUT, GATE, "над каждой таблицей стоит надпись «Таблица N — Название»",
          not bad_tbl, f"таблиц {content_tables}, с надписью {captioned}, "
                       f"без надписи {len(bad_tbl)}: " + "; ".join(bad_tbl[:3]))
    check(S_LAYOUT, INFO, "перечни, свёрстанные таблицей и не требующие номера",
          True, f"{len(exempt)}: " + "; ".join(exempt))
    check(S_LAYOUT, GATE,
          f"шапка повторяется на следующей странице у таблиц от {TALL_TABLE_ROWS} строк "
          "(w:tblHeader на первой строке)",
          not no_repeat_header, "; ".join(no_repeat_header[:3]))

    fig_below = fig_above = 0
    seen_drawing = False
    for node in body.iter():
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
    check(S_LAYOUT, GATE, "подпись рисунка расположена под изображением",
          fig_above == 0,
          f"под изображением {fig_below}, над изображением {fig_above}")

    bad_caps = []
    for p in paras:
        t = para_text(p).strip()
        if not CAPTION_STRICT.match(t):
            continue
        pPr = p.find(f"{{{W}}}pPr")
        jc = pPr.find(f"{{{W}}}jc") if pPr is not None else None
        pind = pPr.find(f"{{{W}}}ind") if pPr is not None else None
        jc_val = jc.get(f"{{{W}}}val") if jc is not None else "(наследует)"
        fl = pind.get(f"{{{W}}}firstLine") if pind is not None else None
        fl_val = int(fl) if fl not in (None, "0") else 0
        want_jc = "left" if t.startswith("Таблица") else "center"
        if jc_val != want_jc or fl_val != 0:
            bad_caps.append(f"{t[:28]}: {jc_val}, отступ {fl_val}")
    check(S_LAYOUT, GATE,
          "подписи таблиц слева и рисунков по центру без абзацного отступа",
          not bad_caps, "; ".join(bad_caps[:4]))

    body_tabs = len(list(doc.iter(f"{{{W}}}tab")))
    text_tabs = sum((t.text or "").count("\t") for t in doc.iter(f"{{{W}}}t"))
    check(S_LAYOUT, GATE, "в документе нет табуляторов",
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
        empty = (not para_text(nxt).strip()
                 and not nxt.findall(f".//{{{W}}}drawing")
                 and not nxt.findall(f".//{{{M}}}oMath"))
        if empty and (after is None or etree.QName(after).localname != "tbl"):
            bad_spacers += 1
    check(S_LAYOUT, GATE, "пустых абзацев после таблиц вне стыка двух таблиц нет",
          not bad_spacers, f"распорок: {bad_spacers}")

    # Оглавление и титульный лист
    has_toc = any(para_text(p).strip().upper().startswith(("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"))
                  for p in paras)
    check(S_LAYOUT, GATE, "в документе есть оглавление", has_toc)
    toc_fields = [t.text or "" for t in doc.iter(f"{{{W}}}instrText")
                  if "TOC" in (t.text or "")]
    check(S_LAYOUT, GATE, "оглавление собрано полем TOC, а не статическим текстом",
          bool(toc_fields), "; ".join(toc_fields[:2]))
    check(S_LAYOUT, INFO, "автообновление полей при открытии (w:updateFields)",
          b"updateFields" in settings)

    # Стили TOC 1/TOC 2: унаследованный от Normal абзацный отступ или
    # выключка «по ширине» выталкивают перенесённую строку записи влево и
    # отрывают номер страницы от текста заголовка.
    bad_toc_style = []
    toc_style_found = False
    for st in styles.findall(f"{{{W}}}style"):
        sid = (st.get(f"{{{W}}}styleId") or "").replace(" ", "").lower()
        if sid not in ("toc1", "toc2"):
            continue
        toc_style_found = True
        pPr = st.find(f"{{{W}}}pPr")
        ind = pPr.find(f"{{{W}}}ind") if pPr is not None else None
        jc = pPr.find(f"{{{W}}}jc") if pPr is not None else None
        fl = int(ind.get(f"{{{W}}}firstLine", "0")) if ind is not None else 0
        jc_val = jc.get(f"{{{W}}}val") if jc is not None else None
        if fl != 0 or jc_val in ("both", "center"):
            bad_toc_style.append(f"{st.get(f'{{{W}}}styleId')}: отступ {fl}, "
                                 f"выравнивание {jc_val or 'наследует'}")
    if toc_style_found:
        check(S_LAYOUT, GATE,
              "стили TOC 1/TOC 2 без абзацного отступа и без выключки по ширине",
              not bad_toc_style, "; ".join(bad_toc_style))
    else:
        check(S_LAYOUT, INFO, "стили TOC 1/TOC 2 не переопределены (стандартные Word)",
              True)

    def starts_new_page(p) -> bool:
        if p.find(f"{{{W}}}pPr/{{{W}}}pageBreakBefore") is not None:
            return True
        if style_break and style_of(p) == "Heading1":
            return True
        return bool(p.findall(f".//{{{W}}}br[@{{{W}}}type='page']"))

    city_idx = toc_idx = None
    for i, p in enumerate(paras):
        t = para_text(p).strip()
        if city_idx is None and CITY_YEAR_RE.match(t):
            city_idx = i
        if toc_idx is None and t.upper() in ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"):
            toc_idx = i
    if toc_idx is None:
        check(S_LAYOUT, GATE,
              "содержание начинается с отдельной страницы после титульного листа",
              False, "заголовок содержания не найден")
    else:
        start = city_idx + 1 if city_idx is not None and city_idx < toc_idx else 0
        separated = any(starts_new_page(paras[k])
                        for k in range(start, toc_idx + 1))
        check(S_LAYOUT, GATE,
              "содержание начинается с отдельной страницы после титульного листа",
              separated,
              f"строка «город — год» параграф {city_idx}, содержание параграф "
              f"{toc_idx}")

    # =====================================================================
    # 3. ФОРМУЛЫ
    # =====================================================================
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
    check(S_MATH, INFO, "нумерованных формульных строк найдено", True, str(len(numbers)))
    check(S_MATH, GATE, "номера формул не повторяются",
          len(numbers) == len(set(numbers)),
          ", ".join(sorted({n for n in numbers if numbers.count(n) > 1})))

    right_aligned = 0
    for _, _, cell, _ in formula_rows:
        jc = cell.find(f".//{{{W}}}p/{{{W}}}pPr/{{{W}}}jc")
        if jc is not None and jc.get(f"{{{W}}}val") in ("right", "end"):
            right_aligned += 1
    check(S_MATH, GATE, "номер формулы прижат к правому краю",
          right_aligned == len(formula_rows),
          f"{right_aligned} из {len(formula_rows)}")

    # Формула с номером верстается таблицей фиксированной ширины во всю полосу
    # набора: иначе Word подгоняет колонки под содержимое и номер уходит от
    # края. Обе ячейки — с вертикальным центрированием, чтобы у высокой формулы
    # номер остался на её уровне.
    if sect:
        sp = sect[-1]
        pgSz, pgMar = sp.find(f"{{{W}}}pgSz"), sp.find(f"{{{W}}}pgMar")
        total = (int(pgSz.get(f"{{{W}}}w")) - int(pgMar.get(f"{{{W}}}left"))
                 - int(pgMar.get(f"{{{W}}}right"))) if pgSz is not None \
            and pgMar is not None else 9355
    else:
        total = 9355
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
        ok_layout = layout is not None and layout.get(f"{{{W}}}type") == "fixed"
        ok_w = (tblW is not None and tblW.get(f"{{{W}}}type") == "dxa"
                and abs(int(tblW.get(f"{{{W}}}w")) - total) <= 60)
        for ci, tc in enumerate(tcs):
            tcPr = tc.find(f"{{{W}}}tcPr")
            va = tcPr.find(f"{{{W}}}vAlign") if tcPr is not None else None
            if va is None or va.get(f"{{{W}}}val") != "center":
                bad_geom.append(f"{right_text}: ячейка {ci} без vAlign=center")
        if not (ok_layout and ok_w):
            bad_geom.append(
                f"{right_text}: layout="
                f"{layout.get(f'{{{W}}}type') if layout is not None else None}, "
                f"tblW={tblW.get(f'{{{W}}}w') if tblW is not None else None}, "
                f"полоса набора {total}")
    check(S_MATH, GATE,
          "формульная строка — фиксированная полоса с вертикальным центрированием",
          not bad_geom, "; ".join(bad_geom[:4]))

    empty_left = [n for _, left, _, n in formula_rows
                  if not "".join(t.text or "" for t in left.iter(f"{{{W}}}t")).strip()
                  and not left.findall(f".//{{{W}}}drawing")
                  and not left.findall(f".//{{{M}}}oMath")]
    check(S_MATH, GATE, "левая ячейка строки нумерации непустая",
          not empty_left, ", ".join(empty_left))

    omath = doc.findall(f".//{{{M}}}oMath")
    omath_para = doc.findall(f".//{{{M}}}oMathPara")
    check(S_MATH, GATE, "число объектов OMML не меньше числа нумерованных формул",
          len(omath) >= len(formula_rows),
          f"m:oMath {len(omath)}, m:oMathPara {len(omath_para)}, "
          f"нумерованных записей {len(formula_rows)}")

    raster_formula = [n for _, left, _, n in formula_rows
                      if left.findall(f".//{{{W}}}drawing")]
    check(S_MATH, GATE, "формула не вставлена растром", not raster_formula,
          f"растром записаны: {', '.join(raster_formula)}" if raster_formula else "")
    check(S_MATH, INFO, "растров в пакете документа", True, str(len(media)))

    # Растр допускается только как самостоятельная иллюстрация с подписью
    # «Рисунок N». Обозначение внутри абзаца растром быть не должно: оно не
    # ищется, не масштабируется и не открывается редактором уравнений.
    inline_raster = []
    for para in paras:
        if not para.findall(f".//{{{W}}}drawing"):
            continue
        if para_text(para).strip():
            inline_raster.append(para_text(para).strip()[:60])
    check(S_MATH, GATE, "обозначения в тексте не сохранены растром",
          not inline_raster,
          f"абзацев с растром внутри текста: {len(inline_raster)}: "
          + "; ".join(inline_raster[:3]))

    inline_math = [om for om in doc.iter(f"{{{M}}}oMath")
                   if om.getparent() is not None
                   and etree.QName(om.getparent()).localname != "oMathPara"]
    check(S_MATH, INFO, "строчных обозначений объектами OMML", True,
          str(len(inline_math)))

    gd_missing = []
    for p in paras:
        t = para_text(p).strip()
        if not t.lower().startswith("где"):
            continue
        if not p.findall(f".//{{{M}}}oMath"):
            gd_missing.append(t[:50])
    check(S_MATH, GATE, "в пояснениях «где» символы набраны объектами OMML",
          not gd_missing, "; ".join(gd_missing[:3]))
    colon_where = [para_text(p).strip()[:40] for p in paras
                   if para_text(p).strip().lower().startswith("где:")]
    check(S_MATH, GATE, "пояснения начинаются словом «где» без двоеточия",
          not colon_where, "; ".join(colon_where[:3]))

    flat_re = re.compile(r"(?<![A-Za-zА-Яа-яЁё])(?:[A-Za-z][LR])(?![A-Za-zА-Яа-яЁё])")
    flat_hits = []
    for tnode in doc.iter(f"{{{W}}}t"):
        if any(etree.QName(a).localname == "oMath" for a in tnode.iterancestors()):
            continue
        flat_hits.extend(flat_re.findall(tnode.text or ""))
    check(S_MATH, GATE, "объекты с индексами не набраны текстом (tL, qR и родня)",
          not flat_hits, ", ".join(sorted(set(flat_hits))[:8]))

    structures: dict[str, int] = {}
    for om in doc.iter(f"{{{M}}}oMath"):
        for node in om.iter():
            q = etree.QName(node).localname
            if q in ("sSub", "sSup", "sSubSup", "f", "nary", "d", "rad"):
                structures[q] = structures.get(q, 0) + 1
    check(S_MATH, INFO, "структурных элементов математики", True,
          ", ".join(f"{k}: {v}" for k, v in sorted(structures.items())))

    # =====================================================================
    # 4. БИБЛИОГРАФИЯ И ССЫЛКИ
    # =====================================================================
    def numbers_of(regex, caps: list[str]) -> list[str]:
        out = []
        for t in caps:
            m = regex.match(t)
            if m:
                out.append(m.group(1))
        return out

    def gaps(keys: list[str]) -> list[str]:
        by_group: dict[str, list[int]] = {}
        for k in keys:
            head, _, tail = k.rpartition(".")
            try:
                by_group.setdefault(head or "сквозная", []).append(int(tail))
            except ValueError:
                continue
        bad = []
        for group, nums in by_group.items():
            nums.sort()
            if nums != list(range(1, len(nums) + 1)):
                bad.append(f"{group}: {nums}")
        return bad

    fig_nums = numbers_of(FIGURE_RE, fig_caps)
    check(S_BIB, GATE, "нумерация рисунков без пропусков",
          not gaps(fig_nums), "; ".join(gaps(fig_nums)))
    tbl_nums = numbers_of(TABLE_RE, tbl_caps)
    check(S_BIB, GATE, "нумерация таблиц без пропусков",
          not gaps(tbl_nums), "; ".join(gaps(tbl_nums)))
    form_nums = [n.strip("()") for n in numbers]
    check(S_BIB, GATE, "нумерация формул без пропусков",
          not gaps(form_nums), "; ".join(gaps(form_nums)))

    def unreferenced(keys: list[str], word_forms) -> list[str]:
        out = []
        for key in keys:
            if not any(re.search(w + r"\w*\s+" + re.escape(key) + r"(?!\d)", all_text)
                       for w in word_forms):
                out.append(key)
        return out

    miss_fig = unreferenced(fig_nums, ["рисун", "Рисун"])
    check(S_BIB, GATE, "на каждый рисунок есть ссылка в тексте",
          not miss_fig, ", ".join(miss_fig))
    miss_tbl = unreferenced(tbl_nums, ["таблиц", "Таблиц"])
    check(S_BIB, GATE, "на каждую таблицу есть ссылка в тексте",
          not miss_tbl, ", ".join(miss_tbl))
    miss_form = [n for n in numbers if all_text.count(n) < 2]
    check(S_BIB, INFO, "формулы без отдельной ссылки по номеру в тексте", True,
          ", ".join(miss_form) if miss_form else "нет")

    cited = set(re.findall(r"\[(\d+)(?:[,;–-]\s*\d+)*\]", all_text))
    if cited and entries:
        over = sorted(int(c) for c in cited if int(c) > entries)
        check(S_BIB, GATE, "затекстовые отсылки указывают на существующие записи",
              not over, f"записей {entries}, отсылки за пределами списка: "
                        + ", ".join(str(o) for o in over[:8]))
    check(S_BIB, INFO, "приложений в документе", True, str(appendices))

    # =====================================================================
    # 5. МЕТАДАННЫЕ
    # =====================================================================
    core_s = core.decode("utf-8", errors="replace")
    app_s = app.decode("utf-8", errors="replace")
    creator = re.search(r"<dc:creator>([^<]*)</dc:creator>", core_s)
    modifier = re.search(r"<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>", core_s)
    cr = creator.group(1).strip() if creator else ""
    md = modifier.group(1).strip() if modifier else ""
    if args.author:
        check(S_META, GATE, f"автор и последний редактор — {args.author}",
              cr == args.author and md == args.author,
              f"creator={cr!r}, lastModifiedBy={md!r}")
    else:
        check(S_META, GATE, "автор и последний редактор заполнены без имени "
                            "генератора",
              bool(cr) and bool(md) and not GENERATOR_RE.search(cr)
              and not GENERATOR_RE.search(md),
              f"creator={cr!r}, lastModifiedBy={md!r}")
    check(S_META, GATE, "в свойствах нет следов автоматизации",
          "python-docx" not in core_s and "generated by" not in core_s)

    created = re.search(r"<dcterms:created[^>]*>([^<]+)</dcterms:created>", core_s)
    modified = re.search(r"<dcterms:modified[^>]*>([^<]+)</dcterms:modified>", core_s)
    rev = re.search(r"<cp:revision>(\d+)</cp:revision>", core_s)
    ok_dates = bool(created and modified and created.group(1) <= modified.group(1)
                    and rev and int(rev.group(1)) > 1)
    check(S_META, GATE if args.strict_metadata else INFO,
          "даты правки последовательны, правок больше одной", ok_dates,
          f"created={created.group(1) if created else None}, "
          f"modified={modified.group(1) if modified else None}, "
          f"revision={rev.group(1) if rev else None}")
    total_time = re.search(r"<TotalTime>(\d+)</TotalTime>", app_s)
    appl = re.search(r"<Application>([^<]*)</Application>", app_s)
    check(S_META, GATE if args.strict_metadata else INFO,
          "приложение — Word для Windows, время редактирования заполнено",
          bool(appl and appl.group(1) == "Microsoft Office Word"
               and total_time and int(total_time.group(1)) > 0),
          f"TotalTime={total_time.group(1) if total_time else None}, "
          f"Application={appl.group(1) if appl else None}")
    check(S_META, GATE, "в пакете нет миниатюры шаблона",
          "docProps/thumbnail.jpeg" not in names)
    title_m = re.search(r"<dc:title>([^<]*)</dc:title>", core_s)
    title_v = title_m.group(1).strip() if title_m else ""
    check(S_META, GATE, "название документа заполнено реальным названием работы",
          len(title_v) >= 20 and title_v.lower() not in
          ("диссертация", "вкр", "отчёт", "документ"),
          repr(title_v[:60]))
    kw = re.search(r"<cp:keywords>([^<]*)</cp:keywords>", core_s)
    kw_count = len([k for k in kw.group(1).split(";") if k.strip()]) if kw else 0
    check(S_META, GATE, "теги документа заполнены (не меньше трёх)",
          kw_count >= 3, f"тегов: {kw_count}")

    # =====================================================================
    # Вывод
    # =====================================================================
    gates = [r for r in results if r[1] == GATE]
    failed = [r for r in gates if not r[3]]
    must_failed = [r for r in failed if r[0] == S_MUST]
    print(f"\n=== Приёмка: {docx.name} ===")
    print(f"профиль {args.profile} ({prof['name']}), вид документа {args.doctype}")
    for section in SECTIONS:
        rows = [r for r in results if r[0] == section]
        if not rows:
            continue
        print(f"\n--- {section} ---")
        for _, kind, name, ok, detail in rows:
            if kind == DEVIATION:
                mark = " откл. "
            elif ok:
                mark = "  OK   "
            elif kind == GATE:
                mark = "  FAIL "
            else:
                mark = "  инфо "
            print(mark + name + (f"  :: {detail}" if detail else ""))
    print(f"\nГейтов: {len(gates)}, нарушено: {len(failed)}"
          f" (из них обязательных элементов: {len(must_failed)})")
    print("ИТОГ:", "PASS" if not failed else "FAIL")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
