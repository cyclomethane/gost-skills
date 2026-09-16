---
name: gost
description: Formats a document to ГОСТ (Russian state standard) from scratch, or reviews a finished one and flags every error with a reference to the standard's clause — dissertation and avtoreferat, thesis (ВКР) and diploma, research report, journal article, abstract and annotation, official/business document, defense presentation, bibliography, units of measurement. Saves weeks of revisions and guards against the work being sent back for rework. Use whenever the user mentions ГОСТ, оформление or проверку document, титульный лист, список литературы, диссертацию, автореферат, научный доклад, ВКР, диплом, реферат, курсовую, статью, отчёт, единицы измерения — even when the standard is not named.
---

# GOST documents: router and shared layer

One skill for all ГОСТ-governed documents. It determines the document type
itself, loads the matching module, and maintains the shared layer for
layout, units, and bibliography — the same across all document types. Work
proceeds in one of two modes: assemble a new document to the standard, or
accept a finished document with an audit report in which every remark
carries a reference to the standard's or guideline's clause.

Requirement priority: first, the requirements of the educational
institution, department, degree program, or edition for the specific
document; then the currently effective national standards within their
scope; then this skill's modules as a normalized working interpretation of
the standards; then general editorial recommendations — only where the
levels above leave a choice. A conflict between requirements is recorded
with both references; the higher-priority source applies.

## Step 1. Document type and what to read

| Document or task | Module |
| --- | --- |
| Dissertation, avtoreferat, scientific report (ГОСТ Р 7.0.11-2011) | `references/dissertation.md`: workflow, full checklists, and reference files in `references/dissertation/`, acceptance script `scripts/gost_acceptance.py` |
| ВКР (thesis), diploma and master's project | `references/vkr.md` |
| Research report (НИР); layout profile when the guideline references ГОСТ 7.32 | `references/nir-report-7.32.md` |
| ЕСКД text document: explanatory note, design documentation | `references/text-documents-2.105.md` |
| Journal or proceedings article (ГОСТ Р 7.0.7-2021) | `references/journal-article-7.0.7.md` |
| Abstract and annotation as a secondary scientific document (ГОСТ Р 7.0.99-2018) | `references/abstract-referat-7.0.99.md` |
| Application, memo, certificate, act, minutes (ГОСТ Р 7.0.97-2025) | `references/official-documents-7.0.97.md` |
| Presentation and talk for a thesis or research defense | `references/defense-presentation.md` |
| Bibliographic references: in-text, footnote, end-of-text | `references/references-7.0.5.md` |
| Bibliography entries (ГОСТ Р 7.0.100-2018) | `references/bibliography-records-7.0.100.md` |
| References to online sources: websites, online articles, repositories, datasets | `references/web-references-7.0.108.md` |
| Numbers and units of measurement in any document (ГОСТ 8.417-2024) | `references/units-8.417.md` |

A type module holds what the shared layer doesn't: document structure, the
institution's title-page template, type-specific font sizes, the boundaries
of the standard's scope. Modules combine. A ВКР (thesis) is assembled as:
`vkr.md` as the base, the layout profile from `nir-report-7.32.md`, the
abstract per `abstract-referat-7.0.99.md`, the bibliography trio plus
`units-8.417.md` on top, and `defense-presentation.md` before the defense. A
journal article is `journal-article-7.0.7.md` plus the same bibliography
trio.

Term papers (курсовая), coursework abstracts with an outline, internship
reports, conference abstracts, monographs, and degree-committee service
documents (review, appraisal, implementation act) don't have a dedicated
module yet. `vkr.md` and `nir-report-7.32.md` partially cover them; a full
module is created via the extension protocol below, and any specific point
without a source gets NEEDS_SOURCE_CHECK.

## Step 2. Shared DOCX layout layer

The mechanical check reads `document.xml` from the unpacked package and does
not modify the file. Use the `docx` skill to unpack. A twip is 1/1440 inch;
font size in `w:sz` is stored in half-points.

Standards disagree on margin values, font size, and page-number placement,
so the shared layer maintains two profiles. Values come from the
document-type module; where no module exists yet, the specific point gets
NEEDS_SOURCE_CHECK.

| Requirement | ГОСТ Р 7.0.11 profile (dissertation) | ГОСТ 7.32 / ВКР profile |
| --- | --- | --- |
| Margins mm: left/right/top/bottom | 25/10/20/20 — `w:pgMar` 1417/567/1134/1134 | 30/15/20/20 — `w:pgMar` 1701/850/1134/1134 |
| Page number | centered in top margin | footer; ГОСТ 7.32 itself doesn't fix the position — the local template sets it |
| Font size | 12–14 pt — `w:sz`/`w:szCs` 24 to 28 | at least 12 pt — `w:sz` from 24; 14 pt is the safe ВКР profile |
| 1.5 line spacing | `w:spacing w:line="360" w:lineRule="auto"` | same |
| 1.25 cm first-line indent | `w:ind w:firstLine="709"`, consistent throughout | same |
| A4 sheet 210 × 297 mm | `w:pgSz w:w="11906" w:h="16838"` | same; large report tables use A3 |

Rules common to both profiles: headings use an explicit font with no theme
override — remove `w:asciiTheme`, `w:hAnsiTheme`, `w:cstheme`,
`w:eastAsiaTheme` from Heading styles, because a theme attribute overrides
an explicit font and reverts to Calibri Light; the table of contents is an
auto-generated field `TOC \o "1-2" \h \z \u` with `w:updateFields
w:val="true"` in settings.xml; captions per ГОСТ 2.105 read "Таблица N —
Название" above the table, left-aligned, and "Рисунок N — Название" below
the figure, centered; cell text is set compactly with no first-line indent;
tab characters in text are copy-paste debris — replace with a space;
metadata carries no trace of automation — `dc:creator` and
`cp:lastModifiedBy` mention no generators, `Application` in app.xml is
Microsoft Office Word, and `docProps/thumbnail.jpeg` is absent from the
package.

The dissertation this skill was calibrated against intentionally diverges
from the ГОСТ Р 7.0.11 text in two places: 30/15 margins instead of 25/10,
and the page number at the bottom instead of the top. Both values match the
ГОСТ 7.32 profile; the acceptance report flags these divergences
separately, without failing the check. When reviewing a third-party
document, apply that document's exact profile values.

## Formulas — OMML

Every formula in any document is an Office Math object: inline is
`m:oMath`, display is `m:oMathPara`. Raster images, EMF, EQ fields,
MathType objects, and plain-text formulas don't pass. Objects with indices
are built with `m:sSub`/`m:sSup` structures, not glued-together text. A
single variable in running text is not a formula; an expression with an
operation or relation is. Formula sources are stored in
`FORMULA_SOURCES.json`; the full formula-line layout and geometry-check
spec is in `references/dissertation/07-formulas.md`.

Modules add punctuation rules on top of the mechanics: per ГОСТ 7.32,
explanations follow immediately after the formula, "где" is written
without a colon, the number sits in parentheses on the right as `(3.1)` or
`(В.1)`, and the sign repeats at the start of the new line when a formula
wraps. These requirements apply together with the OMML rule, not instead
of it.

## Units of measurement

Numbers and units in every document type follow ГОСТ 8.417-2024: unit
symbols are set upright, a space separates the number and the unit (`100
кВт`, `80 %`), no period follows the symbol, and prefix case matters — M
and m, k and K denote different magnitudes. Quantities in text and formulas
are set in italics, units upright. Edge cases, ranges, temperature, angles,
and computing units are covered in `references/units-8.417.md`.

## Bibliography

The bibliographic apparatus spans three modules: how to cite a source in
text — `references/references-7.0.5.md`; how to record a source in the
reference list — `references/bibliography-records-7.0.100.md`; how to cite
online sources — `references/web-references-7.0.108.md`. For a
dissertation, the bibliographic apparatus lives entirely in
`references/dissertation/04-bibliography.md`. Abbreviations in references
follow ГОСТ Р 7.0.12, ГОСТ 7.11, and ГОСТ 7.12; these have no dedicated
module. Check each standard's current status against the "National
Standards" index for the current year; until status is verified, mark
NEEDS_SOURCE_CHECK and don't declare the requirement in force.

## Report

Assign every checklist point a status: PASS, FAIL, N/A, or
NEEDS_SOURCE_CHECK. The report has five sections: critical
non-conformities, technical non-conformities, bibliographic issues,
optional elements, and traceability to standard clauses. If a module
defines its own report format — `vkr.md` with sections A–E, or another
module's machine checklist — use the module's format while keeping the
same statuses.

## Extension protocol: a new document type

1. Get the requirements from the user: a guideline, standard, university
   template, or a finished sample document.
2. Record the source in the module's header: title, edition, date
   obtained, what the source doesn't cover.
3. Create `references/<type>.md`: document structure, title page,
   type-specific margin/font-size/spacing values, deviations from the
   Step 2 shared layer.
4. Compile mechanically checkable requirements into a "requirement — DOCX
   value" table following the Step 2 pattern.
5. Add a row to the Step 1 routing table, and to the profile table if the
   module introduces a new value profile.

A guideline's requirement outranks the general standard when they
diverge — the same way a dissertation council's requirements outrank ГОСТ
for a dissertation (`references/dissertation.md`). Record the divergence in
the report with both references.

## Boundaries

For a type without a module, the skill doesn't invent requirements: a
point without a source gets NEEDS_SOURCE_CHECK.

Modules are a normalized working interpretation of the standards (version
2026-09-16) with traceability to source clauses and official
protect.gost.ru cards in their headers. The skill's core rule — OMML math
and the priority of local requirements — is repeated in every module's
header and takes precedence over the module's content.
