---
name: gost
description: Formats a document to ГОСТ (Russian state standard) from scratch, or reviews a finished one and reports every violation with a reference to the standard's clause — missing mandatory sections, volume out of range, and all visual layout errors. Covers dissertation and avtoreferat, thesis (ВКР) and diploma, research report, term paper, internship report, journal article, abstract and annotation, official/business document, defense presentation, bibliography, units of measurement. Saves weeks of revisions and guards against the work being sent back for rework. Use whenever the user mentions ГОСТ, оформление or проверку document, титульный лист, список литературы, диссертацию, автореферат, научный доклад, ВКР, диплом, реферат, курсовую, статью, отчёт, нормоконтроль, единицы измерения — even when the standard is not named.
---

# GOST documents: router and shared layer

One skill for all ГОСТ-governed documents. It determines the document type
itself, loads the matching module, and keeps every shared rule in exactly two
core files, so a module never repeats a rule it did not change:

- `references/layout-core.md` — **visual layout**, the skill's main job: page,
  text, headings, pagination, TOC, captions, tables, figures, formulas, lists,
  appendices, package metadata. Three profiles (А — ГОСТ Р 7.0.11, Б —
  ГОСТ 7.32, В — ГОСТ Р 2.105) resolve every disagreement between standards;
  each row carries the DOCX value to check.
- `references/mandatory-core.md` — **mandatory elements and volume**: which
  structural elements each document type must have, what goes on the title
  page, the eight elements of the introduction, volume norms, and how to word
  a violation.

Two modes: assemble a new document to the standard, or accept a finished one
with an audit report in which every remark carries a clause reference.

## Requirement priority

1. Requirements of the institution, department, degree program, dissertation
   council, or journal for this specific document.
2. Currently effective national standards within their scope.
3. This skill's modules as a normalized working interpretation.
4. General editorial recommendations — only where the levels above leave a
   choice.

A conflict is recorded with both references; the higher level applies. Never
rebuild a local template on your own initiative.

## Rule statuses and report statuses

Rules: **MUST** — direct requirement of a standard; **SHOULD** —
recommendation or safe practice; **MAY** — permitted option; **LOCAL** — the
value comes from a local act; **COND** — mandatory once a condition holds.

Report: **PASS**, **FAIL**, **N/A**, **NEEDS_SOURCE_CHECK**. A point with no
source gets NEEDS_SOURCE_CHECK and is never declared a violation.

## Step 1. Document type and what to read

| Document or task | Module |
| --- | --- |
| Dissertation, avtoreferat, scientific report (ГОСТ Р 7.0.11-2011) | `references/dissertation.md` + files in `references/dissertation/`, acceptance script `scripts/gost_acceptance.py` |
| ВКР (thesis), diploma, master's project, term paper, internship report | `references/vkr.md` |
| Research report (НИР); source of layout profile Б | `references/nir-report-7.32.md` |
| ЕСКД text document: explanatory note, design documentation | `references/text-documents-2.105.md` |
| Journal or proceedings article (ГОСТ Р 7.0.7-2021) | `references/journal-article-7.0.7.md` |
| Abstract and annotation as a secondary document (ГОСТ Р 7.0.99-2018) | `references/abstract-referat-7.0.99.md` |
| Application, memo, certificate, act, minutes (ГОСТ Р 7.0.97-2025) | `references/official-documents-7.0.97.md` |
| Presentation and talk for a thesis or research defense | `references/defense-presentation.md` |
| Bibliographic references: in-text, footnote, end-of-text | `references/references-7.0.5.md` |
| Bibliography entries (ГОСТ Р 7.0.100-2018) | `references/bibliography-records-7.0.100.md` |
| References to online sources: sites, online articles, repositories, datasets | `references/web-references-7.0.108.md` |
| Numbers and units of measurement in any document (ГОСТ 8.417-2024) | `references/units-8.417.md` |

Always read both core files plus the type module. A module holds only what the
cores don't: the mandatory composition of this type, the title-page template,
type-specific values, the boundaries of the standard's scope.

Modules combine. ВКР = `vkr.md` + profile Б from `nir-report-7.32.md` + the
abstract per `abstract-referat-7.0.99.md` + the bibliography trio +
`units-8.417.md`, and `defense-presentation.md` before the defense. A journal
article = `journal-article-7.0.7.md` + the bibliography trio. A term paper,
internship report, or engineering explanatory note is `vkr.md` reduced by the
local guideline, plus `text-documents-2.105.md` for the ЕСКД part.

Conference abstracts, monographs, and degree-committee service documents
(review, appraisal, implementation act) have no module; a full one is created
via the extension protocol below, and any point without a source gets
NEEDS_SOURCE_CHECK.

## Step 2. Audit procedure

1. Establish the document type and the source of requirements: local act,
   guideline, template, sample. Fix the layout profile (А, Б, or В).
2. Check the standard's current status against the "National Standards" index
   for the current year; an unverified status means NEEDS_SOURCE_CHECK.
3. **Mandatory elements and volume** per `mandatory-core.md`: what is missing,
   what is out of range, what is optional and merely absent.
4. **Mechanical layout check** per `layout-core.md` against the document file.
   For an assembled document run `python scripts/gost_acceptance.py file.docx
   --profile A|B`; for a third-party DOCX use the `docx` skill to unpack and
   compare `document.xml` values with the profile table. Reading only — the
   check never modifies the file.
5. Content check: chain from the title through problem, aim, tasks, chapters,
   method, results, conclusion; every claim traced to its evidence. Also a
   plain sanity pass while reading, no clause number attached: a table column
   nobody explains, a value that doesn't add up, a sentence that reads as a
   leftover from an earlier draft, a structural choice with no apparent
   reason. Flag it as a question for the author even when it isn't a GOST
   violation and even when it clearly predates this session's edits — the
   point is to say what an attentive reader would notice, not just what's
   mechanically checkable.
6. Bibliographic apparatus: in-text references, list entries, online sources.
7. Report.

## Step 3. Output scale — match the report to the ask

Default to short and useful, not exhaustive. The full six-section report below
is for an actual audit request or a `gost_acceptance.py` run — not for every
mention of ГОСТ.

- **Just connecting the skill, or a general ask** ("оформи по ГОСТ", "что тут
  не так", a mention with no explicit audit request) → a short list of what's
  missing or wrong, in plain language, ranked by impact. No clause numbers, no
  six-section structure, no PASS/FAIL table for things that are fine.
- **An explicit audit, review, or "проверь по ГОСТ" request, or the acceptance
  script was run** → the full report below: all six sections, every status
  (PASS/FAIL/N/A/NEEDS_SOURCE_CHECK), clause numbers throughout.
- **Clause numbers and traceability** appear only when the user asked for
  detail, for a formal audit, or the finding is a rejection risk (missing
  mandatory element, structure). A quick recommendation doesn't need `[5.3.1]`
  after every line.

## Step 4. Report

Six sections, in this order:

1. **Mandatory elements** — missing structural elements, incomplete
   introduction, volume out of range, title-page fields absent. Word each one
   as required / found / where / clause / fix.
2. **Critical non-conformities** — anything that gets the work sent back:
   structure, numbering, formulas as images, a missing reference list.
3. **Technical (visual) non-conformities** — margins, font, spacing,
   indents, headings, pagination, captions, tables, figures, formulas, TOC.
4. **Bibliographic issues** — with the referenced standard and a note on its
   status.
5. **Optional elements** — N/A, never counted as violations.
6. **Traceability** — a clause reference for every remark.

If a module defines its own report format — `vkr.md` sections A–E, or a
module's machine checklist — use it while keeping these statuses and keeping
mandatory elements first.

## Extension protocol: a new document type

1. Get the requirements from the user: guideline, standard, template, or a
   finished sample.
2. Record the source in the module header: title, edition, date obtained, what
   the source does not cover.
3. Create `references/<type>.md` with only the deltas: mandatory composition,
   title page, values that differ from the profile tables, scope boundaries.
   Everything else stays a pointer to the two core files.
4. Add the type's column to the matrix in `mandatory-core.md`; add a value
   profile to `layout-core.md` only if the type really introduces one.
5. Add a row to the Step 1 routing table.

## Boundaries

For a type without a module the skill invents nothing: a point without a source
gets NEEDS_SOURCE_CHECK. Modules are a normalized working interpretation of the
standards (version 2026-09-16) with traceability to source clauses and official
protect.gost.ru cards in their headers. The skill's own rules — OMML math and
the priority of local requirements — are repeated in every module header and
outrank the module's content. A dissertation published as a monograph is
outside ГОСТ Р 7.0.11; ВКР has no single national standard at all.
