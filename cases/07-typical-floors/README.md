# Case 07 — Typical floors from the architectural link

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

On a residential project with several buildings, the HVAC engineer needs identical floors: floors whose apartments carry the same set of rooms — combined bathroom (СУ), WC (уборная), walk-in closet (гардеробная). The deliverable is his «Таблица 1»: Floor | No. | Type | Rooms | Total area | Living area | Bathroom | WC | Closet.

Revit only supplies the raw room rows. Every total, group number, and the table itself is calculated in Excel, so a colleague can see the formula and change it. The same task comes back on the next project, with a different architectural team and different parameter names, so those names have to be a setting, not a hardcoded lookup.

## Solution

A Dynamo graph (Revit 2024, CPython3) reads rooms from the **loaded architectural link**, not from the host HVAC model, and writes raw rows to the «Выгрузка» sheet with `Data.ExportExcel`. One row is one room. Nothing is aggregated in Revit.

- If exactly one architectural link is loaded, it is chosen automatically. An optional name fragment narrows the choice when several links are loaded. Matching ignores case, `_` / space / `-` / `.`, and Cyrillic/Latin look-alike letters.
- The building code (`K` + two digits) is taken from the link name. The report is `Типовые_этажи_<code>.xlsx` in the reports folder. If the file does not exist, it is copied from `Типовые_этажи_ШАБЛОН.xlsx`. One building cannot overwrite another, and the folder does not have to be prepared per building. A missing link is written to `Типовые_этажи_ОШИБКА.xlsx`. Missing folder, missing template, and a workbook still open in Excel raise an explicit error.
- A later export replaces only «Выгрузка». Settings and calculated sheets stay and recalculate.
- The template is formulas only (Excel 2010+, no macros): «Настройки» (column map, room-name variants, whether apartment type is part of the signature), hidden calculation sheets, «Типовые этажи» (floor signature → group number, colour by group), «Таблица 1». References into «Выгрузка» go through `INDIRECT`, because ExportExcel deletes and recreates that sheet and a normal reference would become `#REF!`.

`KLWT_*` strings in the Python node are **example parameter names** from one client’s shared parameters. Change the НАСТРОЙКИ block for the next architectural model. Room-name variants for the three columns are edited on «Настройки» without touching formulas.

Limits of the template: 15 apartments per floor, 80 floors, 1500 apartments, 3000 export rows.

## Screenshot

«Типовые этажи» after a LibreOffice recalc of **sample rows** (not a project export). Unused apartment columns are omitted so the three made-up apartments and the group list sit together. Group colour is what the sheet is for. The verified project result is in the table below, not in this picture.

![Typical floors sheet with sample apartments. Floors that share a room signature share a colour. Labeled sample data.](../../docs/img/typical_floors_sample.png)

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested in Excel on one real building |
| Impact | 166 apartments, 16 floors (2–17), collapsed into 10 groups: 2 \| 3 \| 4 \| 5 \| 6 \| 7–12 \| 13–14 \| 15 \| 16 \| 17 |
| Writes model? | **No** — reads the architectural link; writes an `.xlsx` only |
| Repeats? | Every project; parameter names and room-name variants are settings |

## Safety

- The graph does not save, sync, or relinquish the open HVAC model. It reads rooms from the link document.
- Audit/apply rules for batch model writes do not apply here; there is no model write. See [docs/SAFETY.md](../../docs/SAFETY.md) for the rest of the portfolio.
- Do not point the reports folder at a live central or at a path you have not checked. The published graph ships with an empty folder input.
- Link failures go to a separate `…_ОШИБКА.xlsx`, not into a building report.

## Lessons

- The Dynamo **editor** does not re-run a Python node whose code did not change. Run from **Dynamo Player** (or change an input) so the export actually executes.
- `Data.ExportExcel` with overwrite recreates «Выгрузка» and breaks direct references (`#REF!`). The template uses `INDIRECT`.
- The folder picker cannot type a UNC path. Pin the share to Quick access, or replace the Directory input with a String node that holds the path (placeholder: `\\server\share\Project\Scripts\TypicalFloors`).
- Choosing a link by a short name fragment mixed up buildings (`…_01_…` vs `…_02_…`) when letters looked alike. The graph now auto-picks the single loaded architectural link, and a typed fragment is matched without case, separators, or Cyrillic/Latin look-alikes.

## How to run

1. Copy `src/` to a folder Excel and Revit 2024 can both see. Put `Типовые_этажи.dyn` and `Типовые_этажи_ШАБЛОН.xlsx` side by side.
2. Once per PC: Dynamo Player → set that folder, and set the «Папка отчётов» input to it. Leave «Связь АР» empty when the model has one architectural link.
3. Open the HVAC model (architectural link loaded) → Player → Run.
4. Open `Типовые_этажи_Kxx.xlsx`. Press F9 if Excel has not recalculated.

Step-by-step, including the UNC String-node workaround and the settings block: [src/INSTRUCTION.md](src/INSTRUCTION.md) · [src/INSTRUCTION.ru.md](src/INSTRUCTION.ru.md).

Rebuild the template (header only, no project rows): `python build.py empty Типовые_этажи_ШАБЛОН.xlsx` from `src/`.

Code and graph: [`src/`](src/)
