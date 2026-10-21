# Case 07 — Apartment schedule export

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

On a residential project with several buildings, the HVAC engineer needs identical floors: floors whose apartments carry the same set of rooms — combined bathroom (СУ), WC (уборная), walk-in closet (гардеробная). The deliverable is his «Таблица 1»: Floor | No. | Type | Rooms | Total area | Living area | Bathroom | WC | Closet.

Revit only supplies the raw room rows. Every total, group number, and the table itself is calculated in Excel, so a colleague can see the formula and change it. The same task comes back on the next project, with a different architectural team and different parameter names, so those names have to be a setting, not a hardcoded lookup.

The graph is «Выгрузка квартирографии» (Apartment Schedule Export).

## Solution

A Dynamo graph (Revit 2024, CPython3) reads rooms from the **loaded architectural link**, not from the host HVAC model, and writes raw rows to the «Выгрузка» sheet with `Data.ExportExcel`. One row is one room. Nothing is aggregated in Revit.

- One copy of the graph and the template serves every project. They live in a shared office folder. In this repository that path is the placeholder `\\server\share\BIM\Dynamo\Выгрузка квартирографии`. `SCRIPT_DIR` in the Python node points there, and `Квартирография_ШАБЛОН.xlsx` is read from that folder.
- Reports are written beside the open HVAC model, in a `Квартирография` folder. For a workshared model the folder is created next to the **central file**. A model that is still unsaved, a cloud model (BIM 360 / ACC), or a Revit Server model (`RSN://`) has no such folder, and the graph raises an explicit error. There is no Directory input and no per-project reports folder to pick.
- If exactly one architectural link is loaded, it is chosen automatically. The only Player input, «Связь АР», is an optional name fragment for when several links are loaded. Matching ignores case, `_` / space / `-` / `.`, and Cyrillic/Latin look-alike letters.
- The building code (`K` + two digits) is taken from the link name. The report is `Квартирография_<code>.xlsx`. If the file does not exist, it is copied from the template in the script folder. One building cannot overwrite another. A missing link is written to `Квартирография_ОШИБКА.xlsx`. A missing template and a workbook still open in Excel raise an explicit error.
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
| Writes model? | **No** — reads the architectural link; writes an `.xlsx` beside the HVAC model (beside the central file when the model is workshared) |
| Repeats? | Every project, from one shared script folder; parameter names and room-name variants are settings |

## Safety

- The graph does not save, sync, or relinquish the open HVAC model. It reads rooms from the link document.
- It creates `Квартирография` and the building workbook next to that model. On a workshared model that write sits next to the central file, so the account needs permission there.
- Audit/apply rules for batch model writes do not apply here; there is no model write. See [docs/SAFETY.md](../../docs/SAFETY.md) for the rest of the portfolio.
- The published graph ships with the placeholder `SCRIPT_DIR` above. Point it at the shared script folder you actually use before a real run.
- Link failures go to a separate `…_ОШИБКА.xlsx`, not into a building report.

## Lessons

- The Dynamo **editor** does not re-run a Python node whose code did not change. Run from **Dynamo Player** (or change an input) so the export actually executes.
- `Data.ExportExcel` with overwrite recreates «Выгрузка» and breaks direct references (`#REF!`). The template uses `INDIRECT`.
- The reports-folder Directory input is gone. The script folder is `SCRIPT_DIR`; each run creates `Квартирография` next to the model. Pin the shared folder to Quick access so Dynamo Player can open the graph.
- Choosing a link by a short name fragment mixed up buildings (`…_01_…` vs `…_02_…`) when letters looked alike. The graph auto-picks the single loaded architectural link, and a typed fragment is matched without case, separators, or Cyrillic/Latin look-alikes.
- A workshared local copy is the wrong neighbour for the report: the folder is created next to the central. Cloud and Revit Server models stop with an explicit error.

## How to run

1. Put `Выгрузка_квартирографии.dyn` and `Квартирография_ШАБЛОН.xlsx` in the shared script folder. Set `SCRIPT_DIR` in the Python node to that folder (the repository copy uses `\\server\share\BIM\Dynamo\Выгрузка квартирографии`).
2. Once per PC: Dynamo Player → set that folder. Leave «Связь АР» empty when the model has one architectural link.
3. Open the HVAC model (architectural link loaded, model saved on disk or a file share) → Player → Run.
4. Open `Квартирография` next to the model, or next to the central file, and then `Квартирография_Kxx.xlsx`. Press F9 if Excel has not recalculated.

Step-by-step, including the settings block: [src/INSTRUCTION.md](src/INSTRUCTION.md) · [src/INSTRUCTION.ru.md](src/INSTRUCTION.ru.md).

Rebuild the template (header only, no project rows): `python build.py empty Квартирография_ШАБЛОН.xlsx` from `src/`. Rebuild the Word instruction after editing `mk.py`: `python mk.py`.

Code and graph: [`src/`](src/)
