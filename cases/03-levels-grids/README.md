# Case 03 — Levels & Grids Audit / Apply

[← Portfolio hub](../../README.md) · [RU](README.ru.md) · [Sample CSV](../../samples/levels_grids_details_sample.csv)

## Problem

After each Architecture update, base files (BF) and then discipline models must match the exemplar on **levels and grids** (names, elevations, grid position/angle, duplicates, missing/extra). Manual Copy/Monitor review does not scale.

## Solution

Cascade toolkit for RBP. **Audit v1.21.0** reports only. **Apply v1.23.0** writes the rows the operator checks in the fix-plan window.

1. **Audit BF ↔ AR** (exemplar = linked AR) → CSV / Excel
2. **Fix BF from the report** — pick models and issues, then Apply
3. **Audit sections ↔ BF**
4. **Fix sections from that report** — same window, section models only

The menu is [`src/run_cascade.cmd`](src/run_cascade.cmd). Comparison uses the Revit API (not the coordination-review dialog). Geometry is transformed into host coordinates. Tolerances live in `tolerances.cfg`.

The fix-plan window (`choose_fix_plan.ps1`) lists each reported issue. Filters: Levels / Grids / RED only. One checkbox controls Sync after the run. Project Base Point X/Y/Z (millimetres; Z minus is down) is written only for the models you tick, and a zero shift is not written.

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested on real working models |
| Impact | On average about 4–5× faster than manual work; cascade BF↔AR and sections↔BF in report sessions |
| Weekly design load | ~5–6 BF, then up to ~80 disciplines |
| Default tolerances | Level 2 mm · Grid 5 mm · Angle 0.1° |
| Versions | Audit `SCRIPT_VERSION` **1.21.0** · Apply `SCRIPT_VERSION` **1.23.0** |

## Safety

| In scope | Out of scope |
|----------|--------------|
| Names, elevations, grid pose, duplicates, missing/extra, traffic light | Internal Origin, Survey Point, True North, Shared Coordinates, Site |
| Apply: only rows checked in the fix plan. Optional Project Base Point shift on those models | Creating or deleting levels and grids. Blind auto-fix |

Audit never Save / Sync / Relinquish. Apply syncs only when the fix-plan checkbox is on.

Logic: until BF matches AR, running the discipline models is wasted effort.

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC with the matching Revit year and RBP
2. Edit `servers.cfg` (placeholder hosts in this repo)
3. `run_cascade.cmd` → **1** audit BF, or `choose_models_path.cmd` and then RBP on `levels_grids_audit.py` (Create New Local, Detach OFF)
4. After the report: menu **2** (BF) or **4** (sections) opens the fix plan, writes `fix_plan.csv` and `rvt_list.txt`, then run `levels_grids_apply.py`
5. Optional rebuild of the Excel file: `color_latest_report.cmd`

`fix_plan.csv` and `last_selection.cfg` are local run files and are not stored in this repo.

Operator notes: [`src/README.md`](src/README.md)

Geometry checks without Revit: [`src/служебное/tests/test_geometry_diff.py`](src/служебное/tests/test_geometry_diff.py)

Code: [`src/`](src/)
