# Case 03 — Levels & Grids Audit / Apply

[← Portfolio hub](../../README.md) · [RU](README.ru.md) · [Sample CSV](../../samples/levels_grids_details_sample.csv)

## Problem

After each Architecture update, base files (BF) and then ~80 discipline models must match the exemplar on **levels and grids** (names, elevations, grid position/angle, duplicates, missing/extra). Manual Copy/Monitor review does not scale.

## Solution

Cascade toolkit for RBP:

1. **Audit BF ↔ AR** (exemplar = linked AR) → CSV / coloured XLSX  
2. Review → **Apply on BF only** (elevation / grid move-rotate when name match + cfg flags) → Sync  
3. **Audit disciplines ↔ BF** → report (mass Apply on 80 models is *not* MVP)

Compares via Revit API (not auto-Accept Coordination Review UI).

## Impact

| Item | Result |
|------|--------|
| Status | **MVP ready** (not yet piloted live) |
| Weekly design load | ~5–6 BF, then up to ~80 disciplines |
| Default tolerances | Level 2 mm · Grid 5 mm · Angle 0.1° |

## Safety (interview highlight)

| In scope | Out of scope |
|----------|--------------|
| Names, elevations, grid pose, duplicates, missing/extra, traffic light | Internal Origin, PBP/Survey, True North, Shared Coordinates, Site |
| Apply: BF list only; no delete of extras in MVP | Blind auto-fix of any coordination |

Logic: until BF = AR, running 80 disciplines is wasted effort.

## How to run (Audit)

1. Copy `src/` to RBP Scripts  
2. `choose_models_path.cmd` → hosts (BF or disciplines)  
3. `reset_report_session.cmd` / `choose_exemplar_link` → exemplar model  
4. RBP: `levels_grids_audit.py` (Create New Local, Detach OFF)

Apply: put **only BF** in the list → `levels_grids_apply.py`.

Code: [`src/`](src/)
