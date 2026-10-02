# Case 05 — Compact save of workshared models

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

Workshared Revit centrals grow after months of Sync. The Autodesk control is **Compact** in Save / Synchronize with Central — not zip, not Purge Unused. Doing it by hand means: pick the right year of Revit, open the central (or a new local), Compact, Sync, Relinquish, close — then repeat across a folder or Revit Server.

## Solution

An operator window (`Сжатие.vbs` → WinForms). The person picks models (files / folders / `RSN://`), chooses a mode, presses Compact. The toolkit groups the list by Revit year and runs the API job; [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is only the batch host.

| Mode | When | What it does |
|------|------|----------------|
| **Fast** | Anytime; others may stay in the files | Create New Local, close worksets, `SynchronizeWithCentral` with **Compact = True**, Relinquish. |
| **Deep** | Monthly / quarterly | Open the **central with Audit** (no new local, no Detach, worksets closed) → Save As the **same** central with Compact → Sync Compact. |

Excel report (sheets «Модели», «Сводка», «Динамика»): size before / after, run summary, history vs previous runs. Savings vary. On live runs they were often very small (for example 61.03 MB → 61.01 MB); some models shrink more (up to about −11.6% on one early fast-mode model).

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested on real working models; RSN deep mode on a limited number of runs |
| Impact | On average about 4–5× faster than manual work; long lists (e.g. ~125 models) in one unattended run; file-size change varies (often small, e.g. 61.03→61.01 MB; up to about −11.6% on one early fast-mode model) |
| Operator UX | One window; mixed local / UNC / `RSN://` list; years run in sequence |
| Writes model? | **Yes** — Compact on the central (workshared) or Save Compact (plain `.rvt`) |

## Screenshots

Names, paths and server addresses in these pictures are blurred.

![Compact models window with a list of models. Names are blurred.](../../docs/img/compact_window.jpg)

Operator window: model list, source (network or Revit Server), deep mode, none checked.

![Excel sheet Models from a deep compact run. Two models, both OK. Names and paths are blurred.](../../docs/img/compact_report_models.jpg)

Excel sheet «Модели», deep mode, 2026-09-16, two models, both OK.

![Excel sheet Summary from a compact report. Size went from 61.03 MB to 61.01 MB.](../../docs/img/compact_report_summary.jpg)

Excel sheet «Сводка»: 61.03 MB → 61.01 MB, delta −0.02 MB, longest model 15 s.

## Safety

- Compact is a **write**. Deep mode is a maintenance window: backup, warn the team, nobody Syncs during the run.
- Fast mode: Create New Local, Detach **OFF**.
- Deep mode: open central + Audit; never Detach a live central. The window asks you to continue before it opens the central and saves it again. It does not check that other users have closed their local files.
- RSN deep mode needs the v4 RBP patch (`служебное/rbp_rsn_patch`, marker `RSN_PATCH_PS_v4`). The window tries to apply it to the local RBP install. Stock RBP detaches. RSN deep mode was tested on a limited number of runs.
- Skip backup / dated copies (e.g. `RVT\Backup\…`) — their central path can be wrong. The code skips paths containing `Backup`, a dated folder, or a Russian “Резерв” folder name.
- This is not Purge Unused and not a zip of the `.rvt`.

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC that has Revit + RBP
2. Edit `servers.cfg` with your Revit Server hosts (placeholders in this repo)
3. Double-click `Сжатие.vbs` (Russian for “Compact”)
4. Add models → choose Fast or Deep → Compact
5. First Revit launch: if BatchRvt is blocked, use `служебное\разрешить_надстройку_BatchRvt.cmd`

`.ps1` files in this case are **UTF-8 with BOM on purpose**. Windows PowerShell 5.1 needs the BOM to read Cyrillic correctly.

Operator notes (RU): [`src/ИНСТРУКЦИЯ.txt`](src/ИНСТРУКЦИЯ.txt)

Code: [`src/`](src/)
