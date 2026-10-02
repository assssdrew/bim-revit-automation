# Case 05 — Compact save of workshared models

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

Workshared Revit centrals grow after months of Sync. The Autodesk control is **Compact** in Save / Synchronize with Central — not zip, not Purge Unused. Doing it by hand means: pick the right year of Revit, open the central (or a new local), Compact, Sync, Relinquish, close — then repeat across a folder or Revit Server.

## Solution

An operator window (`Сжатие.cmd` → WinForms). The person picks models (files / folders / `RSN://`), chooses a mode, presses Compact. The toolkit groups the list by Revit year and runs the API job; [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is only the batch host.

| Mode | When | What it does |
|------|------|----------------|
| **Fast** | Anytime; others may stay in the files | Create New Local, close worksets, `SynchronizeWithCentral` with **Compact = True**, Relinquish. Does not invalidate other people’s locals. |
| **Deep** | Monthly / quarterly | Everyone out first. Open the **central with Audit** (no new local, no Detach) → Save As central with Compact → Sync Compact. After the run, teammates recreate their locals. |

Excel report: size before / after, run summary, history vs previous runs.

## Project scale / context

| Item | Result |
|------|--------|
| Status | Used in production on one project; no automated tests |
| Operator UX | One window; mixed local / UNC / `RSN://` list; years run in sequence |
| Reported use | Live workshared / Revit Server models (author-reported) |
| Writes model? | **Yes** — Compact on the central (workshared) or Save Compact (plain `.rvt`) |

## Safety

- Compact is a **write**. Deep mode is a maintenance window: backup, warn the team, nobody Syncs during the run.
- Fast mode: Create New Local, Detach **OFF**.
- Deep mode: open central + Audit; never Detach a live central. Confirm “everyone is out” in the UI.
- Skip backup / dated copies (e.g. `RVT\Backup\…`) — their central path can be wrong. The code skips paths containing `Backup`, a dated folder, or a Russian “Резерв” folder name.
- This is not Purge Unused and not a zip of the `.rvt`.

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC that has Revit + RBP
2. Edit `servers.cfg` with your Revit Server hosts (placeholders in this repo)
3. Double-click `Сжатие.cmd` (Russian for “Compact”)
4. Add models → choose Fast or Deep → Compact
5. First Revit launch: if BatchRvt is blocked, use `служебное\разрешить_надстройку_BatchRvt.cmd`

Operator notes (RU): [`src/ИНСТРУКЦИЯ.txt`](src/ИНСТРУКЦИЯ.txt)

Code: [`src/`](src/)
