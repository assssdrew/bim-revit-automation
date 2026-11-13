# Case 06 — Model operations (rename, upgrade, relink, worksets)

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

A discipline park (~50+ workshared models) has to move together: new names, sometimes a new folder or Revit year, then every **RVT link** must point at the new files. Doing Save As / Detach / Load From by hand breaks the chain as soon as one host is saved before its links exist. New models also have to land on the right worksets.

## Solution

Framework **v3.1.0**. One operator window (`Операции с моделями.vbs`) opens `служебное/presets.ps1`. That script dot-sources `start_ops.ps1` and `job_lib.ps1` from the same `служебное` folder. `start_ops.ps1` finds Revit Batch Processor and starts each pass. [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is only the batch host.

The launcher sits in `src/`. All code sits in `src/служебное/`. Job files (`rvt_list.txt`, `job_paths.csv`, `mapping.csv`, …) are written next to the launcher when you run the window, and are gitignored.

Tabs cover rename / move, link retarget, year upgrade, **add links**, and **worksets**. A trial run prints the plan and does not open Revit. Pass 1 is `saveas_central.py` (**v3.1.0**, Create New Local, Detach OFF). An old Revit Server year is detached to a local buffer first (`detach_to_buffer.py`, Detach ON, this PC only). Pass 2 is `update_rvt_links.py`. Workset read/write is a separate pair: `read_worksets.py` then `apply_workset_links.py`.

`saveas_central.py` reports `SCRIPT_VERSION` 3.1.0. `detach_to_buffer.py` and `update_rvt_links.py` still report 3.0.0. `read_worksets.py` is 1.0.1 and `apply_workset_links.py` is 1.1.3.

The older menu `служебное/run_cascade.cmd` is still there.

UNC→UNC with a year upgrade does not need the buffer: a newer Revit can open the old file on the share and Save As.

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested on real working models |
| Impact | On average about 4–5× faster than manual work; staged Save As, move, relink, and worksets from one saved job |
| Scale | Discipline park, mixed UNC / `RSN://` |
| Operator UX | Single window. Older menu: `служебное/run_cascade.cmd` |
| Writes model? | **Yes** — Save As central, link remap, workset edits |

## Safety

- Live destination centrals: Create New Local, Detach **OFF**
- Detach **ON** only for the **local buffer** when leaving an old Revit Server (not the live central)
- Nested links are not rewritten from the host — each model that owns a link is opened
- Delete old files only after Pass 2 is green
- Inserting a link uses shared coordinates. A link whose coordinates do not match is skipped
- Name ops run only if the token is present

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC with the matching Revit years + RBP. Keep the `служебное` subfolder next to the launcher.
2. Edit `servers.cfg` (placeholders in this repo)
3. Double-click `Операции с моделями.vbs` (Russian for “Model operations”). It finds `presets.ps1` in the subfolder. The window will not start without `start_ops.ps1` and `job_lib.ps1` beside `presets.ps1`.
4. Pick models, set the rules, use the trial run, then start the tab.
5. Older menu, still in `служебное`: `run_cascade.cmd`.

## Tests

PowerShell and Python checks live in [`src/служебное/tests/`](src/служебное/tests/). They cover name ops, paths, the move tab, relink, worksets, and the window logic. They do not need Revit. On Windows, run the `.ps1` tests with Windows PowerShell 5.1.

## Third-party

- Needs [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) (GPL-3.0, not included).
- `src/служебное/rbp_open_fail_patch/` holds two **modified copies of RBP scripts** (© 2020 Dan Rumery, BVN) that make RBP continue past failure / dialog prompts on Open. They are GPL-3.0 derivative works, not MIT; see [NOTICE](../../NOTICE) and `LICENSE-GPL-3.0.txt` next to them. `служебное/patch_rbp_open_failures.ps1` copies them into *your* RBP `Scripts` folder.

Example mapping (anonymised): [`src/mapping.example.csv`](src/mapping.example.csv)

Tool notes: [`src/README.md`](src/README.md) · operator notes: [`src/ИНСТРУКЦИЯ.txt`](src/ИНСТРУКЦИЯ.txt)

Code: [`src/`](src/)
