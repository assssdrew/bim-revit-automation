# Case 06 — Model operations (rename, upgrade, relink)

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

A discipline park (~50+ workshared models) has to move together: new names, sometimes a new folder or Revit year, then every **RVT link** must point at the new files. Doing Save As / Detach / Load From by hand breaks the chain as soon as one host is saved before its links exist.

## Solution

A single operator window (`Операции с моделями.vbs`, or `show_ops.vbs` with an ASCII file name) opens `presets.ps1`. That script **depends on `start_ops.ps1` and `job_lib.ps1`**: it dot-sources both, and they must sit in the same folder. `start_ops.ps1` finds Revit Batch Processor and starts each pass. [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is only the batch host.

The window lists models (Model / State / Source / Path). A “will become” column appears only when renaming, and old/new path columns only when moving files. **Details** edits the name rules. **Run** saves the job and starts the passes (buffer / Save As / file move / links / delete old). Before an overwrite it asks whether everyone has left the models.

The older menu `run_cascade.cmd` is still in the folder.

1. Pick models (share / files / `RSN://`)
2. Name ops from a sample (`delete` / `replace` / `insert`) + destination / year
3. **Save job** writes `mapping.csv`, `job_paths.csv`, lists — **before** Revit opens anything
4. Pipeline:
   - **Pass 1** `saveas_central.py` — new centrals (Create New Local, Detach **OFF**)
   - If the source is an old Revit Server year: `detach_to_buffer.py` first (**Detach ON**, buffer on this PC only — not a live central)
   - **Pass 2** `update_rvt_links.py` — remap links from `job_paths.csv`
5. After a green Pass 2: delete old UNC / list old RSN for Admin

UNC→UNC with a year upgrade does not need the buffer: a newer Revit can open the old file on the share and Save As.

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested on real working models |
| Impact | On average about 4–5× faster than manual work; staged Save As, move, and relink from one saved job plan |
| Scale | Discipline park, mixed UNC / `RSN://` |
| Operator UX | Single window: list, optional rename/move columns, Details, Run. Older menu: `run_cascade.cmd` |
| Writes model? | **Yes** — Save As central, then link remap |

## Safety

- Live destination centrals: Create New Local, Detach **OFF**
- Detach **ON** only for the **local buffer** when leaving an old Revit Server (not the live central)
- Nested links are not rewritten from the host — each model that owns a link is opened
- Delete old files only after Pass 2 is green
- Name ops run only if the token is present (exceptions for one specific token are a separate replace)

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC with the matching Revit years + RBP
2. Edit `servers.cfg` (placeholders in this repo)
3. Double-click `Операции с моделями.vbs` (Russian for “Model operations”). `show_ops.vbs` opens the same window. `presets.ps1` will not start without `start_ops.ps1` and `job_lib.ps1` beside it.
4. Pick models, set the name rules, press Run.
5. Older menu, still in the folder: `run_cascade.cmd` → **2** pick models → **1** presets → Save job. *Save job* (`job_lib.ps1`) writes `rbp_checklist.txt` next to `mapping.csv` (not stored in the repo): Revit year, Detach / Create New Local, task script and list file, one RBP run per year. `run_cascade.cmd` prints it. Optional: Windows Move UNC (**3**), then delete old (**4**).

## Third-party

- Needs [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) (GPL-3.0, not included).
- `src/rbp_open_fail_patch/` holds two **modified copies of RBP scripts** (© 2020 Dan Rumery, BVN) that make RBP continue past failure / dialog prompts on Open. They are GPL-3.0 derivative works, not MIT; see [NOTICE](../../NOTICE) and `LICENSE-GPL-3.0.txt` next to them. `patch_rbp_open_failures.ps1` copies them into *your* RBP `Scripts` folder.

Example mapping (anonymised): [`src/mapping.example.csv`](src/mapping.example.csv)

Code: [`src/`](src/)
