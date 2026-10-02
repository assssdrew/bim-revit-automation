# Case 06 — Model park ops (rename, upgrade, relink)

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

A discipline park (~50+ workshared models) has to move together: new names, sometimes a new folder or Revit year, then every **RVT link** must point at the new files. Doing Save As / Detach / Load From by hand breaks the chain as soon as one host is saved before its links exist.

## Solution

An operator presets window (`run_cascade.cmd` → **1**) builds the job, then staged API tasks run it. [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is only the batch host.

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
| Status | v3.0.0 — Used in production on one project; no automated tests |
| Scale | Discipline park, mixed UNC / `RSN://` |
| Operator UX | One presets UI; preview of new names before write |
| Writes model? | **Yes** — Save As central, then link remap |

## Safety

- Live destination centrals: Create New Local, Detach **OFF**
- Detach **ON** only for the **local buffer** when leaving an old Revit Server (not the production central)
- Nested links are not rewritten from the host — each model that owns a link is opened
- Delete old files only after Pass 2 is green
- Name ops run only if the token is present (exceptions for one specific token are a separate replace)

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a PC with the matching Revit years + RBP
2. Edit `servers.cfg` (placeholders in this repo)
3. `run_cascade.cmd` → **2** pick models → **1** presets → Save job
4. Follow `rbp_checklist.txt` (one RBP run = one Revit year). This file is **generated** by *Save job* (`job_lib.ps1`) next to `mapping.csv`; it is not stored in the repo. It lists, in order, which RBP run to start: Revit year, Detach / Create New Local setting, task script and list file. `run_cascade.cmd` prints it.
5. Optional: Windows Move UNC (**3**), then delete old (**4**)

## Third-party

- Needs [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) (GPL-3.0, not included).
- `src/rbp_open_fail_patch/` holds two **modified copies of RBP scripts** (© 2020 Dan Rumery, BVN) that make RBP continue past failure / dialog prompts on Open. They are GPL-3.0 derivative works, not MIT; see [NOTICE](../../NOTICE) and `LICENSE-GPL-3.0.txt` next to them. `patch_rbp_open_failures.ps1` copies them into *your* RBP `Scripts` folder.

Example mapping (anonymised): [`src/mapping.example.csv`](src/mapping.example.csv)

Code: [`src/`](src/)
