# BIM Revit Automation Portfolio

Operator-facing automation for Autodesk Revit: **compact workshared models**, rename / year-upgrade / relink a discipline park, weekly health audits, units / levels alignment, alerts on exchange folders.

I write the **task scripts + Windows toolkits** (pick models → run → coloured Excel). [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) (RBP) is one open-source batch host I use for Revit API jobs — not my product, and not every case needs it (see [Third-party](#third-party)).

| | |
|---|---|
| **Author** | [assssdrew](https://github.com/assssdrew) |
| **Stack** | Revit API · IronPython · PowerShell · WinForms · OpenXML · FTP / ntfy |
| **Focus** | Workshared / Revit Server (`RSN://`) pipelines, safety-first writes |
| **Language** | [Русский README](README.ru.md) |

> **Honesty note.** All cases were written for one real multi-discipline project and then sanitised for this public repo. There are **no automated tests and no CI**. Status labels below say only what is claimed about real use; no time-saving numbers are measured. There are **no screenshots yet**.

---

## Quick start (3 steps)

For the Revit API cases (01, 02, 03, 05, 06). Case 04 is a standalone PowerShell watcher — see its README.

1. **Prepare a PC** with Windows, PowerShell, the Revit year that matches your models, and [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor). Copy the case `src/` folder next to it and edit `servers.cfg` (placeholder Revit Server hosts in this repo).
2. **Build the model list** with the case's picker (`choose_models_path.cmd`; case 05 is `Сжатие.cmd`, case 06 is `run_cascade.cmd`).
3. **Run the task** in RBP (e.g. `health_check.py`) on a **disposable local copy first**, then read the report (CSV / coloured XLSX). Details and the exact RBP settings are in each case README.

---

## Cases

| # | Case | Status (what is actually claimed) | One-liner |
|---|------|-----------------------------------|-----------|
| 01 | [Project Units + RSN](cases/01-project-units/) | Used on real Revit Server models (author-reported); no automated tests | Batch set Length accuracy; open local / UNC / `RSN://`; Sync + Relinquish |
| 02 | [Health Check](cases/02-health-check/) (v1.4.0) | Used on one project; no automated tests; read-only | Model health audit → CSV + coloured XLSX (no Save/Sync) |
| 03 | [Levels & Grids](cases/03-levels-grids/) | **MVP — not piloted on a live project** | Cascade audit of Levels/Grids vs linked exemplar; Apply only on base files |
| 04 | [FTP model alerts](cases/04-ftp-model-alerts/) | Used on one project's FTP exchange; no automated tests | Poll shared FTP exchange folders → phone push (ntfy / Telegram) |
| 05 | [Compact save](cases/05-compact-save/) | Used in production on one project; no automated tests | Operator window: fast or deep Compact of workshared / `RSN://` centrals → size report |
| 06 | [Model ops](cases/06-model-ops/) (v3.0.0) | Used in production on one project; no automated tests | Presets UI: rename, Revit-year upgrade, RVT relink across a discipline park |

New cases are added as folders under `cases/` — see [docs/HOW_TO_ADD_CASE.md](docs/HOW_TO_ADD_CASE.md).
Also in `docs/`: a [draft Navisworks clash-tolerance matrix](docs/navisworks-clash-tolerances.md) (a discussion document, not code and not a company standard).

---

## Featured

**Compact save** ([case 05](cases/05-compact-save/)) is maintenance: shrink live centrals. The operator UI (`Сжатие.cmd`, Russian for “Compact”) has **fast** (Create New Local, Sync Compact; teammates may stay) and **deep** (monthly/quarterly; everyone out; open central with Audit).

**Model ops** ([case 06](cases/06-model-ops/)) moves a discipline park together: new names, optional folder / Revit year, then RVT links. Presets UI → Save job → Pass 1 Save As central → Pass 2 relink. Detach is only for a local buffer when leaving an old Revit Server.

---

## Problem → approach

**Weekly reality on a multi-discipline project:** architecture updates land on the server; base files (BF) and many discipline models must stay aligned on units, model health, levels and grids. Centrals also need periodic Compact. A park move (names / year / links) has to be staged so hosts are not saved before their links exist.

Manual open-check-fix does not scale. These toolkits:

1. Build a model list (folder / files / Revit Server)
2. Run the Revit API job (batch host or a dedicated window)
3. Emit operator-friendly **GREEN / YELLOW / RED** reports — or size before/after for Compact
4. Apply writes only where the risk is understood (units; BF levels/grids; Compact with an explicit mode; model-ops after Save job) — never blind coordinate fixes

---

## Project scale / context

Context of the one project these were written for, as described by the author. These are **not measured results**.

| Item | Context |
|------|---------|
| Weekly cascade | ~5–6 base files (BF↔AR), then up to ~80 discipline models (vs BF) |
| Model ops | A park of ~50+ models, mixed UNC / `RSN://` |
| Units / RSN | Run on real Revit Server models (Sync + Relinquish worked) |
| Compact save | Fast mode any time, deep mode in a maintenance window; Excel reports size before → after |
| Health Check | Read-only batch audit of workshared models |
| Levels & Grids | MVP coded; **not yet piloted on a live project** |
| FTP model alerts | Used for phone push when exchange folders change |

Time saved is **not measured**. The design goal is one report session instead of opening dozens of models by hand.

---

## Safety rules

Automation that can destroy a federated model is worse than no automation.

- **Audit** scripts never Save / Sync / Relinquish
- **Apply** is a separate script / explicit list (e.g. BF only)
- **Compact** is a write: fast uses Create New Local; deep opens the central only after the team is out
- **Model ops** writes only after Save job; Detach is the local upgrade buffer, not the live central
- Coordinates, PBP, Survey Point, True North, Shared Coordinates are **out of scope** here — report only (Health Check), never auto-fix
- Tolerances are mandatory to avoid false RED on floating-point noise

Details: [docs/SAFETY.md](docs/SAFETY.md)

---

## Scope / not covered

What this repository is **not**, so nobody has to guess:

- **No C# / .NET Revit add-in.** Everything is IronPython task scripts + PowerShell / WinForms.
- **No Dynamo** graphs or Dynamo scripts.
- **No AI / LLM / machine-learning** components.
- **No automated tests, no CI.** Verification was manual (see per-case status).
- **No screenshots or demo recordings yet.**
- **No Navisworks automation** (only a draft tolerance matrix document) and no Revit add-in UI inside Revit — operator windows are external WinForms / `.cmd`.
- Not a general BIM framework: the cases target one workflow (workshared / Revit Server park maintenance and audits).
- Some file names and code comments in cases 01, 05 and 06 are in Russian (e.g. `Сжатие.cmd`, `служебное/`); translating them is on the to-do list.

---

## Third-party

- **[Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor)** (Dan Rumery, BVN) is the batch host for the RBP-based cases. It is **not** included in this repository; install it separately. RBP is licensed under **GPL-3.0**.
- `cases/06-model-ops/src/rbp_open_fail_patch/` contains **two modified copies of RBP scripts** (`revit_dialog_util.py`, `revit_failure_handling.py`) with the original “Copyright (c) 2020 Dan Rumery, BVN” notice. They are **derivative works of RBP and are covered by GPL-3.0**, not by this repository's MIT licence.
- `cases/01-project-units/src/rbp_rsn_patch/` (and its copy in case 05) is a PowerShell script that edits **your installed** RBP `Scripts` folder to accept `RSN://` paths (it creates `*.bak_before_rsn` backups); it does not ship RBP code.

See [NOTICE](NOTICE) for details.

---

## Repository layout

```text
bim-revit-automation/
  README.md / README.ru.md     ← portfolio hub
  NOTICE                       ← third-party notices (RBP, GPL-3.0 patch files)
  cases/
    01-project-units/src/
    02-health-check/src/
    03-levels-grids/src/
    04-ftp-model-alerts/src/   ← FTP poller + phone alerts (no Revit API)
    05-compact-save/src/       ← operator UI + Compact task
    06-model-ops/src/          ← presets UI + Save As / relink pipeline
  samples/                     ← anonymised report snippets
  docs/
```

Copy a case `src/` onto a PC with the matching Revit year. Case 04 runs standalone on Windows (Task Scheduler). Case 05 starts from `Сжатие.cmd`. Case 06 starts from `run_cascade.cmd`.

---

## Requirements

- Autodesk Revit (year matching the models; units toolkit targets 2022+)
- Windows + PowerShell
- [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) for the Revit API batch cases (01–03, 05–06)

---

## Disclaimer

Scripts are provided as portfolio / starting points. Validate on a disposable local before any production write. Project paths, names and server hosts are sanitised or replaced with placeholders in this public repo.
