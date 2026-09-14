# BIM Revit Automation Portfolio

Operator-facing automation for Autodesk Revit: **compact workshared models**, rename / year-upgrade / relink a discipline park, weekly health audits, units / levels alignment, alerts on exchange folders.

I write the **task scripts + Windows toolkits** (pick models → run → coloured Excel). [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) is one open-source batch host I use for Revit API jobs — not the product, and not every case needs it.

| | |
|---|---|
| **Author** | [assssdrew](https://github.com/assssdrew) |
| **Stack** | Revit API · IronPython · PowerShell · WinForms · OpenXML · FTP / ntfy |
| **Focus** | Workshared / Revit Server (`RSN://`) pipelines, safety-first writes |
| **Language** | [Русский README](README.ru.md) |

---

## Cases (published as they ship)

| # | Case | Status | One-liner |
|---|------|--------|-----------|
| 01 | [Project Units + RSN](cases/01-project-units/) | **Production-ready** | Batch set Length accuracy; open local / UNC / `RSN://`; Sync + Relinquish |
| 02 | [Health Check](cases/02-health-check/) | **Production-ready** (v1.4.0) | Read-only model health audit → CSV + coloured XLSX (no Save/Sync) |
| 03 | [Levels & Grids](cases/03-levels-grids/) | **MVP ready** | Cascade audit of Levels/Grids vs linked exemplar; Apply only on base files |
| 04 | [FTP model alerts](cases/04-ftp-model-alerts/) | **Production-ready** | Poll shared FTP exchange folders → phone push (ntfy / Telegram) |
| 05 | [Compact save](cases/05-compact-save/) | **Production-ready** | Operator window: fast or deep Compact of workshared / `RSN://` centrals → size report |
| 06 | [Model ops](cases/06-model-ops/) | **Production-ready** (v3.0.0) | Presets UI: rename, Revit-year upgrade, RVT relink across a discipline park |

New cases are added as folders under `cases/` — see [docs/HOW_TO_ADD_CASE.md](docs/HOW_TO_ADD_CASE.md).

---

## Featured

**Compact save** ([case 05](cases/05-compact-save/)) is maintenance: shrink live centrals. The operator UI (`Сжатие.cmd`) has **fast** (Create New Local, Sync Compact; teammates may stay) and **deep** (monthly/quarterly; everyone out; open central with Audit).

**Model ops** ([case 06](cases/06-model-ops/)) moves a discipline park together: new names, optional folder / Revit year, then RVT links. Presets UI → Save job → Pass 1 Save As central → Pass 2 relink. Detach is only for a local buffer when leaving an old Revit Server.

---

## Problem → approach

**Weekly reality on a multi-discipline project:** architecture updates land on the server; base files (BF) and ~80 discipline models must stay aligned on units, model health, levels and grids. Centrals also need periodic Compact. A park move (names / year / links) has to be staged so hosts are not saved before their links exist.

Manual open-check-fix does not scale. These toolkits:

1. Build a model list (folder / files / Revit Server)
2. Run the Revit API job (batch host or a dedicated window)
3. Emit operator-friendly **GREEN / YELLOW / RED** reports — or size before/after for Compact
4. Apply writes only where the risk is understood (units; BF levels/grids; Compact with an explicit mode; model-ops after Save job) — never blind coordinate fixes

---

## Impact (project scale)

| Metric | Value |
|--------|-------|
| Weekly cascade | ~5–6 base files (BF↔AR), then up to ~80 discipline models (vs BF) |
| Compact save | Production window: fast anytime / deep on a maintenance slot; Excel size before→after |
| Model ops | Production: rename / year upgrade / relink a ~50+ model park without breaking the link chain |
| Units / RSN | Confirmed on real Revit Server models (Sync + Relinquish OK) |
| Health Check | Batch read-only audit across workshared models |
| Levels & Grids | MVP coded; not yet piloted on a live project |
| FTP model alerts | Field-tested phone push when exchange folders change (no FileZilla babysitting) |

Exact hour-savings vary by project; the design goal is: **one report session instead of opening dozens of models by hand**.

---

## Safety narrative (why employers care)

Automation that can destroy a federated model is worse than no automation.

- **Audit** scripts never Save / Sync / Relinquish
- **Apply** is a separate script / explicit list (e.g. BF only)
- **Compact** is a write: fast uses Create New Local; deep opens the central only after the team is out
- **Model ops** writes only after Save job; Detach is the local upgrade buffer, not the live central
- Coordinates, PBP, Survey Point, True North, Shared Coordinates are **out of scope** here — report only (Health Check), never auto-fix
- Tolerances are mandatory to avoid false RED on floating-point noise

Details: [docs/SAFETY.md](docs/SAFETY.md)

---

## Repository layout

```text
bim-revit-automation/
  README.md / README.ru.md     ← portfolio hub
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

Copy a case `src/` onto a PC with the matching Revit year. Case 04 runs standalone on Windows (Task Scheduler). Case 05 is `Сжатие.cmd`. Case 06 is `run_cascade.cmd`.

---

## Requirements

- Autodesk Revit (year matching the models; units toolkit targets 2022+)
- Windows + PowerShell
- [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) for Revit API batch cases (01–03, 05–06)

---

## Disclaimer

Scripts are provided as portfolio / starting points. Validate on a disposable local before any production write. Company paths and server names are sanitised in this public repo.
