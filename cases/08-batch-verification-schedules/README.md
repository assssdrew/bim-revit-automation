# Case 08 — Batch verification schedules

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

A discipline model should carry the same verification schedules as the project template: empty values, classifier codes, system names, and the other checks the BIM manager marked for that project. Copying those schedules by hand, then collecting the results into one workbook, does not scale across a model list. The workbook has to be written without Excel installed, and the same tables sometimes need to land in Google Sheets.

## Solution

An operator window (`Проверочные спецификации.vbs`, Russian for “Verification schedules”) plans the run. Code lives under `служебное/`. The window:

- reads a template model and a project distribution (which schedules belong to which section)
- exports schedule data with [`служебное/скрипты/export_schedules.py`](src/служебное/скрипты/export_schedules.py) (read-only)
- writes `.xlsx` with [DocumentFormat.OpenXml](src/служебное/lib/LICENSE-MIT.txt) — Excel is not required
- can POST the same tables to a Google Sheet through [`служебное/google_push/`](src/служебное/google_push/) (Apps Script `Code.gs` plus `push_google.ps1`)
- can insert the template schedules into the host with [`служебное/скрипты/insert_schedules.py`](src/служебное/скрипты/insert_schedules.py) and set assembly codes with [`служебное/скрипты/set_assembly_code.py`](src/служебное/скрипты/set_assembly_code.py)

Section codes (ОВ, ВК, ЭОМ, …) are in [`служебное/скрипты/разделы.txt`](src/служебное/скрипты/разделы.txt). Project names, model prefixes, sheet URLs, and the classifier path are filled in the window. In this repository the sample project key is `PROJ` and the classifier path is `\\server\share\BIM\classifier.txt`.

## Inputs and outputs

| | |
|--|--|
| Inputs | Template `.rvt`, model list (disk, UNC, or `RSN://`), `servers.cfg`, optional `служебное/google_push/google_push.cfg` copied from the example |
| Outputs | `.xlsx` workbook, text ledgers under `отчеты/` (created at runtime, not in git), optional Google Sheet update |
| Writes model? | **Export: no.** Insert schedules and the classifier pass **yes** (transactions; workshared models sync) |

## Project scale / context

| Item | Result |
|------|--------|
| Status | Tested on real working models |
| Impact | One window for the template, the model list, the xlsx, and the optional Google Sheet |
| Excel | Not required to build the workbook |
| Google Sheets | Optional. The web-app URL and secret are not in this repo |

## Safety

- Run insert and the classifier pass on a disposable copy first. Insert deletes and recreates verification schedules, then syncs a workshared model.
- Export does not Save, Sync, or Relinquish.
- `google_push.cfg` holds a URL and a shared secret. Only `google_push.cfg.example` is published; both fields are empty. `Code.gs` ships with `PUSH_SECRET = 'REPLACE_ME_WITH_LONG_RANDOM_SECRET'`.
- Do not put a real spreadsheet id, web-app URL, or secret into this public tree.
- Revit Server hosts in `servers.cfg` are placeholders (`revit-server-20xx.example.local`).

Details: [docs/SAFETY.md](../../docs/SAFETY.md)

## How to run

1. Copy `src/` to a Windows PC with the matching Revit year and [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor). `служебное/разработка/` is tests and notes; the window does not need it.
2. Edit `servers.cfg`.
3. For Google Sheets, copy `служебное/google_push/google_push.cfg.example` to `google_push.cfg` on that PC and fill the URL and secret. Deploy `Code.gs` as a web app. Leave the secret out of git.
4. Double-click `Проверочные спецификации.vbs`. It finds `ui_schedules.ps1` under the service folder.
5. Pick the template, the models, and the sections. Export writes the workbook. Insert and the classifier are separate actions.

`служебное/rbp_rsn_patch/apply_rsn_support.ps1` is the same kind of local RBP patch as case 01: it edits your installed RBP so `RSN://` paths pass `File.Exists`. It does not ship RBP source.

Operator notes (RU): [`src/ИНСТРУКЦИЯ.txt`](src/ИНСТРУКЦИЯ.txt)

## Tests

[`src/служебное/разработка/tests/`](src/служебное/разработка/tests/) — Python autotests for the chain, the xlsx layout, section codes, the Google payload, and the window logic. Many are skipped unless `pwsh` is on `PATH`. They do not need Revit.

The published tests do not include client schedule-filter dumps or a client template catalog. Two tests skip when those dumps are absent. Grouping is covered with a short synthetic list in `test_dist_groups.py`. A 30-row stand-in catalog is [`src/служебное/разработка/данные/samplesite_template_specs.tsv`](src/служебное/разработка/данные/samplesite_template_specs.tsv).

Code: [`src/`](src/)
