# Changelog

All notable changes to this portfolio are recorded here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This repository is not one versioned package. Each case keeps its own script version, so the history below is grouped by date and, where a change belongs to one case, by case. Entries come from the git history and merged pull requests #1–#9. They do not add claims the commits do not support.

## [Unreleased]

### Added

- **Case 08.** Batch verification schedules (`Проверочные спецификации`). Operator window, OpenXML workbook (no Excel), optional Google Sheets push, and autotests. Sample project key `PROJ`. Classifier path, Revit Server hosts, sheet URL, and push secret in the published tree are placeholders. Client filter dumps and the client template catalog are not included.

### Changed

- **Case 02.** Health check **v1.4.0 → v1.5.4** (`SCRIPT_VERSION` in `health_check.py`; last commit of that file on the working repo is 2026-08-31). The Excel report adds a Metrics sheet. Link rows include status, IFC, and nested RVT links.
- **Case 03.** Audit **v1.0.2 → v1.21.0** (file last updated 2026-11-06). Apply **v1.0.2 → v1.23.0**. Fix-plan window for base files and for sections. Project Base Point X/Y/Z moves only when the operator checks those axes. Survey Point and True North stay untouched. Apply does not create or delete levels or grids.
- **Case 05.** `compact_save.py` from the working repo's default branch (last commit of that file 2026-10-26, about 86 KB; the script has no `SCRIPT_VERSION`). Report sheets are Models, Summary, and Projects. The Dynamics sheet is not written. Optional parallel run, one BatchRvt per year. `служебное/naming.ps1` and `служебное/CLI.txt` name the exports. Revit Server 2025 is the placeholder `revit-server-2025.example.local`. The RSN deep-compact patch header stays **4.0** (`RSN_PATCH_PS_v4`).
- **Case 06.** Framework **v3.0.0 → v3.1.0**. `saveas_central.py` reports 3.1.0. Code moved under `служебное/`. Workset read/write passes and autotests are included. The single launcher is `Операции с моделями.vbs`.

### Added

- This changelog.

## [2026-11-13]

### Changed

- **Case 07** (#9). Apartment schedule export v7. One building per run. Python writes the workbook with the standard library, areas go to hundredths, and the report adds common-area sheets. Errors stay in the Player result text. The published path and the example link name are placeholders.

## [2026-10-29]

### Changed

- **Docs** (#8). The cases table and the project-scale section are one table. One health-check screenshot stays at the top. Compact scale leads with the ~125-model unattended run. Both hub READMEs share a Russian filename and UI glossary and the same positioning line.
- **Case 07** (#7). The case folder is `07-apartment-schedule-export`. The shared-script graph writes reports in `Квартирография` next to the HVAC model. The internal UNC path and the example link name in the published files are placeholders.

## [2026-10-19]

### Added

- **Case 07** (#6). Typical floors: a Dynamo export of raw rooms from the architectural link into an Excel template that groups identical floors. Paths, project names, and link filenames in the commit are placeholders. Later commits rename this case and replace the graph; see 2026-10-29 and 2026-11-13.

## [2026-10-07]

### Changed

- **Docs** (#5). The repeated speed claim is stated once. Filename glosses are added. Case 01 wording no longer names a workstation.

## [2026-10-03]

### Changed

- **Docs** (#4). Hub READMEs gain license, language, and “tested on real working models” badges, plus the health-check screenshot. Case toolkit scripts under `cases/01`–`cases/06` gain a file header. Revit Batch Processor copies in `rbp_open_fail_patch/` are not part of that header pass.

## [2026-10-02]

### Changed

- **Cases 05 and 06** (#3). Scripts updated from the post-2026-09-14 copies: case 05 uses `Сжатие.vbs` and a rewritten compact UI and deep-save path; case 06 adds the single operator window (`start_ops.ps1`, `show_ops.vbs`, `Операции с моделями.vbs`). Doubled CR line endings in case 06 and in `servers.cfg` are normalized to CR LF. Screenshots are added. Later commits in the same pull request set the status wording to testing on real working models, state the 4–5× batch comparison once, rename case 06 to Model operations / Операции с моделями, and replace “coloured Excel” with “Excel report”.
- **Docs** (#2). Status labels drop “Production-ready”. Hub READMEs gain Quick start, and a third-party notice for Revit Batch Processor (GPL-3.0). `NOTICE` and `LICENSE-GPL-3.0.txt` sit with the modified RBP scripts. Interview wording is removed. Example names in the repo are anonymised (`PROJ_ANNEX` and similar). A follow-up commit anonymises leftover picker labels (comments and dialog text only).

## [2026-09-14]

### Added

- **Case 05.** Compact save of workshared models. The hub README is rewritten so the repository is BIM automation, not only Revit Batch Processor.
- **Case 06.** Model operations: rename, year upgrade, and RVT relink.

### Changed

- **Case 03.** Levels and grids is marked MVP only. The imminent-pilot wording is removed.

## [2026-08-10]

### Added

- **Docs** (#1). Draft Navisworks clash-tolerance matrix for stages P and R (`docs/navisworks-clash-tolerances.md` and the Russian twin). Documentation only. The commit describes the millimetre figures as a starting point for agreement, not a digitised company DOCX.

## [2026-08-05]

### Added

- **Case 04.** FTP model-exchange alerts via ntfy or Telegram. A PowerShell watcher polls shared FTP folders and sends a phone push when exchange files change.

## [2026-08-03]

### Added

- Portfolio published with three Revit Batch Processor cases: project units and `RSN://` (case 01), health check (case 02), and levels and grids (case 03). Scripts in that commit are sanitized, and each case has an English and a Russian README.
