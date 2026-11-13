# Changelog

All notable changes to this portfolio are recorded here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This repository is not one versioned package. Each case keeps its own script version, so the history below is grouped by date and, where a change belongs to one case, by case. Entries come from the git history and merged pull requests #1–#9. They do not add claims the commits do not support.

## [Unreleased]

### Changed

- **Case 03.** The case READMEs and `src/README.md` now say **v1.0.2**, the `SCRIPT_VERSION` already written by `levels_grids_audit.py` and `levels_grids_apply.py`. The shared helper `lg_lib.py` header matches that. The previous `src/README.md` said `1.0.x`.
- **Case 05.** The RSN deep-compact patch header now says **4.0**, matching the marker `RSN_PATCH_PS_v4` and the case READMEs. The header had said `2.0`, which is the case 01 patch. The launcher `установить_поддержку_RSN_в_RBP.cmd` uses the same header. The marker string in the script is unchanged.
- **Case 06.** Toolkit headers that said `1.0` now say **3.0.0**, the `SCRIPT_VERSION` already written by `saveas_central.py`, `detach_to_buffer.py`, and `update_rvt_links.py`, and the version already named in the hub README. Both case READMEs name v3.0.0. The saved-job field `version = "1.0.0"` inside `presets.ps1` is unchanged.

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
