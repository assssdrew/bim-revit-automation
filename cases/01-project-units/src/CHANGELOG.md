# Changelog

## 2026-07-27

- Unified model picker: local folder/files + Revit Server (`RSN://`) in one `choose_models.ps1`.
- Fixed RSN full path build (nested folders are preserved in `rvt_list.txt`).
- Added automatic RBP RSN patch from picker (menu `3`), plus manual launcher:
  - `rbp_rsn_patch/apply_rsn_support.ps1`
  - `install_RSN_support_in_RBP.cmd`
- RSN patch v2: Cyrillic-safe path checks for IronPython (no `str(path).upper()` crash).
- Updated `choose_accuracy.cmd` for reliable UNC execution from `cmd.exe` (`pushd`, ASCII-safe output).
- Added `smoke_test.cmd` for quick toolkit sanity check.
