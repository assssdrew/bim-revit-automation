# Safety rules

These toolkits were written for a real Revit Server / workshared environment. The public code keeps the same safety boundaries.

## Hard rules

1. **Create New Local** for workshared / `RSN://` models. Never Detach live centrals for Apply jobs.
   Exceptions: compact-save **deep** mode opens the central with Audit (no new local, no Detach) and Save As the same central with Compact. **Fast** compact still uses Create New Local. RSN deep mode needs RBP patch v4 (`RSN_PATCH_PS_v4`); tested on a limited number of runs. Model-ops **Detach ON** only for a local upgrade buffer when leaving an old Revit Server — never as the live central.
2. **Audit ≠ Apply.** Report scripts must not Save, Sync, or Relinquish.
3. **Apply scope is narrow.** Prefer an explicit model list (e.g. base files only), not “all 80 disciplines”.
4. **No automatic coordinate surgery.** Internal Origin, PBP / Survey Point, True North, Shared Coordinates / Site stay manual or report-only.
5. **Tolerances required.** Without mm/degree tolerances, floating noise becomes false RED.
6. **Do not auto-delete** levels/grids in MVP Apply. Missing items → report; extras → report.
7. **XLSX without Excel COM** on UNC (OpenXML + local output folder).
8. **RSN patch** (units toolkit) is optional infrastructure: stock [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) (GPL-3.0) rejects `RSN://` on `File.Exists`. Patch only your installed RBP Scripts after understanding the change.
