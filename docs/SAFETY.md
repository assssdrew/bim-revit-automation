# Safety rules

These toolkits are built for **production BIM servers**. The public code keeps the same safety boundaries.

## Hard rules

1. **Create New Local** for workshared / `RSN://` models. Never Detach live centrals for Apply jobs.
2. **Audit ≠ Apply.** Report scripts must not Save, Sync, or Relinquish.
3. **Apply scope is narrow.** Prefer an explicit model list (e.g. base files only), not “all 80 disciplines”.
4. **No automatic coordinate surgery.** Internal Origin, PBP / Survey Point, True North, Shared Coordinates / Site stay manual or report-only.
5. **Tolerances required.** Without mm/degree tolerances, floating noise becomes false RED.
6. **Do not auto-delete** levels/grids in MVP Apply. Missing items → report; extras → report.
7. **XLSX without Excel COM** on UNC (OpenXML + local output folder).
8. **RSN patch** (units toolkit) is optional infrastructure: stock RBP rejects `RSN://` on `File.Exists`. Patch only your installed RBP Scripts after understanding the change.

## Why this matters in interviews

It shows process automation with **risk control**, not “script that changes everything”.
