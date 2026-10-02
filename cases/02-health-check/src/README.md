# Health Check — operator notes

Read-only audit for Revit Batch Processor. Script version **1.4.0**. No Save / Sync / Relinquish.

## Quick start

1. `choose_models_path.cmd`
2. `reset_report_session.cmd` → local folder for Excel reports
3. RBP: `health_check.py` + `rvt_list.txt` (Create New Local, Detach OFF)
4. Rebuild the Excel report if needed: `color_latest_report.cmd`

Tune traffic lights in `thresholds.cfg`.
