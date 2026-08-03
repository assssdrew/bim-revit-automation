# Project Units — operator notes

Batch Length accuracy for Revit Batch Processor. Supports local / UNC / `RSN://` in one `rvt_list.txt`.

Depends on [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor).

## Quick start

1. Edit `servers.cfg` (Revit Server hosts)
2. `choose_accuracy.cmd` → e.g. `0.1`
3. `choose_models_path.cmd` (menu 3 auto-applies RSN patch)
4. RBP: task `set_length_accuracy.py`, list `rvt_list.txt`, Create New Local ON, Detach OFF

Manual RSN patch: `install_RSN_support_in_RBP.cmd` (see `rbp_rsn_patch/`).

Add-in trust helper: `allow_BatchRvt_addin.cmd`.

## After edit

| Path type | Behaviour |
|-----------|-----------|
| `RSN://` / workshared | Sync + Relinquish |
| Plain file | SaveAs in place |

Requires Revit 2022+ (`SpecTypeId` / `UnitTypeId`).
