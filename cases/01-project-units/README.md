# Case 01 — Project Units + Revit Server (`RSN://`)

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

Length rounding must be consistent across dozens of workshared models. Stock [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) rejects `RSN://` paths on `File.Exists`, so Server models never reach the Revit API.

## Solution

Operator toolkit + RBP task script:

- Unified picker: local folder / files / Revit Server → `rvt_list.txt`
- Automatic (or manual) **RSN patch v2** for installed RBP Scripts (Cyrillic-safe IronPython)
- `set_length_accuracy.py` sets Length accuracy (mm), then **SynchronizeWithCentral + Relinquish** for workshared / RSN; SaveAs for plain files

## Impact

| Item | Result |
|------|--------|
| Status | Production-ready |
| Proof | Real RSN run: open local of central → set accuracy → Sync OK |
| Operator UX | One list can mix disk and `RSN://` paths |

## Safety

- Live centrals: **Create New Local**, Detach **OFF**
- Changes only Length accuracy (optionally force mm) — not coordinates
- Smoke test script before batch

## How to run

1. Copy `src/` into your RBP Scripts folder
2. Edit `servers.cfg` with your Revit Server hosts
3. `choose_accuracy.cmd` → e.g. `0.1`
4. `choose_models_path.cmd` → build list (menu 3 applies RSN patch)
5. RBP: task = `set_length_accuracy.py`, list = `rvt_list.txt`

Code: [`src/`](src/)
