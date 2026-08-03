# Levels & Grids — operator notes

Version **1.0.x**. Compare host Levels/Grids to an exemplar RVT link. Audit = report only; Apply = BF list only.

## Model list

`choose_models_path.cmd` → `rvt_list.txt`

## Exemplar

`choose_exemplar_link.cmd` / session reset → exemplar cfg. Script finds the matching RVT link in the host.

| Run | Host list | Exemplar |
|-----|-----------|----------|
| BF ↔ AR | BF models | AR model |
| Disciplines ↔ BF | Disciplines | BF model |

## Out of scope

Internal Origin, PBP/Survey, True North, Shared Coordinates, Site.

## Audit

1. Choose models  
2. Reset session + exemplar  
3. RBP: `levels_grids_audit.py`

## Apply (BF only)

1. `rvt_list.txt` = BF only  
2. Exemplar = AR  
3. RBP: `levels_grids_apply.py`  

Apply does not create/delete levels or grids.
