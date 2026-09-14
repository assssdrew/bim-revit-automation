# How to add the next case

Publish new automation as a numbered folder. Keep the hub README in sync (EN + RU).

## Checklist

1. Create `cases/NN-short-name/`
2. Add `README.md` + `README.ru.md` (problem → solution → impact → safety → how to run)
3. Put sanitised scripts in `src/`
   - No corporate UNC / internal IPs / personal paths
   - Use `*.example.cfg` or placeholder hosts in `servers.cfg`
4. Add anonymised sample under `samples/` if there is a report
5. Update the cases table in root `README.md` and `README.ru.md`
6. Commit with a message like: `Add case NN: short name (MVP|production)`

## Suggested numbering

| NN | Topic |
|----|--------|
| 05 | Compact save (published) |
| 06 | Model ops (detach / save-as central / link remap) |
| 07 | Navisworks rename / publish helpers |

Do not commit live `rvt_list.txt`, session cfg, real project reports, FTP `credentials.xml` / `config.json` / `state.json`.
