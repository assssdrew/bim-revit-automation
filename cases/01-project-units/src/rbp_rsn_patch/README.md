# Патч RBP: поддержка Revit Server (`RSN://`)

Обычный Revit Batch Processor **отклоняет** пути `RSN://...` на проверке `File.Exists` (как будто файла нет).  
API Revit при этом умеет открывать их через **Create New Local**.

Этот патч правит 2 скрипта в папке `Scripts` установленного RBP:

| Файл | Что меняется |
|------|----------------|
| `revit_file_list.py` | `RSN://` считается существующим; год из `_R22`; без `str()` на кириллических путях (IronPython) |
| `revit_script_host.py` | не падать на `File.Exists`; всегда Create New Local для RSN |

## Установка

**Автоматически:** при выборе пункта **3 — Revit Server** в `choose_models_path.cmd` патч применяется сам (Python не нужен). Если RBP открыт, папка `Scripts` определяется по процессу `BatchRvt.exe`.

В консоли должно быть `RBP RSN patch: OK`. Если ошибка — текст причины выводится сразу под этой строкой.

Локалка / шара — пункты 1–2 — работают как раньше.

**Вручную** (после обновления RBP или если авто не нашло Scripts):

1. На ПК, где запускается RBP, запустите `установить_поддержку_RSN_в_RBP.cmd`
2. Или:

```powershell
powershell -ExecutionPolicy Bypass -File rbp_rsn_patch\apply_rsn_support.ps1 -ScriptsPath "C:\Path\To\RevitBatchProcessor\Scripts"
```

Создаются бэкапы `*.bak_before_rsn`. Повторный запуск безопасен.

## Запуск батча

- Central File Processing = **Create New Local**
- Detach = **OFF**
- Версия Revit = год моделей (`*_R22` → 2022)
- В `rvt_list.txt` можно смешивать `C:\...`, `\\server\...`, `RSN://...`

Локали создаются в `C:\REVIT_LOCAL2022\` (и аналогично для других годов) — стандарт RBP.
