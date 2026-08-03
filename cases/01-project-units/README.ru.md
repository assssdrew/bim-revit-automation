# Кейс 01 — Единицы проекта + Revit Server (`RSN://`)

[← К хабу](../../README.ru.md) · [EN](README.md)

## Проблема

Округление Длины должно быть одинаковым в десятках workshared-моделей. Стоковый [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor) отбрасывает `RSN://` на `File.Exists` — до API дело не доходит.

## Решение

Операторский toolkit + task-скрипт RBP:

- Единый picker: папка / файлы / Revit Server → `rvt_list.txt`
- Авто (или ручной) **RSN-патч v2** Scripts RBP (кириллица в IronPython)
- `set_length_accuracy.py` задаёт точность Length, затем **Sync + Relinquish** для workshared/RSN; SaveAs для обычных файлов

## Эффект

| | |
|--|--|
| Статус | В проде |
| Доказательство | Реальный прогон RSN: local of central → accuracy → Sync OK |
| UX | В одном списке можно смешивать диск и `RSN://` |

## Безопасность

- Живые centrals: **Create New Local**, Detach **OFF**
- Меняется только округление Длины — не координаты
- Перед батчем — `smoke_test.cmd`

## Запуск

1. Скопировать `src/` в Scripts RBP
2. Прописать хосты в `servers.cfg`
3. `choose_accuracy.cmd` → например `0.1`
4. `choose_models_path.cmd` (пункт 3 ставит RSN-патч)
5. RBP: task = `set_length_accuracy.py`, list = `rvt_list.txt`

Код: [`src/`](src/)
