# Health Check для Revit Batch Processor

Аудит без `Save` / `Sync` / `Relinquish`. Версия скрипта **1.5.4**.

## Канон отчёта

1. **Каждый новый прогон RBP** — новый `health_check_….csv` / XLSX (старые файлы не трогаются)
2. **XLSX один раз** — после **последней** модели прогона
3. **Общая папка** — `%USERPROFILE%\Documents\doc\script\rbp\reports` (создаётся сама при первом отчёте)
4. **Имя файла** — по имени скрипта: `health_check_YYYY-MM-DD_HHmm.csv` / `.xlsx`
5. Лист **«Сводка»** — узкий обзор на один экран
6. Лист **«Связи»** — хост + плюсик со связями
7. Лист **«Метрики»** — остальные колонки
8. Сборка XLSX без Excel COM (OpenXML)

## Выбор моделей

`choose_models_path.cmd` → пункт **1**:
1. Открыть **родительскую** папку (например `…\4_Разработка проекта`) — путь запоминается
2. Отметить папки дисциплин (**Ctrl / Space**), напр. `5_1_Электрика` … `9_ПТ`
3. Скрипт ищет папки `RVT` на любой глубине и берёт только `*_R##.rvt` (R22/R24…)
4. Игнор: `Backup`, `*_backup`, `Revit_temp`, `Резерв`, `Семейства`, `.0001.rvt`, `.rfa`
5. Список сразу пишется в `rvt_list.txt` (без второго чекбокса)

Пункты 2–3: отдельные файлы / Revit Server — как раньше.

## Метрики (волны A/B/C)

Почти всё — **колонки отчёта**. В светофор — только пороги из `thresholds.cfg`.

| Волна | В отчёте |
|-------|----------|
| A | CAD Import / CAD Link, In-place, Groups (типы+экземпляры), Design Options, Views/Sheets (+ без шаблона, пустые листы), Rooms/Spaces unplaced, Phases, Project Locations |
| B | Loadable families + types + unused types (оценка), Images, Detail Lines, Filled Regions, pinned FamilyInstance, Workset1 elements, MEP open connectors (оценка, capped), IFC links |
| C | Max geom distance + far elements, Schedules, Materials, Line patterns |

**Не в batch (дорого / API):** вес отдельных RFA, CAD внутри семейств, nested shared depth, точный Purge Unused, поиск дубликатов геометрии.

## Запуск

1. `choose_models_path.cmd`
2. RBP: `health_check.py` + `rvt_list.txt` (`Create New Local`, Detach OFF)
3. После последней модели: CSV и XLSX в `Documents\doc\script\rbp\reports` (`health_check_…`)

`reset_report_session.cmd` — необязателен (папки / сброс маркера вручную).

Пересобрать XLSX: `color_latest_report.cmd`

Скопируйте на ПК RBP: `C:\Users\Public\Documents\doc\script\rbp\batch_health_check\`
