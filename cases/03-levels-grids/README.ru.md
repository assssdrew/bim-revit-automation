# Кейс 03 — Контроль уровней и осей (Audit / Apply)

[← К хабу](../../README.ru.md) · [EN](README.md) · [Пример CSV](../../samples/levels_grids_details_sample.csv)

## Проблема

После обновления АР нужно согласовать **уровни и оси** сначала в БФ, затем в ~80 разделах (имена, отметки, позиция/угол осей, дубли, missing/extra). Ручной Copy/Monitor не масштабируется.

## Решение

Каскадный toolkit для RBP:

1. **Audit БФ ↔ АР** (эталон = связь АР) → CSV / цветной XLSX  
2. Разбор → **Apply только БФ** (elevation / move-rotate осей при match по имени + флаги cfg) → Sync  
3. **Audit разделы ↔ БФ** → отчёт (массовый Apply на 80 моделей — *не* MVP)

Сравнение через Revit API (не auto-Accept Coordination Review).

## Эффект

| | |
|--|--|
| Статус | **MVP готов** · живой пилот со дня на день |
| Недельная нагрузка | ~5–6 БФ, затем до ~80 разделов |
| Допуски по умолчанию | Уровень 2 мм · Ось 5 мм · Угол 0.1° |

## Безопасность (акцент для собеседования)

| В scope | Вне scope |
|---------|-----------|
| Имена, отметки, позы осей, дубли, missing/extra, светофор | Internal Origin, PBP/Survey, True North, Shared Coordinates, Site |
| Apply: только список БФ; в MVP не удаляем «лишние» | Слепая автокоординация |

Логика: пока БФ ≠ АР, гонять 80 разделов бессмысленно.

## Запуск (Audit)

1. Скопировать `src/` в Scripts RBP  
2. `choose_models_path.cmd`  
3. `reset_report_session` / `choose_exemplar_link` — эталон  
4. RBP: `levels_grids_audit.py`

Apply: в списке **только БФ** → `levels_grids_apply.py`.

Код: [`src/`](src/)
