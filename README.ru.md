# Портфолио: автоматизация BIM / Revit

Пакетная автоматизация моделей Autodesk Revit через [Revit Batch Processor (RBP)](https://github.com/bvn-architecture/RevitBatchProcessor) — открытая утилита, не моя разработка.

Я пишу **task-скрипты и операторские toolkit’и**: выбор моделей → прогон → цветной Excel-отчёт (и, где безопасно, Apply + Sync).

| | |
|---|---|
| **Автор** | [assssdrew](https://github.com/assssdrew) |
| **Стек** | Revit API · IronPython (RBP) · PowerShell · OpenXML |
| **Фокус** | Workshared / Revit Server (`RSN://`), безопасный Apply |
| **EN** | [English README](README.md) |

---

## Кейсы (публикуются по мере готовности)

| # | Кейс | Статус | Суть |
|---|------|--------|------|
| 01 | [Единицы проекта + RSN](cases/01-project-units/) | **В проде** | Пакетная точность Length; локальные / UNC / `RSN://`; Sync + Relinquish |
| 02 | [Health Check](cases/02-health-check/) | **В проде** (v1.4.0) | Аудит здоровья модели без Save/Sync → CSV + цветной XLSX |
| 03 | [Уровни и оси](cases/03-levels-grids/) | **MVP готов** · пилот со дня на день | Каскадная сверка Levels/Grids с эталоном; Apply только для БФ |

Новые кейсы — папки в `cases/`, инструкция: [docs/HOW_TO_ADD_CASE.md](docs/HOW_TO_ADD_CASE.md).

---

## Проблема → подход

**Типовая неделя:** пришла обновлённая АР; нужно держать в согласованности БФ и до ~80 разделов (единицы, здоровье модели, уровни/оси).

Вручную это не масштабируется. Toolkit’и:

1. Собирают список моделей (папка / файлы / Revit Server)
2. Идут через RBP с **Create New Local** (без Detach на живых centrals)
3. Отдают отчёты **GREEN / YELLOW / RED**
4. Пишут в модель только там, где риск понятен (единицы; уровни/оси БФ) — без «автокоординации»

---

## Эффект (масштаб проекта)

| Метрика | Значение |
|---------|----------|
| Недельный каскад | ~5–6 БФ (БФ↔АР), затем до ~80 разделов (↔БФ) |
| Units / RSN | Подтверждено на реальных моделях Revit Server |
| Health Check | Пакетный read-only аудит workshared-моделей |
| Levels & Grids | MVP в коде; живой пилот ожидается со дня на день |

Цель: **одна сессия отчёта вместо ручного открытия десятков моделей**.

---

## Безопасность (важно на собеседовании)

Автоматизация, которая ломает федеративную модель, хуже отсутствия автоматизации.

- **Audit** никогда не делает Save / Sync / Relinquish
- **Apply** — отдельный скрипт / явный список (например, только БФ)
- Координаты, PBP, Survey, True North, Shared Coordinates — **не трогаем автоматом**
- Допуски обязательны, иначе ложные RED

Подробнее: [docs/SAFETY.ru.md](docs/SAFETY.ru.md)

---

## Структура репозитория

```text
bim-revit-automation/
  README.md / README.ru.md
  cases/
    01-project-units/src/
    02-health-check/src/
    03-levels-grids/src/
  samples/
  docs/
```

Папку `cases/*/src` копируют в каталог Scripts RBP.

---

## Требования

- Autodesk Revit (год = году моделей; units — 2022+)
- [Revit Batch Processor](https://github.com/bvn-architecture/RevitBatchProcessor)
- Windows + PowerShell

---

## Дисклеймер

Код — портфолио / стартовая точка. Перед Apply — проверка на тестовой копии. Корпоративные пути в публичном репо санитизированы.
