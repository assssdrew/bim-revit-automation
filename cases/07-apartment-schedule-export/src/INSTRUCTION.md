# Apartment schedule export v7 — instructions

[Русский](INSTRUCTION.ru.md)

The graph exports apartment rooms (combined bathroom, WC, walk-in closet) and common-area rooms (stair, airlock, airlock in a fire-safe zone, corridor, lift lobby) from the linked architectural model of **one building** into an Excel file. When Excel opens the file it finds identical floors, fills «Таблица 1» (Table 1), and fills the «МОП» sheet. Revit does not modify the model. Excel is not started for the export: Python’s standard library (`zipfile` and XML) writes the workbook directly.

What is new in v7: areas to hundredths (`AREA_DECIMALS = 2`); a «МОП» sheet; floors in «Таблица 1» separated by a line and a fill; each typical-floor group has its own colour (up to 40); the result and every error are text in the Player «Результат» output. There is no `…_ОШИБКА.xlsx` file, and the export does not leave an Excel process running.

## 1. What is where

The script folder is shared by every project. Placeholder path in this repository (substitute your own folder):

`\\server\share\Dynamo\Выгрузка квартирографии`

| File | What it is |
|------|------------|
| `Выгрузка_квартирографии.dyn` | Dynamo graph. This is what you run. |
| `Квартирография_ШАБЛОН.xlsx` | Report template. Do not open it to edit, rename, or delete it. |
| `Выгрузка_квартирографии.py` | Copy of the node code (reference only; not required to run). |
| `Инструкция_Выгрузка_квартирографии.docx` | The same instructions as a Word file. |
| `build.py` | Rebuilds the template. Not required to run. |
| `mk.py` | Rebuilds the docx. Not required to run. |

The reports folder is created next to the open HVAC model:

- A normal model: the model folder `\Квартирография`.
- A workshared model: the central-file folder `\Квартирография`.

| File in «Квартирография» | What it is |
|--------------------------|------------|
| `Квартирография_K01.xlsx`, `…_K04.xlsx`, … | One report per building. Created on the first export of that building. |
| `Квартирография_K01_старый_ДАТА.xlsx` | A copy of a v6 report. Created once, the first time v7 runs against an old file (the old file has no «МОП» sheet). Safe to delete. |

Reports do not appear in the script folder. The HVAC model has to be saved on disk or on a file share. A cloud model (BIM 360 / ACC) or a Revit Server model (`RSN://`) has no folder beside it, and the graph stops.

`SCRIPT_DIR` in the Python node must point at the script folder that holds the template.

## 2. First-time setup (once per computer)

1. Revit → Manage → Dynamo Player.
2. In the player, click the folder icon (change folder).
3. Select the script folder (the path above). The player remembers it, and «Выгрузка_квартирографии» stays in the list.

If the network folder is not visible in the picker:

- Open the folder in File Explorer → right-click «Выгрузка квартирографии» → Pin to Quick access.
- Open the picker again. The folder appears under Quick access.

To stop Dynamo warning about a network file: Dynamo → Preferences → Security → Trusted Locations → add this folder.

The only Player input is «Связь АР» (architectural link). The «Результат» output shows the run text.

## 3. Export a building (every time)

One run exports one building (one architectural link).

1. Close every `Квартирография_*.xlsx` file in Excel.
2. Open the HVAC model of the building. Its architectural link must be loaded, and the model must be saved.
3. Manage → Dynamo Player → Run on the «Выгрузка_квартирографии» row.
4. Wait until it finishes and expand «Результат»: `ГОТОВО: корпус K04, Файл: …` and the room counts. If the text starts with `ОШИБКА:`, see section 7.
5. Open the «Квартирография» folder next to the HVAC model (next to the central file when the model is workshared) and `Квартирография_Kxx.xlsx`. Example: `Квартирография_K04.xlsx`.

The building code is taken from the link name. Fictional link name: `PRJ_AR_01_K04_R24.rvt` → `K04`. If the folder or the building file does not exist yet, both are created; the file is a copy of the template in the script folder.

Leave «Связь АР» empty when one architectural link is loaded. Fill it only when several architectural links are loaded (section 7).

Run the graph from Dynamo Player. If you open the same graph in the Dynamo editor and press Run without changing the Python node, the editor can skip the node, and the export does not refresh.

## 4. What to read in the report

| Sheet | Contents |
|-------|----------|
| Типовые этажи | Identical floors share a group number and a colour (one colour per group, up to 40 groups). On the right, a summary sorted by group: group, floor list, count. Export date and time are at the top. |
| Таблица 1 | Floor \| No. \| Type \| Rooms \| Total area \| Living area \| Bathroom \| WC \| Closet. Each new floor is separated by a line; every other floor is filled grey. |
| МОП | Common-area areas by floor: one column per name from «Настройки», a floor total on the right, and a building total in row 3. |
| Выгрузка МОП | Raw common-area rows (floor, name, area). Do not edit them by hand. |
| Выгрузка | Raw apartment rows from Revit. Do not edit them by hand: the next export clears this sheet and writes it again. |
| Настройки | Building, whether apartment type counts, apartment room names, and common-area names (section 6). |
| Инструкция | Short note inside the workbook. |

Formulas recalculate when the file opens (`fullCalcOnLoad`). Excel may ask to save changes on close; answering Yes is fine. Areas are exported to 0.01 m².

## 5. Exporting the same building again

The building file is not rewritten as a whole: «Выгрузка» and «Выгрузка МОП» are cleared and filled again, so old rows do not remain. «Типовые этажи», «Таблица 1», and «МОП» recalculate from the new rows. Edits on «Настройки» stay. Other buildings are untouched, because the file name comes from the link name.

Close the building file before you run. If it is open, the script writes nothing and «Результат» says the file is busy.

Template formulas read «Выгрузка» through `INDIRECT`, so replacing that sheet does not turn them into `#REF!`.

The first v7 run against a v6 file (no «МОП» sheet) renames it to `Квартирография_<code>_старый_<date>.xlsx` and creates a new file from the template.

## 6. Another project, or different architectural parameters

Room names live on «Настройки» (do not edit the formulas):

- The bathroom / WC / closet columns list apartment room-name variants. If the architectural model uses another name (for example «С/у»), type it into an empty cell of that column.
- Column F, «МОП», lists common-area names for the «МОП» sheet, up to 10. The same name has to be in `MOP_NAMES` in the script.
- «Учитывать тип квартиры…» (include apartment type): «нет» compares room sets only; «да» also compares apartment types.

Parameter names in the script, when the architectural model uses different ones:

1. Open `Выгрузка_квартирографии.dyn` in Dynamo.
2. Double-click the «Выгрузка АР» node.
3. Edit only the НАСТРОЙКИ block at the top. The `KLWT_*` strings are example shared-parameter names from one client:

```python
P_FLOOR = "KLWT_Этаж"                     # floor
P_TYPE  = "KLWT_Тип квартиры"             # apartment type
P_NUM   = "KLWT_Номер квартиры на этаже"  # apartment number on the floor
P_SOBSH = "KLWT_Площадь квартиры общая"   # total area
P_SZH   = "KLWT_Площадь квартиры жилая"   # living area
ROOM_NAMES = ["Совмещенный санузел", "Санузел", "Ванная", "Уборная", "Туалет", "Гардеробная"]
MOP_NAMES  = ["Лестничная клетка", "Тамбур-шлюз", "Тамбур-шлюз (ПБЗ)", "Коридор", "Лифтовой холл"]
AREA_DECIMALS = 2                         # decimal places
```

4. Accept changes, then Ctrl+S.

A new room name (apartment or common area) has to be added in two places: `ROOM_NAMES` / `MOP_NAMES` in the script (otherwise the room is not exported) and «Настройки» (otherwise it is not counted in the right column). For new buildings, edit the template in the script folder (one file for every project). For workbooks that already exist, edit those files in «Квартирография».

A new project does not need its own copy of the graph. Reports land next to the open HVAC model. If the script folder moves, change `SCRIPT_DIR` in the НАСТРОЙКИ block and save the graph.

## 7. Errors

Every message is text in the Dynamo Player «Результат» output.

| What you see | What to do |
|--------------|------------|
| ОШИБКА: Загружено несколько связей АР … | Type part of the link name into «Связь АР», for example `K04`, or unload the extra links (Manage Links → Unload). Hiding a link in a view is not enough. Matching ignores case, separators, and Cyrillic/Latin look-alikes. |
| ОШИБКА: Загруженная связь АР не найдена … | Load the architectural link (Manage Links → Reload / Load), or type part of its name into «Связь АР». The message lists the links in the model. |
| ОШИБКА: Файл занят: … закройте его | The building workbook is open in Excel. Close it and run again. If nothing has it open, end `EXCEL.EXE` in Task Manager. |
| ОШИБКА: Нет шаблона … | The template was deleted or renamed, or `SCRIPT_DIR` points elsewhere. Check that Explorer can open the folder. |
| ОШИБКА: Модель ОВ не сохранена | Save the HVAC model and run again. |
| ОШИБКА: Модель ОВ хранится в облаке / на Revit Server | A reports folder cannot be created beside a cloud or Revit Server model. Open a model saved on a file share. |
| ОШИБКА: Не удалось создать папку отчётов … | The account cannot create a folder next to the HVAC model (or next to the central file). |
| ВНИМАНИЕ: нет параметров в связи … | Change the parameter names in the НАСТРОЙКИ block (section 6). |
| ВНИМАНИЕ: пропущено помещений без этажа/номера | Those rooms have no floor or apartment number in the architectural model. |
| Zeros or blanks in the report | Check that the architectural parameters named in НАСТРОЙКИ are filled, and that room names match «Настройки». |

## 8. Limits

- Excel is not required for the export. It is required to read the report.
- One run exports one building (one architectural link).
- The «МОП» sheet: up to 10 room kinds; floors must be numbers (`01`, `1`, `2`, …). A common-area room whose floor is not a number is left off the sheet.
- Group colours repeat only when there are more than 40 groups.
- Up to 15 apartments per floor, 80 floors, 1500 apartments, 3000 export rows per building.
- No macros. The `.xlsx` opens in Excel 2010 and newer.
- The script only reads rooms in the architectural link. It does not save or sync the HVAC model. It does write the «Квартирография» folder and the building workbook.
