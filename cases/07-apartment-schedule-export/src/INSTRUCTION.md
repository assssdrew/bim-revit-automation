# Apartment schedule export — instructions

[Русский](INSTRUCTION.ru.md)

The graph exports apartment rooms (combined bathroom, WC, walk-in closet) from the linked architectural model into Excel. Excel finds identical floors and fills «Таблица 1» (Table 1). Revit does not calculate anything, and the HVAC model is not modified.

## 1. What is where

The script folder is shared by every project. Placeholder path in this repository (substitute your own folder):

`\\server\share\BIM\Dynamo\Выгрузка квартирографии`

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
| `Квартирография_ОШИБКА.xlsx` | Created only when the architectural link was not found. The reason is inside. |

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

There is no reports-folder input. The only Player input is «Связь АР» (architectural link).

## 3. Export a building (every time)

1. Close every `Квартирография_*.xlsx` file in Excel.
2. Open the HVAC model of the building. Its architectural link must be loaded, and the model must be saved.
3. Manage → Dynamo Player → Run on the «Выгрузка_квартирографии» row.
4. Wait until it finishes.
5. Open the «Квартирография» folder next to the HVAC model (next to the central file when the model is workshared) and `Квартирография_Kxx.xlsx`. Example: `Квартирография_K04.xlsx`.

The building code is taken from the link name. Fictional link name: `PRJ_AR_01_K04_R24.rvt` → `K04`. If the folder or the building file does not exist yet, both are created; the file is a copy of the template in the script folder.

Leave «Связь АР» empty when one architectural link is loaded. Fill it only when several architectural links are loaded (section 7).

Run the graph from Dynamo Player. If you open the same graph in the Dynamo editor and press Run without changing the Python node, the editor can skip the node, and the export does not refresh.

## 4. What to read in the report

| Sheet | Contents |
|-------|----------|
| Типовые этажи | Floors grouped: identical floors share a colour and a group number (for example 7–12). Export date and time are at the top right. |
| Таблица 1 | Floor \| No. \| Type \| Rooms \| Total area \| Living area \| Bathroom \| WC \| Closet. |
| Выгрузка | Raw rows from Revit. Do not edit them by hand: the next export replaces this sheet. |
| Настройки | Building, whether apartment type counts, room-name variants (section 6). |
| Инструкция | Short note inside the workbook. |

If the numbers do not update, press F9.

## 5. Exporting the same building again

The building file is not rewritten as a whole: only the «Выгрузка» sheet is replaced. «Типовые этажи» and «Таблица 1» recalculate from the new rows. Edits on «Настройки» stay. Other buildings are untouched, because the file name comes from the link name.

Close the building file in Excel before you run.

`Data.ExportExcel` with overwrite deletes «Выгрузка» and creates it again. Ordinary references to that sheet become `#REF!`. Template formulas read «Выгрузка» through `INDIRECT`, so they survive the new sheet.

## 6. Another project, or different architectural parameters

Room names live on «Настройки» (do not edit the formulas):

- The bathroom / WC / closet columns list name variants. If the architectural model uses another name (for example «С/у»), type it into an empty cell of that column.
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
```

4. Accept changes, then Ctrl+S.

A new room name has to be added in two places: `ROOM_NAMES` in the script (otherwise the room is not exported) and «Настройки» (otherwise it is not counted in the bathroom / WC / closet column). For new buildings, edit the template in the script folder (one file for every project). For workbooks that already exist, edit those files in «Квартирография».

A new project does not need its own copy of the graph. Reports land next to the open HVAC model. If the script folder moves, change `SCRIPT_DIR` in the НАСТРОЙКИ block and save the graph.

## 7. Errors

| What you see | What to do |
|--------------|------------|
| `Квартирография_ОШИБКА.xlsx` appeared | The architectural link was not found. «Выгрузка» lists the reason and the links in the model. Load the link (Manage Links → Reload / Load). |
| «Загружено несколько связей АР» (several architectural links) | Type part of the link name into «Связь АР», for example `K04`. Matching ignores case, separators (`_`, space, `-`, `.`), and Cyrillic/Latin look-alikes (А/A, К/K, …). |
| «Закройте файл … и запустите снова» (close the file and run again) | The building workbook is open in Excel. Close it and run again. |
| «Нет шаблона …» (template missing) | `Квартирография_ШАБЛОН.xlsx` is not in the folder named by `SCRIPT_DIR`. The path should end with `\Dynamo\Выгрузка квартирографии`. |
| «Модель ОВ не сохранена» (HVAC model is not saved) | Save the HVAC model and run again. |
| «Модель ОВ хранится в облаке / на Revit Server» | A reports folder cannot be created beside a cloud or Revit Server model. Open a model saved on a file share. |
| «Не удалось создать папку отчётов …» | The account cannot create a folder next to the HVAC model (or next to the central file). |
| Zeros or blanks in the report | Press F9. Check that the architectural parameters named in НАСТРОЙКИ are filled, and that room names match «Настройки». |
| Warning «Use ExportToExcel instead» | It does not change the result. |

## 8. Limits

- Microsoft Excel must be installed (`Data.ExportExcel`).
- Revit 2024, Python node engine CPython3.
- Up to 15 apartments per floor, 80 floors, 1500 apartments, 3000 export rows per building.
- No macros. The `.xlsx` opens in Excel 2010 and newer.
- The script only reads rooms in the architectural link. It does not save or sync the HVAC model. It does write the «Квартирография» folder and the building workbook.
