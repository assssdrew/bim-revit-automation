# Typical floors — instructions

[Русский](INSTRUCTION.ru.md)

The graph exports apartment rooms (combined bathroom, WC, walk-in closet) from the linked architectural model into Excel. Excel finds identical floors and fills «Таблица 1» (Table 1). Revit does not calculate anything, and the model is not modified.

## 1. What is in the folder

Example reports folder (substitute your own path; this repository has no corporate path):

`\\server\share\Project\Scripts\TypicalFloors`

| File | What it is |
|------|------------|
| `Типовые_этажи.dyn` | Dynamo graph. This is what you run. |
| `Типовые_этажи_ШАБЛОН.xlsx` | Report template. Do not open it to edit, rename, or delete it. |
| `Export_AR_to_Excel.py` | Copy of the node code (reference only; not required to run). |
| `build.py` | Rebuilds the template. Not required to run. |
| `Типовые_этажи_K01.xlsx`, `…_K04.xlsx`, … | One report per building. Created on the first export of that building. |
| `Типовые_этажи_ОШИБКА.xlsx` | Created only when the architectural link was not found. The reason is inside. |

Keep these files in one folder. The graph looks for the template in the folder set on the «Папка отчётов» (reports folder) input.

## 2. First-time setup (once per computer)

1. Revit → Manage → Dynamo Player.
2. In the player, click the folder icon (change folder) and select the folder that contains `Типовые_этажи.dyn`. The player remembers it.
3. On the «Типовые_этажи» graph, set «Папка отчётов» to that same folder (it must contain `Типовые_этажи_ШАБЛОН.xlsx`). Leave «Связь АР» (architectural link) empty.

If the network folder is not visible in the picker:

- Open the folder in File Explorer → right-click → Pin to Quick access.
- Open the picker again. The folder appears under Quick access.

The folder picker cannot type a UNC path. If pinning does not work, open the `.dyn` in Dynamo, replace the «Папка отчётов» Directory node with a **String** node, and paste the path as a single line with no extra spaces and no line break:

`\\server\share\Project\Scripts\TypicalFloors`

Mark the node as an input, wire it to the second input of the «Выгрузка АР» node, and save (Ctrl+S).

To stop Dynamo warning about a network file: Dynamo → Preferences → Security → Trusted Locations → add this folder.

## 3. Export a building (every time)

1. Close every `Типовые_этажи_*.xlsx` file in Excel.
2. Open the HVAC model of the building. Its architectural link must be loaded.
3. Manage → Dynamo Player → Run on the «Типовые_этажи» row.
4. Wait until it finishes.
5. Open `Типовые_этажи_Kxx.xlsx`, where `Kxx` is the building code taken from the link name.

Fictional link name: `PRJ_AR_01_K04_R24.rvt` → file `Типовые_этажи_K04.xlsx`. If that file does not exist yet, it is copied from the template. Nothing has to be prepared in advance.

Fill «Связь АР» only when several architectural links are loaded (see section 7). With a single loaded architectural link, the script picks it.

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

The building file is not rewritten as a whole: only the «Выгрузка» sheet is replaced. «Типовые этажи» and «Таблица 1» recalculate from the new rows. Edits on «Настройки» stay. Other buildings are untouched, because the file name comes from the link name — one building’s report cannot overwrite another’s.

Close the building file in Excel before you run.

`Data.ExportExcel` with overwrite deletes «Выгрузка» and creates it again. Ordinary references to that sheet become `#REF!`. Template formulas read «Выгрузка» through `INDIRECT`, so they survive the new sheet.

## 6. Another project, or different architectural parameters

Room names live on «Настройки» (do not edit the formulas):

- The bathroom / WC / closet columns list name variants. If the architectural model uses another name (for example «С/у»), type it into an empty cell of that column.
- «Учитывать тип квартиры…» (include apartment type): «нет» compares room sets only; «да» also compares apartment types.

Parameter names in the script, when the architectural model uses different ones:

1. Open `Типовые_этажи.dyn` in Dynamo.
2. Double-click the «Выгрузка АР» node.
3. Edit only the НАСТРОЙКИ block at the top. The `KLWT_*` strings are example shared-parameter names from one client, not a required standard:

```python
P_FLOOR = "KLWT_Этаж"                     # floor
P_TYPE  = "KLWT_Тип квартиры"             # apartment type
P_NUM   = "KLWT_Номер квартиры на этаже"  # apartment number on the floor
P_SOBSH = "KLWT_Площадь квартиры общая"   # total area
P_SZH   = "KLWT_Площадь квартиры жилая"   # living area
ROOM_NAMES = ["Совмещенный санузел", "Санузел", "Ванная", "Уборная", "Туалет", "Гардеробная"]
```

4. Accept changes, then Ctrl+S.

A new room name has to be added in two places: `ROOM_NAMES` in the script (otherwise the room is not exported) and «Настройки» (otherwise it is not counted in the bathroom / WC / closet column). For new buildings, edit the template. For workbooks that already exist, edit those files.

A different reports folder: change the «Папка отчётов» input (or the String node, if you replaced it as in section 2).

## 7. Errors

| What you see | What to do |
|--------------|------------|
| `Типовые_этажи_ОШИБКА.xlsx` appeared | The architectural link was not found. «Выгрузка» lists the reason and the links in the model. Load the link (Manage Links → Reload / Load). |
| «Загружено несколько связей АР» (several architectural links) | Type part of the link name into «Связь АР», for example `K04`. Matching ignores case, separators (`_`, space, `-`, `.`), and Cyrillic/Latin look-alikes (А/A, К/K, …). A fragment such as `…_01_…` does not attach to `…_02_…` because of a look-alike letter, and a single loaded architectural link is chosen with no filter. |
| «Закройте файл … и запустите снова» (close the file and run again) | The building workbook is open in Excel. Close it and run again. |
| «Нет шаблона …» (template missing) | `Типовые_этажи_ШАБЛОН.xlsx` is not in the folder you set: it was deleted or renamed, or the path points elsewhere. |
| «Папка отчётов не найдена» (reports folder not found) | The folder is unreachable or the path is mistyped. Check that Explorer can open it. |
| Zeros or blanks in the report | Press F9. Check that the architectural parameters named in НАСТРОЙКИ are filled, and that room names match «Настройки». |
| Warning «Use ExportToExcel instead» | Not a failure. It does not change the result. |

## 8. Limits

- Microsoft Excel must be installed (`Data.ExportExcel`).
- Revit 2024, Python node engine CPython3.
- Up to 15 apartments per floor, 80 floors, 1500 apartments, 3000 export rows per building.
- No macros. The `.xlsx` opens in Excel 2010 and newer.
- The script only reads rooms in the architectural link. It does not save or sync the HVAC model.
