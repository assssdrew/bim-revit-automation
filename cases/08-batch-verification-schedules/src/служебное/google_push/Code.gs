/**
 * Один раз: Расширения → Apps Script → вставить этот файл → Развернуть → Веб-приложение.
 * Доступ: только вы (или все с аккаунтом). POST из PowerShell push_google.ps1.
 * В google_push.cfg укажите URL веб-приложения и тот же secret.
 *
 * SECRET: лучше не хранить в коде. Проект Apps Script → Настройки проекта → Свойства скрипта →
 * добавить PUSH_SECRET = <длинная случайная строка>. Значение в коде (ниже) — запасной вариант.
 * Не коммитьте реальный secret в git.
 *
 * Идемпотентность: каждая отправка полностью пересоздаёт содержимое листов (значения, формат,
 * объединения и группировку строк), поэтому повтор отправки даёт тот же результат.
 *
 * Ответ — всегда JSON: {"ok":true,...} или {"ok":false,"error":"..."}.
 */
var PUSH_SECRET = 'REPLACE_ME_WITH_LONG_RANDOM_SECRET';
var SCRIPT_VERSION = 3;
var MAX_GROUP_DEPTH = 8;
// Доступ к таблице по ссылке для всех, у кого она есть (без входа в аккаунт Google не нужен, если политика домена позволяет).
// Это МИНИМУМ: уже выставленный более широкий доступ (например «Редактор» по ссылке) скрипт НИКОГДА не понижает.
// 'VIEW' — минимум «Читатель», 'EDIT' — «Редактор» (нужен, чтобы проектировщики пользовались списком уровней 1/2/3 в A1),
// '' — доступ не трогать. Таблицу целиком перезаписывает скрипт, правки в листах разделов при следующей отправке теряются.
var LINK_SHARING = 'VIEW';
// целое без ведущих нулей — число, остальное остаётся текстом как есть (как в xlsx)
var INT_RE = /^(0|[1-9][0-9]{0,10})$/;

function getPushSecret_() {
  var v = '';
  try {
    v = PropertiesService.getScriptProperties().getProperty('PUSH_SECRET') || '';
  } catch (e) {}
  return v || PUSH_SECRET;
}

function permRank_(p) {
  if (p === DriveApp.Permission.EDIT) {
    return 3;
  }
  if (p === DriveApp.Permission.COMMENT) {
    return 2;
  }
  if (p === DriveApp.Permission.VIEW) {
    return 1;
  }
  return 0;
}

// Идемпотентно и без понижения: раньше setSharing(ANYONE_WITH_LINK, VIEW) превращал уже выставленного «Редактора»
// по ссылке в «Читателя» — у проектировщиков пропадали правка A1 (список уровней) и работа с группами.
// Ошибку (политика домена, нет прав владельца) не бросает, а возвращает текстом: отправка данных не падает из-за доступа.
function ensureLinkSharing_(ss) {
  if (!LINK_SHARING) {
    return 'skipped';
  }
  try {
    var file = DriveApp.getFileById(ss.getId());
    var want = LINK_SHARING === 'EDIT' ? DriveApp.Permission.EDIT : DriveApp.Permission.VIEW;
    var access = file.getSharingAccess();
    var perm = file.getSharingPermission();
    if (access === DriveApp.Access.ANYONE_WITH_LINK || access === DriveApp.Access.ANYONE) {
      if (permRank_(perm) >= permRank_(want)) {
        return 'ok';
      }
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, want);
      return 'ok';
    }
    if ((access === DriveApp.Access.DOMAIN_WITH_LINK || access === DriveApp.Access.DOMAIN) && permRank_(perm) > permRank_(want)) {
      // у домена права выше, чем дал бы «все по ссылке»: не понижаем и не расширяем до публичных правок молча
      return 'ok: доступ по домену (' + perm + ') оставлен; чтобы открыть всем по ссылке с правкой, задайте LINK_SHARING = \'EDIT\' или откройте вручную';
    }
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, want);
    return 'ok';
  } catch (e) {
    return 'error: ' + (e && e.message ? e.message : e);
  }
}

// Диагностика «у проектировщиков не работают «+» и список 1/2/3»: запустить вручную из редактора (Выполнить → diagnoseAccess),
// смотреть журнал. Показывает доступ по ссылке, защиты листов/диапазонов и проверку данных в A1. Ничего не меняет.
function diagnoseAccess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [];
  try {
    var file = DriveApp.getFileById(ss.getId());
    lines.push('Доступ по ссылке: ' + file.getSharingAccess() + ' / ' + file.getSharingPermission());
    lines.push('Редакторов: ' + ss.getEditors().length + ', читателей: ' + ss.getViewers().length);
  } catch (e) {
    lines.push('Доступ не прочитан: ' + e);
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var info = 'Лист «' + sh.getName() + '»';
    try {
      var sp = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      var rp = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      info += ': защит листа ' + sp.length + ', защит диапазонов ' + rp.length;
      for (var k = 0; k < sp.length; k++) {
        info += sp[k].isWarningOnly() ? ' [лист: только предупреждение]' : ' [ЛИСТ ЗАЩИЩЁН, редакторов ' + sp[k].getEditors().length + ']';
      }
    } catch (e2) {
      info += ': защиты не прочитаны (' + e2 + ')';
    }
    if (sh.getName() !== 'Сводка') {
      try {
        info += sh.getRange(1, 1).getDataValidation() ? ', список в A1 есть' : ', списка в A1 НЕТ';
      } catch (e3) {}
    }
    lines.push(info);
  }
  lines.push('Если у листа есть защита (не «только предупреждение»), редакторы без права на неё не смогут раскрывать группы строк и выбирать уровень: Данные → Защищённые листы и диапазоны.');
  lines.push('Если доступ по ссылке «Читатель» — список в A1 недоступен (читатель не может менять ячейки): задайте LINK_SHARING = \'EDIT\' или используйте Вид → Группа → Развернуть/Свернуть все группы строк.');
  var text = lines.join('\n');
  Logger.log(text);
  return text;
}

// Запустить ОДИН РАЗ вручную из редактора Apps Script (Выполнить → setLinkSharing): выдаёт скрипту право на Диск
// и сразу открывает таблицу по ссылке. После этого — «Развернуть → Управление развёртываниями → новая версия».
function setLinkSharing() {
  var res = ensureLinkSharing_(SpreadsheetApp.getActiveSpreadsheet());
  Logger.log(res);
  return res;
}

// Время последней успешной выгрузки (мс с 1970, UTC) — в свойствах скрипта. Страница просмотра по нему узнаёт о новых данных.
// Сбой записи свойства не должен ронять отправку данных.
var LAST_PUSH_KEY_ = 'LAST_PUSH_MS';

function markPushed_() {
  try {
    PropertiesService.getScriptProperties().setProperty(LAST_PUSH_KEY_, String(new Date().getTime()));
  } catch (e) {}
}

function viewerLastPush_() {
  try {
    var n = Number(PropertiesService.getScriptProperties().getProperty(LAST_PUSH_KEY_));
    return isNaN(n) || n < 0 ? 0 : n;
  } catch (e) {
    return 0;
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Цвет Excel (.Interior.Color): число R + G*256 + B*65536, то есть R — МЛАДШИЙ байт.
function bgrToHex(bgr) {
  if (bgr === null || bgr === undefined || bgr === '' || bgr === 16777215) {
    return null;
  }
  var n = parseInt(bgr, 10);
  if (isNaN(n) || n < 0) {
    return null;
  }
  var r = n & 0xff;
  var g = (n >> 8) & 0xff;
  var b = (n >> 16) & 0xff;
  var hx = function (x) {
    var s = x.toString(16);
    return s.length < 2 ? '0' + s : s;
  };
  return '#' + hx(r) + hx(g) + hx(b);
}

function colLetter_(c) {
  var s = '';
  while (c > 0) {
    var rem = (c - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function padSheetRows(rows) {
  var colCount = 0;
  for (var r = 0; r < rows.length; r++) {
    colCount = Math.max(colCount, rows[r] ? rows[r].length : 0);
  }
  for (var r = 0; r < rows.length; r++) {
    var line = rows[r] || [];
    while (line.length < colCount) {
      line.push('');
    }
    rows[r] = line;
  }
  return rows;
}

// Приводит ячейки к типам: целые — числом, всё остальное — строкой.
function typeRows_(rows) {
  var out = [];
  for (var r = 0; r < rows.length; r++) {
    var line = [];
    for (var c = 0; c < rows[r].length; c++) {
      var v = rows[r][c];
      if (v === null || v === undefined) {
        v = '';
      }
      var s = String(v);
      line.push(INT_RE.test(s) ? Number(s) : s);
    }
    out.push(line);
  }
  return out;
}

function setNumberFormatChunked_(sh, a1List, fmt) {
  for (var i = 0; i < a1List.length; i += 200) {
    sh.getRangeList(a1List.slice(i, i + 200)).setNumberFormat(fmt);
  }
}

// clear() НЕ снимает группировку строк: раньше она копилась с каждой отправкой,
// и после ~5 отправок Google отвечал "row group depth cannot exceed 8".
// Удаление строк убирает их группы; строка 1 группой не бывает.
function resetSheet_(sh) {
  try {
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);
  } catch (e) {}
  var maxR = sh.getMaxRows();
  var maxC = sh.getMaxColumns();
  try {
    sh.getRange(1, 1, maxR, maxC).breakApart();
  } catch (e) {}
  sh.clear();
  sh.clearFormats();
  // clear() не убирает проверку данных (выпадающий список 1/2/3 из A1). Она потом «размножается» на все новые строки
  // (insertRowsAfter копирует правила строки выше), и setValues падает на именах моделей: лист оставался пустым.
  try {
    sh.getRange(1, 1, maxR, maxC).clearDataValidations();
  } catch (e) {}
  if (maxR > 1) {
    sh.deleteRows(2, maxR - 1);
  }
  try {
    sh.getRange(1, 1, 1, 1).shiftRowGroupDepth(-MAX_GROUP_DEPTH);
  } catch (e) {}
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) {
    sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < cols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  }
}

function applyCellFormats(sh, formats) {
  if (!formats || formats.length < 1) {
    return;
  }
  var buckets = {};
  for (var i = 0; i < formats.length; i++) {
    var f = formats[i];
    var key = (f.bg || '') + '|' + (f.b ? 'b' : '') + '|' + (f.wrap ? 'w' : '') + '|' + (f.ha || '') + '|' + (f.rot || '') + '|' + (f.dim ? 'd' : '');
    if (!buckets[key]) {
      buckets[key] = { f: f, a1: [] };
    }
    buckets[key].a1.push(colLetter_(f.c) + f.r);
  }
  for (var key in buckets) {
    var item = buckets[key];
    var f = item.f;
    for (var i = 0; i < item.a1.length; i += 200) {
      var rl = sh.getRangeList(item.a1.slice(i, i + 200));
      if (f.bg) {
        var hex = bgrToHex(f.bg);
        if (hex) {
          rl.setBackground(hex);
        }
      }
      if (f.b) {
        rl.setFontWeight('bold');
      }
      if (f.dim) {
        // выборочная выгрузка: пара «модель + спецификация» в этой сборке не проверялась (данные прошлых проверок)
        rl.setFontColor('#999999');
      }
      if (f.wrap) {
        rl.setWrap(true);
      }
      if (f.ha === 'center') {
        rl.setHorizontalAlignment('center');
      }
      if (f.rot) {
        // вертикальный текст (названия спецификаций в шапке Сводки)
        rl.setTextRotation(parseInt(f.rot, 10));
        rl.setHorizontalAlignment('center');
        rl.setVerticalAlignment('bottom');
      }
    }
  }
}

function applyColumnWidths(sh, widths) {
  if (!widths || widths.length < 1) {
    return;
  }
  for (var c = 0; c < widths.length; c++) {
    var w = parseInt(widths[c], 10);
    if (!isNaN(w) && w > 0) {
      sh.setColumnWidth(c + 1, Math.min(700, Math.round(w * 7 + 5)));
    }
  }
}

function applySectionRowGroups(sh, rowCount) {
  if (rowCount < 4) {
    return;
  }
  var colCount = Math.max(6, sh.getLastColumn());
  // Кнопка +/- над группой (на строке модели/спецификации). По умолчанию Google рисует её ПОД группой,
  // и плюсы выглядели смещёнными на одну группу вниз.
  try {
    sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch (e) {}
  var values = sh.getRange(1, 1, rowCount, colCount).getValues();
  // колонка «Спецификация» — по заголовку (строка 2): в прежней раскладке с колонкой «Замечаний» это 3-я, теперь 2-я
  var specCol = 1;
  for (var hc = 0; hc < colCount; hc++) {
    if (values[1][hc] === 'Спецификация') {
      specCol = hc;
      break;
    }
  }
  var r = 3;
  while (r <= rowCount) {
    var model = values[r - 1][0];
    if (model) {
      var outerFirst = r + 1;
      var rr = outerFirst;
      while (rr <= rowCount) {
        var nextModel = values[rr - 1][0];
        if (nextModel && rr > r) {
          break;
        }
        var spec = values[rr - 1][specCol];
        if (!nextModel && spec) {
          var innerFirst = rr + 1;
          var k = innerFirst;
          while (k <= rowCount) {
            if (values[k - 1][0] || values[k - 1][specCol]) {
              break;
            }
            k++;
          }
          var innerLast = k - 1;
          if (innerLast >= innerFirst) {
            sh.getRange(innerFirst, 1, innerLast - innerFirst + 1, colCount).shiftRowGroupDepth(1);
          }
          rr = k;
        } else {
          rr++;
        }
      }
      var outerLast = rr - 1;
      if (outerLast >= outerFirst) {
        sh.getRange(outerFirst, 1, outerLast - outerFirst + 1, colCount).shiftRowGroupDepth(1);
      }
      r = rr;
    } else {
      r++;
    }
  }
  setOutlineLevel_(sh, 2);
}

// Уровень раскрытия как кнопки 1/2/3 в Excel: 1 — только модели, 2 — + спецификации, 3 — + элементы.
function setOutlineLevel_(sh, level) {
  var depth = Math.max(1, Math.min(3, parseInt(level, 10) || 2)) - 1;
  try {
    sh.expandRowGroupsUpToDepth(depth);
  } catch (e) {
    if (depth === 0) {
      try {
        sh.collapseAllRowGroups();
      } catch (e2) {}
    }
  }
}

var SECTION_NOTE_ = 'Вариант 1. Лист раздела. Выберите уровень раскрытия в ячейке слева: 1 — только модели, 2 — спецификации, 3 — элементы. Плюс слева раскрывает одну строку. Если списка нет (режим просмотра): Вид → Группа → Развернуть/Свернуть все группы строк.';

// A1 — выпадающий список 1/2/3 (аналог кнопок уровней в Excel), заметка переезжает в B1.
function applyLevelPicker_(sh, mergeCols) {
  var picker = sh.getRange(1, 1);
  picker.setNumberFormat('0');
  picker.setValue(2);
  picker.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3'], true).setAllowInvalid(false)
      .setHelpText('Уровень раскрытия: 1 — модели, 2 — спецификации, 3 — элементы').build()
  );
  picker.setBackground('#ffff00');
  picker.setFontWeight('bold');
  picker.setHorizontalAlignment('center');
  var note = sh.getRange(1, 2, 1, mergeCols - 1);
  note.merge();
  note.setValue(SECTION_NOTE_);
  note.setBackground('#ffff00');
  note.setWrap(true);
}

// Простой триггер: срабатывает при ручном выборе в A1 на листе раздела.
function onEdit(e) {
  try {
    if (!e || !e.range) {
      return;
    }
    var sh = e.range.getSheet();
    if (sh.getName() === 'Сводка' || e.range.getRow() !== 1 || e.range.getColumn() !== 1) {
      return;
    }
    var level = parseInt(e.range.getValue(), 10);
    if (level >= 1 && level <= 3) {
      setOutlineLevel_(sh, level);
    }
  } catch (err) {}
}

function applySummaryLayout(sh, rowCount, colCount, meta) {
  if (colCount < 1 || rowCount < 2) {
    return;
  }
  // Первая колонка (модели) закреплена. Google не даёт закрепить часть объединённой ячейки,
  // поэтому заметку переносим из A1 в объединённую B1:…, а в A1 — короткий заголовок.
  if (colCount > 2) {
    var noteText = sh.getRange(1, 1).getValue();
    var noteRange = sh.getRange(1, 2, 1, colCount - 1);
    noteRange.merge();
    sh.getRange(1, 2).setValue(noteText);
    noteRange.setBackground('#ffff00').setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, 1).setValue('Сводка').setFontWeight('bold').setVerticalAlignment('middle');
  } else {
    sh.getRange(1, 1, 1, colCount).merge();
  }
  sh.setRowHeight(1, 36);
  // высота шапки под самое длинное вертикальное название (пункты Excel -> пиксели)
  var hh = meta && parseInt(meta.headerHeight, 10);
  sh.setRowHeight(2, hh > 0 ? Math.round(hh * 1.34) : 90);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(1);
  if (meta && meta.dateCols) {
    var dc = parseInt(meta.dateCols, 10);
    if (!isNaN(dc) && dc > 0) {
      var firstSpec = 2 + dc;
      if (firstSpec <= colCount && !(hh > 0)) {
        sh.getRange(2, firstSpec, 1, colCount - firstSpec + 1).setWrap(true);
      }
    }
  }
}

// Пары «…_Пустые» / «…_Без значения» одной проверки: в шапке (строка 2) убираем границу между колонками пары.
// Вызывать ПОСЛЕ applyBorders. meta.pairCols — номера левых колонок пар.
function applyHeaderPairs_(sh, meta) {
  var cols = meta && meta.pairCols;
  if (!cols || !cols.length) {
    return;
  }
  for (var i = 0; i < cols.length; i++) {
    var c = parseInt(cols[i], 10);
    if (isNaN(c) || c < 1) {
      continue;
    }
    sh.getRange(2, c, 1, 1).setBorder(null, null, null, false, null, null);
    sh.getRange(2, c + 1, 1, 1).setBorder(null, false, null, null, null, null);
  }
}

function applySectionLayout(sh, rowCount, colCount) {
  if (colCount < 1 || rowCount < 2) {
    return;
  }
  var mergeCols = Math.max(colCount, 6);
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(2);
  applySectionRowGroups(sh, rowCount);
  applyLevelPicker_(sh, mergeCols);
}

function applyBorders(sh, rowCount, colCount, firstRow) {
  if (rowCount < firstRow || colCount < 1) {
    return;
  }
  sh.getRange(firstRow, 1, rowCount - firstRow + 1, colCount)
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ---------- Веб-страница для просмотра отчёта (только чтение) ----------
// Отдельное развёртывание «Веб-приложение» (выполнять от имени владельца, доступ — все): doGet отдаёт Viewer.html,
// данные страница берёт через google.script.run (viewerGetSection / viewerGetSummary). Ничего не пишет в таблицу:
// только getValues-подобные вызовы. Раскрытие строк у каждого своё — хранится в браузере (localStorage).
// Развёртывание для отправки данных (doPost) не затрагивается; ping=1 по-прежнему отвечает «ok».
var VIEWER_SKIP_SHEETS_ = { 'Сводка': true, 'Динамика': true };

// Листы разделов: все видимые, кроме «Сводка» и «Динамика».
function viewerSectionNames_(ss) {
  var names = [];
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var n = all[i].getName();
    if (VIEWER_SKIP_SHEETS_[n]) {
      continue;
    }
    try {
      if (all[i].isSheetHidden()) {
        continue;
      }
    } catch (e) {}
    names.push(n);
  }
  return names;
}

function viewerBaseUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.ping === '1') {
    return ContentService.createTextOutput('ok');
  }
  var init = { sections: [], section: String(p.section || ''), baseUrl: viewerBaseUrl_() };
  try {
    init.sections = viewerSectionNames_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    return HtmlService.createHtmlOutput('Не удалось открыть таблицу: ' + String(err && err.message ? err.message : err));
  }
  var tpl = HtmlService.createTemplateFromFile('Viewer');
  // JSON встраивается в <script>: «<» экранируем, чтобы имя листа не могло закрыть тег
  tpl.init = JSON.stringify(init).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return tpl
    .evaluate()
    .setTitle('Проверочные спецификации — отчёт')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Лист -> компактный объект для страницы. Только чтение. Цвета — палитрой (индекс 0 — без заливки), флаги ячейки:
// 1 — полужирный, 2 — серый шрифт («не проверялось в этой сборке»). Для листа раздела у каждой строки уровень lv:
// 0 — модель, 1 — спецификация, 2 — элемент (та же структура, что у групп строк на листе).
function viewerReadSheet_(sh, summary) {
  var out = {
    name: sh.getName(), kind: summary ? 'summary' : 'section', cols: 0, specCol: 1, note: '', empty: false,
    palette: [''], widths: [], header: null, rows: [], rot: null
  };
  var lastR = sh.getLastRow();
  var lastC = sh.getLastColumn();
  if (lastR < 1 || lastC < 1) {
    out.empty = true;
    return out;
  }
  var rng = sh.getRange(1, 1, lastR, lastC);
  var disp = rng.getDisplayValues();
  var bgs = rng.getBackgrounds();
  var wts = rng.getFontWeights();
  var fcs = rng.getFontColors();
  var index = { '': 0, '#ffffff': 0 };
  var encode = function (r) {
    var v = [];
    var b = [];
    var f = [];
    var anyB = false;
    var anyF = false;
    for (var c = 0; c < lastC; c++) {
      v.push(disp[r][c]);
      var hx = String(bgs[r][c] || '').toLowerCase();
      if (index[hx] === undefined) {
        index[hx] = out.palette.length;
        out.palette.push(hx);
      }
      b.push(index[hx]);
      if (index[hx]) {
        anyB = true;
      }
      var fl = (wts[r][c] === 'bold' ? 1 : 0) | (String(fcs[r][c] || '').toLowerCase() === '#999999' ? 2 : 0);
      f.push(fl);
      if (fl) {
        anyF = true;
      }
    }
    var o = { v: v };
    if (anyB) {
      o.b = b;
    }
    if (anyF) {
      o.f = f;
    }
    return o;
  };
  out.cols = lastC;
  for (var c = 1; c <= lastC; c++) {
    var w = 0;
    try {
      w = sh.getColumnWidth(c);
    } catch (e) {}
    out.widths.push(w || 0);
  }
  if (summary) {
    out.note = String(disp[0][1] || disp[0][0] || '');
  } else if (lastR < 3) {
    // «Замечаний нет (выгрузка …)»: лист раздела очищен
    var a1 = String(disp[0][0] || '');
    out.note = /^[123]$/.test(a1) ? '' : a1;
  }
  if (lastR < 2) {
    out.empty = true;
    return out;
  }
  out.header = encode(1);
  if (summary) {
    try {
      var rots = sh.getRange(2, 1, 1, lastC).getTextRotations();
      out.rot = [];
      for (var rc = 0; rc < lastC; rc++) {
        out.rot.push(rots[0][rc].getDegrees() !== 0);
      }
    } catch (e2) {
      out.rot = null;
    }
  } else {
    for (var hc = 0; hc < lastC; hc++) {
      if (disp[1][hc] === 'Спецификация') {
        out.specCol = hc;
        break;
      }
    }
  }
  for (var r = 2; r < lastR; r++) {
    var blank = true;
    for (var k = 0; k < lastC; k++) {
      if (disp[r][k] !== '') {
        blank = false;
        break;
      }
    }
    if (blank) {
      continue;
    }
    var row = encode(r);
    if (!summary) {
      row.lv = disp[r][0] !== '' ? 0 : (disp[r][out.specCol] !== '' ? 1 : 2);
    }
    out.rows.push(row);
  }
  return out;
}

// Вызывается страницей через google.script.run. Отдаёт только листы разделов (без «Сводка», «Динамика», скрытых).
function viewerGetSection(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  name = String(name);
  if (viewerSectionNames_(ss).indexOf(name) < 0) {
    throw new Error('Нет раздела «' + name + '»');
  }
  var version = viewerLastPush_();   // версия читается ДО листа: если выгрузка придёт во время чтения, страница лишь лишний раз предложит обновиться
  var data = viewerReadSheet_(ss.getSheetByName(name), false);
  data.version = version;
  return data;
}

function viewerGetSummary() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Сводка');
  if (!sh) {
    throw new Error('Нет листа «Сводка»');
  }
  var version = viewerLastPush_();
  var data = viewerReadSheet_(sh, true);
  data.version = version;
  return data;
}

// Лёгкий вызов для страницы (раз в 5 минут): только время последней выгрузки, листы не читаются. 0 — выгрузки с этой версией Code.gs ещё не было.
function viewerGetVersion() {
  return viewerLastPush_();
}

function validateSheets_(sheets) {
  if (Object.prototype.toString.call(sheets) !== '[object Array]') {
    throw new Error('sheets must be an array');
  }
  for (var i = 0; i < sheets.length; i++) {
    var spec = sheets[i];
    if (!spec || !spec.name) {
      throw new Error('sheet #' + (i + 1) + ' has no name');
    }
    var rows = spec.rows || [];
    if (Object.prototype.toString.call(rows) !== '[object Array]') {
      throw new Error('sheet ' + spec.name + ': rows must be an array');
    }
    for (var r = 0; r < rows.length; r++) {
      if (rows[r] !== null && Object.prototype.toString.call(rows[r]) !== '[object Array]') {
        throw new Error('sheet ' + spec.name + ': row ' + (r + 1) + ' is not an array (rows must be array of arrays)');
      }
    }
  }
}

function writeSheet_(ss, spec) {
  var name = spec.name;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  resetSheet_(sh);
  var rows = padSheetRows(spec.rows || []);
  var rowCount = rows.length;
  var colCount = 0;
  for (var ri = 0; ri < rowCount; ri++) {
    colCount = Math.max(colCount, rows[ri] ? rows[ri].length : 0);
  }
  if (rowCount > 0 && colCount > 0) {
    // +1 строка: закреплённые строки не могут занимать весь лист (иначе «cannot freeze all rows»)
    ensureSize_(sh, rowCount + 1, colCount);
    var typed = typeRows_(rows);
    // весь диапазон — «обычный текст», чтобы Google не превращал "2026-10-05" в дату,
    // "007" в 7, а "=..." в формулу; целые числа переводим в числовой формат
    sh.getRange(1, 1, rowCount, colCount).setNumberFormat('@');
    var numeric = [];
    for (var r = 0; r < typed.length; r++) {
      for (var c = 0; c < typed[r].length; c++) {
        if (typeof typed[r][c] === 'number') {
          numeric.push(colLetter_(c + 1) + (r + 1));
        }
      }
    }
    setNumberFormatChunked_(sh, numeric, '0');
    sh.getRange(1, 1, rowCount, colCount).setValues(typed);
  }
  applyColumnWidths(sh, spec.columnWidths);
  applyCellFormats(sh, spec.formats);
  if (name === 'Сводка') {
    applySummaryLayout(sh, rowCount, colCount, spec.meta || {});
    applyBorders(sh, rowCount, colCount, 2);
    applyHeaderPairs_(sh, spec.meta || {});
  } else {
    applySectionLayout(sh, rowCount, colCount);
    applyBorders(sh, rowCount, colCount, 2);
  }
  return { name: name, rows: rowCount, cols: colCount };
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var data = JSON.parse(e.postData.contents);
    var secret = getPushSecret_();
    if (!secret || /^REPLACE_ME/.test(secret)) {
      return jsonOut_({ ok: false, error: 'PUSH_SECRET is not configured in Code.gs / script properties' });
    }
    if (!data || data.secret !== secret) {
      return jsonOut_({ ok: false, error: 'forbidden' });
    }
    var sheets = data.sheets || [];
    validateSheets_(sheets);
    // две одновременные отправки (двойной клик, повтор по таймауту) не должны мешать друг другу
    lock.waitLock(60000);
    locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = [];
    for (var i = 0; i < sheets.length; i++) {
      try {
        written.push(writeSheet_(ss, sheets[i]));
      } catch (werr) {
        throw new Error('лист ' + sheets[i].name + ': ' + (werr && werr.message ? werr.message : werr));
      }
    }
    // разделы, которых нет в этой выгрузке: очищаем, чтобы не остались данные прошлой отправки
    var cleared = [];
    var clear = data.clear || [];
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    for (var ci = 0; ci < clear.length; ci++) {
      var csh = ss.getSheetByName(clear[ci]);
      if (csh) {
        resetSheet_(csh);
        csh.getRange(1, 1).setNumberFormat('@').setValue('Замечаний нет (выгрузка ' + stamp + ')');
        cleared.push(clear[ci]);
      }
    }
    var hide = data.hide || [];
    for (var h = 0; h < hide.length; h++) {
      var hidden = ss.getSheetByName(hide[h]);
      if (!hidden) {
        continue;
      }
      var visible = 0;
      var all = ss.getSheets();
      for (var s = 0; s < all.length; s++) {
        if (!all[s].isSheetHidden()) {
          visible++;
        }
      }
      if (visible > 1 || hidden.isSheetHidden()) {
        hidden.hideSheet();
      }
    }
    SpreadsheetApp.flush();
    var sharing = ensureLinkSharing_(ss);
    markPushed_();
    return jsonOut_({ ok: true, version: SCRIPT_VERSION, sheets: written, cleared: cleared, sharing: sharing });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}
