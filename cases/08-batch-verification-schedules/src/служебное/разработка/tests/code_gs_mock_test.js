// Мок-тест служебное/google_push/Code.gs без Google: имитирует SpreadsheetApp достаточно, чтобы
// проверить многократную отправку (глубина группировки строк <= 8), идемпотентность, цвета, типы значений.
// Запуск: node code_gs_mock_test.js   (код выхода 0 — всё хорошо)
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const CODE = path.join(__dirname, '..', '..', '..', 'служебное', 'google_push', 'Code.gs');
const FIX = path.join(__dirname, 'fixtures');

function a1ToRC(a1) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return [parseInt(m[2], 10), c];
}

// как Google Sheets разбирает введённое значение в ячейке НЕ с форматом "@"
function autoParse(v) {
  if (typeof v !== 'string') return v;
  if (v.charAt(0) === '=') return { formula: v };
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return { date: v };
  if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
  return v;
}

function makeEnv(opts) {
  opts = opts || {};
  const MAX_DEPTH = 8;
  class Sheet {
    constructor(name) {
      this.name = name; this.rows = 1000; this.cols = 26;
      this.cells = new Map(); // "r,c" -> {v, nf, bg, bold, wrap, ha}
      this.depth = new Array(this.rows).fill(0); // глубина группы по строкам (clear() её НЕ трогает — как в Google)
      this.merges = []; this.frozen = 0; this.hidden = false; this.collapsed = false;
    }
    cell(r, c) { const k = r + ',' + c; if (!this.cells.has(k)) this.cells.set(k, {}); return this.cells.get(k); }
    getName() { return this.name; }
    getProtections() { return []; }
    getMaxRows() { return this.rows; }
    getMaxColumns() { return this.cols; }
    getLastColumn() { let m = 0; for (const [k, v] of this.cells) if (v.v !== undefined && v.v !== '') m = Math.max(m, +k.split(',')[1]); return m; }
    setFrozenRows(n) { if (n >= this.rows) throw new Error('cannot freeze all rows'); this.frozen = n; }
    setFrozenColumns() {}
    clear() { const keep = new Map(); for (const [k, v] of this.cells) if (v.dv) keep.set(k, { dv: v.dv }); this.cells = keep; this.merges = []; } // проверка данных clear() не снимает
    clearFormats() { for (const v of this.cells.values()) { delete v.bg; delete v.bold; delete v.wrap; delete v.ha; delete v.nf; } }
    deleteRows(start, n) {
      if (start <= this.frozen) throw new Error('cannot delete frozen rows');
      if (n >= this.rows - start + 1 && start === 1) throw new Error('cannot delete all rows');
      const nc = new Map();
      for (const [k, v] of this.cells) {
        const [r, c] = k.split(',').map(Number);
        if (r >= start && r < start + n) continue;
        nc.set((r >= start + n ? r - n : r) + ',' + c, v);
      }
      this.cells = nc; this.depth.splice(start - 1, n); this.rows -= n;
    }
    insertRowsAfter(after, n) {
      this.depth.splice(after, 0, ...new Array(n).fill(0)); this.rows += n;
      // как в Google: новые строки наследуют формат и проверку данных строки выше
      for (let c = 1; c <= this.cols; c++) { const src = this.cells.get(after + ',' + c); if (src && src.dv) for (let r = after + 1; r <= after + n; r++) this.cell(r, c).dv = src.dv; }
    }
    insertColumnsAfter(after, n) { this.cols += n; }
    getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
    getRangeList(list) { return new RangeList(this, list); }
    setRowHeight(r, h) { (this.heights = this.heights || {})[r] = h; }
    setColumnWidth(c, w) { (this.widths = this.widths || {})[c] = w; }
    collapseAllRowGroups() { this.collapsed = true; this.level = 0; }
    expandRowGroupsUpToDepth(d) { this.level = d; }
    isSheetHidden() { return this.hidden; }
    hideSheet() { this.hidden = true; }
  }
  class Range {
    constructor(sh, r, c, nr, nc) {
      if (r < 1 || c < 1 || r + nr - 1 > sh.rows || c + nc - 1 > sh.cols) {
        throw new Error('Range out of bounds: ' + [r, c, nr, nc] + ' sheet ' + sh.rows + 'x' + sh.cols);
      }
      Object.assign(this, { sh, r, c, nr, nc });
    }
    each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.sh.cell(this.r + i, this.c + j), this.r + i, this.c + j); }
    setNumberFormat(f) { this.each(x => { x.nf = f; }); return this; }
    clearDataValidations() { this.each(x => { delete x.dv; }); }
    checkDv(x, v, r, c) {
      if (x.dv && !x.dv.allowInvalid && v !== '' && !x.dv.list.map(String).includes(String(v))) {
        throw new Error('The data that you entered in cell ' + r + ',' + c + ' violates the data validation rules set on this cell.');
      }
    }
    setValues(vs) {
      if (vs.length !== this.nr || vs.some(row => row.length !== this.nc)) throw new Error('setValues: dimension mismatch');
      this.each((x, r, c) => { const v = vs[r - this.r][c - this.c]; this.checkDv(x, v, r, c); x.v = x.nf === '@' && typeof v === 'string' ? v : autoParse(v); });
    }
    setValue(v) { this.each((x, r, c) => { this.checkDv(x, v, r, c); x.v = x.nf === '@' && typeof v === 'string' ? v : autoParse(v); }); return this; }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = this.sh.cell(this.r + i, this.c + j).v; row.push(v === undefined ? '' : v); } out.push(row); } return out; }
    breakApart() { this.sh.merges = []; }
    merge() { this.sh.merges.push([this.r, this.c, this.nr, this.nc]); if (this.nr * this.nc === 1) return; }
    getValue() { return this.getValues()[0][0]; }
    setBackground(h) { this.each(x => { x.bg = h; }); return this; }
    setFontWeight(w) { this.each(x => { x.bold = w === 'bold'; }); return this; }
    setFontColor(c) { this.each(x => { x.fc = c; }); return this; }
    setWrap(b) { this.each(x => { x.wrap = b; }); return this; }
    setHorizontalAlignment(a) { this.each(x => { x.ha = a; }); }
    setTextRotation(d) { this.each(x => { x.rot = d; }); }
    setVerticalAlignment() { return this; }
    setBorder(t, l, b, r) { (this.sh.borderCalls = this.sh.borderCalls || []).push({ r: this.r, c: this.c, nr: this.nr, nc: this.nc, t, l, b, rr: r }); }
    getDataValidation() { return this.sh.cell(this.r, this.c).dv || null; }
    setDataValidation(v) { this.each(x => { x.dv = v; }); }
    shiftRowGroupDepth(n) {
      for (let i = 0; i < this.nr; i++) {
        const idx = this.r - 1 + i; const d = this.sh.depth[idx] + n;
        if (d > MAX_DEPTH) throw new Error('Row group depth cannot exceed ' + MAX_DEPTH + '.');
        this.sh.depth[idx] = Math.max(0, d);
      }
    }
  }
  class RangeList {
    constructor(sh, list) { this.sh = sh; this.ranges = list.map(a1 => { const [r, c] = a1ToRC(a1); return new Range(sh, r, c, 1, 1); }); }
  }
  for (const m of ['setNumberFormat', 'setBackground', 'setFontWeight', 'setFontColor', 'setWrap', 'setHorizontalAlignment', 'setTextRotation', 'setVerticalAlignment']) {
    RangeList.prototype[m] = function (...a) { this.ranges.forEach(r => r[m](...a)); return this; };
  }
  const sheets = [];
  const ss = {
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = new Sheet(n); sheets.push(s); return s; },
    getSheets: () => sheets,
    getId: () => 'SHEET_ID',
    getEditors: () => [1, 2], getViewers: () => [],
  };
  const drive = { access: opts.access || 'PRIVATE', permission: opts.permission || 'NONE', calls: 0, fail: opts.driveFail || null };
  const props = opts.props || {};
  const sandbox = {
    console, JSON, Math, Object, String, Number, parseInt, isNaN, Date, Error, RegExp, Array,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {}, ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
      newDataValidation: () => { const v = {}; const b = { requireValueInList(l) { v.list = l; return b; }, setAllowInvalid(a) { v.allowInvalid = a; return b; }, setHelpText() { return b; }, build() { return v; } }; return b; }, BorderStyle: { SOLID_MEDIUM: 'SOLID_MEDIUM' } },
    ContentService: { MimeType: { JSON: 'JSON' }, createTextOutput: t => ({ text: t, setMimeType() { return this; }, getContent() { return t; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: { formatDate: () => '2026-01-01 00:00' },
    Session: { getScriptTimeZone: () => 'UTC' },
    SpreadsheetAppProtection: null,
    Logger: { log() {} },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', ANYONE: 'ANYONE', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', DOMAIN: 'DOMAIN', PRIVATE: 'PRIVATE' }, Permission: { VIEW: 'VIEW', COMMENT: 'COMMENT', EDIT: 'EDIT', NONE: 'NONE' },
      getFileById: () => ({
        getSharingAccess: () => drive.access, getSharingPermission: () => drive.permission,
        setSharing: (a, p) => { if (drive.fail) throw new Error(drive.fail); drive.calls++; drive.access = a; drive.permission = p; },
      }),
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(CODE, 'utf8'), sandbox, { filename: 'Code.gs' });
  // Code.gs объявляет var на верхнем уровне — доступны как свойства sandbox
  const post = (body) => {
    const out = sandbox.doPost({ postData: { contents: typeof body === 'string' ? body : JSON.stringify(body) } });
    return JSON.parse(out.text || out.getContent());
  };
  const edit = (sheet, r, c, value) => { sheet.cell(r, c).v = value; sandbox.onEdit({ range: { getSheet: () => sheet, getRow: () => r, getColumn: () => c, getValue: () => value } }); };
  return { sandbox, ss, sheets, post, edit, drive };
}

const SECRET = 'test-secret-123';
const payload = JSON.parse(fs.readFileSync(path.join(FIX, 'payload_multi.json'), 'utf8'));
const body = (extra) => Object.assign({ secret: SECRET, hide: ['Динамика'], sheets: payload.sheets.filter(s => s.name !== 'Динамика') }, extra || {});
function snapshot(sh) {
  const cells = [...sh.cells.entries()].map(([k, v]) => [k, JSON.stringify(v)]).sort((a, b) => a[0] < b[0] ? -1 : 1);
  return JSON.stringify({ cells, depth: sh.depth, merges: sh.merges, rows: sh.rows, frozen: sh.frozen });
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('12 отправок подряд: ошибки "row group depth" нет, глубина <= 2', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  for (let i = 1; i <= 12; i++) {
    const res = env.post(body());
    assert.strictEqual(res.ok, true, 'отправка #' + i + ': ' + JSON.stringify(res));
  }
  const eom = env.ss.getSheetByName('ЭОМ');
  assert.ok(Math.max(...eom.depth) === 2, 'ожидалась глубина 2, получено ' + Math.max(...eom.depth));
});

test('время последней выгрузки: пишется только при успешной отправке (для страницы просмотра)', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const get = () => env.sandbox.viewerGetVersion();
  assert.strictEqual(get(), 0);
  assert.strictEqual(env.post({ secret: 'wrong', sheets: [] }).ok, false);
  assert.strictEqual(env.post('{bad json').ok, false);
  assert.strictEqual(env.post({ secret: SECRET, sheets: [{ name: 'X', rows: 'нет' }] }).ok, false);
  assert.strictEqual(get(), 0, 'неудачные отправки время не меняют');
  const before = Date.now();
  const res = env.post(body());
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const t1 = get();
  assert.ok(t1 >= before && t1 <= Date.now(), String(t1));
  assert.deepStrictEqual(Object.keys(res).sort(), ['cleared', 'ok', 'sharing', 'sheets', 'version']);   // ответ doPost прежний
});

test('сбой записи свойства не роняет отправку данных', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.sandbox.PropertiesService.getScriptProperties = () => ({ getProperty: k => (k === 'PUSH_SECRET' ? SECRET : null), setProperty() { throw new Error('quota'); } });
  assert.strictEqual(env.post(body()).ok, true);
  assert.strictEqual(env.sandbox.viewerGetVersion(), 0);
});

test('идемпотентность: состояние после 1-й и 6-й отправки одинаково', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const first = env.sheets.map(snapshot);
  for (let i = 0; i < 5; i++) env.post(body());
  assert.deepStrictEqual(env.sheets.map(snapshot), first);
});

test('мок воспроизводит старую ошибку: clear() без удаления строк копит группы до отказа', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const sh = env.ss.insertSheet('X');
  assert.throws(() => { for (let i = 0; i < 9; i++) { sh.clear(); sh.getRange(5, 1, 3, 1).shiftRowGroupDepth(1); } }, /depth cannot exceed 8/);
});

test('цвета: Excel BGR-число -> #RRGGBB (R — младший байт)', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const g = env.sandbox.bgrToHex;
  assert.strictEqual(g(0x0000FF), '#ff0000'); // Excel: красный
  assert.strictEqual(g(0xFF0000), '#0000ff'); // синий
  assert.strictEqual(g(0x00FFFF), '#ffff00'); // жёлтый
  assert.strictEqual(g(0x80C0FF), '#ffc080'); // оранжевый «есть замечания»
  assert.strictEqual(g(0xCEEFC6), '#c6efce'); // зелёный
  assert.strictEqual(g(16777215), null);
  env.post(body());
  const sum = env.ss.getSheetByName('Сводка');
  const bgs = new Set([...sum.cells.values()].map(c => c.bg).filter(Boolean));
  assert.ok(bgs.has('#ffff00') && bgs.has('#ffc080'), 'в Сводке должны быть жёлтая шапка и оранжевые ячейки: ' + [...bgs]);
});

test('значения не искажаются: дата, 007, формула остаются текстом; целые — числом', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const res = env.post({ secret: SECRET, sheets: [{ name: 'ЭОМ', rows: [['Заметка', '', '', '', '', '', ''], ['Модель', 'Замечаний', 'Спец', 'П', 'С', 'Н', 'З'], ['M', '5', '2026-10-05', '007', '=1+1', 'Кириллица ёЁ', ' пробел ']] }] });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const sh = env.ss.getSheetByName('ЭОМ');
  const v = sh.getRange(3, 1, 1, 7).getValues()[0];
  assert.deepStrictEqual(v, ['M', 5, '2026-10-05', '007', '=1+1', 'Кириллица ёЁ', ' пробел ']);
});

test('неверный secret / не заданный secret -> ok:false, листы не тронуты', () => {
  let env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  let res = env.post(body({ secret: 'bad' }));
  assert.strictEqual(res.ok, false); assert.strictEqual(res.error, 'forbidden');
  assert.strictEqual(env.sheets.length, 0);
  env = makeEnv(); // нет Script Property -> остаётся REPLACE_ME в коде
  res = env.post(body({ secret: 'REPLACE_ME_WITH_LONG_RANDOM_SECRET' }));
  assert.strictEqual(res.ok, false); assert.match(res.error, /not configured/);
});

test('битый payload (rows — не массив массивов): ok:false, ни один лист не изменён', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const before = env.sheets.map(snapshot);
  const res = env.post({ secret: SECRET, sheets: [{ name: 'Сводка', rows: ['просто строка'] }, { name: 'ЭОМ', rows: [['a']] }] });
  assert.strictEqual(res.ok, false); assert.match(res.error, /array of arrays/);
  assert.deepStrictEqual(env.sheets.map(snapshot), before);
});

test('невалидный JSON -> ok:false (а не HTML/исключение)', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const res = env.post('{not json');
  assert.strictEqual(res.ok, false);
});

test('clear: раздел, которого нет в выгрузке, очищается', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const ss = env.ss.getSheetByName('ЭОМ');
  assert.ok(ss.getRange(3, 1, 1, 1).getValues()[0][0] !== '');
  const res = env.post({ secret: SECRET, clear: ['ЭОМ'], sheets: payload.sheets.filter(s => s.name === 'Сводка') });
  assert.strictEqual(res.ok, true, JSON.stringify(res)); assert.deepStrictEqual(res.cleared, ['ЭОМ']);
  assert.ok(String(ss.getRange(1, 1, 1, 1).getValues()[0][0]).startsWith('Замечаний нет'));
  assert.strictEqual(ss.getMaxRows(), 1, 'старые строки должны быть удалены');
  assert.ok(Math.max(...ss.depth) === 0);
});

test('payload из одной строки (пустая выгрузка) принимается', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const empty = JSON.parse(fs.readFileSync(path.join(FIX, 'payload_empty.json'), 'utf8'));
  const res = env.post({ secret: SECRET, sheets: empty.sheets.filter(s => s.name === 'Сводка') });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
});

test('A1 на листе раздела — список 1/2/3, по умолчанию 2; смена уровня работает как кнопки Excel', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const sh = env.ss.getSheetByName('ЭОМ');
  const a1 = sh.cell(1, 1);
  assert.strictEqual(JSON.stringify(a1.dv.list), '["1","2","3"]');
  assert.strictEqual(a1.v, 2);
  assert.strictEqual(sh.level, 1, 'по умолчанию раскрыты спецификации (глубина 1)');
  assert.ok(String(sh.cell(1, 2).v).includes('1 — только модели'));
  assert.deepStrictEqual(sh.merges.filter(m => m[0] === 1), [[1, 2, 1, 6]]);
  env.edit(sh, 1, 1, '1'); assert.strictEqual(sh.level, 0);
  env.edit(sh, 1, 1, '3'); assert.strictEqual(sh.level, 2);
  const sv = env.ss.getSheetByName('Сводка');
  env.edit(sv, 1, 1, '1'); assert.strictEqual(sv.level, undefined, 'Сводка не затрагивается');
  // повторная отправка возвращает вид по умолчанию
  env.post(body()); assert.strictEqual(sh.level, 1); assert.strictEqual(sh.cell(1, 1).v, 2);
});

test('ширины колонок из payload (по содержимому) применяются, длинные не режутся на 400px', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const res = env.post({ secret: SECRET, sheets: [{ name: 'ЭОМ', rows: [['n'], ['a', 'b']], columnWidths: [90, 12] }] });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const w = env.ss.getSheetByName('ЭОМ').widths;
  assert.ok(w[1] > 400 && w[1] <= 700, String(w[1]));
  assert.ok(w[2] < 100, String(w[2]));
});

test('Сводка: вертикальные заголовки спецификаций и высота шапки из meta', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const sv = env.ss.getSheetByName('Сводка');
  const rotated = [...sv.cells.entries()].filter(([k, v]) => v.rot === 90).map(([k]) => k);
  assert.ok(rotated.length >= 1 && rotated.every(k => k.startsWith('2,')), rotated.join(' '));
  const meta = payload.sheets.find(s => s.name === 'Сводка').meta;
  assert.ok(meta.headerHeight > 60);
  assert.strictEqual(sv.heights[2], Math.round(meta.headerHeight * 1.34));
});

test('регресс: 2-я и следующие отправки не падают из-за проверки данных в A1 (лист ЭОМ оставался пустым)', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  for (let i = 1; i <= 4; i++) {
    const res = env.post(body());
    assert.strictEqual(res.ok, true, 'отправка #' + i + ': ' + JSON.stringify(res));
  }
  const sh = env.ss.getSheetByName('ЭОМ');
  assert.strictEqual(sh.getRange(3, 1, 1, 1).getValues()[0][0], 'PROJ_IOS_EOM_ORG_K01_R24');
  assert.ok(sh.cell(1, 1).dv && !sh.cell(5, 1).dv, 'список только в A1');
});

test('доступ по ссылке: после отправки «все со ссылкой», повтор ничего не меняет', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  let res = env.post(body());
  assert.strictEqual(res.sharing, 'ok');
  assert.strictEqual(env.drive.access, 'ANYONE_WITH_LINK');
  assert.strictEqual(env.drive.permission, 'VIEW');
  env.post(body()); env.post(body());
  assert.strictEqual(env.drive.calls, 1, 'setSharing вызывается только когда доступ отличается');
  // если кто-то закрыл доступ — следующая отправка вернёт его
  env.drive.access = 'PRIVATE'; env.post(body());
  assert.strictEqual(env.drive.access, 'ANYONE_WITH_LINK');
});

test('запрет домена на публичный доступ не роняет отправку данных, а сообщается', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET }, driveFail: 'Access denied: DriveApp (sharing policy)' });
  const res = env.post(body());
  assert.strictEqual(res.ok, true);
  assert.match(res.sharing, /^error: .*policy/);
  assert.strictEqual(env.ss.getSheetByName('ЭОМ').getRange(3, 1, 1, 1).getValues()[0][0], 'PROJ_IOS_EOM_ORG_K01_R24');
});

test('шапка Сводки: у пар «Пустые»/«Без значения» убрана граница между колонками, после общих границ', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const sum = JSON.parse(JSON.stringify(payload.sheets.find(s => s.name === 'Сводка')));
  sum.meta.pairCols = [4];
  const res = env.post({ secret: SECRET, sheets: [sum] });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const calls = env.ss.getSheetByName('Сводка').borderCalls;
  const all = calls.findIndex(x => x.r === 2 && x.nc > 1);                       // общие границы таблицы
  const right = calls.findIndex(x => x.r === 2 && x.c === 4 && x.nc === 1 && x.rr === false);
  const left = calls.findIndex(x => x.r === 2 && x.c === 5 && x.nc === 1 && x.l === false);
  assert.ok(all >= 0 && right > all && left > all, JSON.stringify(calls));
  assert.ok(calls[right].l === null && calls[right].t === null);
});

test('доступ по ссылке не понижается: «Редактор» по ссылке остаётся при LINK_SHARING=VIEW', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET }, access: 'ANYONE_WITH_LINK', permission: 'EDIT' });
  const res = env.post(body());
  assert.strictEqual(res.sharing, 'ok');
  assert.strictEqual(env.drive.permission, 'EDIT');
  assert.strictEqual(env.drive.calls, 0, 'setSharing не вызывался');
});

test('читатель по ссылке повышается до минимума; домен с правкой не трогается', () => {
  let env = makeEnv({ props: { PUSH_SECRET: SECRET }, access: 'ANYONE_WITH_LINK', permission: 'VIEW' });
  env.sandbox.LINK_SHARING = 'EDIT';
  env.post(body());
  assert.strictEqual(env.drive.permission, 'EDIT');
  env = makeEnv({ props: { PUSH_SECRET: SECRET }, access: 'DOMAIN_WITH_LINK', permission: 'EDIT' });
  const res = env.post(body());
  assert.match(res.sharing, /^ok: доступ по домену/);
  assert.strictEqual(env.drive.access, 'DOMAIN_WITH_LINK');
  assert.strictEqual(env.drive.calls, 0);
});

test('diagnoseAccess читает доступ, защиты и список в A1, ничего не меняет', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET }, access: 'ANYONE_WITH_LINK', permission: 'VIEW' });
  env.post(body());
  const text = env.sandbox.diagnoseAccess();
  assert.match(text, /Доступ по ссылке: ANYONE_WITH_LINK \/ VIEW/);
  assert.match(text, /Лист «ЭОМ»: защит листа 0.*список в A1 есть/);
  assert.strictEqual(env.drive.calls, 0);
});

test('dim: ячейки непроверенных пар (выборочная выгрузка) получают серый шрифт, остальные нет', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const sum = JSON.parse(JSON.stringify(payload.sheets.find(s => s.name === 'Сводка')));
  sum.formats = sum.formats.concat([{ r: 3, c: 2, ha: 'center', dim: true }, { r: 3, c: 3, ha: 'center' }]);
  const res = env.post({ secret: SECRET, sheets: [sum] });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const sh = env.ss.getSheetByName('Сводка');
  assert.strictEqual(sh.cell(3, 2).fc, '#999999');
  assert.strictEqual(sh.cell(3, 3).fc, undefined);
});

test('выборочная отправка: лист, не попавший в payload и не названный в clear, не меняется', () => {
  const env = makeEnv({ props: { PUSH_SECRET: SECRET } });
  env.post(body());
  const eom = env.ss.getSheetByName('ЭОМ');
  const before = snapshot(eom);
  const res = env.post({ secret: SECRET, partial: true, clear: [], sheets: payload.sheets.filter(s => s.name === 'Сводка') });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  assert.deepStrictEqual(res.cleared, []);
  assert.strictEqual(snapshot(eom), before, 'лист ЭОМ не должен меняться');
});

test('лист раздела без колонки «Замечаний»: группы строк строятся по заголовку «Спецификация», как в прежней раскладке', () => {
  const oldEnv = makeEnv({ props: { PUSH_SECRET: SECRET } });
  oldEnv.post(body());
  const dropCol2 = (sheet) => Object.assign({}, sheet, {
    rows: sheet.rows.map(r => r.filter((_, i) => i !== 1)),
    formats: (sheet.formats || []).filter(f => f.c !== 2).map(f => Object.assign({}, f, { c: f.c > 2 ? f.c - 1 : f.c })),
    columnWidths: (sheet.columnWidths || []).filter((_, i) => i !== 1),
  });
  const newEnv = makeEnv({ props: { PUSH_SECRET: SECRET } });
  const eom = payload.sheets.find(s => s.name === 'ЭОМ');
  const res = newEnv.post({ secret: SECRET, sheets: [dropCol2(eom)] });
  assert.strictEqual(res.ok, true, JSON.stringify(res));
  const a = oldEnv.ss.getSheetByName('ЭОМ'), b = newEnv.ss.getSheetByName('ЭОМ');
  assert.deepStrictEqual(b.depth, a.depth, 'вложенность групп должна совпасть');
  assert.ok(Math.max(...b.depth) >= 1);
  assert.strictEqual(b.cell(1, 1).v, 2);                       // A1 — уровень по умолчанию 2
});

let failed = 0;
for (const [name, fn] of tests) {
  try { fn(); console.log('ok   - ' + name); } catch (e) { failed++; console.log('FAIL - ' + name + '\n       ' + (e && e.message)); }
}
console.log(failed ? failed + ' тест(ов) упало' : 'все ' + tests.length + ' тестов прошли');
process.exit(failed ? 1 : 0);
