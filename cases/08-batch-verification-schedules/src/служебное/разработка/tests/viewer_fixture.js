// Подставные листы и окружение Apps Script для тестов веб-страницы (Code.gs + Viewer.html). Только чтение.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = path.join(__dirname, '..', '..', '..', 'служебное', 'google_push');
const CODE = fs.readFileSync(path.join(DIR, 'Code.gs'), 'utf8');

// разрешены только методы чтения; любой другой вызов у листа/диапазона — провал теста
const READ_SHEET = new Set(['getName', 'getLastRow', 'getLastColumn', 'getRange', 'getColumnWidth', 'isSheetHidden']);
const READ_RANGE = new Set(['getDisplayValues', 'getBackgrounds', 'getFontWeights', 'getFontColors', 'getTextRotations']);
const calls = [];
function guard(obj, allowed, label) {
  return new Proxy(obj, {
    get(t, k) {
      if (typeof k === 'symbol' || k === 'then' || k === 'toJSON') return t[k];
      if (typeof t[k] === 'function' || !(k in t)) {
        if (!allowed.has(k)) throw new Error('ЗАПИСЬ/неразрешённый вызов ' + label + '.' + String(k));
        calls.push(label + '.' + String(k));
      }
      return t[k];
    },
    set() { throw new Error('запись свойства в ' + label); },
  });
}

function makeSheet(name, grid, opts) {
  opts = opts || {};
  const R = grid.length, C = Math.max.apply(null, grid.map(r => r.length));
  const at = (r, c) => (grid[r] && grid[r][c]) || {};
  const sheet = {
    getName: () => name,
    getLastRow: () => R,
    getLastColumn: () => C,
    isSheetHidden: () => !!opts.hidden,
    getColumnWidth: c => 100 + c,
    getRange(r, c, nr, nc) {
      const mat = fn => { const o = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(fn(at(r - 1 + i, c - 1 + j))); o.push(row); } return o; };
      return guard({
        getDisplayValues: () => mat(x => (x.v === undefined ? '' : String(x.v))),
        getBackgrounds: () => mat(x => x.bg || '#ffffff'),
        getFontWeights: () => mat(x => (x.bold ? 'bold' : 'normal')),
        getFontColors: () => mat(x => x.fc || '#000000'),
        getTextRotations: () => mat(x => ({ getDegrees: () => (x.rot || 0) })),
      }, READ_RANGE, name + '.range');
    },
  };
  return guard(sheet, READ_SHEET, name);
}

function section(models) {
  // как лист раздела: A1 — список уровня, B1 — заметка; строка 2 — шапка; затем модель / спецификация / элементы
  const g = [[{ v: 2 }, { v: 'Заметка' }], [{ v: 'Модель', bg: '#cccccc', bold: true }, { v: 'Спецификация', bg: '#cccccc', bold: true }, { v: 'Параметр' }, { v: 'Наименование' }]];
  for (const m of models) {
    g.push([{ v: m.name, bg: '#99ccff', bold: true }, { v: '' }, { v: '' }, { v: '' }]);
    for (const s of m.specs) {
      g.push([{ v: '' }, { v: s.name, bg: '#ffffcc' }, { v: s.param }, { v: '' }]);
      for (const it of s.items) g.push([{ v: '' }, { v: '' }, { v: '' }, { v: it, fc: s.dim ? '#999999' : undefined }]);
    }
  }
  return g;
}

function env(sheets) {
  const evaluated = [];
  const props = {};
  const ss = { getSheets: () => sheets, getSheetByName: n => sheets.find(s => s.getName() === n) || null, getId: () => 'ID' };
  const outputs = [];
  const sandbox = {
    console, JSON, Math, Object, String, Number, parseInt, isNaN, Date, Error, RegExp, Array,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (props[k] === undefined ? null : props[k]), setProperty: (k, v) => { props[k] = String(v); } }) },
    ContentService: { MimeType: { JSON: 'JSON' }, createTextOutput: t => ({ text: t, setMimeType() { return this; }, getContent() { return t; } }) },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ABC/exec' }) },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutput: t => ({ html: t }),
      createTemplateFromFile: f => {
        const tpl = { file: f };
        tpl.evaluate = () => {
          let html = fs.readFileSync(path.join(DIR, f + '.html'), 'utf8').replace('<?!= init ?>', tpl.init);
          const out = { html, setTitle(t) { out.title = t; return out; }, addMetaTag(a, b) { out.meta = [a, b]; return out; }, setXFrameOptionsMode(m) { out.frame = m; return out; } };
          evaluated.push(out);
          return out;
        };
        return tpl;
      },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: 'Code.gs' });
  return { sandbox, evaluated, props };
}


module.exports = { DIR, CODE, calls, guard, makeSheet, section, env };
