// Мок-тест серверной части веб-страницы (doGet, viewerGetSection, viewerGetSummary в Code.gs) без Google.
// Проверяет: ping, список разделов, параметр ?section=, структуру строк (модель/спецификация/элемент), цвета и флаги,
// отказ для «Сводка»/скрытых/несуществующих листов в viewerGetSection и — главное — что нет НИ ОДНОГО вызова записи в таблицу.
// Запуск: node viewer_mock_test.js (код выхода 0 — всё хорошо)
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const { DIR, CODE, calls, makeSheet, section, env } = require('./viewer_fixture');

const EOM = makeSheet('ЭОМ', section([
  { name: 'PROJ_IOS_EOM_K01', specs: [
    { name: 'П        5.7_Ширина_Лотки', param: 'ADSK_Ширина', items: ['Лоток 1', 'Лоток 2'] },
    { name: 'П        6.7_Высота_Лотки', param: 'ADSK_Высота', items: ['Лоток 3'], dim: true },
  ] },
  { name: 'PROJ_IOS_EOM_K02', specs: [{ name: 'П        5.7_Ширина_Лотки', param: 'ADSK_Ширина', items: ['Лоток 9'] }] },
]));
const SS = makeSheet('СС', [[{ v: 'Замечаний нет (выгрузка 2026-10-21 10:00)' }]]);
const OV = makeSheet('ОВ', section([{ name: 'M', specs: [{ name: 'S', param: 'P', items: ['x'] }] }]));
const HIDDEN = makeSheet('ПД', section([]), { hidden: true });
const DYN = makeSheet('Динамика', section([]), { hidden: true });
const SUM = makeSheet('Сводка', [
  [{ v: 'Сводка' }, { v: 'Заметка сводки' }],
  [{ v: 'Модель', bg: '#cccccc', bold: true }, { v: 'П 5.1_Ширина', rot: 90, bg: '#cccccc' }, { v: 'Проверено' }],
  [{ v: 'PROJ_IOS_EOM_K01', bg: '#99ccff' }, { v: 3, bg: '#ff9999' }, { v: '2026-10-20' }],
]);
const all = [SUM, EOM, SS, OV, HIDDEN, DYN];

let failures = 0;
function t(name, fn) {
  try { fn(); console.log('ok   - ' + name); } catch (e) { failures++; console.log('FAIL - ' + name + '\n' + (e && e.stack || e)); }
}

t('ping по-прежнему отвечает ok (проверить_google.ps1)', () => {
  const { sandbox } = env(all);
  assert.strictEqual(sandbox.doGet({ parameter: { ping: '1' } }).getContent(), 'ok');
});

t('doGet отдаёт страницу: разделы без Сводка, Динамика и скрытых, параметр section, ссылка', () => {
  const { sandbox, evaluated } = env(all);
  const out = sandbox.doGet({ parameter: { section: 'ЭОМ' } });
  assert.strictEqual(out.frame, 'ALLOWALL');
  assert.ok(out.html.includes('<title>'));
  const m = /var INIT = (\{.*?\});\r?\n/s.exec(out.html);
  const init = JSON.parse(m[1]);
  assert.deepStrictEqual(Array.from(init.sections), ['ЭОМ', 'СС', 'ОВ']);
  assert.strictEqual(init.section, 'ЭОМ');
  assert.ok(init.baseUrl.endsWith('/exec'));
  assert.strictEqual(evaluated.length, 1);
  // без параметра и без e (запуск из редактора)
  assert.strictEqual(JSON.parse(/var INIT = (\{.*?\});\r?\n/s.exec(sandbox.doGet({ parameter: {} }).html)[1]).section, '');
  assert.ok(sandbox.doGet().html);
});

t('имя листа с «<» не ломает встраиваемый JSON', () => {
  const bad = makeSheet('</script><b>', section([]));
  const { sandbox } = env([bad]);
  const html = sandbox.doGet({ parameter: { section: '</script>' } }).html;
  assert.ok(!html.includes('</script><b>'));
  const init = JSON.parse(/var INIT = (\{.*?\});\r?\n/s.exec(html)[1]);
  assert.strictEqual(init.sections[0], '</script><b>');
});

t('viewerGetSection: уровни, цвета, флаги, шапка, ширины', () => {
  const { sandbox } = env(all);
  const d = JSON.parse(JSON.stringify(sandbox.viewerGetSection('ЭОМ')));
  assert.strictEqual(d.kind, 'section');
  assert.strictEqual(d.cols, 4);
  assert.strictEqual(d.specCol, 1);
  assert.deepStrictEqual(d.header.v, ['Модель', 'Спецификация', 'Параметр', 'Наименование']);
  assert.deepStrictEqual(d.rows.map(r => r.lv), [0, 1, 2, 2, 1, 2, 0, 1, 2]);
  assert.strictEqual(d.rows[0].v[0], 'PROJ_IOS_EOM_K01');
  assert.strictEqual(d.palette[d.rows[0].b[0]], '#99ccff');         // цвет строки модели
  assert.strictEqual(d.palette[d.rows[1].b[1]], '#ffffcc');         // цвет строки спецификации
  assert.strictEqual(d.rows[0].f[0], 1);                              // полужирный
  assert.strictEqual(d.rows[5].f[3], 2);                              // серый шрифт — не проверялось
  assert.strictEqual(d.rows[2].b, undefined);                         // без заливки — без массива
  assert.strictEqual(d.widths.length, 4);
  assert.strictEqual(d.palette[0], '');
});

t('viewerGetSection: «Замечаний нет» и пустой раздел', () => {
  const { sandbox } = env(all);
  const d = JSON.parse(JSON.stringify(sandbox.viewerGetSection('СС')));
  assert.ok(d.empty && d.note.startsWith('Замечаний нет'));
  assert.strictEqual(d.rows.length, 0);
});

t('viewerGetSection отказывает для Сводка, Динамика, скрытых и неизвестных листов', () => {
  const { sandbox } = env(all);
  for (const n of ['Сводка', 'Динамика', 'ПД', 'Нет такого', '', 'constructor']) {
    assert.throws(() => sandbox.viewerGetSection(n), /Нет раздела/, n);
  }
});

t('viewerGetSummary: плоская таблица, цвета, поворот шапки, заметка', () => {
  const { sandbox } = env(all);
  const d = JSON.parse(JSON.stringify(sandbox.viewerGetSummary()));
  assert.strictEqual(d.kind, 'summary');
  assert.strictEqual(d.note, 'Заметка сводки');
  assert.deepStrictEqual(d.header.v, ['Модель', 'П 5.1_Ширина', 'Проверено']);
  assert.deepStrictEqual(d.rot, [false, true, false]);
  assert.strictEqual(d.rows.length, 1);
  assert.strictEqual(d.rows[0].v[1], '3');
  assert.strictEqual(d.palette[d.rows[0].b[1]], '#ff9999');
  assert.strictEqual(d.rows[0].lv, undefined);
});

t('чтение 2000 строк: единицы вызовов API, а не по ячейкам', () => {
  const models = [];
  for (let m = 0; m < 40; m++) models.push({ name: 'M' + m, specs: [1, 2, 3, 4].map(s => ({ name: 'S' + s, param: 'P', items: new Array(12).fill('e') })) });
  const big = makeSheet('ВК', section(models));
  const { sandbox } = env([big]);
  calls.length = 0;
  const d = sandbox.viewerGetSection('ВК');
  assert.ok(d.rows.length > 2000, String(d.rows.length));
  assert.ok(calls.length < 30, 'вызовов API: ' + calls.length);
});

t('никаких записей: запретные методы не вызывались (guard бросил бы исключение)', () => {
  const { sandbox } = env(all);
  sandbox.viewerGetSection('ЭОМ'); sandbox.viewerGetSummary(); sandbox.doGet({ parameter: {} });
  const src = CODE.slice(CODE.indexOf('// ---------- Веб-страница'), CODE.indexOf('function validateSheets_'));
  assert.ok(!/\.(setValue|setValues|clear|merge|insertSheet|deleteRows|hideSheet|setSharing|appendRow)\(/.test(src));
});

t('viewerGetVersion: только время выгрузки из свойств, листы не читаются', () => {
  const { sandbox, props } = env(all);
  assert.strictEqual(sandbox.viewerGetVersion(), 0);
  props.LAST_PUSH_MS = '1792580000000';
  calls.length = 0;
  assert.strictEqual(sandbox.viewerGetVersion(), 1792580000000);
  assert.strictEqual(calls.length, 0, 'обращений к листам быть не должно: ' + calls.join());
  props.LAST_PUSH_MS = 'мусор';
  assert.strictEqual(sandbox.viewerGetVersion(), 0);
  props.LAST_PUSH_MS = '-5';
  assert.strictEqual(sandbox.viewerGetVersion(), 0);
});

t('данные раздела и Сводки несут версию выгрузки (читается до листа)', () => {
  const { sandbox, props } = env(all);
  props.LAST_PUSH_MS = '1792580000000';
  assert.strictEqual(sandbox.viewerGetSection('ЭОМ').version, 1792580000000);
  assert.strictEqual(sandbox.viewerGetSummary().version, 1792580000000);
  assert.strictEqual(sandbox.viewerGetSection('СС').version, 1792580000000);
});

t('Viewer.html вызывает на сервере только viewerGet*', () => {
  const html = fs.readFileSync(path.join(DIR, 'Viewer.html'), 'utf8');
  const fns = new Set((html.match(/\.viewer\w+\(/g) || []).map(x => x.slice(1, -1)));
  assert.deepStrictEqual(Array.from(fns).sort(), ['viewerGetSection', 'viewerGetSummary', 'viewerGetVersion']);
  assert.ok(!/google\.script\.run[^;]*\.(push|doPost|write)/.test(html));
});

if (failures) { console.log(failures + ' failed'); process.exit(1); }
console.log('all ok');
