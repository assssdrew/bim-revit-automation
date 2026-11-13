// Браузерный тест страницы Viewer.html (Chromium через Playwright) на данных из настоящего Code.gs с подставными листами:
// дерево модель -> спецификация -> элементы, уровни 1/2/3, поиск, цвета, закреплённая шапка, Сводка, состояние только в браузере
// (у каждого пользователя своё, сохраняется между перезагрузками), быстрая загрузка ~2000 строк, текст ячеек не исполняется как HTML.
// Нет playwright/Chromium — тест пропускается (код выхода 0). Запуск: node viewer_browser_test.js
'use strict';
const path = require('path');
const assert = require('assert');
const { execSync } = require('child_process');
const { makeSheet, section, env } = require('./viewer_fixture');

function loadPlaywright() {
  for (const spec of ['playwright', 'playwright-core']) { try { return require(spec); } catch (e) { /* дальше */ } }
  try { return require(path.join(execSync('npm root -g').toString().trim(), 'playwright')); } catch (e) { return null; }
}
const pw = loadPlaywright();
if (!pw) { console.log('SKIP - нет playwright'); process.exit(0); }

const models = [];
for (let m = 0; m < 40; m++) models.push({ name: 'PROJ_IOS_VK_K' + m, specs: [1, 2, 3, 4].map(s => ({ name: 'П        ' + s + '_Спека', param: 'P' + s, items: new Array(12).fill(0).map((_, i) => 'Элемент ' + i) })) });
const BIG = makeSheet('ВК', section(models));
const EOM_DATA = section([
  { name: 'PROJ_IOS_EOM_K01', specs: [
    { name: 'П        5.7_Ширина_Лотки', param: 'ADSK_Ширина', items: ['Лоток 1', 'Лоток 2'] },
    { name: 'П        6.7_Высота_Лотки', param: 'ADSK_Высота', items: ['Лоток 3'], dim: true }] },
  { name: 'PROJ_IOS_EOM_K02', specs: [{ name: 'П        5.7_Ширина_Лотки', param: 'ADSK_Ширина', items: ['Лоток 9'] }] },
]);
const EOM = makeSheet('ЭОМ', EOM_DATA);
const XSS = makeSheet('ОВ', section([{ name: '<img src=x onerror="window.__xss=1">', specs: [{ name: '<b>S</b>', param: 'P', items: ['<script>window.__xss=2</script>'] }] }]));
const SUM = makeSheet('Сводка', [
  [{ v: 'Сводка' }, { v: 'Заметка сводки' }],
  [{ v: 'Модель', bg: '#cccccc', bold: true }, { v: 'П 5.1_Ширина_Воздуховоды_Без значения', rot: 90, bg: '#cccccc' }, { v: 'Проверено' }],
  [{ v: 'PROJ_IOS_EOM_K01', bg: '#99ccff' }, { v: 3, bg: '#ff9999' }, { v: '2026-10-20' }],
  [{ v: 'PROJ_IOS_EOM_K02', bg: '#99ccff' }, { v: 0 }, { v: '2026-10-19' }],
]);
const { sandbox, props } = env([SUM, EOM, BIG, XSS]);
const T0 = Date.UTC(2026, 9, 21, 10, 30);          // 21.10.2026 10:30 UTC = 17:30 по Asia/Saigon
props.LAST_PUSH_MS = String(T0);
const serverCalls = [];
const dialogs = [];
const pageErrors = [];
let failOnce = false;
let failVersion = false;
const URL = 'https://gas.test/macros/s/ABC/exec';

async function newPage(browser, context, query, opts) {
  opts = opts || {};
  const page = await context.newPage();
  page.on('dialog', d => { dialogs.push(d.message()); d.dismiss(); });
  page.on('pageerror', e => pageErrors.push(String(e)));
  await page.route('https://gas.test/**', route => {
    const u = new URL_(route.request().url());
    const section = u.searchParams.get('section') || '';
    const html = sandbox.doGet({ parameter: section ? { section } : {} }).html;
    route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.exposeFunction('__serverCall', (fn, arg) => {
    serverCalls.push(fn);
    if (fn === 'viewerGetVersion' && failVersion) { throw new Error('quota exceeded'); }
    if (failOnce) { failOnce = false; throw new Error('Нет раздела «ВК»'); }
    return JSON.parse(JSON.stringify(sandbox[fn](arg)));          // как google.script.run: данные как JSON, ошибки — исключением
  });
  await page.addInitScript(o => {
    window.BVS_POLL_MS = o.poll || 300000;
    window.__hidden = false;
    Object.defineProperty(document, 'hidden', { get: () => window.__hidden, configurable: true });
    window.google = { script: { run: (function make(ok, fail) {
      const o = {
        withSuccessHandler: f => make(f, fail), withFailureHandler: f => make(ok, f),
      };
      return new Proxy(o, { get: (t, k) => t[k] || (arg => { window.__serverCall(k, arg).then(r => setTimeout(() => ok && ok(r), 40), e => fail && fail(new Error(String(e && e.message || e).replace(/^Error: /, '')))); }) });
    })(null, null) } };
  }, { poll: opts.poll });
  await page.goto(URL + (query || ''));
  return page;
}
const URL_ = globalThis.URL;

const visible = page => page.$$eval('#grid tbody tr', trs => trs.filter(t => !t.hidden).map(t => t.className + '|' + t.cells[0].textContent.trim() + '|' + t.cells[1].textContent.trim()));
const count = fn => serverCalls.filter(x => x === fn).length;
const waitGrid = page => page.waitForSelector('#grid:not([hidden]) tbody tr', { timeout: 15000 });

let failures = 0;
async function t(name, fn) {
  try { await fn(); console.log('ok   - ' + name); } catch (e) { failures++; console.log('FAIL - ' + name + '\n' + (e && e.stack || e)); }
}

(async () => {
  let browser;
  try { browser = await pw.chromium.launch(); }
  catch (e) { try { browser = await pw.chromium.launch({ executablePath: '/opt/pw-browsers/chromium' }); } catch (e2) { console.log('SKIP - нет Chromium: ' + e2.message.split('\n')[0]); process.exit(0); } }
  const ctxA = await browser.newContext();

  await t('?section=ЭОМ открывает раздел; по умолчанию уровень 2: модели и спецификации, без элементов', async () => {
    const page = await newPage(browser, ctxA, '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'ЭОМ');
    const tabs = await page.$$eval('.tab', els => els.map(e => e.textContent));
    assert.deepStrictEqual(tabs, ['ЭОМ', 'ВК', 'ОВ', 'Сводка']);
    const v = await visible(page);
    assert.deepStrictEqual(v.map(x => x.split('|')[0]), ['r l0', 'r l1', 'r l1', 'r l0', 'r l1'].map(x => x.replace('r ', '')));
    assert.strictEqual(await page.$eval('#lvls [data-lv="2"]', e => e.className), 'on');
    await page.close();
  });

  await t('уровни 1/2/3, развернуть/свернуть всё', async () => {
    const page = await newPage(browser, ctxA, '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    await page.click('#lvls [data-lv="3"]');
    assert.strictEqual((await visible(page)).length, 9);
    await page.click('#lvls [data-lv="1"]');
    assert.deepStrictEqual((await visible(page)).map(x => x.split('|')[0]), ['l0', 'l0']);
    await page.click('#expand');
    assert.strictEqual((await visible(page)).length, 9);
    await page.click('#collapse');
    assert.strictEqual((await visible(page)).length, 2);
    await page.click('#lvls [data-lv="2"]');
    assert.strictEqual((await visible(page)).length, 5);
    await page.close();
  });

  await t('плюсы/минусы на каждом узле; состояние сохраняется в браузере и переживает перезагрузку', async () => {
    const page = await newPage(browser, ctxA, '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    await page.click('#lvls [data-lv="1"]');
    assert.strictEqual(await page.$eval('tr.l0 .tg', e => e.textContent), '+');
    await page.click('tr.l0 .tg');                                         // раскрыли первую модель: появились её спецификации
    assert.strictEqual(await page.$eval('tr.l0 .tg', e => e.textContent), '−');
    assert.deepStrictEqual((await visible(page)).map(x => x.split('|')[0]), ['l0', 'l1', 'l1', 'l0']);
    await page.click('tr.l1 .tg');                                         // раскрыли первую спецификацию: появились элементы
    assert.deepStrictEqual((await visible(page)).map(x => x.split('|')[0]), ['l0', 'l1', 'l2', 'l2', 'l1', 'l0']);
    await page.reload();
    await waitGrid(page);
    assert.deepStrictEqual((await visible(page)).map(x => x.split('|')[0]), ['l0', 'l1', 'l2', 'l2', 'l1', 'l0']);
    await page.close();
  });

  await t('у другого пользователя (другой браузерный контекст) раскрытие своё', async () => {
    const ctxB = await browser.newContext();
    const page = await newPage(browser, ctxB, '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    assert.strictEqual((await visible(page)).length, 5);                    // по умолчанию, состояние первого не видно
    await page.click('#lvls [data-lv="3"]');
    await page.close();
    const page2 = await newPage(browser, ctxA, '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page2);
    assert.strictEqual((await visible(page2)).length, 6);                   // у первого осталось своё (не 9)
    await page2.close();
    await ctxB.close();
  });

  await t('поиск по имени модели', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    await page.fill('#q', 'k02');
    await page.waitForTimeout(300);
    const v = await visible(page);
    assert.deepStrictEqual(v.map(x => x.split('|')[1].replace(/^[+\u2212]/, '')), ['PROJ_IOS_EOM_K02', ''], JSON.stringify(v));
    assert.ok((await page.textContent('#stat')).includes('1 из 2'));
    await page.fill('#q', '');
    await page.waitForTimeout(300);
    assert.strictEqual((await visible(page)).length, 5);
    await page.close();
  });

  await t('цвета ячеек как на листе, закреплённая шапка, серый шрифт', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    assert.strictEqual(await page.$eval('tr.l0 td', e => getComputedStyle(e).backgroundColor), 'rgb(153, 204, 255)');
    assert.strictEqual(await page.$eval('tr.l1 td:nth-child(2)', e => getComputedStyle(e).backgroundColor), 'rgb(255, 255, 204)');
    assert.strictEqual(await page.$eval('thead th', e => getComputedStyle(e).position), 'sticky');
    assert.strictEqual(await page.$eval('thead th', e => getComputedStyle(e).backgroundColor), 'rgb(204, 204, 204)');
    await page.click('#expand');
    assert.ok(await page.$$eval('td.d', els => els.length) >= 1);
    await page.close();
  });

  await t('Сводка: плоская таблица, цвета, вертикальная шапка, закреплённая первая колонка, поиск, без уровней', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('Сводка'));
    await waitGrid(page);
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'Сводка');
    assert.strictEqual(await page.$$eval('#grid tbody tr', r => r.length), 2);
    assert.strictEqual(await page.$$eval('.tg', r => r.length), 0);
    assert.strictEqual(await page.$eval('table.sum tbody tr td:nth-child(2)', e => getComputedStyle(e).backgroundColor), 'rgb(255, 153, 153)');
    assert.strictEqual(await page.$eval('table.sum th .rot', e => getComputedStyle(e).writingMode), 'vertical-rl');
    assert.strictEqual(await page.$eval('table.sum td:first-child', e => getComputedStyle(e).position), 'sticky');
    assert.strictEqual(await page.$eval('#lvls', e => getComputedStyle(e).display), 'none');
    assert.strictEqual(await page.textContent('#note'), 'Заметка сводки');
    await page.fill('#q', 'k01');
    await page.waitForTimeout(300);
    assert.strictEqual((await page.$$eval('#grid tbody tr', trs => trs.filter(x => !x.hidden).length)), 1);
    await page.close();
  });

  await t('вкладки переключаются без перезагрузки; данные кэшируются (повторный клик не ходит на сервер)', async () => {
    const page = await newPage(browser, await browser.newContext(), '');
    await waitGrid(page);
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'ЭОМ');          // первый раздел по умолчанию
    const before = serverCalls.length;
    await page.click('.tab[data-tab="Сводка"]');
    await page.waitForSelector('table.sum');
    await page.click('.tab[data-tab="ЭОМ"]');
    await page.waitForSelector('#grid:not(.sum) tbody tr');
    await page.click('.tab[data-tab="Сводка"]');
    await page.waitForSelector('table.sum');
    assert.strictEqual(serverCalls.length - before, 1);                      // только первая загрузка «Сводки»
    await page.close();
  });

  await t('неизвестный раздел в ссылке: открывается первый, показано предупреждение', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('Нет такого'));
    await waitGrid(page);
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'ЭОМ');
    assert.ok((await page.textContent('#note')).includes('Нет такого'));
    await page.close();
  });

  await t('текст ячеек не исполняется как HTML', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ОВ'));
    await waitGrid(page);
    await page.click('#expand');
    assert.strictEqual(await page.evaluate(() => window.__xss), undefined);
    assert.ok((await page.textContent('#grid')).includes('<img src=x'));
    assert.strictEqual(await page.$$eval('#grid img, #grid script, #grid b', e => e.length), 0);
    await page.close();
  });

  await t('раздел на 2000+ строк: открывается быстро, уровни переключаются быстро', async () => {
    const page = await newPage(browser, await browser.newContext(), '');
    await waitGrid(page);
    const t0 = Date.now();
    await page.click('.tab[data-tab="ВК"]');
    await page.waitForFunction(() => document.querySelectorAll('#grid tbody tr').length > 2000, null, { timeout: 15000 });
    const open = Date.now() - t0;
    const rows = await page.$$eval('#grid tbody tr', r => r.length);
    assert.ok(rows > 2000, String(rows));
    const t1 = Date.now();
    await page.click('#expand');
    await page.waitForFunction(() => [...document.querySelectorAll('#grid tbody tr')].every(r => !r.hidden));
    const lvl = Date.now() - t1;
    console.log('       строк ' + rows + ': открытие ' + open + ' мс (с задержкой мок-сервера 40 мс), уровень 3 — ' + lvl + ' мс');
    assert.ok(open < 3000, 'открытие ' + open);
    assert.ok(lvl < 1500, 'уровень ' + lvl);
    await page.close();
  });

  await t('ошибка сервера показывается; «Обновить» загружает заново', async () => {
    const page = await newPage(browser, await browser.newContext(), '');
    await waitGrid(page);
    failOnce = true;
    await page.click('.tab[data-tab="ВК"]');
    await page.waitForSelector('#msg.err');
    assert.ok((await page.textContent('#msg')).includes('Нет раздела'));
    assert.strictEqual(await page.$eval('#grid', e => e.hidden), true);
    await page.click('#reload');
    await page.waitForFunction(() => document.querySelectorAll('#grid tbody tr').length > 100);
    assert.strictEqual(await page.$eval('#msg', e => getComputedStyle(e).display), 'none');
    await page.close();
  });


  await t('в шапке «Данные от ДД.ММ.ГГГГ ЧЧ:ММ» по Asia/Saigon; без выгрузки — «Время выгрузки неизвестно»', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page);
    assert.strictEqual(await page.textContent('#ver'), 'Данные от 21.10.2026 17:30');
    await page.click('.tab[data-tab="Сводка"]');
    await page.waitForSelector('table.sum');
    assert.strictEqual(await page.textContent('#ver'), 'Данные от 21.10.2026 17:30');
    await page.close();
    props.LAST_PUSH_MS = '0';
    const page2 = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'));
    await waitGrid(page2);
    assert.strictEqual(await page2.textContent('#ver'), 'Время выгрузки неизвестно');
    await page2.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('новая выгрузка: раз в N минут проверка, плашка, клик обновляет раздел без перезагрузки и сохраняет состояние', async () => {
    const ctx = await browser.newContext();
    const page = await newPage(browser, ctx, '?section=' + encodeURIComponent('ЭОМ'), { poll: 300 });
    await waitGrid(page);
    await page.evaluate(() => { window.__marker = 'same page'; });
    await page.click('#lvls [data-lv="1"]');
    await page.click('tr.l0 .tg');                                         // раскрыли первую модель
    const before = count('viewerGetVersion');
    await page.waitForFunction(() => true);
    await page.waitForTimeout(1000);
    assert.ok(count('viewerGetVersion') >= before + 2, 'проверка версии должна повторяться');
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), true);   // данные те же — плашки нет
    // приходит новая выгрузка с изменёнными данными
    EOM_DATA[4][3].v = 'Лоток НОВЫЙ';
    props.LAST_PUSH_MS = String(T0 + 3600 * 1000);
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 5000 });
    assert.ok((await page.textContent('#fresh')).includes('Данные обновились'));
    assert.strictEqual(await page.textContent('#ver'), 'Данные от 21.10.2026 17:30');     // пока показаны старые
    const callsBefore = count('viewerGetSection');
    await page.click('#fresh');
    await page.waitForFunction(() => document.querySelector('#ver').textContent === 'Данные от 21.10.2026 18:30');
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), true);
    assert.strictEqual(await page.evaluate(() => window.__marker), 'same page');           // страница не перезагружалась
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'ЭОМ');
    assert.strictEqual(count('viewerGetSection') - callsBefore, 1);
    assert.deepStrictEqual((await visible(page)).map(x => x.split('|')[0]), ['l0', 'l1', 'l1', 'l0']);   // раскрытие сохранилось
    await page.click('#expand');
    assert.ok((await page.textContent('#grid')).includes('Лоток НОВЫЙ'));                  // данные новые
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), true);                    // плашка не вернулась (версия та же)
    await page.close();
    EOM_DATA[4][3].v = 'Лоток 1';
    props.LAST_PUSH_MS = String(T0);
  });

  await t('«Сводка»: обновление по плашке перечитывает Сводку и остаётся на ней', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('Сводка'), { poll: 300 });
    await page.waitForSelector('table.sum');
    props.LAST_PUSH_MS = String(T0 + 60000);
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 5000 });
    const before = count('viewerGetSummary');
    await page.click('#fresh');
    await page.waitForFunction(() => document.querySelector('#ver').textContent === 'Данные от 21.10.2026 17:31');
    assert.strictEqual(count('viewerGetSummary') - before, 1);
    assert.strictEqual(await page.$eval('.tab.on', e => e.textContent), 'Сводка');
    await page.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('вкладка скрыта: проверок нет; при возврате на вкладку — сразу', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'), { poll: 200 });
    await waitGrid(page);
    await page.evaluate(() => { window.__hidden = true; });
    await page.waitForTimeout(400);                          // дать уже запущенной проверке завершиться
    const frozen = count('viewerGetVersion');
    props.LAST_PUSH_MS = String(T0 + 7200 * 1000);
    await page.waitForTimeout(1200);
    assert.strictEqual(count('viewerGetVersion'), frozen, 'пока вкладка скрыта, версию не спрашиваем');
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), true);
    await page.evaluate(() => { window.__hidden = false; document.dispatchEvent(new Event('visibilitychange')); });
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 150 + 1000 });
    assert.ok(count('viewerGetVersion') > frozen);
    await page.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('возврат на вкладку проверяет сразу, не дожидаясь таймера (интервал 1 час)', async () => {
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'), { poll: 3600000 });
    await waitGrid(page);
    const before = count('viewerGetVersion');
    props.LAST_PUSH_MS = String(T0 + 1000 * 60);
    await page.evaluate(() => { document.dispatchEvent(new Event('visibilitychange')); });
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 2000 });
    assert.strictEqual(count('viewerGetVersion') - before, 1);
    await page.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('ошибки проверки версии молча игнорируются: без окон, без плашки и сообщений; потом работает', async () => {
    dialogs.length = 0; pageErrors.length = 0;
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'), { poll: 200 });
    await waitGrid(page);
    failVersion = true;
    const before = count('viewerGetVersion');
    props.LAST_PUSH_MS = String(T0 + 99999);
    await page.waitForTimeout(1200);
    assert.ok(count('viewerGetVersion') > before, 'проверки продолжались');
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), true);
    assert.strictEqual(await page.$eval('#msg', e => getComputedStyle(e).display), 'none');
    assert.strictEqual(await page.textContent('#note'), '');
    assert.deepStrictEqual(dialogs, []);
    assert.deepStrictEqual(pageErrors, []);
    assert.strictEqual((await visible(page)).length, 5);       // страница работает
    failVersion = false;
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 3000 });
    await page.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('ошибка при обновлении по плашке: старые данные остаются, плашка возвращается, окон нет', async () => {
    dialogs.length = 0;
    const page = await newPage(browser, await browser.newContext(), '?section=' + encodeURIComponent('ЭОМ'), { poll: 200 });
    await waitGrid(page);
    props.LAST_PUSH_MS = String(T0 + 5000);
    await page.waitForSelector('#fresh:not([hidden])', { timeout: 3000 });
    failOnce = true;
    await page.click('#fresh');
    await page.waitForFunction(() => document.querySelector('#note').textContent.includes('Не удалось обновить'));
    assert.strictEqual(await page.$eval('#fresh', e => e.hidden), false);
    assert.strictEqual((await visible(page)).length, 5);
    assert.deepStrictEqual(dialogs, []);
    await page.click('#fresh');
    await page.waitForFunction(() => document.querySelector('#ver').textContent === 'Данные от 21.10.2026 17:30');
    await page.close();
    props.LAST_PUSH_MS = String(T0);
  });

  await t('на сервер уходят только viewerGetSection / viewerGetSummary', async () => {
    const names = Array.from(new Set(serverCalls)).sort();
    assert.deepStrictEqual(names, ['viewerGetSection', 'viewerGetSummary', 'viewerGetVersion']);
  });

  await browser.close();
  if (failures) { console.log(failures + ' failed'); process.exit(1); }
  console.log('all ok');
})();
