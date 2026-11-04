/**
 * WB → «Таблица конверсий». Вставить целиком в привязанный Apps Script.
 * Ключ: свойство скрипта WB_API_TOKEN. Установка: wbInstall.
 * Обновляет только данные существующих строк по заголовкам первой строки.
 * Не добавляет строки/столбцы и не перезаписывает формулы, заголовки, фильтры.
 * AO:AP — ручные показы и их период; API воронки их не предоставляет.
 * Это снимок выбранного периода, а не накопительный журнал по дням.
 */
const WB = Object.freeze({
  get bookId() {
    const saved = PropertiesService.getScriptProperties().getProperty('WB_SPREADSHEET_ID');
    if (saved) return saved;
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active) throw new Error('Запустите wbInstall из привязанного к таблице проекта Apps Script.');
    return active.getId();
  },
  sheet: 'Рабочая табличка',
  timezone: 'Europe/Moscow',
  url: 'https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products',
  pageSize: 1000,
  intervalMs: 21000,
  budgetMs: 260000,
  headers: [
    'WB: период', 'WB: обновлено (МСК)', 'WB: переходы', 'WB: корзины',
    'WB: заказы', 'WB: выкупы', 'WB: выручка выкупов, ₽',
    'WB: средняя цена, ₽', 'WB: в корзину, %', 'WB: в заказ, %',
    'WB: выкуп, %', 'WB: заказ / переход, %',
    'WB: расчётная прибыль по марже, ₽', 'WB: остаток WB',
    'WB: остаток продавца', 'WB: статус',
    'WB: показы (ручной ввод)', 'WB: период показов (ручной ввод)',
    'WB: CTR, %', 'WB: заказы / 10 тыс. показов',
    'WB: прогноз выкупов / 10 тыс.', 'WB: прогноз прибыли / 10 тыс., ₽',
    'WB: ранг прогноза', 'WB: оценка', 'WB: название товара'
  ]
});
const WB_LAYOUT = new WeakMap();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('WB Аналитика')
    .addItem('Обновить все данные WB', 'wbUpdateAll')
    .addItem('Период: вчера', 'wbYesterday')
    .addItem('Период: последние 7 дней', 'wbWeek')
    .addItem('Период: последние 30 дней', 'wbMonth')
    .addSeparator()
    .addItem('Только товары с остатком', 'wbShowInStock')
    .addItem('Показать товары без ограничения остатков', 'wbShowAllStock')
    .addSeparator()
    .addItem('Включить ежедневное обновление', 'wbInstall')
    .addItem('Отключить ежедневное обновление', 'wbDisable')
    .addToUi();
}

function wbInstall() {
  PropertiesService.getScriptProperties().setProperty('WB_SPREADSHEET_ID', WB.bookId);
  const book = SpreadsheetApp.openById(WB.bookId);
  wbSetupForecastSheet_(book);
  const props = PropertiesService.getScriptProperties();
  const pendingRetry = props.getProperty('WB_RETRY_TRIGGER_ID');
  const old = wbOwnTriggers_().filter(trigger => trigger.getUniqueId() !== pendingRetry);
  ScriptApp.newTrigger('wbUpdateAll').timeBased().everyDays(1)
    .atHour(4).inTimezone(WB.timezone).create();
  ScriptApp.newTrigger('wbCheckForCodeChanges').timeBased().everyMinutes(1).create();
  old.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  onOpen();
  wbToast_('Пустая вкладка прогноза подготовлена. Данные загрузятся при следующем обновлении WB.');
}

function wbSetupForecastSheet_(book) {
  const out = book.getSheetByName('Прогноз по конверсиям');
  const src = book.getSheetByName(WB.sheet);
  if (!out || !src) throw new Error('Проверьте наличие вкладок «Прогноз по конверсиям» и «' + WB.sheet + '».');
  const h = ['Артикул WB','Товар','Тип товара','Показы (ввод)','CTR, %','Переходы',
    'Корзины','Переход → корзина, %','Заказы','Корзина → заказ, %','Выкупы',
    'Заказ → выкуп, %','Себестоимость за единицу, ₽ (ввод)',
    'Реклама за период, ₽ (ввод)','Выручка выкупов, ₽',
    'Прибыль после себестоимости и рекламы, ₽','Прибыль на выкуп, ₽','Маржинальность, %',
    'Заказы / 10 тыс. показов','Выкупы / 10 тыс. показов','Период данных WB','Сравнить'];
  const oldRows = Math.max(out.getLastRow(), 1), oldCols = Math.max(out.getLastColumn(), h.length);
  const old = out.getRange(1, 1, oldRows, oldCols).getValues();
  const oldHeaders = new Map((old[0] || []).map((v, i) => [String(v).trim(), i]));
  const prior = new Map();
  for (let r = 1; r < old.length; r++) {
    const idCol = oldHeaders.has('Артикул WB') ? oldHeaders.get('Артикул WB') : 0;
    const id = String(old[r][idCol] || '').trim();
    if (/^\d+$/.test(id)) prior.set(id, old[r]);
  }
  const end = src.getLastRow();
  if (end < 2) throw new Error('Сначала заполните артикулы на вкладке «' + WB.sheet + '».');
  const sourceRows = src.getRange(2, 1, end - 1, Math.max(3, src.getLastColumn())).getValues();
  const products = new Map();
  sourceRows.forEach(r => {
    const id = String(r[2] || '').trim();
    if (/^\d+$/.test(id) && !products.has(id)) products.set(id, [id, r[1] || '', r[0] || '']);
  });
  if (!products.size) throw new Error('На вкладке «' + WB.sheet + '» не найдены артикулы WB.');
  if (out.getMaxColumns() < 26) out.insertColumnsAfter(out.getMaxColumns(), 26 - out.getMaxColumns());
  const output = [h];
  products.forEach((product, id) => {
    const oldRow = prior.get(id);
    const value = (heading, fallback) => {
      const i = oldHeaders.get(heading);
      return oldRow && i !== undefined && oldRow[i] !== '' ? oldRow[i] : fallback;
    };
    output.push([id, product[1], product[2], value('Показы (ввод)', ''), value('CTR, %', ''),
      value('Переходы', ''), value('Корзины', ''), value('Переход → корзина, %', ''),
      value('Заказы', ''), value('Корзина → заказ, %', ''), value('Выкупы', ''),
      value('Заказ → выкуп, %', ''), value('Себестоимость за единицу, ₽ (ввод)', ''),
      value('Реклама за период, ₽ (ввод)', 0), value('Выручка выкупов, ₽', ''),
      value('Прибыль после себестоимости и рекламы, ₽', ''), value('Прибыль на выкуп, ₽', ''),
      value('Маржинальность, %', ''), value('Заказы / 10 тыс. показов', ''),
      value('Выкупы / 10 тыс. показов', ''), value('Период данных WB', ''), value('Сравнить', false)]);
  });
  const priorPeriod = out.getRange('W2').getValue();
  out.getRange(1, 1, out.getMaxRows(), 22).clearContent();
  out.getRange(1, 1, output.length, 22).setValues(output);
  if (out.getFilter()) out.getFilter().remove();
  out.getRange(1, 1, output.length, 22).createFilter();
  out.getRange(2, 22, output.length - 1, 1).insertCheckboxes();
  out.getRange(2, 22, output.length - 1, 1).setValues(output.slice(1).map(r => [r[21] === true]));
  out.setFrozenRows(1);
  out.getBandings().forEach(b => b.remove());
  const banding = out.getRange(1, 1, output.length, 22)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banding.setHeaderRowColor('#243b53');
  banding.setFirstRowColor('#e7eef5');
  banding.setSecondRowColor('#dce8f2');
  out.setHiddenGridlines(true);
  out.getRange(1, 1, output.length, 22).setFontFamily('Arial').setVerticalAlignment('middle');
  out.getRange(1, 1, 1, 22).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  out.setRowHeight(1, 58);
  out.setRowHeights(2, Math.max(output.length - 1, 1), 28);
  out.setColumnWidth(1, 138); out.setColumnWidth(2, 250); out.setColumnWidth(3, 155);
  out.setColumnWidths(4, 9, 142); out.setColumnWidth(13, 218); out.setColumnWidth(14, 190);
  out.setColumnWidths(15, 4, 190); out.setColumnWidths(19, 3, 170); out.setColumnWidth(22, 95);
  const periodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Вчера','Последние 7 дней','Последние 30 дней','Последние 90 дней','Последние 365 дней'], true)
    .setAllowInvalid(false).build();
  out.getRange('W1').setValue('Период обновления');
  out.getRange('W2').setValue(priorPeriod || 'Вчера').setDataValidation(periodRule);
  out.getRange('W1').setBackground('#243b53').setFontColor('#ffffff').setFontWeight('bold');
  out.getRange('W2').setBackground('#dce8f2').setFontColor('#243b53');
  out.setColumnWidth(23, 190);
  const expected = ['WB ID из файла себестоимости','Себестоимость из файла, ₽','Последняя дата в файле'];
  const costHeaders = out.getRange(1, 24, 1, 3).getValues()[0];
  if (costHeaders.every(v => v === '')) out.getRange(1, 24, 1, 3).setValues([expected]);
  else if (!expected.every((v, i) => costHeaders[i] === v)) throw new Error('X:Z заняты данными; справочник себестоимости не изменён.');
  const costBand = out.getRange(1, 24, out.getMaxRows(), 3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  costBand.setHeaderRowColor('#243b53'); costBand.setFirstRowColor('#e7eef5'); costBand.setSecondRowColor('#dce8f2');
  out.getRange(1, 24, 1, 3).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  out.setColumnWidths(24, 2, 205); out.setColumnWidth(26, 150);
  const n = output.length - 1;
  if (n > 0) {
    [[5],[8],[10],[12],[18]].forEach(x => out.getRange(2, x[0], n, 1).setNumberFormat('0.00%'));
    [13,14,15,16,17].forEach(c => out.getRange(2,c,n,1).setNumberFormat('#,##0.00'));
    [4,6,7,9,11,19,20].forEach(c => out.getRange(2,c,n,1).setNumberFormat('#,##0.##'));
    const rules = [];
    [15,17,18,19,20].forEach(c => {
      const L = String.fromCharCode(64+c), rng = out.getRange(2,c,n,1), endRow=n+1;
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${L}2<>"",${L}2<=PERCENTILE($${L}$2:$${L}$${endRow},0.33))`).setBackground('#f4cccc').setRanges([rng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${L}2<>"",${L}2>PERCENTILE($${L}$2:$${L}$${endRow},0.33),${L}2<PERCENTILE($${L}$2:$${L}$${endRow},0.67))`).setBackground('#fff2cc').setRanges([rng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${L}2<>"",${L}2>=PERCENTILE($${L}$2:$${L}$${endRow},0.67))`).setBackground('#d9ead3').setRanges([rng]).build());
    });
    out.setConditionalFormatRules(rules);
  }
  out.getRange(2,4,Math.max(n,1),1).setNote('Показы выбранного периода. Заполнятся автоматически, если WB вернёт метрику показов.');
  out.getRange(2,13,Math.max(n,1),1).setNote('Себестоимость одной выкупленной единицы. Ручное значение сохраняется.');
  out.getRange(2,14,Math.max(n,1),1).setNote('Рекламные расходы за период; по умолчанию 0 ₽.');
  return out;
}

function wbDisable() {
  wbOwnTriggers_().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  wbToast_('Ежедневное обновление этого пользователя отключено.');
}

function wbOwnTriggers_() {
  // Только обработчики этой интеграции и ранее выданного скрипта артикулов.
  const names = ['wbUpdateAll', 'wbCheckForCodeChanges', 'refreshArticles', 'setupArticleFilter'];
  return ScriptApp.getProjectTriggers().filter(trigger =>
    names.indexOf(trigger.getHandlerFunction()) !== -1);
}

function wbYesterday() { wbSetDays_(1); }
function wbWeek() { wbSetDays_(7); }
function wbMonth() { wbSetDays_(30); }

function wbSetDays_(days) {
  PropertiesService.getScriptProperties().setProperty('WB_REPORT_DAYS', String(days));
  wbSetForecastPeriod_(days);
  wbRun_(days);
}

function wbUpdateAll() {
  const sheet = SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  const selected = sheet ? String(sheet.getRange('W2').getDisplayValue()).trim() : '';
  const fromSheet = {
    'Вчера': 1, 'Последние 7 дней': 7, 'Последние 30 дней': 30,
    'Последние 90 дней': 90, 'Последние 365 дней': 365
  }[selected];
  const days = fromSheet || Number(PropertiesService.getScriptProperties()
    .getProperty('WB_REPORT_DAYS') || 1);
  PropertiesService.getScriptProperties().setProperty('WB_REPORT_DAYS', String(days));
  return wbRun_(days);
}

function wbCheckForCodeChanges() {
  const props = PropertiesService.getScriptProperties();
  wbProbePortalAccess_();
  wbJamCtrTick_(props);
  const loaded = props.getProperty('WB_LOADED_RELEASE_SHA');
  if (!loaded || loaded === props.getProperty('WB_ACTIVE_RELEASE_SHA')) return;
  // Отмечаем выпуск перед обращением к WB, чтобы не запускать один и тот же
  // неудачный запрос каждую минуту. При HTTP 429 сработает штатный отложенный повтор.
  props.setProperty('WB_ACTIVE_RELEASE_SHA', loaded);
  wbToast_('Новый код WB найден. Запускаю обновление данных.');
  // Existing values and formulas are not rebuilt when a release changes.
  return wbUpdateAll();
}

function wbSetForecastPeriod_(days) {
  const sheet = SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  if (!sheet) return;
  const labels = {1: 'Вчера', 7: 'Последние 7 дней', 30: 'Последние 30 дней',
    90: 'Последние 90 дней', 365: 'Последние 365 дней'};
  if (labels[days]) sheet.getRange('W2').setValue(labels[days]);
}

function wbRun_(days) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Обновление уже выполняется. Повторный запуск не требуется.');
  }
  const props = PropertiesService.getScriptProperties();
  try {
    if ([1, 7, 30, 90, 365].indexOf(days) === -1) {
      throw new Error('Выберите период 1, 7, 30, 90 или 365 дней.');
    }
    const token = (props.getProperty('WB_API_TOKEN') || '').trim();
    if (!token) throw new Error('Добавьте WB_API_TOKEN в свойства скрипта.');
    wbProbeAdvertisingAccess_(token);
    const sheet = wbSheet_();
    wbCheckLayout_(sheet);
    const period = wbPeriod_(days);
    wbJamCtrQueue_(period);
    const state = { deadline: Date.now() + WB.budgetMs };
    const products = wbFetchProducts_(token, period, state);
    if (!products.length) {
      throw new Error('WB вернул пустой список. Существующие данные не изменены.');
    }
    // Все страницы загружаются и проверяются до записи в таблицу.
    const records = products.map(item => wbRecord_(item, period));
    // Те же ответы, что используются для загрузки; дополнительных запросов нет.
    // Только имена полей, без ключа, значений или персональных данных.
    const selectedFields = products[0].statistic.selected;
    console.log('WB поля воронки: ' + Object.keys(selectedFields).sort().join(', '));
    console.log('WB поля CTR/показов: ' + wbMetricPaths_(products[0]).join(', '));
    const ids = new Set();
    records.forEach(record => {
      if (ids.has(record.id)) {
        throw new Error('WB вернул повторяющиеся карточки. Повторите обновление позже.');
      }
      ids.add(record.id);
    });
    const written = wbWrite_(sheet, records, period);
    SpreadsheetApp.flush();
    const now = Utilities.formatDate(new Date(), WB.timezone, 'yyyy-MM-dd HH:mm:ss');
    props.setProperties({ WB_LAST_SUCCESS: now, WB_LAST_ERROR: '' });
    wbClearRetry_(props);
    wbJamCtrQueue_(period);
    wbToast_('Обновлено строк: ' + written + '. Период: ' + wbPeriodLabel_(period));
  } catch (error) {
    // Не записываем тела ответов WB, ключ или заголовки запросов в журнал.
    const message = error.message || 'Ошибка обновления WB.';
    props.setProperty('WB_LAST_ERROR', message.slice(0, 1000));
    if (error.wbRetryAt) {
      wbScheduleRetry_(props, error.wbRetryAt);
      wbToast_(message + ' Автоматический повтор запланирован. Данные пока не обновлены.');
      return { deferred: true };
    }
    console.error(message);
    wbToast_(message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function wbSheet_() {
  const sheet = SpreadsheetApp.openById(WB.bookId).getSheetByName(WB.sheet);
  if (!sheet) throw new Error('Не найдена вкладка «' + WB.sheet + '».');
  return sheet;
}

function wbColumnShift_(sheet) {
  if (!WB_LAYOUT.has(sheet)) {
    WB_LAYOUT.set(sheet, sheet.getRange('A1').getDisplayValue().trim() === 'Тип товара' ? 1 : 0);
  }
  return WB_LAYOUT.get(sheet);
}

// Координаты остального кода остаются логическими — до вставки столбца.
function wbRange_(sheet, row, column, height, width) {
  if (typeof row === 'string') {
    const match = row.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error('Некорректная ссылка на ячейку.');
    column = wbColumnNumber_(match[1]);
    row = Number(match[2]);
  }
  return sheet.getRange(row, column + wbColumnShift_(sheet), height || 1, width || 1);
}

function wbColumnNumber_(letters) {
  return Array.from(letters).reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0);
}

function wbColumnLetters_(number) {
  let result = '';
  while (number > 0) {
    number--;
    result = String.fromCharCode(65 + number % 26) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

function wbFormula_(sheet, formula) {
  if (!wbColumnShift_(sheet)) return formula;
  // Только ссылки, используемые в наших формулах; текстовые строки — русские.
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?\d+|(?=[:,)]))/g,
    (_, dollar, letters, row) => dollar +
      wbColumnLetters_(wbColumnNumber_(letters) + 1) + row);
}

function wbEnsureTypeColumn_(sheet) {
  if (wbColumnShift_(sheet)) return;
  if (sheet.getRange('A1').getDisplayValue().trim() !== 'Товар') {
    throw new Error('Ожидается заголовок «Товар» в A1. Столбец не добавлен.');
  }
  sheet.insertColumnBefore(1); // Sheets сам сдвигает ссылки в существующих формулах.
  sheet.getRange('A1').setValue('Тип товара');
  WB_LAYOUT.set(sheet, 1);
  sheet.getRange('A1').setNote(
    'Тип берётся из subjectName WB. Обувь и канцелярия объединяются по словарю; для остальных используется категория WB. Фильтр — в заголовке A1.'
  );
}

function wbProductType_(subjectName) {
  const name = String(subjectName || '').trim();
  const shoeTypes = ['Обувь', 'Ботинки', 'Ботильоны', 'Балетки', 'Босоножки',
    'Ботфорты', 'Валенки', 'Кеды', 'Кроссовки', 'Лоферы', 'Мокасины',
    'Полуботинки', 'Сабо', 'Сандалии', 'Сапоги', 'Сланцы', 'Тапочки',
    'Туфли', 'Угги', 'Чешки', 'Эспадрильи', 'Резиновые сапоги'];
  const stationeryTypes = ['Канцелярия', 'Тетради', 'Блокноты', 'Ежедневники',
    'Записные книжки', 'Ручки шариковые', 'Ручки гелевые',
    'Карандаши чернографитные', 'Карандаши цветные', 'Фломастеры',
    'Маркеры перманентные', 'Маркеры текстовыделители', 'Ластики',
    'Линейки', 'Пеналы', 'Скрепки', 'Степлеры', 'Скобы для степлеров', 'Точилки'];
  const matches = list => list.some(value => value.toLowerCase() === name.toLowerCase());
  if (matches(shoeTypes)) return 'Обувь';
  if (matches(stationeryTypes)) return 'Канцелярия';
  return wbSafeText_(name || 'Тип не указан');
}

function wbCheckLayout_(sheet) {
  if (wbRange_(sheet, 'B1').getDisplayValue().trim() !== 'Артикул WB') {
    throw new Error('Не найден заголовок «Артикул WB» в ожидаемом столбце. Заголовок не изменён.');
  }
  if (!/^Прибыль с выкупа/.test(wbRange_(sheet, 'I1').getDisplayValue().trim())) {
    throw new Error('Не найден заголовок «Прибыль с выкупа» в ожидаемом столбце.');
  }
  const width = Math.max(0, Math.min(WB.headers.length, sheet.getMaxColumns() - 23 - wbColumnShift_(sheet)));
  if (!width) return;
  const range = wbRange_(sheet, 1, 24, Math.max(sheet.getLastRow(), 1), width);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const ours = width === WB.headers.length && WB.headers.every((header, i) =>
    values[0][i] === header ||
    (i === 12 && values[0][i] === 'WB: расчётная прибыль по I, ₽'));
  if (!ours && values.some((row, r) => row.some((value, c) =>
    value !== '' || formulas[r][c] !== ''))) {
    throw new Error('Столбцы блока WB уже заняты другими данными. Код ничего не перезаписал.');
  }
}

function wbPeriod_(days) {
  const today = Utilities.formatDate(new Date(), WB.timezone, 'yyyy-MM-dd');
  const end = new Date(today + 'T12:00:00Z');
  end.setUTCDate(end.getUTCDate() - 1);
  const start = new Date(end.getTime());
  start.setUTCDate(start.getUTCDate() - days + 1);
  return { start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10) };
}

function wbPeriodLabel_(period) { return period.start + ' — ' + period.end; }

function wbFetchProducts_(token, period, state) {
  let offset = 0;
  const all = [];
  while (true) {
    const data = wbRequest_(token, {
      selectedPeriod: period, nmIds: [], brandNames: [], subjectIds: [], tagIds: [],
      skipDeletedNm: true, limit: WB.pageSize, offset: offset
    }, state);
    if (!data || !data.data || !Array.isArray(data.data.products)) {
      throw new Error('Неожиданный формат ответа API воронки WB. Данные не записаны.');
    }
    const page = data.data.products;
    all.push.apply(all, page);
    if (page.length < WB.pageSize) return all;
    offset += page.length;
  }
}

function wbRequest_(token, payload, state) {
  const props = PropertiesService.getScriptProperties();
  for (let attempt = 0; attempt < 3; attempt++) {
    const next = wbNextRequestAt_(props);
    const wait = Math.max(0, next - Date.now());
    if (wait > 30000 || (wait > 0 && Date.now() + wait + 15000 > state.deadline)) {
      throw wbDeferredError_(next, props.getProperty('WB_LAST_HTTP_STATUS'));
    }
    if (Date.now() + 15000 > state.deadline) {
      throw new Error('Не хватило времени на загрузку всех страниц WB. Запись в таблицу не выполнялась.');
    }
    if (wait) Utilities.sleep(wait);
    let response;
    // Сохраняем интервал и между ручным и ежедневным запуском.
    props.setProperty('WB_NEXT_REQUEST_AT_V2', String(Date.now() + WB.intervalMs));
    try {
      response = UrlFetchApp.fetch(WB.url, {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: token }, payload: JSON.stringify(payload),
        muteHttpExceptions: true, followRedirects: false, validateHttpsCertificates: true
      });
    } catch (_) {
      if (attempt < 2) continue;
      throw new Error('Не удалось связаться с API WB. Данные не записаны.');
    }
    const status = response.getResponseCode();
    props.setProperty('WB_LAST_HTTP_STATUS', String(status));
    if (status === 200) {
      try { return JSON.parse(response.getContentText()); }
      catch (_) { throw new Error('WB вернул некорректный JSON. Данные не записаны.'); }
    }
    if (status === 401) throw new Error('WB HTTP 401: проверьте токен и срок его действия.');
    if (status === 403) throw new Error('WB HTTP 403: проверьте права категории «Аналитика».');
    if (status === 429 || status >= 500) {
      const delay = Math.max(WB.intervalMs * (attempt + 1),
        wbRateLimitDelay_(response.getAllHeaders(), Date.now()));
      const until = Date.now() + delay;
      props.setProperty('WB_NEXT_REQUEST_AT_V2', String(until));
      console.log('WB HTTP ' + status + ': ожидание ' + Math.ceil(delay / 1000) + ' сек.');
      if (delay > 30000 || attempt === 2) throw wbDeferredError_(until, status);
      if (attempt < 2) continue;
    }
    throw new Error('API воронки WB: HTTP ' + status + '. Данные не записаны.');
  }
}

function wbNextRequestAt_(props) {
  const saved = props.getProperty('WB_NEXT_REQUEST_AT_V2');
  if (saved !== null) {
    const value = Number(saved);
    return Number.isFinite(value) && value >= 0 ? value : 0;
  }
  // В старой версии абсолютный timestamp заголовка мог быть ошибочно
  // прибавлен к текущему времени как длительность. Такое значение не переносим.
  const legacy = Number(props.getProperty('WB_NEXT_REQUEST_AT') || 0);
  const now = Date.now();
  const valid = Number.isFinite(legacy) && legacy > now && legacy <= now + 86400000;
  const next = valid ? legacy : 0;
  if (legacy > now && !valid) {
    console.log('Исправлено некорректное старое состояние ожидания; лимит будет уточнён по ответу WB.');
  }
  props.setProperty('WB_NEXT_REQUEST_AT_V2', String(next));
  props.deleteProperty('WB_NEXT_REQUEST_AT');
  return next;
}

function wbRateLimitDelay_(headers, now) {
  let delay = 0;
  Object.keys(headers).forEach(name => {
    const key = name.toLowerCase();
    if (['retry-after', 'x-ratelimit-retry', 'x-ratelimit-reset'].indexOf(key) === -1) return;
    const raw = String(headers[name]).trim();
    const value = Number(raw);
    let milliseconds = 0;
    if (raw !== '' && Number.isFinite(value) && value >= 0) {
      if (key === 'retry-after') milliseconds = value * 1000; // HTTP: длительность в секундах.
      else if (value >= 1e12) milliseconds = value - now; // Unix milliseconds.
      else if (value >= 1e9) milliseconds = value * 1000 - now; // Unix seconds.
      else milliseconds = value * 1000;
    } else if (key === 'retry-after') {
      const date = Date.parse(raw); // HTTP допускает дату вместо числа секунд.
      if (Number.isFinite(date)) milliseconds = date - now;
    }
    if (Number.isFinite(milliseconds)) delay = Math.max(delay, milliseconds, 0);
  });
  return delay;
}

function wbDeferredError_(until, status) {
  const date = Utilities.formatDate(new Date(until), WB.timezone, 'yyyy-MM-dd HH:mm:ss');
  const http = Number(status);
  const limited = http === 429 || http >= 500;
  const error = new Error('Ожидание API WB' + (limited ? ' (HTTP ' + status + ')' : '') +
    ' до ' + date + ' МСК.');
  if (until <= Date.now() + 31 * 86400000) error.wbRetryAt = until;
  return error;
}

function wbClearRetry_(props) {
  const id = props.getProperty('WB_RETRY_TRIGGER_ID');
  if (id) {
    ScriptApp.getProjectTriggers().filter(trigger => trigger.getUniqueId() === id)
      .forEach(trigger => ScriptApp.deleteTrigger(trigger));
    props.deleteProperty('WB_RETRY_TRIGGER_ID');
  }
}

function wbScheduleRetry_(props, until) {
  // Использует существующий обработчик загрузчика: его менять не требуется.
  const previous = props.getProperty('WB_RETRY_TRIGGER_ID');
  const trigger = ScriptApp.newTrigger('wbUpdateAll').timeBased()
    .after(Math.max(60000, until - Date.now())).create();
  if (previous) {
    ScriptApp.getProjectTriggers().filter(item => item.getUniqueId() === previous)
      .forEach(item => ScriptApp.deleteTrigger(item));
  }
  props.setProperty('WB_RETRY_TRIGGER_ID', trigger.getUniqueId());
}

function wbNumber_(value, field) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
    throw new Error('Некорректное или отсутствующее поле WB: ' + field + '. Данные не записаны.');
  }
  return value;
}

function wbMetricPaths_(root) {
  const found = [], stack = [{value: root, path: ''}], seen = new Set();
  while (stack.length) {
    const node=stack.pop();
    if (!node.value || typeof node.value !== 'object' || seen.has(node.value)) continue;
    seen.add(node.value);
    Object.keys(node.value).forEach(key => {
      const path=node.path ? node.path+'.'+key : key;
      if (/(ctr|impression|show|view|click)/i.test(key)) found.push(path);
      if (node.value[key] && typeof node.value[key] === 'object') stack.push({value:node.value[key],path:path});
    });
  }
  return found.sort().slice(0,100);
}

function wbMetric_(root, aliases) {
  const wanted = new Set(aliases.map(x => x.toLowerCase()));
  const stack = [root];
  while (stack.length) {
    const value = stack.pop();
    if (!value || typeof value !== 'object') continue;
    Object.keys(value).forEach(key => {
      if (wanted.has(key.toLowerCase()) && typeof value[key] === 'number' && Number.isFinite(value[key])) {
        stack.push({__metricFound: value[key]});
      } else if (value && typeof value[key] === 'object') stack.push(value[key]);
    });
    const found = stack.find(x => x && typeof x === 'object' && Object.prototype.hasOwnProperty.call(x, '__metricFound'));
    if (found) return found.__metricFound;
  }
  return undefined;
}

function wbRecord_(item, period) {
  const product = item && item.product;
  const stats = item && item.statistic && item.statistic.selected;
  if (!product || !stats || !stats.period ||
      stats.period.start !== period.start || stats.period.end !== period.end ||
      !Number.isSafeInteger(product.nmId) || product.nmId <= 0 || typeof product.title !== 'string') {
    throw new Error('Некорректная карточка или период в ответе WB. Данные не записаны.');
  }
  const fields = ['openCount','cartCount','orderCount','buyoutCount','buyoutSum','avgPrice'];
  const numbers = fields.map(field => wbNumber_(stats[field], field));
  const conversions = stats.conversions || {};
  const ratios = ['addToCartPercent','cartToOrderPercent','buyoutPercent']
    .map(field => wbNumber_(conversions[field], field) / 100);
  const rawCtr = wbMetric_(item, ['ctr','ctrPercent','clickThroughRate','clickRate']);
  const rawShows = wbMetric_(item, ['impressions','impressionCount','showCount','showsCount','showCardCount','viewCount','views']);
  const ctr = rawCtr === undefined ? undefined : (rawCtr > 1 ? rawCtr / 100 : rawCtr);
  const stocks = product.stocks || {};
  return {
    id: String(product.nmId), title: wbSafeText_(product.title), type: wbProductType_(product.subjectName),
    numbers: numbers, ratios: ratios, ctr: ctr, impressions: rawShows,
    stockWB: wbNumber_(stocks.wb, 'stocks.wb'), stockMP: wbNumber_(stocks.mp, 'stocks.mp')
  };
}

function wbSafeText_(text) { return /^[=+@-]/.test(text) ? "'" + text : text; }

function wbWrite_(sheet, records, period) {
  // Режим пользователя: только данные под существующими заголовками первой
  // строки. Не вставляем строки/столбцы, не меняем формулы, фильтры и оформление.
  const rows = sheet.getLastRow();
  const columns = sheet.getLastColumn();
  if (rows < 2) throw new Error('В таблице нет строк с артикулами. Структура не изменена.');
  const range = sheet.getRange(1, 1, rows, columns);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const headers = new Map(values[0].map((value, index) => [String(value).trim(), index]));
  const idColumn = headers.get('Артикул WB');
  if (idColumn === undefined) throw new Error('Не найден заголовок «Артикул WB».');
  const byId = new Map(records.map(record => [record.id, record]));
  const updates = new Map();
  const now = Utilities.formatDate(new Date(), WB.timezone, 'yyyy-MM-dd HH:mm:ss');
  const periodText = wbPeriodLabel_(period);
  let matched = 0;
  function put(row, header, value) {
    const column = headers.get(header);
    if (column === undefined || formulas[row][column] !== '' || value === undefined) return;
    if (values[row][column] === value) return;
    if (!updates.has(column)) updates.set(column, []);
    updates.get(column).push([row + 1, value]);
  }
  for (let r = 1; r < rows; r++) {
    const id = String(values[r][idColumn]).trim();
    if (!/^\d+$/.test(id)) continue;
    const record = byId.get(id);
    if (!record) {
      const periodColumn = headers.get('WB: период');
      if (periodColumn !== undefined && values[r][periodColumn] !== '') {
        put(r, 'WB: статус', 'Не получен в текущем ответе — данные прежнего периода');
      }
      continue;
    }
    matched++;
    const n = record.numbers;
    const ratios = record.ratios;
    const clickOrder = n[0] > 0 ? n[2] / n[0] : '';
    const facts = {
      'Тип товара': record.type,
      'Товар': record.title,
      'В корзину из кликов, %': ratios[0],
      'В заказ из корзины, %': ratios[1],
      'Выкуп, %': ratios[2],
      // avgPrice — средняя цена заказов, не текущий прайс. При отсутствии
      // заказов WB возвращает 0: сохранённую цену в таком случае не затираем.
      'Цена, ₽': n[5] > 0 ? n[5] : undefined,
      'Переходы': n[0], 'Корзины': n[1], 'Заказы': n[2], 'Выкупы': n[3],
      'Выручка, ₽': n[4], 'Конверсия клик→заказ, %': clickOrder,
      'WB: период': periodText, 'WB: обновлено (МСК)': now,
      'WB: переходы': n[0], 'WB: корзины': n[1],
      'WB: заказы': n[2], 'WB: выкупы': n[3],
      'WB: выручка выкупов, ₽': n[4], 'WB: средняя цена, ₽': n[5],
      'WB: в корзину, %': ratios[0], 'WB: в заказ, %': ratios[1],
      'WB: выкуп, %': ratios[2], 'WB: заказ / переход, %': clickOrder,
      'WB: остаток WB': record.stockWB, 'WB: остаток продавца': record.stockMP,
      'WB: статус': 'Обновлено', 'WB: название товара': record.title
    };
    Object.keys(facts).forEach(header => put(r, header, facts[header]));
  }
  if (!matched) throw new Error('Артикулы из столбца «Артикул WB» не найдены в ответе WB. Ничего не изменено.');
  // Пишем только изменившиеся ячейки без формул, объединяя смежные по столбцу.
  updates.forEach((cells, column) => {
    let start = 0;
    while (start < cells.length) {
      let end = start + 1;
      while (end < cells.length && cells[end][0] === cells[end - 1][0] + 1) end++;
      sheet.getRange(cells[start][0], column + 1, end - start, 1)
        .setValues(cells.slice(start, end).map(cell => [cell[1]]));
      start = end;
    }
  });
  wbSyncForecast_(records, period);
  console.log('Заполнены существующие строки: ' + matched +
    '. Формулы, заголовки, структура и фильтры сохранены.');
  return matched;
}

function wbSyncForecast_(records, period) {
  if (PropertiesService.getScriptProperties().getProperty('WB_COMPARISON_JOB')) return;
  const book = SpreadsheetApp.openById(WB.bookId);
  const sheet = book.getSheetByName('Прогноз по конверсиям');
  if (!sheet || sheet.getRange('A1').getValue() !== 'Артикул WB' || sheet.getLastRow() < 2) return;
  const headers = sheet.getRange(1, 1, 1, 22).getValues()[0];
  const col = new Map(headers.map((value, index) => [String(value).trim(), index + 1]));
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,22).getValues();
  const byId = new Map(records.map(record => [record.id, record]));
  const costRows = sheet.getRange(2,24,Math.max(sheet.getLastRow()-1,1),3).getValues();
  const costsById = new Map();
  costRows.forEach(r => { const id=String(r[0]).trim(); if (/^\d+$/.test(id) && typeof r[1]==='number') costsById.set(id,r[1]); });
  const output = rows.map(row => row.slice());
  let matches=0;
  output.forEach((row,i) => {
    const id=String(row[0]).trim(), record=byId.get(id);
    if (!record) return;
    matches++;
    const n=record.numbers, ratios=record.ratios;
    if (record.title) row[1]=record.title;
    if (record.type) row[2]=record.type;
    if (record.impressions !== undefined && record.impressions >= 0) row[3]=record.impressions;
    if (record.ctr !== undefined) row[4]=record.ctr;
    else if (Number(row[3])>0) row[4]=n[0]/Number(row[3]);
    row[5]=n[0]; row[6]=n[1]; row[7]=ratios[0]; row[8]=n[2]; row[9]=ratios[1]; row[10]=n[3]; row[11]=ratios[2];
    if ((row[12]==='' || row[12]===null) && costsById.has(id)) row[12]=costsById.get(id);
    if (row[13]==='' || row[13]===null) row[13]=0;
    row[14]=n[4];
    const cost=Number(row[12]), ads=Number(row[13])||0;
    if (row[12]!=='' && Number.isFinite(cost)) {
      row[15]=n[4]-n[3]*cost-ads;
      row[16]=n[3]>0?(n[4]-ads)/n[3]-cost:'';
      row[17]=n[4]>0?row[15]/n[4]:'';
    } else { row[15]=''; row[16]=''; row[17]=''; }
    row[18]=Number(row[3])>0?10000*n[2]/Number(row[3]):'';
    row[19]=Number(row[3])>0?10000*n[3]/Number(row[3]):'';
    row[20]=wbPeriodLabel_(period);
  });
  if (matches) {
    const preserved=sheet.getRange(2,1,output.length,22).getFormulas();
    output.forEach((row,r)=>row.forEach((value,c)=>{if(preserved[r][c]) output[r][c]=preserved[r][c];}));
    sheet.getRange(2,1,output.length,22).setValues(output);
  }
  console.log('Обновлено строк прогноза: '+matches+' из '+output.length+'.');
}

function wbEmptyFormula_(sheet, row, column, formula) {
  const cell = wbRange_(sheet, row, column);
  if (cell.getValue() === '' && cell.getFormula() === '') cell.setFormula(wbFormula_(sheet, formula));
}

function wbExpandFilter_(sheet, end, columns) {
  const filter = sheet.getFilter();
  const criteria = [];
  if (filter) {
    const range = filter.getRange();
    const neededEnd = Math.max(end, range.getLastRow());
    if (range.getRow() === 1 && range.getColumn() === 1 &&
        range.getLastRow() >= end && range.getLastColumn() >= columns) return;
    for (let col = range.getColumn(); col <= range.getLastColumn(); col++) {
      const rule = filter.getColumnFilterCriteria(col);
      if (rule) criteria.push([col, rule]);
    }
    end = neededEnd;
    columns = Math.max(columns, range.getLastColumn());
    filter.remove();
  }
  const next = sheet.getRange(1, 1, Math.max(end, 2), columns).createFilter();
  criteria.forEach(item => next.setColumnFilterCriteria(item[0], item[1]));
}

function wbShowInStock() {
  const sheet = wbSheet_();
  wbCheckInstalled_(sheet);
  wbExpandFilter_(sheet, Math.max(sheet.getLastRow(), 2), 48 + wbColumnShift_(sheet));
  sheet.getFilter().setColumnFilterCriteria(37 + wbColumnShift_(sheet),
    SpreadsheetApp.newFilterCriteria()
      .whenFormulaSatisfied(wbFormula_(sheet, '=AND($AM2="Обновлено",($AK2+$AL2)>0)')).build());
  wbToast_('Показаны товары с положительным остатком WB или продавца.');
}

function wbShowAllStock() {
  const sheet = wbSheet_();
  wbCheckInstalled_(sheet);
  if (sheet.getFilter()) sheet.getFilter().removeColumnFilterCriteria(37 + wbColumnShift_(sheet));
}

function wbCheckInstalled_(sheet) {
  if (sheet.getMaxColumns() < 48 + wbColumnShift_(sheet) || wbRange_(sheet, 'X1').getValue() !== WB.headers[0]) {
    throw new Error('Сначала выполните wbInstall.');
  }
}

function wbToast_(message) {
  console.log(message);
  try { SpreadsheetApp.openById(WB.bookId).toast(message, 'WB Аналитика', 10); }
  catch (_) { /* Журнал выполнения остаётся доступен и без интерфейса таблицы. */ }
}


// Extended seller CSV is generated asynchronously; one request per minute tick.
function wbJamCtrNote_(message) {
  const sheet=SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  if (sheet) sheet.getRange('E1').setNote(message);
  console.log(message);
}
function wbJamCtrQueue_(period) {
  const props=PropertiesService.getScriptProperties();
  if (props.getProperty('WB_JAM_CTR_PERIOD') === wbPeriodLabel_(period)) return;
  const pending=props.getProperty('WB_JAM_CTR_JOB');
  if (pending && wbPeriodLabel_(JSON.parse(pending).period) === wbPeriodLabel_(period)) return;
  props.setProperty('WB_JAM_CTR_JOB',JSON.stringify({id:Utilities.getUuid(),period:period,stage:'create',attempts:0,next:0}));
  wbJamCtrNote_('CTR: запрошена автоматическая проверка расширенного CSV Джем за '+wbPeriodLabel_(period));
}
function wbJamCtrTick_(props) {
  const raw=props.getProperty('WB_JAM_CTR_JOB');
  if (!raw) return;
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(500)) return;
  try {
    const job=JSON.parse(props.getProperty('WB_JAM_CTR_JOB')||'null');
    if (!job || Date.now() < job.next) return;
    if (job.attempts >= 30) {props.deleteProperty('WB_JAM_CTR_JOB');wbJamCtrNote_('CTR: отчёт Джем не получен за 30 попыток.');return;}
    job.attempts++;
    const base='https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads';
    const options={method:job.stage==='create'?'post':'get',headers:{Authorization:props.getProperty('WB_API_TOKEN')},muteHttpExceptions:true,followRedirects:false};
    let url=base;
    if (job.stage==='create') {
      options.contentType='application/json';
      options.payload=JSON.stringify({id:job.id,reportType:'DETAIL_HISTORY_REPORT',userReportName:'CTR automatic check',params:{nmIDs:[],startDate:job.period.start,endDate:job.period.end,timezone:WB.timezone,aggregationLevel:'day',skipDeletedNm:true}});
    } else if (job.stage==='status') url+='?filter%5BdownloadIds%5D='+encodeURIComponent(job.id);
    else url+='/file/'+encodeURIComponent(job.id);
    const response=UrlFetchApp.fetch(url,options), status=response.getResponseCode();
    if (status===429 || status>=500) {
      job.next=Date.now()+Math.max(60000,wbRateLimitDelay_(response.getAllHeaders(),Date.now()));
    } else if (status<200 || status>=300) {
      props.deleteProperty('WB_JAM_CTR_JOB');wbJamCtrNote_('CTR: API расширенного отчёта Джем вернул HTTP '+status+'.');return;
    } else if (job.stage==='create') {job.stage='status';job.next=Date.now()+60000;}
    else if (job.stage==='status') {
      const data=JSON.parse(response.getContentText()).data;
      const item=Array.isArray(data)?data.find(x=>x.id===job.id):null;
      if (item && String(item.status).trim()==='SUCCESS') job.stage='file';
      else if (item && item.status==='FAILED') {props.deleteProperty('WB_JAM_CTR_JOB');wbJamCtrNote_('CTR: генерация отчёта Джем завершилась FAILED.');return;}
      job.next=Date.now()+60000;
    } else {
      const blob=response.getBlob(), bytes=blob.getBytes();
      const texts=bytes[0]===80 && bytes[1]===75?Utilities.unzip(blob).filter(b=>/\.csv$/i.test(b.getName())).map(b=>b.getDataAsString('UTF-8')):[response.getContentText('UTF-8')];
      wbJamCtrApply_(texts,job.period);
      props.setProperty('WB_JAM_CTR_PERIOD',wbPeriodLabel_(job.period));
      props.deleteProperty('WB_JAM_CTR_JOB');return;
    }
    props.setProperty('WB_JAM_CTR_JOB',JSON.stringify(job));
  } catch (_) {wbJamCtrNote_('CTR: ошибка обработки расширенного CSV; повтор при следующей проверке.');}
  finally {lock.releaseLock();}
}
function wbJamCtrApply_(texts,period) {
  if (PropertiesService.getScriptProperties().getProperty('WB_COMPARISON_JOB')) return;
  const totals=new Map(), fields=new Set();
  texts.forEach(text=>{
    text=text.replace(/^\uFEFF/,'');
    const first=text.split(/\r?\n/)[0], separator=first.split(';').length>first.split(',').length?';':',';
    const rows=Utilities.parseCsv(text,separator), h=(rows.shift()||[]).map(x=>x.trim().toLowerCase().replace(/[\s_%-]/g,''));
    h.forEach(x=>fields.add(x));
    const index=names=>h.findIndex(x=>names.indexOf(x)>=0);
    const id=index(['nmid','артикулwb']), shows=index(['impressions','impressioncount','showcount','showscount','shows','показы','количествопоказов']);
    const opens=index(['opencardcount','opencard','clicks','переходывкарточку']), ctr=index(['ctr','ctrpercent']);
    if (id<0 || shows<0) return;
    rows.forEach(r=>{
      const key=String(r[id]).trim(), n=x=>Number(String(x||'').replace(/\s/g,'').replace(',','.'));
      const impressions=n(r[shows]);
      if (!/^\d+$/.test(key) || !Number.isFinite(impressions) || impressions<0) return;
      const clicks=opens>=0?n(r[opens]):ctr>=0?n(r[ctr])*impressions/100:NaN;
      if (!Number.isFinite(clicks) || clicks<0) return;
      const t=totals.get(key)||{shows:0,clicks:0};t.shows+=impressions;t.clicks+=clicks;totals.set(key,t);
    });
  });
  if (!totals.size) {wbJamCtrNote_('Расширенный CSV Джем получен. Поля показов/CTR не найдены. Фактические заголовки: '+Array.from(fields).join(', '));return;}
  const sheet=SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,22).getValues();let updated=0;
  rows.forEach((r,i)=>{
    const t=totals.get(String(r[0]));
    if (!t || r[20]!==wbPeriodLabel_(period)) return;
    sheet.getRange(i+2,4,1,2).setValues([[t.shows,t.shows>0?t.clicks/t.shows:'']]);
    sheet.getRange(i+2,19,1,2).setValues([[t.shows>0?10000*Number(r[8])/t.shows:'',t.shows>0?10000*Number(r[10])/t.shows:'']]);updated++;
  });
  wbJamCtrNote_('Показы и CTR получены из расширенного CSV Джем. Обновлено артикулов: '+updated+'. Период '+wbPeriodLabel_(period));
}


// Read-only permission check using the existing token inside Apps Script.
function wbProbeAdvertisingAccess_(token) {
  const props=PropertiesService.getScriptProperties();
  const today=Utilities.formatDate(new Date(),WB.timezone,'yyyy-MM-dd');
  if (props.getProperty('WB_ADVERT_PROBE_DATE')===today) return;
  let message;
  try {
    const response=UrlFetchApp.fetch('https://advert-api.wildberries.ru/adv/v1/promotion/count',{
      headers:{Authorization:token},muteHttpExceptions:true,followRedirects:false});
    const status=response.getResponseCode();
    if (status===200) {
      const data=JSON.parse(response.getContentText());
      const total=typeof data.all==='number'?data.all:'не указан';
      message='Проверка имеющегося WB_API_TOKEN: API Продвижение доступен (HTTP 200). Кампаний: '+total+'. Показы рекламных кампаний доступны отдельным методом /adv/v3/fullstats; это не общие показы воронки.';
    } else message='Проверка имеющегося WB_API_TOKEN: API Продвижение вернул HTTP '+status+'. Ключ не изменялся.';
    if (status!==429 && status<500) props.setProperty('WB_ADVERT_PROBE_DATE',today);
  } catch (_) {message='Проверка имеющегося WB_API_TOKEN: не удалось связаться с API Продвижение.';}
  props.setProperty('WB_ADVERT_PROBE_STATUS',message);
  const sheet=SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  if (sheet) sheet.getRange('D1').setNote(message);
  console.log(message);
}


function wbProbePortalAccess_() {
  const props=PropertiesService.getScriptProperties();
  if (props.getProperty('WB_PORTAL_PROBE_V1')) return;
  const token=props.getProperty('WB_API_TOKEN');
  if (!token) return;
  let message;
  try {
    const period=wbPeriod_(1);
    const response=UrlFetchApp.fetch('https://seller-content.wildberries.ru/ns/analytics-api/content-analytics/api/v1/sales-funnel/report/details',{
      method:'post',contentType:'application/json',headers:{Authorization:token},
      payload:JSON.stringify({currentPeriod:period,subjects:[],brands:[],nms:[],tagIds:[],skipDeletedNm:true,orderBy:{field:'openCardCount',mode:'desc'},limit:1,offset:0}),
      muteHttpExceptions:true,followRedirects:false,validateHttpsCertificates:true});
    const status=response.getResponseCode();
    message='Прямой запрос к общей воронке кабинета с существующим WB_API_TOKEN: HTTP '+status+'.';
    if (status===200) {
      try {message+=' Поля ответа: '+wbMetricPaths_(JSON.parse(response.getContentText())).join(', ');}
      catch (_) {message+=' Ответ не JSON.';}
    }
    if (status!==429 && status<500) props.setProperty('WB_PORTAL_PROBE_V1',message);
  } catch (_) {message='Прямой запрос к общей воронке: сетевой сбой.';}
  const sheet=SpreadsheetApp.openById(WB.bookId).getSheetByName('Прогноз по конверсиям');
  if (sheet) sheet.getRange('F1').setNote(message);
  console.log(message);
}
