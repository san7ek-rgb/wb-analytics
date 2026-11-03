/**
 * WB → «Таблица конверсий». Вставить целиком в привязанный Apps Script.
 * Ключ: свойство скрипта WB_API_TOKEN. Установка: wbInstall.
 * Слева автоматически добавляется A «Тип товара». Старые столбцы сдвигаются.
 * B:X существующих строк не перезаписываются. Y:AW — блок интеграции.
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
  sheet: 'Таблица конверсий',
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
  // Сначала успешная загрузка; затем регистрация ежедневного запуска.
  wbUpdateAll();
  const old = wbOwnTriggers_();
  ScriptApp.newTrigger('wbUpdateAll').timeBased().everyDays(1)
    .atHour(4).inTimezone(WB.timezone).create();
  old.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  onOpen();
  wbToast_('Готово: обновление ежедневно между 04:00 и 05:00 МСК.');
}

function wbDisable() {
  wbOwnTriggers_().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  wbToast_('Ежедневное обновление этого пользователя отключено.');
}

function wbOwnTriggers_() {
  // Только обработчики этой интеграции и ранее выданного скрипта артикулов.
  const names = ['wbUpdateAll', 'refreshArticles', 'setupArticleFilter'];
  return ScriptApp.getProjectTriggers().filter(trigger =>
    names.indexOf(trigger.getHandlerFunction()) !== -1);
}

function wbYesterday() { wbSetDays_(1); }
function wbWeek() { wbSetDays_(7); }
function wbMonth() { wbSetDays_(30); }

function wbSetDays_(days) {
  wbRun_(days);
  PropertiesService.getScriptProperties().setProperty('WB_REPORT_DAYS', String(days));
}

function wbUpdateAll() {
  const days = Number(PropertiesService.getScriptProperties()
    .getProperty('WB_REPORT_DAYS') || 1);
  wbRun_(days);
}

function wbRun_(days) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Обновление уже выполняется. Повторный запуск не требуется.');
  }
  const props = PropertiesService.getScriptProperties();
  try {
    if ([1, 7, 30].indexOf(days) === -1) {
      throw new Error('WB_REPORT_DAYS должен быть равен 1, 7 или 30.');
    }
    const token = (props.getProperty('WB_API_TOKEN') || '').trim();
    if (!token) throw new Error('Добавьте WB_API_TOKEN в свойства скрипта.');
    const sheet = wbSheet_();
    wbCheckLayout_(sheet);
    const period = wbPeriod_(days);
    const state = { deadline: Date.now() + WB.budgetMs };
    const products = wbFetchProducts_(token, period, state);
    if (!products.length) {
      throw new Error('WB вернул пустой список. Существующие данные не изменены.');
    }
    // Все страницы загружаются и проверяются до записи в таблицу.
    const records = products.map(item => wbRecord_(item, period));
    const ids = new Set();
    records.forEach(record => {
      if (ids.has(record.id)) {
        throw new Error('WB вернул повторяющиеся карточки. Повторите обновление позже.');
      }
      ids.add(record.id);
    });
    wbEnsureTypeColumn_(sheet);
    wbWrite_(sheet, records, period);
    SpreadsheetApp.flush();
    const now = Utilities.formatDate(new Date(), WB.timezone, 'yyyy-MM-dd HH:mm:ss');
    props.setProperties({ WB_LAST_SUCCESS: now, WB_LAST_ERROR: '' });
    wbToast_('Обновлено товаров: ' + records.length + '. Период: ' + wbPeriodLabel_(period));
  } catch (error) {
    // Не записываем тела ответов WB, ключ или заголовки запросов в журнал.
    const message = error.message || 'Ошибка обновления WB.';
    props.setProperty('WB_LAST_ERROR', message.slice(0, 1000));
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
    const next = Number(props.getProperty('WB_NEXT_REQUEST_AT') || 0);
    const wait = Math.max(0, next - Date.now());
    if (Date.now() + wait + 15000 > state.deadline) {
      throw new Error('Достигнут лимит времени запуска. Выберите меньший объём или повторите позже.');
    }
    if (wait) Utilities.sleep(wait);
    let response;
    // Сохраняем интервал и между ручным и ежедневным запуском.
    props.setProperty('WB_NEXT_REQUEST_AT', String(Date.now() + WB.intervalMs));
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
    if (status === 200) {
      try { return JSON.parse(response.getContentText()); }
      catch (_) { throw new Error('WB вернул некорректный JSON. Данные не записаны.'); }
    }
    if (status === 401) throw new Error('WB HTTP 401: проверьте токен и срок его действия.');
    if (status === 403) throw new Error('WB HTTP 403: проверьте права категории «Аналитика».');
    if (status === 429 || status >= 500) {
      let seconds = 0;
      const headers = response.getAllHeaders();
      Object.keys(headers).forEach(key => {
        if (/^(retry-after|x-ratelimit-retry|x-ratelimit-reset)$/i.test(key)) {
          const value = Number(headers[key]);
          if (Number.isFinite(value) && value > seconds) seconds = value;
        }
      });
      props.setProperty('WB_NEXT_REQUEST_AT', String(Date.now() +
        Math.max(WB.intervalMs * (attempt + 1), seconds * 1000)));
      if (attempt < 2) continue;
    }
    throw new Error('API воронки WB: HTTP ' + status + '. Данные не записаны.');
  }
}

function wbNumber_(value, field) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
    throw new Error('Некорректное или отсутствующее поле WB: ' + field + '. Данные не записаны.');
  }
  return value;
}

function wbRecord_(item, period) {
  const product = item && item.product;
  const stats = item && item.statistic && item.statistic.selected;
  if (!product || !stats || !stats.period ||
      stats.period.start !== period.start || stats.period.end !== period.end ||
      !Number.isSafeInteger(product.nmId) || product.nmId <= 0 ||
      typeof product.title !== 'string') {
    throw new Error('Некорректная карточка или период в ответе WB. Данные не записаны.');
  }
  const fields = ['openCount', 'cartCount', 'orderCount', 'buyoutCount', 'buyoutSum', 'avgPrice'];
  const numbers = fields.map(field => wbNumber_(stats[field], field));
  const conversions = stats.conversions || {};
  const ratios = ['addToCartPercent', 'cartToOrderPercent', 'buyoutPercent']
    .map(field => wbNumber_(conversions[field], field) / 100);
  const stocks = product.stocks || {};
  return {
    id: String(product.nmId), title: wbSafeText_(product.title),
    type: wbProductType_(product.subjectName),
    numbers: numbers, ratios: ratios,
    stockWB: wbNumber_(stocks.wb, 'stocks.wb'),
    stockMP: wbNumber_(stocks.mp, 'stocks.mp')
  };
}

function wbSafeText_(text) { return /^[=+@-]/.test(text) ? "'" + text : text; }

function wbWrite_(sheet, records, period) {
  const requiredCols = 23 + WB.headers.length;
  const actualCols = requiredCols + wbColumnShift_(sheet);
  if (sheet.getMaxColumns() < actualCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), actualCols - sheet.getMaxColumns());
  }
  const last = Math.max(sheet.getLastRow(), 1);
  const grid = wbRange_(sheet, 1, 1, last, requiredCols).getValues();
  const formulas = wbRange_(sheet, 1, 1, last, 23).getFormulas();
  const rowById = new Map();
  const available = [];
  for (let r = 1; r < grid.length; r++) {
    const id = String(grid[r][1]).trim();
    if (/^\d+$/.test(id)) {
      if (rowById.has(id)) {
        throw new Error('Повторяющийся артикул в B: ' + id + '. Устраните дубликат перед обновлением.');
      }
      rowById.set(id, r + 1);
    } else if (grid[r].every(value => value === '') &&
               formulas[r][0] === '' && formulas[r][1] === '') {
      available.push(r + 1);
    }
  }
  let append = last + 1;
  const planned = records.map(record => ({
    record: record,
    row: rowById.get(record.id) || available.shift() || append++,
    isNew: !rowById.has(record.id)
  }));
  const end = Math.max(last, ...planned.map(item => item.row));
  if (end > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), end - sheet.getMaxRows());
  }
  wbRange_(sheet, 1, 24, 1, WB.headers.length).setValues([WB.headers]);
  const label = wbPeriodLabel_(period);
  const now = Utilities.formatDate(new Date(), WB.timezone, 'yyyy-MM-dd HH:mm:ss');
  const received = new Set(records.map(record => record.id));
  planned.forEach(item => {
    const r = item.row;
    const record = item.record;
    if (item.isNew) wbRange_(sheet, r, 1, 1, 2).setValues([[record.title, Number(record.id)]]);
    if (wbColumnShift_(sheet)) sheet.getRange(r, 1).setValue(record.type);
    const clickOrder = record.numbers[0] > 0
      ? record.numbers[2] / record.numbers[0] : '';
    wbRange_(sheet, r, 24, 1, 12).setValues([[
      label, now, ...record.numbers, ...record.ratios, clickOrder
    ]]);
    wbRange_(sheet, r, 37, 1, 3).setValues([[
      record.stockWB, record.stockMP, 'Обновлено'
    ]]);
    wbRange_(sheet, r, 48).setValue(record.title);
    // В AN:AO ничего не записываем: это входные данные пользователя.
    wbEmptyFormula_(sheet, r, 36,
      '=IF(AND(AM' + r + '="Обновлено",ISNUMBER(I' + r + ')),AC' + r + '*I' + r + ',"")');
    const valid = 'AND(AM' + r + '="Обновлено",ISNUMBER(AN' + r + '),AN' + r + '>0,AO' + r + '=X' + r + ')';
    const calculated = [
      '=IF(' + valid + ',Z' + r + '/AN' + r + ',"")',
      '=IF(' + valid + ',10000*AB' + r + '/AN' + r + ',"")',
      '=IF(AQ' + r + '="","",AQ' + r + '*AH' + r + ')',
      '=IF(AND(AR' + r + '<>"",ISNUMBER(I' + r + ')),AR' + r + '*I' + r + ',"")',
      '=IF(AS' + r + '="","",RANK(AS' + r + ',$AS$2:$AS,0))',
      '=IF(AS' + r + '="","",IF(AS' + r + '=MAX($AS$2:$AS),"Лидер по прогнозу","Сравнить с лидером"))'
    ];
    calculated.forEach((formula, index) => wbEmptyFormula_(sheet, r, 42 + index, formula));
  });
  rowById.forEach((r, id) => {
    if (!received.has(id) && grid[r - 1][23] !== '') {
      wbRange_(sheet, r, 39).setValue('Не получен в текущем ответе — данные прежнего периода');
    }
  });
  wbRange_(sheet, 1, 40).setNote(
    'Общие показы воронка WB не возвращает. Введите показы из сопоставимого источника за период в столбце «WB: период» этой строки. Не смешивайте рекламные показы с общими переходами.'
  );
  wbRange_(sheet, 1, 41).setNote(
    'Для расчётов скопируйте сюда текст из столбца «WB: период» той же строки. При смене периода старые показы автоматически исключаются из прогнозов.'
  );
  wbRange_(sheet, 1, 36).setNote(
    'Расчёт: фактические выкупы WB × прибыль из столбца «Прибыль с выкупа». Это оценка по вашей марже, не финансовый отчёт WB. Маржа заполняется вручную.'
  );
  wbRange_(sheet, 1, 34).setNote(
    'Процент выкупа WB исключает ещё доставляемые товары. Выкупы воронки относятся к выбранному периоду заказов и могут уточняться позднее.'
  );
  if (end > 1) {
    wbRange_(sheet, 2, 32, end - 1, 4).setNumberFormat('0.00%');
    wbRange_(sheet, 2, 42, end - 1, 1).setNumberFormat('0.00%');
    wbRange_(sheet, 2, 43, end - 1, 3).setNumberFormat('0.00');
  }
  wbExpandFilter_(sheet, end, actualCols);
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
