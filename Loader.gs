/**
 * Постоянный загрузчик. Один раз заменить им весь код привязанного Apps Script.
 * Затем выполнить wbInstall. Ключ WB_API_TOKEN остаётся в свойствах скрипта.
 * Исполняет код из указанного доверенного репозитория с правами владельца триггера.
 * Не менять OWNER/REPO на сторонний репозиторий и не давать посторонним право записи.
 */
const WB_LOADER = Object.freeze({
  base: 'https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/main/',
  handlers: ['wbInstall', 'wbUpdateAll', 'wbYesterday', 'wbWeek', 'wbMonth',
    'wbShowInStock', 'wbShowAllStock', 'wbDisable']
});

// Простой триггер не обращается к сети и не требует авторизации.
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
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Откройте этот проект через Расширения → Apps Script в таблице.');
  PropertiesService.getScriptProperties().setProperty('WB_SPREADSHEET_ID', active.getId());
  wbRunRemote_('wbInstall');
}

function wbUpdateAll() { wbRunRemote_('wbUpdateAll'); }
function wbYesterday() { wbRunRemote_('wbYesterday'); }
function wbWeek() { wbRunRemote_('wbWeek'); }
function wbMonth() { wbRunRemote_('wbMonth'); }
function wbShowInStock() { wbRunRemote_('wbShowInStock'); }
function wbShowAllStock() { wbRunRemote_('wbShowAllStock'); }

// Отключение работает даже при недоступности GitHub.
function wbDisable() {
  const handlers = ['wbUpdateAll', 'refreshArticles', 'setupArticleFilter'];
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log('Ежедневное обновление этого пользователя отключено.');
}

function wbReadPublic_(path) {
  const response = UrlFetchApp.fetch(WB_LOADER.base + path, {
    method: 'get', muteHttpExceptions: true,
    followRedirects: false, validateHttpsCertificates: true
  });
  const status = response.getResponseCode();
  if (status !== 200) {
    throw new Error('Не удалось загрузить версию из GitHub: HTTP ' + status +
      '. Данные не обновлены; повторите запуск позже.');
  }
  return response.getContentText();
}

function wbRunRemote_(handler) {
  if (WB_LOADER.handlers.indexOf(handler) === -1) {
    throw new Error('Неизвестная команда загрузчика.');
  }
  let release;
  try { release = JSON.parse(wbReadPublic_('release.json')); }
  catch (error) { throw new Error('Не удалось прочитать описание версии. ' + error.message); }
  if (release.schema !== 1 || !/^[a-f0-9]{64}$/.test(release.sha256) ||
      release.path !== 'releases/' + release.sha256 + '.gs') {
    throw new Error('Некорректное описание версии. Код не запущен.');
  }
  const source = wbReadPublic_(release.path);
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8
  ).map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
  if (digest !== release.sha256) {
    throw new Error('Контрольная сумма версии не совпадает. Код не запущен.');
  }
  console.log('Версия WB: ' + release.sha256.slice(0, 12));
  // Ключ WB не передаётся в GitHub. Исполняется только проверенный выпуск
  // из фиксированного репозитория; entry point выбран из закрытого списка.
  eval(source + '\n' + handler + '();');
}
