/** Add once as a separate file in the spreadsheet's bound Apps Script.
 * Set WB_COST_RECEIVER_TOKEN in Script Properties to a random secret (32+ chars).
 * Deploy as web app executing as owner. Never put the token in this file.
 * Only accepts costs, never code, formulas, arbitrary sheet IDs or ranges.
 */
function doPost(e) {
  const reply = value => ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty('WB_COST_RECEIVER_TOKEN');
  let request;
  try {
    if (!e || !e.postData || e.postData.contents.length > 1000000) return reply({ok:false,error:'Invalid request'});
    request = JSON.parse(e.postData.contents);
  } catch (_) { return reply({ok:false,error:'Invalid JSON'}); }
  if (!secret || secret.length < 32 || request.token !== secret) return reply({ok:false,error:'Unauthorized'});
  if (request.action === 'funnel') return wbReceiveFunnel_(request, reply, props);
  if (request.schema !== 1 || !Array.isArray(request.costs) || !request.costs.length || request.costs.length > 5000) {
    return reply({ok:false,error:'Invalid schema'});
  }
  const ids = new Set();
  for (const cost of request.costs) {
    if (!cost || !/^\d{1,16}$/.test(String(cost.id)) || ids.has(String(cost.id)) ||
        typeof cost.amount !== 'number' || !Number.isFinite(cost.amount) || cost.amount <= 0 ||
        cost.amount > 10000000 || !/^\d{4}-\d{2}-\d{2}$/.test(cost.date) ||
        !Number.isFinite(Date.parse(cost.date))) return reply({ok:false,error:'Invalid cost row'});
    ids.add(String(cost.id));
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return reply({ok:false,error:'Busy; retry later'});
  try {
    const bookId = props.getProperty('WB_SPREADSHEET_ID');
    if (!bookId) return reply({ok:false,error:'Run wbInstall first'});
    const sheet = SpreadsheetApp.openById(bookId).getSheetByName('Прогноз по конверсиям');
    if (!sheet) return reply({ok:false,error:'Forecast sheet missing'});
    const expected = ['WB ID из файла себестоимости','Себестоимость из файла, ₽','Последняя дата в файле'];
    if (sheet.getMaxColumns() < 26 || !expected.every((v,i) => sheet.getRange(1,24+i).getValue() === v)) {
      return reply({ok:false,error:'Cost helper headers mismatch'});
    }
    const count = Math.max(1,sheet.getLastRow()-1);
    const helper = sheet.getRange(2,24,count,3).getValues();
    const previous = new Map(helper.map(r => [String(r[0]),r[1]]));
    const byId = new Map();
    helper.forEach((r,i) => { if (/^\d+$/.test(String(r[0]))) byId.set(String(r[0]),i); });
    let imported=0, skippedOlder=0;
    request.costs.forEach(cost => {
      const id=String(cost.id), index=byId.get(id);
      const existing=index === undefined ? null : helper[index];
      const oldDate=existing && existing[2] instanceof Date ? Utilities.formatDate(existing[2],'Europe/Moscow','yyyy-MM-dd') : String(existing ? existing[2] : '');
      if (oldDate && oldDate > cost.date) { skippedOlder++; return; }
      const row=[id,cost.amount,cost.date];
      if (index === undefined) { byId.set(id,helper.length); helper.push(row); }
      else helper[index]=row;
      imported++;
    });
    if (helper.length+1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),helper.length+1-sheet.getMaxRows());
    sheet.getRange(2,24,helper.length,3).setValues(helper);
    let updated=0;
    if (sheet.getRange('M1').getValue() === 'Себестоимость за единицу, ₽ (ввод)') {
      const costs = new Map(helper.map(r => [String(r[0]),r[1]]));
      const rows = sheet.getRange(2,1,count,22).getValues();
      const formulas = sheet.getRange(2,13,count,1).getFormulas();
      rows.forEach((r,i) => {
        const id=String(r[0]);
        if (!ids.has(id) || !costs.has(id) || formulas[i][0] ||
            (r[12] !== '' && r[12] !== previous.get(id))) return;
        const row=i+2;
        sheet.getRange(row,13).setValue(costs.get(id));
        sheet.getRange(row,16,1,3).setFormulas([[
          '=IF(OR(M'+row+'="",O'+row+'=""),"",O'+row+'-K'+row+'*M'+row+'-N'+row+')',
          '=IFERROR(P'+row+'/K'+row+',"")',
          '=IFERROR(P'+row+'/O'+row+',"")'
        ]]);
        updated++;
      });
    }
    SpreadsheetApp.flush();
    return reply({ok:true,imported:imported,updated:updated,skippedOlder:skippedOlder});
  } catch (_) {
    return reply({ok:false,error:'Write failed; inspect Apps Script execution log'});
  } finally { lock.releaseLock(); }
}

/** Accepts only normalized funnel metrics and writes existing cells on the two analytics tabs. */
function wbReceiveFunnel_(request, reply, props) {
  if (request.schema !== 1 || !/^\d{4}-\d{2}-\d{2}$/.test(request.start) ||
      !/^\d{4}-\d{2}-\d{2}$/.test(request.end) || request.start > request.end ||
      !Array.isArray(request.rows) || !request.rows.length || request.rows.length > 1000) {
    return reply({ok:false,error:'Invalid funnel schema'});
  }
  const seen = new Set();
  for (const row of request.rows) {
    if (!row || !/^\d{1,16}$/.test(String(row.id)) || seen.has(String(row.id)) ||
        !Number.isSafeInteger(row.impressions) || row.impressions < 0 || row.impressions > 1000000000 ||
        !Number.isFinite(row.ctr) || row.ctr < 0 || row.ctr > 1 ||
        !['clicks','carts','orders','buyouts'].every(k => Number.isSafeInteger(row[k]) && row[k] >= 0 && row[k] <= 1000000000)) {
      return reply({ok:false,error:'Invalid funnel row'});
    }
    seen.add(String(row.id));
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return reply({ok:false,error:'Busy; retry later'});
  try {
    const bookId = props.getProperty('WB_SPREADSHEET_ID');
    if (!bookId) return reply({ok:false,error:'Run wbInstall first'});
    const book = SpreadsheetApp.openById(bookId);
    const forecast = book.getSheetByName('Прогноз по конверсиям');
    const source = book.getSheetByName('Рабочая табличка');
    if (!forecast || !source) return reply({ok:false,error:'Required analytics tab missing'});
    let updatedForecast = 0, updatedSource = 0;
    const byId = new Map(request.rows.map(row => [String(row.id), row]));

    // Forecast layout uses fixed headers established by wbInstall. Only existing
    // product rows and non-formula cells are eligible for updates.
    if (forecast.getLastRow() >= 2) {
      const n = forecast.getLastRow() - 1;
      const values = forecast.getRange(1, 1, n + 1, 22).getValues();
      const formulas = forecast.getRange(1, 1, n + 1, 22).getFormulas();
      const expected = ['Артикул WB','Показы (ввод)','CTR, %','Переходы','Корзины','Заказы','Выкупы','Период данных WB'];
      const cols = [0,3,4,5,6,8,10,20];
      if (!expected.every((h,i) => String(values[0][cols[i]]).trim() === h)) {
        return reply({ok:false,error:'Forecast headers mismatch; nothing written'});
      }
      for (let r = 1; r < values.length; r++) {
        const item = byId.get(String(values[r][0]).trim());
        if (!item) continue;
        const writes = [[3,item.impressions],[4,item.ctr],[5,item.clicks],[6,item.carts],
          [8,item.orders],[10,item.buyouts],[20,request.start+' — '+request.end]];
        writes.forEach(([c,v]) => {
          if (!formulas[r][c] && values[r][c] !== v) {
            forecast.getRange(r + 1, c + 1).setValue(v);
            updatedForecast++;
          }
        });
      }
    }

    // Source-tab impressions/CTR are updated only when a recognized existing
    // heading is present. Other funnel figures continue through the official WB API.
    if (source.getLastRow() >= 2) {
      const rows = source.getLastRow(), cols = source.getLastColumn();
      const values = source.getRange(1, 1, rows, cols).getValues();
      const formulas = source.getRange(1, 1, rows, cols).getFormulas();
      const headers = new Map(values[0].map((v,i) => [String(v).trim(),i]));
      const idCol = headers.get('Артикул WB');
      if (idCol === undefined) return reply({ok:false,error:'Source product ID header missing'});
      const impressionCol = ['Показы','Показы, шт','Показы карточки','Количество показов']
        .map(h => headers.get(h)).find(c => c !== undefined);
      const ctrCol = ['CTR','CTR, %','CTR (%)'].map(h => headers.get(h)).find(c => c !== undefined);
      for (let r = 1; r < values.length; r++) {
        const item = byId.get(String(values[r][idCol]).trim());
        if (!item) continue;
        [[impressionCol,item.impressions],[ctrCol,item.ctr]].forEach(([c,v]) => {
          if (c !== undefined && !formulas[r][c] && values[r][c] !== v) {
            source.getRange(r + 1, c + 1).setValue(v);
            updatedSource++;
          }
        });
      }
    }
    SpreadsheetApp.flush();
    return reply({ok:true,received:request.rows.length,updatedForecast:updatedForecast,updatedSource:updatedSource});
  } catch (_) {
    return reply({ok:false,error:'Funnel write failed; inspect Apps Script execution log'});
  } finally { lock.releaseLock(); }
}
