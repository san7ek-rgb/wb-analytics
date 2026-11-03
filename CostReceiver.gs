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
