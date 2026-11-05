/** Add once as a separate file in the spreadsheet's bound Apps Script.
 * Set WB_COST_RECEIVER_TOKEN in Script Properties to a random secret (32+ chars).
 * Deploy as web app executing as owner. Never put the token in this file.
 * Accepts validated costs, funnel metrics and comparison jobs; never executable code or arbitrary ranges.
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
  if (request.action === 'task' || request.action === 'jobError' || request.action === 'jobDefer') return wbComparisonTask_(request, reply, props);
  if (request.action === 'catalog') return wbReceiveCatalog_(request, reply, props);
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
    const helperColumn=wbComparisonCostColumn_(sheet);
    const expected = ['WB ID из файла себестоимости','Себестоимость из файла, ₽','Последняя дата в файле'];
    if (sheet.getMaxColumns() < helperColumn+2 || !expected.every((v,i) => sheet.getRange(1,helperColumn+i).getValue() === v)) {
      return reply({ok:false,error:'Cost helper headers mismatch'});
    }
    const count = Math.max(1,sheet.getLastRow()-1);
    const helper = sheet.getRange(2,helperColumn,count,3).getValues();
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
    sheet.getRange(2,helperColumn,helper.length,3).setValues(helper);
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
        ['cartRate','orderRate','buyoutRate'].some(k=>row[k]!==undefined && row[k]!==null && (!Number.isFinite(row[k]) || row[k]<0 || row[k]>1)) ||
        (row.revenue !== undefined && row.revenue !== null && (!Number.isFinite(row.revenue) || row.revenue < 0 || row.revenue > 1000000000000)) ||
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
    wbEnsureComparisonLayout_(forecast);
    const job=JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
    if (request.jobId && (!job || job.id!==request.jobId || ['queued','running'].indexOf(job.status)<0 ||
        job.start!==request.start || job.end!==request.end || request.rows.some(r=>job.ids.indexOf(String(r.id))<0))) {
      return reply({ok:false,error:'Comparison job or period mismatch'});
    }
    let updatedForecast = 0, updatedSource = 0;
    const byId = new Map(request.rows.map(row => [String(row.id), row]));

    // Forecast layout uses fixed headers established by wbInstall. Only existing
    // product rows and non-formula cells are eligible for updates.
    if (forecast.getLastRow() >= 2 && (request.jobId || !job)) {
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
        const cost=typeof values[r][12]==='number' && values[r][12]>=0?values[r][12]:null;
        const ads=typeof values[r][13]==='number'?values[r][13]:0;
        const revenue=typeof item.revenue==='number'?item.revenue:null;
        const profit=cost!==null && revenue!==null?revenue-item.buyouts*cost-ads:null;
        const rates=wbConversionRates_(item);
        const writes = [[3,item.impressions],[4,item.ctr],[5,item.clicks],[6,item.carts],
          [7,rates[1]===null?'':rates[1]],[8,item.orders],
          [9,rates[2]===null?'':rates[2]],[10,item.buyouts],
          [11,rates[3]===null?'':rates[3]],[20,request.start+' — '+request.end],
          [18,item.impressions>0?10000*item.orders/item.impressions:''],
          [19,item.impressions>0?10000*item.buyouts/item.impressions:'']];
        if (request.jobId || item.revenue!==undefined) writes.push([14,revenue===null?'':revenue],
          [15,profit===null?'':profit],[16,profit!==null && item.buyouts>0?profit/item.buyouts:''],
          [17,profit!==null && revenue>0?profit/revenue:'']);
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
    if (source.getLastRow() >= 2 && !request.jobId) {
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
    if (request.jobId) {
      wbWriteComparisonAbc_(forecast,job.ids,request.rows,false);
      SpreadsheetApp.flush();
      wbSaveComparisonRows_(request.rows,props);
      wbComparisonClearWake_();
      job.status='completed'; job.completed=new Date().toISOString();
      job.message='Сравнение готово: получено '+request.rows.length+' из '+job.ids.length+' товаров.';
      props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));
    }
    return reply({ok:true,received:request.rows.length,updatedForecast:updatedForecast,updatedSource:updatedSource});
  } catch (_) {
    return reply({ok:false,error:'Funnel write failed; inspect Apps Script execution log'});
  } finally { lock.releaseLock(); }
}

/** Read existing product IDs only; no caller-selected sheets or ranges. */
function wbReceiveCatalog_(request, reply, props) {
  if (request.schema !== 1) return reply({ok:false,error:'Invalid schema'});
  const bookId = props.getProperty('WB_SPREADSHEET_ID');
  if (!bookId) return reply({ok:false,error:'Run wbInstall first'});
  const book = SpreadsheetApp.openById(bookId);
  const ids = new Set();
  ['Прогноз по конверсиям','Рабочая табличка'].forEach(name => {
    const sheet = book.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headings = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    const column = headings.findIndex(h => h.trim() === 'Артикул WB');
    if (column < 0) return;
    sheet.getRange(2,column+1,sheet.getLastRow()-1,1).getDisplayValues().forEach(row => {
      const id = String(row[0]).trim();
      if (/^\d{1,16}$/.test(id)) ids.add(id);
    });
  });
  return reply({ok:true,ids:Array.from(ids).slice(0,300),total:ids.size,limit:300});
}

// GENERATED COMPARISON UI — edit src/Comparison.gs and src/ComparisonSidebar.html
/** Comparison UI and queue. Included in the standalone receiver bundle. */
function wbReceiverUpgrade_(releaseHash) {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('WB_COMPARISON_STYLE_RELEASE')===releaseHash) return;
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000)) return;
  try {
    if(props.getProperty('WB_COMPARISON_STYLE_RELEASE')!==releaseHash) {
      wbColorComparisonColumns_();
      props.setProperty('WB_COMPARISON_STYLE_RELEASE',releaseHash);
    }
  } finally {lock.releaseLock();}
}

function wbComparisonSheet_() {
  const props = PropertiesService.getScriptProperties();
  const bookId = props.getProperty('WB_SPREADSHEET_ID');
  if (!bookId) throw new Error('Не задан WB_SPREADSHEET_ID.');
  const sheet = SpreadsheetApp.openById(bookId).getSheetByName('Прогноз по конверсиям');
  if (!sheet || sheet.getRange('A1').getValue() !== 'Артикул WB' || sheet.getRange('V1').getValue() !== 'Сравнить') {
    throw new Error('Не найдена ожидаемая структура «Прогноз по конверсиям».');
  }
  return sheet;
}

function wbInstallComparison() {
  wbComparisonSheet_();
  wbColorComparisonColumns_();
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'wbComparisonOnOpen');
  if (!triggers.length) ScriptApp.newTrigger('wbComparisonOnOpen').forSpreadsheet(wbComparisonSheet_().getParent()).onOpen().create();
  wbComparisonOnOpen();
  wbOpenComparison();
}

function wbComparisonOnOpen() {
  SpreadsheetApp.getUi().createMenu('Сравнение WB').addItem('Открыть панель сравнения', 'wbOpenComparison').addToUi();
}

function wbColorComparisonColumns_() {
  const sheet=wbComparisonSheet_(),height=Math.max(2,sheet.getLastRow());
  wbEnsureComparisonLayout_(sheet);
  sheet.getBandings().forEach(b=>{const r=b.getRange();if(r.getColumn()<=22 && r.getLastColumn()<=22)b.remove();});
  const backgrounds=[];
  for(let row=0;row<height;row++) {
    const cells=[];
    for(let col=0;col<22;col++) {
      const manual=col===12||col===13,selection=col===21;
      cells.push(row===0?(manual?'#805e30':selection?'#2b665b':'#294e70'):
        manual?(row%2?'#ecdfc6':'#e3d3b6'):selection?(row%2?'#d1e6db':'#c1dbce'):(row%2?'#e0eaf3':'#d2e0ec'));
    }
    backgrounds.push(cells);
  }
  sheet.getRange(1,1,height,22).setBackgrounds(backgrounds);
  sheet.getRange(1,1,1,22).setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('M1').setNote('Ручной ввод: себестоимость единицы. Автоматическая загрузка WB этот столбец не меняет.');
  sheet.getRange('N1').setNote('Ручной ввод: расходы на рекламу за выбранный период. Для включения в сравнение подтвердите период в правой панели.');
  sheet.getRange('V1').setNote('Выбор товаров для сравнения. Панель «Сравнение WB» → «Открыть панель сравнения».');
  sheet.getRange(1,1,1,12).setNote('Автоматические данные. Меняйте товары и период через панель сравнения; значения воронки вручную не редактируйте.');
  sheet.getRange(1,15,1,7).setNote('Автоматические данные и расчёты. Формулы сохраняются.');
  const abcBackgrounds=Array.from({length:height},(_,r)=>[r===0?'#294e70':r%2?'#e0eaf3':'#d2e0ec',r===0?'#294e70':r%2?'#e0eaf3':'#d2e0ec']);
  sheet.getRange(1,23,height,2).setBackgrounds(abcBackgrounds);
  sheet.getRange(1,23,1,2).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.getRange('W1').setNote('ABC только по четырём конверсиям с равным весом: CTR, переход → корзина, корзина → заказ, заказ → выкуп. Сравниваются выбранные товары за один период.');
  sheet.getRange('X1').setNote('Относительное отклонение сводного индекса конверсий от среднего среди выбранных товаров. +20% означает индекс на 20% выше среднего, не разницу в процентных пунктах.');
  const props=PropertiesService.getScriptProperties(),job=JSON.parse(props.getProperty('WB_COMPARISON_JOB')||'null');
  if(job) wbWriteComparisonAbc_(sheet,job.ids,job.status==='completed'?wbComparisonFacts_(props):[],['queued','running'].indexOf(job.status)>=0);
}

function wbComparisonCostColumn_(sheet) {
  const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const expected=['WB ID из файла себестоимости','Себестоимость из файла, ₽','Последняя дата в файле'];
  const matches=[];
  h.forEach((v,i)=>{if(expected.every((name,j)=>h[i+j]===name))matches.push(i+1);});
  if(matches.length!==1) throw new Error('Не удалось однозначно найти справочник себестоимости. Данные не перезаписаны.');
  return matches[0];
}

function wbEnsureComparisonLayout_(sheet) {
  let costColumn=wbComparisonCostColumn_(sheet);
  const height=Math.max(1,sheet.getLastRow());
  if(costColumn<=24 && costColumn+2>=23) {
    let destination=27;
    while(true) {
      if(sheet.getMaxColumns()<destination+2) sheet.insertColumnsAfter(sheet.getMaxColumns(),destination+2-sheet.getMaxColumns());
      const range=sheet.getRange(1,destination,height,3),values=range.getValues(),formulas=range.getFormulas();
      if(values.every((row,r)=>row.every((v,c)=>v==='' && !formulas[r][c]))) break;
      destination+=3;
    }
    sheet.getRange(1,costColumn,height,3).moveTo(sheet.getRange(1,destination,height,3));
    sheet.getBandings().forEach(b=>{const r=b.getRange();if(r.getColumn()===costColumn && r.getLastColumn()===costColumn+2)b.remove();});
    costColumn=destination;
  }
  if(sheet.getRange('W1').getValue()!=='ABC по конверсиям' || sheet.getRange('X1').getValue()!=='Отклонение от среднего, %') {
    // The user explicitly replaced the old W/X controls; the cost block has
    // already been moved intact, including formulas, before these are cleared.
    sheet.getRange(2,23,Math.max(1,sheet.getMaxRows()-1),2).clearDataValidations();
    sheet.getRange(2,23,Math.max(1,height-1),2).clearContent();
    sheet.getRange(1,23,1,2).setValues([['ABC по конверсиям','Отклонение от среднего, %']]);
  }
  sheet.getRange(2,24,Math.max(1,sheet.getMaxRows()-1),1).setNumberFormat('+0.0%;-0.0%;0.0%');
  sheet.setColumnWidth(23,170);sheet.setColumnWidth(24,230);
  const filter=sheet.getFilter();
  if(filter) {
    const old=filter.getRange();
    if(old.getRow()===1 && old.getColumn()===1 && old.getLastColumn()<24) {
      const criteria=[];
      for(let c=1;c<=old.getLastColumn();c++){const criterion=filter.getColumnFilterCriteria(c);if(criterion)criteria.push([c,criterion]);}
      const end=Math.max(2,height,old.getLastRow());filter.remove();
      const next=sheet.getRange(1,1,end,24).createFilter();
      criteria.forEach(([c,criterion])=>next.setColumnFilterCriteria(c,criterion));
    }
  }
  return costColumn;
}

function wbComparisonFacts_(props) {
  const facts=[];
  for(let i=0;i<Number(props.getProperty('WB_COMPARISON_ROWS_COUNT')||0);i++) facts.push.apply(facts,JSON.parse(props.getProperty('WB_COMPARISON_ROWS_'+i)||'[]'));
  return facts;
}

function wbConversionRates_(fact) {
  const ratio=(value,n,d)=>typeof value==='number' && Number.isFinite(value)?value:d>0?n/d:null;
  return [ratio(fact.ctr,fact.clicks,fact.impressions),ratio(fact.cartRate,fact.carts,fact.clicks),
    ratio(fact.orderRate,fact.orders,fact.carts),ratio(fact.buyoutRate,fact.buyouts,fact.orders)];
}

function wbConversionAbc_(facts) {
  const rated=facts.map(f=>({id:f.id,rates:wbConversionRates_(f),eligible:f.impressions>0 && f.clicks>0 && f.carts>0 && f.orders>0}));
  rated.forEach(r=>{r.eligible=r.eligible && r.rates.every(v=>typeof v==='number' && Number.isFinite(v) && v>=0 && v<=1);});
  const eligible=rated.filter(r=>r.eligible),means=[0,0,0,0];
  if(eligible.length<2) return new Map(rated.map(r=>[r.id,{category:'Нет данных',difference:null,score:null,rates:r.rates}]));
  eligible.forEach(r=>r.rates.forEach((v,i)=>means[i]+=v/eligible.length));
  eligible.forEach(r=>{r.score=r.rates.reduce((sum,v,i)=>sum+(means[i]>0?v/means[i]:1),0)/4;});
  const scores=eligible.map(r=>r.score).sort((a,b)=>a-b);
  const quantile=p=>{if(!scores.length)return null;const i=(scores.length-1)*p,lo=Math.floor(i),hi=Math.ceil(i);return scores[lo]+(scores[hi]-scores[lo])*(i-lo);};
  const low=quantile(1/3),high=quantile(2/3),average=scores.length?scores.reduce((a,b)=>a+b,0)/scores.length:0;
  return new Map(rated.map(r=>{
    if(!r.eligible) return [r.id,{category:'Нет данных',difference:null,score:null,rates:r.rates}];
    let category='B';
    if(Math.abs(scores[scores.length-1]-scores[0])>1e-12) {
      if(Math.abs(high-low)<1e-12) category=r.score>high+1e-12?'A':r.score<low-1e-12?'C':'B';
      else category=r.score<=low?'C':r.score>=high?'A':'B';
    }
    return [r.id,{category:category,difference:average>0?r.score/average-1:0,score:r.score,rates:r.rates}];
  }));
}

function wbWriteComparisonAbc_(sheet,ids,facts,loading) {
  if(sheet.getRange('W1').getValue()!=='ABC по конверсиям') wbEnsureComparisonLayout_(sheet);
  const n=Math.max(1,sheet.getLastRow()-1),rows=sheet.getRange(2,1,n,24).getValues(),formulas=sheet.getRange(2,23,n,2).getFormulas();
  const selected=new Set(ids),abc=wbConversionAbc_(facts),out=[],backgrounds=[];
  rows.forEach((r,i)=>{
    const id=String(r[0]).trim(),rated=abc.get(id),chosen=selected.has(id);
    const label=chosen?(loading?'Загрузка…':rated?rated.category:'Нет данных'):'';
    const difference=chosen && !loading && rated?rated.difference:null;
    out.push([formulas[i][0]||label,formulas[i][1]||(difference===null?'':difference)]);
    const color=label==='A'?'#c6e1cf':label==='B'?'#eee1b7':label==='C'?'#edc9c7':i%2?'#d2e0ec':'#e0eaf3';backgrounds.push([color,color]);
  });
  sheet.getRange(2,23,n,2).setValues(out).setBackgrounds(backgrounds);
}

function wbOpenComparison() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(WB_COMPARISON_HTML).setTitle('Сравнение товаров WB'));
}

function wbComparisonInit() {
  const sheet = wbComparisonSheet_();
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,22).getValues() : [];
  const seen = new Set(), products = [];
  rows.forEach(r => {
    const id = String(r[0]).trim();
    if (/^\d{1,16}$/.test(id) && !seen.has(id)) {
      seen.add(id);
      products.push({id:id,name:String(r[1] || ''),type:String(r[2] || ''),selected:r[21] === true});
    }
  });
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(),'Europe/Moscow','yyyy-MM-dd');
  const yesterday = new Date(today+'T12:00:00Z'); yesterday.setUTCDate(yesterday.getUTCDate()-1);
  return {products:products,historyStart:props.getProperty('WB_HISTORY_START') || '',
    yesterday:yesterday.toISOString().slice(0,10),job:wbComparisonStatus()};
}

function wbComparisonDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const date = new Date(s+'T12:00:00Z');
  return !isNaN(date.getTime()) && date.toISOString().slice(0,10) === s;
}

function wbQueueComparison(config) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const old = JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
    if (old && ['queued','running'].indexOf(old.status) >= 0) throw new Error('Сравнение уже выполняется. Дождитесь результата или отмените его.');
    if (!config || !Array.isArray(config.ids) || config.ids.length < 2 || config.ids.length > 300) throw new Error('Выберите от 2 до 300 артикулов.');
    const ids = Array.from(new Set(config.ids.map(String)));
    const sheet = wbComparisonSheet_();
    wbEnsureComparisonLayout_(sheet);
    const rows = sheet.getRange(2,1,Math.max(1,sheet.getLastRow()-1),22).getValues();
    const known = new Set(rows.map(r => String(r[0]).trim()));
    if (ids.length !== config.ids.length || ids.some(id => !/^\d{1,16}$/.test(id) || !known.has(id))) throw new Error('В списке есть неизвестные или повторные артикулы.');
    if (config.metric!=='conversionScore' && !config.includeAds && rows.some(r=>ids.indexOf(String(r[0]).trim())>=0 && typeof r[13]==='number' && r[13]!==0)) throw new Error('В N есть расходы на рекламу. Подтвердите, что они относятся к выбранному периоду, либо вручную установите 0.');
    const start = config.mode === 'all' ? config.historyStart : config.start;
    const end = config.end;
    if (!wbComparisonDate_(start) || !wbComparisonDate_(end) || start > end) throw new Error('Укажите корректные даты начала и конца. Для всей истории нужна дата начала работы магазина.');
    const today = Utilities.formatDate(new Date(),'Europe/Moscow','yyyy-MM-dd');
    if (end > today || start < '2000-01-01') throw new Error('Допустимы даты с 2000 года по сегодняшний день.');
    const metrics = ['conversionScore','orders10000','buyouts10000','revenue10000','profitUnit','margin'];
    if (metrics.indexOf(config.metric) < 0) throw new Error('Неизвестный показатель сравнения.');
    const job = {id:Utilities.getUuid(),status:'queued',ids:ids,start:start,end:end,metric:config.metric,
      includeAds:config.includeAds === true,mode:config.mode==='all'?'all':'dates',created:new Date().toISOString(),message:'Команда принята. Сервер начинает загрузку после получения сигнала.'};
    if (wbComparisonDate_(config.historyStart)) props.setProperty('WB_HISTORY_START',config.historyStart);
    // Only the existing comparison checkboxes are synchronized; formulas are preserved.
    const formulas = sheet.getRange(2,22,rows.length,1).getFormulas();
    const toggles=rows.map((r,i)=>[formulas[i][0] || (/^\d+$/.test(String(r[0]).trim())?ids.indexOf(String(r[0]).trim())>=0:r[21])]);
    sheet.getRange(2,22,rows.length,1).setValues(toggles);
    wbWriteComparisonAbc_(sheet,ids,[],true);
    props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));
    wbComparisonWake_(job.id);
    return job;
  } finally { lock.releaseLock(); }
}

function wbCancelComparison() {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props=PropertiesService.getScriptProperties(), job=JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
    if (job && ['queued','running'].indexOf(job.status)>=0) {
      job.status='cancelled'; job.message='Сравнение отменено.';
      wbWriteComparisonAbc_(wbComparisonSheet_(),job.ids,[],false);
      props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));
      wbComparisonClearWake_();
    }
    return job;
  } finally {lock.releaseLock();}
}

function wbComparisonWake_(id) {
  try {CacheService.getScriptCache().put('wb-comparison-pending',id,21600);} catch (_) { /* Property is still the authoritative queue. */ }
}
function wbComparisonClearWake_() {
  try {CacheService.getScriptCache().remove('wb-comparison-pending');} catch (_) {}
}

function wbComparisonTask_(request, reply, props) {
  if (request.schema !== 1) return reply({ok:false,error:'Invalid schema'});
  // Hold the control connection, never the script lock, while idle. This avoids
  // frequent properties reads and lets a new command wake the worker promptly.
  if (request.action==='task' && Number.isInteger(request.waitSeconds) && request.waitSeconds>0) {
    const peek=JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
    if (!peek || ['queued','running'].indexOf(peek.status)<0) {
      const deadline=Date.now()+Math.min(request.waitSeconds,20)*1000;
      while(Date.now()<deadline) {
        try {if(CacheService.getScriptCache().get('wb-comparison-pending')) break;} catch (_) {break;}
        Utilities.sleep(500);
      }
    }
  }
  const lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return reply({ok:false,error:'Busy'});
  try {
    const job = JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
    if (!job || ['queued','running'].indexOf(job.status)<0) {wbComparisonClearWake_();return reply({ok:true,job:null});}
    if (request.action==='jobDefer') {
      if(request.jobId!==job.id) return reply({ok:false,error:'Job mismatch'});
      job.status='queued';job.attempts=Math.max(0,Number(job.attempts||0)-1);
      job.message='Команда принята. Ожидает завершения текущей загрузки WB.';
      props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));wbComparisonWake_(job.id);
      return reply({ok:true});
    }
    if (request.action === 'jobError') {
      if (request.jobId !== job.id) return reply({ok:false,error:'Job mismatch'});
      wbWriteComparisonAbc_(wbComparisonSheet_(),job.ids,[],false);
      job.status='error'; job.message='Сервер не смог получить полный отчёт. '+String(request.message || 'Проверьте журнал сервера.').slice(0,250);
      props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));
      wbComparisonClearWake_();
      return reply({ok:true});
    }
    job.attempts=Number(job.attempts || 0)+1;
    if(job.attempts>3){wbWriteComparisonAbc_(wbComparisonSheet_(),job.ids,[],false);job.status='error';job.message='Сервер трижды прервал сравнение. Проверьте доступность периода и журнал сервера.';props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));return reply({ok:true,job:null});}
    job.status='running'; job.message='Загружается выбранный период. Большой или недоступный WB период может завершиться ошибкой.';
    props.setProperty('WB_COMPARISON_JOB',JSON.stringify(job));
    wbComparisonClearWake_();
    return reply({ok:true,job:job});
  } finally {lock.releaseLock();}
}

function wbSaveComparisonRows_(rows, props) {
  const old = Number(props.getProperty('WB_COMPARISON_ROWS_COUNT') || 0);
  const chunks=[]; let batch=[];
  rows.forEach(row => {
    if (JSON.stringify(batch.concat([row])).length>6000 && batch.length) {chunks.push(batch);batch=[];}
    batch.push(row);
  });
  if(batch.length) chunks.push(batch);
  chunks.forEach((chunk,i)=>props.setProperty('WB_COMPARISON_ROWS_'+i,JSON.stringify(chunk)));
  for(let i=chunks.length;i<old;i++) props.deleteProperty('WB_COMPARISON_ROWS_'+i);
  props.setProperty('WB_COMPARISON_ROWS_COUNT',String(chunks.length));
}

function wbComparisonStatus() {
  const props=PropertiesService.getScriptProperties(),job=JSON.parse(props.getProperty('WB_COMPARISON_JOB') || 'null');
  if(!job || job.status!=='completed') return job;
  const facts=wbComparisonFacts_(props),abc=wbConversionAbc_(facts);
  const sheet=wbComparisonSheet_(),rows=sheet.getRange(2,1,Math.max(1,sheet.getLastRow()-1),22).getValues();
  const byId=new Map(rows.map(r=>[String(r[0]).trim(),r]));
  const results=facts.map(f=>{
    const r=byId.get(f.id)||[],cost=typeof r[12]==='number' && Number.isFinite(r[12]) && r[12]>=0?r[12]:null;
    const ads=typeof r[13]==='number'?r[13]:0;
    const revenue=typeof f.revenue==='number'?f.revenue:null;
    const profit=cost!==null && revenue!==null?revenue-f.buyouts*cost-ads:null;
    const rating=abc.get(f.id),rates=rating.rates;
    return {id:f.id,name:String(r[1]||''),impressions:f.impressions,ctr:rates[0],orders:f.orders,buyouts:f.buyouts,
      cartRate:rates[1],orderRate:rates[2],buyoutRate:rates[3],abc:rating.category,difference:rating.difference,
      conversionScore:rating.score===null?null:100*rating.score,
      grade:rating.category==='A'?'green':rating.category==='B'?'yellow':rating.category==='C'?'red':'none',
      revenue:revenue,profit:profit,cost:cost,orders10000:f.impressions>0?10000*f.orders/f.impressions:null,
      buyouts10000:f.impressions>0?10000*f.buyouts/f.impressions:null,
      revenue10000:f.impressions>0 && revenue!==null?10000*revenue/f.impressions:null,
      profitUnit:f.buyouts>0 && profit!==null?profit/f.buyouts:null,
      margin:revenue>0 && profit!==null?profit/revenue:null};
  });
  results.sort((a,b)=>(b.conversionScore===null?-Infinity:b.conversionScore)-(a.conversionScore===null?-Infinity:a.conversionScore));
  job.metric='conversionScore';job.results=results;job.missing=job.ids.filter(id=>!facts.some(r=>r.id===id));
  return job;
}

const WB_COMPARISON_HTML = "<!doctype html><html lang=\"ru\"><head><meta charset=\"utf-8\"><style>\n[hidden]{display:none!important}body{font:13px Arial,sans-serif;background:#e5edf4;color:#21394e;margin:12px}h3{margin:14px 0 8px}input,select,button{box-sizing:border-box;font:inherit;border:1px solid #91a8bb;border-radius:6px;padding:8px;background:#d8e5ef;color:#17364e;width:100%;margin:4px 0}button{cursor:pointer;background:#285b7c;color:#fff;font-weight:bold}button:disabled{opacity:.5;cursor:wait}.muted{font-size:11px;line-height:1.5;color:#486074}.products{max-height:200px;overflow:auto;border:1px solid #91a8bb;border-radius:6px}.product{display:block;padding:7px;border-bottom:1px solid #bbccd9}.product:nth-child(even){background:#d5e2ed}input[type=checkbox]{width:auto;margin-right:6px}.row{display:flex;gap:6px}.status{padding:9px;background:#cfdeeb;white-space:pre-wrap;border-radius:6px;margin-top:10px}.card{padding:10px;margin:8px 0;border-radius:6px}.green{background:#c6e1cf}.yellow{background:#eee1b7}.red{background:#edc9c7}.none{background:#cbd6df}.error{background:#edc9c7}.secondary{background:#59768d}label{display:block;margin-top:7px}.value{font-size:18px;font-weight:bold;margin:6px 0}\n</style></head><body>\n<p class=\"muted\">Цвета столбцов: песочный — ручной ввод (M, N); голубой — автоматические данные и расчёты; мятный — выбор товаров.</p><h3>Артикулы для сравнения</h3><input id=\"search\" placeholder=\"Поиск по артикулу, товару, типу\"><div class=\"row\"><button class=\"secondary\" id=\"selectVisible\">Выбрать найденные</button><button class=\"secondary\" id=\"clear\">Снять всё</button></div><div id=\"count\"></div><div class=\"products\" id=\"products\"></div>\n<h3>Период</h3><select id=\"mode\"><option value=\"dates\">Любые даты</option><option value=\"all\">С начала работы магазина</option></select>\n<label id=\"historyLabel\">Дата начала работы магазина<input type=\"date\" id=\"historyStart\"></label>\n<label id=\"startLabel\">С даты<input type=\"date\" id=\"start\"></label><label>По дату включительно<input type=\"date\" id=\"end\"></label>\n<p class=\"muted\">Можно выбрать прошлый год. Сегодня — неполный день. Доступность истории зависит от WB: при отказе период не подменяется. Дата начала магазина задаётся вами — автоматически она пока не определена.</p>\n<h3>ABC по конверсиям</h3><select id=\"metric\"><option value=\"conversionScore\">Четыре конверсии — равный вес</option></select>\n<input type=\"checkbox\" id=\"includeAds\" hidden>\n<p class=\"muted\">CTR, переход → корзина, корзина → заказ, заказ → выкуп: каждую конверсию делим на среднюю среди выбранных сопоставимых товаров и усредняем с равным весом. A — выше, B — средние, C — ниже. W — категория, X — отклонение индекса от среднего. Финансы в ABC не входят. Это относительная оценка, не статистический A/B-тест.</p>\n<label>Показать оценки<select id=\"gradeFilter\"><option value=\"all\">Все</option><option value=\"green\">A — выше</option><option value=\"yellow\">B — средние</option><option value=\"red\">C — ниже</option><option value=\"none\">Нет данных</option></select></label><button id=\"compare\" disabled>Сравнить</button><button id=\"cancel\" class=\"secondary\" hidden>Отменить ожидание</button><div class=\"status\" id=\"status\">Загрузка…</div><div id=\"results\"></div>\n<script>\nconst $=id=>document.getElementById(id);let products=[],selected=new Set(),currentJob=null,poller=null;\nconst names={conversionScore:'Индекс конверсий (среднее 100)'};\nfunction text(el,value){el.textContent=value;return el;}\nfunction error(e){$('status').className='status error';$('status').textContent=e.message||String(e);$('compare').disabled=false;}\nfunction visible(){const q=$('search').value.trim().toLocaleLowerCase();return products.filter(p=>(p.id+' '+p.name+' '+p.type).toLocaleLowerCase().includes(q));}\nfunction renderProducts(){const root=$('products');root.textContent='';visible().forEach(p=>{const label=document.createElement('label');label.className='product';const box=document.createElement('input');box.type='checkbox';box.checked=selected.has(p.id);box.onchange=()=>{box.checked?selected.add(p.id):selected.delete(p.id);$('count').textContent='Выбрано: '+selected.size+' / 300';};label.append(box,document.createTextNode(p.id+' · '+p.name));root.appendChild(label);});$('count').textContent='Выбрано: '+selected.size+' / 300';}\nfunction mode(){const all=$('mode').value==='all';$('historyLabel').hidden=!all;$('startLabel').hidden=all;}\nfunction fmt(v,percent=false){return v===null||v===undefined?'нет данных':(percent?v*100:v).toLocaleString('ru-RU',{maximumFractionDigits:2})+(percent?'%':'');}\nfunction showJob(job){currentJob=job;const active=job&&['queued','running'].includes(job.status);$('compare').disabled=!!active;$('cancel').hidden=!active;$('status').className='status'+(job&&job.status==='error'?' error':'');$('status').textContent=job?job.message:'Выберите товары и период.';if(poller){clearTimeout(poller);poller=null;}if(active)poller=setTimeout(refresh,1000);$('results').textContent='';if(!job||job.status!=='completed')return;\n text($('status'),job.start+' — '+job.end+'\\nПолучено '+job.results.length+' из '+job.ids.length+' товаров. '+(job.missing.length?'Нет полных сопоставимых данных WB: '+job.missing.join(', '):''));\n const legend=document.createElement('p');legend.className='muted';legend.textContent='ABC сравнивает только конверсии среди выбранных карточек. Для оценки нужны хотя бы две карточки с полной сопоставимой воронкой. Равные значения получают одинаковую категорию; серый — недостаточно данных.';$('results').appendChild(legend);\n job.results.filter(r=>$('gradeFilter').value==='all'||r.grade===$('gradeFilter').value).forEach(r=>{const card=document.createElement('div');card.className='card '+r.grade;const title=document.createElement('strong');title.textContent=(r.abc==='Нет данных'?'Нет данных':'ABC '+r.abc)+': '+r.id+' · '+r.name;const value=document.createElement('div');value.className='value';value.textContent=names[job.metric]+': '+fmt(r[job.metric],job.metric==='margin');const details=document.createElement('div');details.textContent='От среднего: '+(r.difference===null?'нет данных':(r.difference>0?'+':'')+fmt(r.difference,true))+'\\nCTR: '+fmt(r.ctr,true)+'\\nПереход → корзина: '+fmt(r.cartRate,true)+'\\nКорзина → заказ: '+fmt(r.orderRate,true)+'\\nЗаказ → выкуп: '+fmt(r.buyoutRate,true)+'\\nПоказы: '+fmt(r.impressions)+' · Заказы: '+fmt(r.orders)+' · Выкупы: '+fmt(r.buyouts);details.style.whiteSpace='pre-wrap';card.append(title,value,details);$('results').appendChild(card);});}\nfunction refresh(){google.script.run.withSuccessHandler(showJob).withFailureHandler(error).wbComparisonStatus();}\n$('gradeFilter').onchange=()=>{if(currentJob)showJob(currentJob);};\n$('search').oninput=renderProducts;$('mode').onchange=mode;$('selectVisible').onclick=()=>{visible().forEach(p=>{if(selected.size<300)selected.add(p.id);});renderProducts();};$('clear').onclick=()=>{selected.clear();renderProducts();};\n$('compare').onclick=()=>{if(selected.size<2||selected.size>300){error({message:'Выберите от 2 до 300 товаров.'});return;}$('compare').disabled=true;$('status').className='status';$('status').textContent='Отправляем команду на сравнение…';$('results').textContent='';google.script.run.withSuccessHandler(showJob).withFailureHandler(error).wbQueueComparison({ids:Array.from(selected),mode:$('mode').value,start:$('start').value,end:$('end').value,historyStart:$('historyStart').value,metric:$('metric').value,includeAds:$('includeAds').checked});};\n$('cancel').onclick=()=>google.script.run.withSuccessHandler(showJob).withFailureHandler(error).wbCancelComparison();\ngoogle.script.run.withSuccessHandler(data=>{products=data.products;selected=new Set(products.filter(p=>p.selected).map(p=>p.id));$('historyStart').value=data.historyStart;$('start').value=data.yesterday;$('end').value=data.yesterday;if(data.job){$('mode').value=data.job.mode||'dates';$('start').value=data.job.start;$('end').value=data.job.end;$('metric').value='conversionScore';$('includeAds').checked=data.job.includeAds;}renderProducts();mode();showJob(data.job);}).withFailureHandler(error).wbComparisonInit();\n</script></body></html>\n";
const WB_RECEIVER_EXPORTS = Object.freeze({doPost,wbInstallComparison,wbComparisonOnOpen,wbOpenComparison,wbComparisonInit,wbQueueComparison,wbCancelComparison,wbComparisonStatus});
