/** Comparison UI and queue. Included in the standalone receiver bundle. */
function wbReceiverUpgrade_(releaseHash) {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('WB_COMPARISON_STYLE_RELEASE')===releaseHash) return;
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000)) return;
  try {
    if(props.getProperty('WB_COMPARISON_STYLE_RELEASE')!==releaseHash) {
      wbColorComparisonColumns_();
      wbFillSourceUnitProfit_();
      wbEnsureHistoryImport_(props);
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

function wbComparisonOnOpen(e) {
  if(e && e.range) {
    if(e.range.getSheet().getSheetId()!==1537466773 || e.range.getRow()>5 || e.range.getLastRow()<5)return;
    const result=wbPhotosApply_({schema:1,firstColumn:e.range.getColumn(),lastColumn:e.range.getLastColumn()},v=>v,PropertiesService.getScriptProperties());
    if(!result.ok)console.log('WB photo update: '+JSON.stringify(result));
    return;
  }
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
    yesterday:yesterday.toISOString().slice(0,10),history:wbHistoryStatus(),job:wbComparisonStatus()};
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
      includeAds:config.includeAds === true,source:config.historyOnly?'history':'wb',mode:config.mode==='all'?'all':'dates',created:new Date().toISOString(),message:'Команда принята. Сервер начинает загрузку после получения сигнала.'};
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
    const history=request.historyWorker===1?JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB') || 'null'):null;
    const historyReady=history && ['queued','running'].indexOf(history.status)>=0 && (!history.notBefore || history.notBefore<=Date.now());
    if ((!peek || ['queued','running'].indexOf(peek.status)<0) && !historyReady) {
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
    if (!job || ['queued','running'].indexOf(job.status)<0) {
      if(request.action==='task' && request.historyWorker===1) {
        const history=wbClaimHistoryImport_(props);if(history)return reply({ok:true,job:history});
      }
      wbComparisonClearWake_();return reply({ok:true,job:null});
    }
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

function wbHistoryStatus() {
  const props=PropertiesService.getScriptProperties();
  const summary=JSON.parse(props.getProperty('WB_HISTORY_STATUS') || 'null');
  const job=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB') || 'null');
  return summary?Object.assign(summary,{importJob:job}):{dailyRows:0,days:0,articles:0,periodRows:0,importJob:job,unconfirmed:true};
}

function wbCreateHistoryImport_(props,start,end) {
  if(!wbComparisonDate_(start)||!wbComparisonDate_(end)||start>end)throw new Error('Укажите корректные даты истории.');
  const total=Math.round((Date.parse(end+'T12:00:00Z')-Date.parse(start+'T12:00:00Z'))/86400000)+1;
  const today=Utilities.formatDate(new Date(),'Europe/Moscow','yyyy-MM-dd');
  if(total>5000||end>=today)throw new Error('История загружается только за завершённые дни, до 5000 дней за задание.');
  const catalog=wbReceiveCatalog_({schema:1},v=>v,props);
  if(!catalog.ok||!catalog.ids.length)throw new Error('Не найдены артикулы для истории.');
  const job={id:Utilities.getUuid(),kind:'history',status:'queued',ids:catalog.ids,start:start,end:end,cursor:start,
    processedDays:0,totalDays:total,unavailableStreak:0,attempts:0,created:new Date().toISOString(),
    message:'Дневная история поставлена в очередь с '+start+' по '+end+'.'};
  props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));wbComparisonWake_(job.id);return job;
}

function wbEnsureHistoryImport_(props) {
  if(props.getProperty('WB_HISTORY_IMPORT_INITIALIZED'))return;
  if(!props.getProperty('WB_HISTORY_IMPORT_JOB')) {
    const today=Utilities.formatDate(new Date(),'Europe/Moscow','yyyy-MM-dd');
    const end=new Date(today+'T12:00:00Z');end.setUTCDate(end.getUTCDate()-1);
    // Initial date explicitly requested by the user; never inferred as a WB limit.
    wbCreateHistoryImport_(props,'2025-05-01',end.toISOString().slice(0,10));
  }
  props.setProperty('WB_HISTORY_IMPORT_INITIALIZED','1');
}

function wbStartHistoryImport(config) {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const props=PropertiesService.getScriptProperties(),old=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB')||'null');
    if(old && ['queued','running'].indexOf(old.status)>=0)throw new Error('Импорт уже идёт. Сначала приостановите его.');
    const job=wbCreateHistoryImport_(props,config.start,config.end);
    props.setProperty('WB_HISTORY_IMPORT_INITIALIZED','1');return job;
  } finally {lock.releaseLock();}
}

function wbPauseHistoryImport() {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const props=PropertiesService.getScriptProperties(),job=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB')||'null');
    if(job && ['queued','running'].indexOf(job.status)>=0){job.status='paused';job.message='Импорт приостановлен. Сохранённые дни остаются в базе.';props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));}
    return job;
  } finally {lock.releaseLock();}
}

function wbResumeHistoryImport() {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const props=PropertiesService.getScriptProperties(),job=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB')||'null');
    if(job && ['paused','error'].indexOf(job.status)>=0){job.status='queued';job.attempts=0;job.notBefore=0;job.message='Продолжение с '+job.cursor;props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));wbComparisonWake_(job.id);}
    return job;
  } finally {lock.releaseLock();}
}

function wbClaimHistoryImport_(props) {
  const job=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB')||'null');
  if(!job||['queued','running'].indexOf(job.status)<0||(job.notBefore&&job.notBefore>Date.now()))return null;
  job.attempts=Number(job.attempts||0)+1;
  if(job.attempts>3){job.status='error';job.message='Три прерывания импорта. Можно продолжить с '+job.cursor;props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));return null;}
  job.status='running';job.message='Загружается дневная история с '+job.cursor;
  props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));return job;
}

function wbHistoryImportAction_(request,reply,props) {
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return reply({ok:false,error:'Busy'});
  try {
    const job=JSON.parse(props.getProperty('WB_HISTORY_IMPORT_JOB')||'null');
    if(request.schema!==1||!job||request.jobId!==job.id)return reply({ok:false,error:'History job mismatch'});
    if(['queued','running'].indexOf(job.status)<0)return reply({ok:true,continue:false});
    if(request.action==='historyError') {
      job.status='error';job.message='Импорт остановлен на '+job.cursor+'. '+String(request.message||'Проверьте журнал сервера.').slice(0,180);
    } else if(request.action==='historyDefer') {
      job.status='queued';job.attempts=Math.max(0,Number(job.attempts||0)-1);job.message='Ожидается свободный профиль WB.';
    } else {
      const last=new Date(job.end+'T12:00:00Z');last.setUTCDate(last.getUTCDate()+1);const after=last.toISOString().slice(0,10);
      if(!wbComparisonDate_(request.cursor)||request.cursor<job.cursor||request.cursor>after)return reply({ok:false,error:'Invalid history cursor'});
      job.cursor=request.cursor;job.processedDays=Math.round((Date.parse(job.cursor+'T12:00:00Z')-Date.parse(job.start+'T12:00:00Z'))/86400000);
      job.attempts=0;job.unavailableStreak=Number.isInteger(request.unavailableStreak)?Math.min(100,Math.max(0,request.unavailableStreak)):0;
      job.notBefore=Date.now()+Math.min(86400,Math.max(0,Number(request.delaySeconds)||0))*1000;
      job.lastRetryDelay=request.retryDelay?Math.min(86400,Math.max(0,Number(request.retryDelay)||0)):0;
      job.status=job.cursor>job.end?'completed':request.yield?'queued':'running';
      job.message=job.status==='completed'?'Обход истории завершён. Недоступные и неполные дни отмечены отдельно.':
        'Обработано '+job.processedDays+' из '+job.totalDays+' дней. Следующая дата: '+job.cursor+(request.delaySeconds?' (пауза перед продолжением)':'');
    }
    props.setProperty('WB_HISTORY_IMPORT_JOB',JSON.stringify(job));
    if(request.summary)wbReceiveHistoryStatus_({schema:1,summary:request.summary},v=>v,props);
    if(job.status==='queued')wbComparisonWake_(job.id);
    const compare=JSON.parse(props.getProperty('WB_COMPARISON_JOB')||'null');
    return reply({ok:true,continue:['queued','running'].indexOf(job.status)>=0,comparisonWaiting:compare&&['queued','running'].indexOf(compare.status)>=0});
  } finally {lock.releaseLock();}
}

// Fill missing unit margins only; preserve the owner's entered margins/formulas.
function wbFillSourceUnitProfit_() {
  const forecast=wbComparisonSheet_(),book=forecast.getParent();
  const sheet=book.getSheetByName('Рабочая табличка');
  if(!sheet || sheet.getLastRow()<2)return;
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(v=>v.trim());
  const id=headers.indexOf('Артикул WB')+1,price=headers.indexOf('Цена, ₽')+1;
  const profit=headers.findIndex(v=>/^Прибыль с выкупа/.test(v))+1;
  if(!id||!price||!profit)return;
  const helper=wbComparisonCostColumn_(forecast);
  function letter(n){let value='';while(n){n--;value=String.fromCharCode(65+n%26)+value;n=Math.floor(n/26);}return value;}
  const source="'"+forecast.getName().replace(/'/g,"''")+"'!";
  const ids=source+'$'+letter(helper)+'$2:$'+letter(helper);
  const costs=source+'$'+letter(helper+1)+'$2:$'+letter(helper+1);
  const height=sheet.getLastRow()-1,range=sheet.getRange(2,profit,height,1);
  const values=range.getValues(),formulas=range.getFormulas(),articles=sheet.getRange(2,id,height,1).getValues();
  for(let i=0;i<height;i++) {
    if(values[i][0]!==''||formulas[i][0]||!/^\d+$/.test(String(articles[i][0]).trim()))continue;
    const row=i+2,priceRef='$'+letter(price)+row,idRef='$'+letter(id)+row;
    sheet.getRange(row,profit).setFormula('=LET(unitCost,IFERROR(INDEX('+costs+',MATCH(TO_TEXT('+idRef+'),ARRAYFORMULA(TO_TEXT('+ids+')),0)),""),IF(AND(ISNUMBER('+priceRef+'),ISNUMBER(unitCost)),'+priceRef+'-unitCost,""))');
  }
  sheet.getRange(1,profit).setNote('Прибыль с одной единицы. Пустые ячейки рассчитываются как цена минус себестоимость по артикулу из справочника. Без комиссий, логистики, налогов и рекламы. Ранее введённые значения сохранены. Без цены или себестоимости расчёт пустой.');
}
