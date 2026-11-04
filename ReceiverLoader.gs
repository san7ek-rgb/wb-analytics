/** Replace the existing CostReceiver.gs ONCE with this permanent loader.
 * Keep the separate Loader.gs and script properties. Deploy a new version of
 * the existing web app once, then run wbInstallComparison once.
 * Only public code is downloaded. Secrets, rows and cookies never go to GitHub.
 * SHA-256 checks integrity; repository writers are trusted code publishers.
 */
const WB_RECEIVER_LOADER = Object.freeze({
  base: 'https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/main/',
  manifest: 'receiver-release.json',
  lastGood: 'WB_RECEIVER_LAST_GOOD'
});

function doPost(e) {
  const reply = value => ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string' || e.postData.contents.length > 1000000) return reply({ok:false,error:'Invalid request'});
    const request = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('WB_COST_RECEIVER_TOKEN');
    if (!secret || secret.length < 32 || !request || request.token !== secret) return reply({ok:false,error:'Unauthorized'});
  } catch (_) { return reply({ok:false,error:'Invalid JSON'}); }
  try { return wbReceiverRpc('doPost',e); }
  catch (_) { return reply({ok:false,error:'Receiver code unavailable; retry later or inspect Apps Script execution log'}); }
}

function wbInstallComparison() {
  // Explicit references also allow Apps Script to infer the initial UI/trigger scopes.
  SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers();
  return wbReceiverRpc('wbInstallComparison');
}
function wbAuthorizeComparisonUi() {
  // A static call is necessary: getUi() alone does not infer container.ui.
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('Подключение панели сравнения…'));
  return wbInstallComparison();
}
function wbComparisonOnOpen(e) { return wbReceiverRpc('wbComparisonOnOpen',e); }
function wbOpenComparison() { return wbReceiverRpc('wbOpenComparison'); }
function wbComparisonInit() { return wbReceiverRpc('wbComparisonInit'); }
function wbQueueComparison(config) { return wbReceiverRpc('wbQueueComparison',config); }
function wbCancelComparison() { return wbReceiverRpc('wbCancelComparison'); }
function wbComparisonStatus() { return wbReceiverRpc('wbComparisonStatus'); }

function wbReceiverCacheGet_(key) {
  try { return CacheService.getScriptCache().get(key); } catch (_) { return null; }
}
function wbReceiverCachePut_(key,value,ttl) {
  try { CacheService.getScriptCache().put(key,value,ttl); } catch (_) { /* Fetch again if cache is unavailable or code is large. */ }
}
function wbReceiverFetch_(path) {
  const response=UrlFetchApp.fetch(WB_RECEIVER_LOADER.base+path,{
    muteHttpExceptions:true,followRedirects:false,validateHttpsCertificates:true
  });
  if(response.getResponseCode()!==200) throw new Error('Receiver download failed.');
  return response.getContentText();
}
function wbReceiverValidateManifest_(value) {
  if (!value || value.schema!==1 || !/^[a-f0-9]{64}$/.test(value.sha256) ||
      value.path!=='receiver-releases/'+value.sha256+'.gs') throw new Error('Invalid receiver release.');
  return value;
}
function wbReceiverManifest_() {
  const cached=wbReceiverCacheGet_('wb-receiver-manifest-v1');
  if(cached) return wbReceiverValidateManifest_(JSON.parse(cached));
  const text=wbReceiverFetch_(WB_RECEIVER_LOADER.manifest+'?v='+Date.now());
  const manifest=wbReceiverValidateManifest_(JSON.parse(text));
  wbReceiverCachePut_('wb-receiver-manifest-v1',JSON.stringify(manifest),60);
  return manifest;
}
function wbReceiverCompile_(manifest) {
  manifest=wbReceiverValidateManifest_(manifest);
  const key='wb-receiver-source-'+manifest.sha256;
  const source=wbReceiverCacheGet_(key) || wbReceiverFetch_(manifest.path);
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,source,Utilities.Charset.UTF_8)
    .map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
  if(digest!==manifest.sha256) throw new Error('Receiver checksum mismatch.');
  // Compile before invocation: download/integrity/syntax failures may fall back.
  // A handler that has started is NEVER retried against another version.
  const invoke=new Function('argument','action','releaseHash',source+
    '\nif (!Object.prototype.hasOwnProperty.call(WB_RECEIVER_EXPORTS,action)) throw new Error("Unsupported receiver action");'+
    '\nif(typeof wbReceiverUpgrade_ === "function") wbReceiverUpgrade_(releaseHash);'+
    '\nreturn WB_RECEIVER_EXPORTS[action](argument);');
  wbReceiverCachePut_(key,source,21600);
  return invoke;
}
function wbReceiverRpc(action,argument) {
  if(typeof action!=='string' || action.length>80) throw new Error('Invalid receiver action.');
  const props=PropertiesService.getScriptProperties();
  let manifest,invoke;
  try {
    manifest=wbReceiverManifest_();
    invoke=wbReceiverCompile_(manifest);
  } catch (_) {
    const previous=props.getProperty(WB_RECEIVER_LOADER.lastGood);
    if(!previous) throw new Error('Receiver release unavailable. Check GitHub access.');
    manifest=wbReceiverValidateManifest_(JSON.parse(previous));
    invoke=wbReceiverCompile_(manifest);
    console.log('Receiver: using previous verified release '+manifest.sha256.slice(0,12));
  }
  const result=invoke(argument,action,manifest.sha256);
  props.setProperty(WB_RECEIVER_LOADER.lastGood,JSON.stringify(manifest));
  return result;
}
