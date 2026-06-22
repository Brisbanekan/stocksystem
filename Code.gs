/************************************************************
 * 年節禮盒庫存系統 — Google Apps Script 後端 + LINE Bot
 * 架構：Google Sheets(資料庫) + Apps Script(API/Webhook)
 *
 * 工作表(分頁)：
 *  1) Items  欄位：ID | 分類 | 禮盒名稱 | 規格 | 年分 | 節日 | 單位 | 即時庫存
 *     分類：A=公益、B=聖保羅、C=鬥茶王
 *  2) Log    欄位：時間 | 類型 | 禮盒名稱 | 年分 | 節日 | 數量 | 結存 | 人員/備註 | 來源
 *     類型：進料 / 出料 / 刪除
 *
 * 部署：擴充功能 > Apps Script > 貼上 > 部署為「網頁應用程式」
 ************************************************************/

// ====== 設定 ======
var SHEET_ID   = '';
var ITEM_SHEET = 'Items';
var LOG_SHEET  = 'Log';
var LINE_TOKEN = 'YOUR_CHANNEL_ACCESS_TOKEN';
var CATS = {A:'公益', B:'聖保羅', C:'鬥茶王'};
var FESTS = ['過年','端午','中秋'];

// ====== 共用 ======
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function itemSheet_() { return ss_().getSheetByName(ITEM_SHEET); }
function logSheet_()  { return ss_().getSheetByName(LOG_SHEET); }
function label_(it) { return it.name + '（' + it.year + it.fest + '）'; }

// ====== 資料存取 ======
// Items 欄位順序：0 ID,1 分類,2 名稱,3 規格,4 年分,5 節日,6 單位,7 即時庫存
function rowToItem_(r, rowIndex) {
  return {id:r[0], cat:String(r[1]), name:String(r[2]), spec:String(r[3]),
          year:r[4], fest:String(r[5]), unit:String(r[6]), qty:Number(r[7])||0, rowIndex:rowIndex};
}
function getItems_() {
  var rows = itemSheet_().getDataRange().getValues(); rows.shift();
  return rows.filter(function(r){return r[0]!=='';}).map(function(r){return rowToItem_(r);});
}
// 出貨明細：取 Log 中所有「出料」紀錄。Log 欄位：0時間,1類型,2禮盒,3年分,4節日,5數量,6結存,7人員/備註(=出貨對象),8來源
function getShipments_() {
  var rows = logSheet_().getDataRange().getValues(); rows.shift();
  return rows.filter(function(r){return r[1]==='出料';}).reverse().map(function(r){
    return {time:r[0], name:String(r[2]), year:r[3], fest:String(r[4]), qty:Number(r[5])||0, target:String(r[7]||'')};
  });
}
// 順序子序列比對：字按順序出現即可，如「公益禮盒」對到「公益月餅禮盒」
function isSubseq_(a, b){ var i=0; for(var j=0;j<b.length;j++){ if(b.charAt(j)===a.charAt(i)) i++; } return i===a.length; }
// 以禮盒名稱模糊比對，回傳依分數排序的候選 [{item,s}]：完全相符100 > 連續子字串80 > 子序列60
function matchItems_(key) {
  key = String(key).trim().toLowerCase().replace(/\s+/g,'');
  var out=[]; if(!key) return out;
  var rows = itemSheet_().getDataRange().getValues();
  for (var i=1;i<rows.length;i++){
    if(rows[i][0]==='') continue;
    var name=String(rows[i][2]).toLowerCase().replace(/\s+/g,'');
    var s=0;
    if(name===key) s=100; else if(name.indexOf(key)>=0) s=80; else if(isSubseq_(key,name)) s=60;
    if(s>0) out.push({item:rowToItem_(rows[i], i+1), s:s});
  }
  out.sort(function(a,b){return b.s-a.s;});
  return out;
}
// 回傳 {status:'ok'|'none'|'ambiguous', item, candidates}
function resolveItem_(key) {
  var m=matchItems_(key);
  if(!m.length) return {status:'none'};
  if(m.length===1 || m[0].s>m[1].s) return {status:'ok', item:m[0].item};
  return {status:'ambiguous', candidates:m.filter(function(x){return x.s===m[0].s;}).map(function(x){return x.item;})};
}
function ambigMsg_(cands) {
  return '🔍 找到多個符合的禮盒，請輸入更完整的名稱：\n'+cands.map(function(i){return '· '+i.name+'（'+i.year+i.fest+'）';}).join('\n');
}

// ====== 核心：庫存異動（進料 / 出料）======
function applyTransaction_(type, key, qty, who, source) {
  qty = Number(qty);
  if (!qty || qty<=0) return {ok:false, msg:'數量需為正整數'};
  if (['進料','出料'].indexOf(type)<0) return {ok:false, msg:'類型錯誤'};

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var res = resolveItem_(key);
    if (res.status==='none') return {ok:false, msg:'找不到禮盒「'+key+'」'};
    if (res.status==='ambiguous') return {ok:false, ambiguous:true, msg:ambigMsg_(res.candidates)};
    var it = res.item;
    var newQty = (type==='進料') ? it.qty+qty : it.qty-qty;
    if (newQty < 0) return {ok:false, msg:it.name+' 庫存不足，目前僅 '+it.qty+' '+it.unit};

    itemSheet_().getRange(it.rowIndex, 8).setValue(newQty); // 第8欄=即時庫存
    logSheet_().appendRow([new Date(), type, it.name, it.year, it.fest, qty, newQty, who||'-', source||'web']);
    return {ok:true, item:it.name, label:label_(it), unit:it.unit, qty:newQty, type:type};
  } finally {
    lock.releaseLock();
  }
}

// ====== 刪除禮盒（含異動紀錄留存）======
function deleteItem_(id, reason, who) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rows = itemSheet_().getDataRange().getValues();
    for (var i=1;i<rows.length;i++){
      if (String(rows[i][0])===String(id)) {
        var it = rowToItem_(rows[i], i+1);
        logSheet_().appendRow([new Date(), '刪除', it.name, it.year, it.fest, it.qty, '—',
                               (who?who+' / ':'')+'原因：'+(reason||'未填原因'), 'web']);
        itemSheet_().deleteRow(it.rowIndex);
        return {ok:true, item:it.name, leftover:it.qty, unit:it.unit};
      }
    }
    return {ok:false, msg:'找不到該禮盒'};
  } finally {
    lock.releaseLock();
  }
}

// ====== 新增禮盒 ======
function addItem_(cat, name, spec, year, fest, unit, init) {
  if (!CATS[cat]) return {ok:false, msg:'分類需為 A/B/C'};
  if (FESTS.indexOf(fest)<0) return {ok:false, msg:'節日需為 過年/端午/中秋'};
  if (!name) return {ok:false, msg:'禮盒名稱為必填'};
  var id = 'G' + new Date().getTime();
  itemSheet_().appendRow([id, cat, name, spec, Number(year), fest, unit||'盒', Number(init)||0]);
  if (Number(init)>0) logSheet_().appendRow([new Date(),'進料',name,Number(year),fest,Number(init),Number(init),'期初','web']);
  return {ok:true, id:id};
}

// ====== Web App：GET ======
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var out;
  if (action==='items') out = {ok:true, items:getItems_()};
  else if (action==='log') {
    var rows = logSheet_().getDataRange().getValues(); rows.shift();
    out = {ok:true, log: rows.slice(-50).reverse()};
  }
  else if (action==='shipments') out = {ok:true, shipments:getShipments_()};
  else out = {ok:true, msg:'gift-box inventory API'};
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ====== Web App：POST（前端 API + LINE Webhook）======
function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  if (body.events) { // LINE Webhook
    body.events.forEach(function(ev){
      if (ev.type==='message' && ev.message.type==='text') {
        replyLine_(ev.replyToken, handleCommand_(ev.message.text));
      }
    });
    return ContentService.createTextOutput('OK');
  }

  var r;
  if (body.action==='tx')       r = applyTransaction_(body.type, body.key, body.qty, body.who, 'web');
  else if (body.action==='add') r = addItem_(body.cat, body.name, body.spec, body.year, body.fest, body.unit, body.init);
  else if (body.action==='del') r = deleteItem_(body.id, body.reason, body.who);
  else r = {ok:false, msg:'unknown action'};
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// ====== LINE 指令解析（與網頁原型同邏輯）======
function handleCommand_(text) {
  var parts = String(text).trim().split(/\s+/);
  var cmd = parts[0];

  if (/^(說明|help|\?|？)$/i.test(cmd)) {
    return '📖 指令說明\n進料 <禮盒名> <數量> [備註]\n出料 <禮盒名> <數量> [出貨對象]\n庫存 <禮盒名>　查單一禮盒\n庫存 <節日>　依過年/端午/中秋查\n庫存　　　　列出全部\n規格 <禮盒名>　查規格';
  }
  if (cmd==='庫存') {
    var arg = parts.slice(1).join(' ');
    var items = getItems_();
    if (arg && FESTS.indexOf(arg)>=0) {
      var list = items.filter(function(i){return i.fest===arg;});
      if(!list.length) return '查無「'+arg+'」禮盒';
      return '🎁 '+arg+'禮盒庫存：\n'+list.map(function(i){return '· ['+i.cat+'] '+i.name+' '+i.qty+' '+i.unit+'（'+i.year+'）';}).join('\n');
    }
    if (arg) {
      var r1=resolveItem_(arg);
      if(r1.status==='none') return '❓ 找不到禮盒「'+arg+'」';
      if(r1.status==='ambiguous') return ambigMsg_(r1.candidates);
      var it=r1.item;
      return '🎁 '+it.name+'\n分類：'+it.cat+' '+CATS[it.cat]+'\n年分/節日：'+it.year+' '+it.fest+'\n規格：'+it.spec+'\n即時庫存：'+it.qty+' '+it.unit;
    }
    if(!items.length) return '目前無任何禮盒';
    return '🎁 全部禮盒庫存：\n'+items.map(function(i){return '· ['+i.cat+'] '+i.name+' '+i.qty+i.unit+'（'+i.year+i.fest+'）';}).join('\n');
  }
  if (cmd==='規格') {
    var r2=resolveItem_(parts.slice(1).join(' '));
    if(r2.status==='none') return '❓ 找不到禮盒';
    if(r2.status==='ambiguous') return ambigMsg_(r2.candidates);
    var fs=r2.item;
    return '📐 '+fs.name+'\n規格：'+fs.spec+'\n分類：'+fs.cat+' '+CATS[fs.cat]+'\n單位：'+fs.unit;
  }
  if (['進料','出料'].indexOf(cmd)>=0) {
    if (parts.length<3) return '格式：'+cmd+' <禮盒名> <數量> [備註]';
    var qtyIdx=-1;
    for(var i=1;i<parts.length;i++){ if(/^\d+$/.test(parts[i])){qtyIdx=i;break;} }
    if(qtyIdx<0) return '❓ 請輸入數量，例如：'+cmd+' 公益月餅禮盒 100';
    var name=parts.slice(1,qtyIdx).join(' ');
    var qty=parseInt(parts[qtyIdx],10);
    var extra=parts.slice(qtyIdx+1).join(' ');
    var r=applyTransaction_(cmd, name, qty, extra, 'line');
    if(!r.ok) return r.ambiguous ? r.msg : '❌ '+r.msg;
    var sign=cmd==='進料'?'+':'−';
    var xlabel=cmd==='出料'?'出貨對象':'備註';
    return '✅ '+cmd+'登記成功\n'+r.label+' '+sign+qty+' '+r.unit+'\n目前庫存：'+r.qty+' '+r.unit+(extra?('\n'+xlabel+'：'+extra):'');
  }
  return '🤔 看不懂指令，輸入「說明」查看用法';
}

// ====== 回覆 LINE ======
function replyLine_(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method:'post', contentType:'application/json',
    headers:{Authorization:'Bearer '+LINE_TOKEN},
    payload:JSON.stringify({replyToken:replyToken, messages:[{type:'text', text:text}]})
  });
}

// ====== 一次性：建立工作表與範例資料（手動執行一次）======
function setup_() {
  var ss=ss_();
  var items=ss.getSheetByName(ITEM_SHEET)||ss.insertSheet(ITEM_SHEET);
  items.clear();
  items.appendRow(['ID','分類','禮盒名稱','規格','年分','節日','單位','即時庫存']);
  items.appendRow(['G1','A','公益月餅禮盒','8入/盒',2026,'中秋','盒',120]);
  items.appendRow(['G2','B','聖保羅蛋黃酥禮盒','12入/盒',2026,'中秋','盒',80]);
  items.appendRow(['G3','C','鬥茶王烏龍禮盒','150g×2罐',2026,'過年','盒',60]);
  var log=ss.getSheetByName(LOG_SHEET)||ss.insertSheet(LOG_SHEET);
  log.clear();
  log.appendRow(['時間','類型','禮盒名稱','年分','節日','數量','結存','人員/備註','來源']);
}
