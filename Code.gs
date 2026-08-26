/*************************************************************
 * 年節禮盒庫存系統 — 後端（Google Apps Script）v2
 * 一支程式同時服務：
 *   ① LINE Bot（webhook）
 *   ② 網頁（index.html）讀寫 API
 * 資料全部存在這份 Google 試算表，三邊共用同一份資料。
 *
 * 安裝／更新步驟（詳見「網頁連動_更新說明.md」）：
 *   1. 把 CHANNEL_ACCESS_TOKEN 換成你的長期 token
 *   2. 執行一次 initSheets()（非破壞性，只建立表頭與必要結構）
 *   3. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署
 *      （網址不變；首次部署則用「新增部署作業 → 網頁應用程式」）
 *   4. 把網址貼到 index.html 的 API_URL，以及 LINE 後台 Webhook URL
 *************************************************************/

/*** ❶ 貼上你的 Channel access token（長期，很長那一串） ***/
const CHANNEL_ACCESS_TOKEN = '在這裡貼上你的_CHANNEL_ACCESS_TOKEN';

/*** 工作表名稱 ***/
const SHEET_ITEMS = '禮盒';
const SHEET_TX    = '異動';
const SHEET_PLAN  = '送禮規劃';
const SHEET_CO    = '公司';
const SHEET_HIST  = '規劃歷史';

const CATS  = { A: '公益', B: '聖保羅', C: '鬥茶王' };
const FESTS = ['過年', '端午', '中秋'];

/* ❷ 送禮總表頁網址（GitHub Pages，截圖模式）；換成你自己的網址即可 */
const TOTAL_URL = 'https://brisbanekan.github.io/stocksystem/?view=total&shot=1';

/* 送禮規劃欄位順序（程式內部 key）
 * a/b/c 只保留公益／聖保羅／鬥茶王三個內建類別的數字，方便在試算表用肉眼看；
 * cats 才是完整資料（JSON，含各公司自訂的禮盒類別），前端一律以 cats 為準，a/b/c 為輔助顯示。 */
const PLAN_KEYS = ['id','co','year','fest','dept','target','plan','a','b','c','note','signer','status','shipped','txRefs','delivered','date','cats','deliveredCats'];
const PLAN_HEADERS = ['id','公司','年分','節日','部門','送禮對象','預計','公益','聖保羅','鬥茶王','備註','簽核','狀態','已扣庫存','出貨紀錄id','配送明細','配送日期','類別JSON(含自訂類別)','各類別送達明細JSON'];

/* ======================================================
 *  入口：doPost（LINE 與 網頁 共用） / doGet
 * ====================================================== */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}

  // 有 events → 來自 LINE
  if (body && body.events) {
    try { body.events.forEach(handleEvent); }
    catch (err) { console.error('LINE error: ' + err); }
    return jsonOut({ ok: true });
  }
  // 否則 → 來自網頁的 API 請求
  return handleApi(body || {});
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) return handleApi(e.parameter);
  return ContentService.createTextOutput('年節禮盒系統後端運作中 ✅');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================================================
 *  網頁 API
 * ====================================================== */
function handleApi(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (req.action) {
      case 'load':       return jsonOut(loadAll());
      case 'applyTx':    return jsonOut(apiApplyTx(req));
      case 'adjustQty':  return jsonOut(apiAdjustQty(req));
      case 'addItem':    return jsonOut(apiAddItem(req));
      case 'deleteItem': return jsonOut(apiDeleteItem(req));
      case 'removeTxs':  return jsonOut(apiRemoveTxs(req));
      case 'savePlan':   return jsonOut(apiSavePlan(req));
      case 'listVersions': return jsonOut(apiListVersions());
      case 'loadVersion':  return jsonOut(apiLoadVersion(req));
      default:           return jsonOut({ ok: false, error: 'unknown action: ' + req.action });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function loadAll() {
  return {
    ok: true,
    items: readItems().map(function (i) {
      return { id: i.id, cat: i.cat, name: i.name, spec: i.spec, year: i.year, fest: i.fest, unit: i.unit, qty: i.qty };
    }),
    txs: readTxs(),
    plan: readPlan(),
    companies: readCompanies()
  };
}

function apiApplyTx(req) {
  const items = readItems();
  const it = items.find(function (i) { return String(i.id) === String(req.itemId); });
  if (!it) return { ok: false, error: 'item not found' };
  const qty = Number(req.qty) || 0;
  if (qty <= 0) return { ok: false, error: 'qty must be > 0' };
  if (req.type === '出料' && qty > it.qty) return { ok: false, error: '庫存不足', itemQty: it.qty };
  const tx = applyTx(req.type, it, qty, req.who || '');
  return { ok: true, itemQty: it.qty, tx: tx };
}

function apiAdjustQty(req) {
  const items = readItems();
  const it = items.find(function (i) { return String(i.id) === String(req.itemId); });
  if (!it) return { ok: false, error: 'item not found' };
  const nq = it.qty + (Number(req.delta) || 0);
  sheet(SHEET_ITEMS).getRange(it.row, 8).setValue(nq);
  return { ok: true, itemQty: nq };
}

function apiAddItem(req) {
  const it = req.item || {};
  const items = readItems();
  const id = items.reduce(function (m, i) { return Math.max(m, Number(i.id) || 0); }, 0) + 1;
  sheet(SHEET_ITEMS).appendRow([id, it.cat, it.name, it.spec, Number(it.year) || new Date().getFullYear(),
    it.fest, it.unit || '盒', Number(it.qty) || 0]);
  return { ok: true, id: id };
}

function apiDeleteItem(req) {
  const sh = sheet(SHEET_ITEMS);
  const items = readItems();
  const it = items.find(function (i) { return String(i.id) === String(req.itemId); });
  if (!it) return { ok: false, error: 'item not found' };
  const id = newId();
  const time = nowStamp();
  const who = '原因：' + (req.reason || '未填原因');
  sheet(SHEET_TX).appendRow([time, '刪除', label(it), it.qty, '—', who, it.year, it.fest, it.cat, '', id]);
  sh.deleteRow(it.row);
  return { ok: true, tx: { id: id, time: time, type: '刪除', name: label(it), qty: it.qty, bal: '—',
    who: who, year: it.year, fest: it.fest, cat: it.cat, target: '' } };
}

function apiRemoveTxs(req) {
  const ids = {};
  (req.ids || []).forEach(function (x) { ids[String(x)] = 1; });
  const sh = sheet(SHEET_TX);
  const data = sh.getDataRange().getValues();
  for (var r = data.length - 1; r >= 1; r--) {     // 由下往上刪，避免列號位移
    if (ids[String(data[r][10])]) sh.deleteRow(r + 1);
  }
  return { ok: true };
}

function apiSavePlan(req) {
  // 送禮規劃（整張覆蓋；此資料僅網頁使用，LINE 不會動）
  const sh = sheet(SHEET_PLAN) || ss().insertSheet(SHEET_PLAN);
  sh.clearContents();
  sh.getRange(1, 1, 1, PLAN_HEADERS.length).setValues([PLAN_HEADERS]).setFontWeight('bold');
  const rows = (req.plan || []).map(function (r) {
    return PLAN_KEYS.map(function (k) {
      if (k === 'shipped') return r.shipped ? true : false;
      var v = r[k];
      return (v === undefined || v === null) ? '' : v;
    });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, PLAN_KEYS.length).setValues(rows);
  sh.setFrozenRows(1);

  // 公司名稱＋各公司自訂禮盒類別（第3欄為 categories 的 JSON，例如 {"cat_xxx":"特別禮盒"}）
  const cs = sheet(SHEET_CO) || ss().insertSheet(SHEET_CO);
  cs.clearContents();
  cs.getRange(1, 1, 1, 3).setValues([['key', '名稱', '自訂類別JSON']]).setFontWeight('bold');
  const co = req.companies || {};
  const crows = Object.keys(co).map(function (k) {
    return [k, (co[k] && co[k].name) || k, JSON.stringify((co[k] && co[k].categories) || {})];
  });
  if (crows.length) cs.getRange(2, 1, crows.length, 3).setValues(crows);
  cs.setFrozenRows(1);

  // 版本快照（每次儲存留一版；與上一版相同則略過）
  snapshotPlan(req.plan || [], req.companies || {}, req.editor || '');

  return { ok: true };
}

/* ======================================================
 *  規劃版本紀錄（每次儲存留版本；可查閱／還原）
 * ====================================================== */
function snapshotPlan(planArr, companies, editor) {
  var sh = sheet(SHEET_HIST);
  if (!sh) sh = ss().insertSheet(SHEET_HIST);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 5).setValues([['版本時間', '編輯者', '對象數', '規劃JSON', '公司JSON']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var planJSON = JSON.stringify(planArr || []);
  var last = sh.getLastRow();
  if (last >= 2 && sh.getRange(last, 4).getValue() === planJSON) return; // 無變化不留版本
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([ts, editor || '', (planArr || []).length, planJSON, JSON.stringify(companies || {})]);
}

function apiListVersions() {
  var sh = sheet(SHEET_HIST);
  if (!sh) return { ok: true, versions: [] };
  var data = sh.getDataRange().getValues();
  data.shift();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== '') out.push({ row: i + 2, ts: data[i][0], editor: data[i][1], count: data[i][2] });
  }
  out.reverse();
  return { ok: true, versions: out.slice(0, 300) };
}

function apiLoadVersion(req) {
  var sh = sheet(SHEET_HIST);
  if (!sh) return { ok: false, error: 'no history' };
  var row = Number(req.row);
  if (!row || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'bad version' };
  var v = sh.getRange(row, 1, 1, 5).getValues()[0];
  var plan = [], companies = {};
  try { plan = JSON.parse(v[3] || '[]'); } catch (e) {}
  try { companies = JSON.parse(v[4] || '{}'); } catch (e) {}
  return { ok: true, ts: v[0], editor: v[1], plan: plan, companies: companies };
}

/* ======================================================
 *  試算表存取
 * ====================================================== */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name) { return ss().getSheetByName(name); }

function readItems() {
  const sh = sheet(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();
  data.shift();
  return data
    .filter(function (r) { return r[0] !== '' && r[2] !== ''; })
    .map(function (r, idx) {
      return { row: idx + 2, id: r[0], cat: r[1], name: r[2], spec: r[3],
        year: Number(r[4]), fest: r[5], unit: r[6] || '盒', qty: Number(r[7]) || 0 };
    });
}

function readTxs() {
  const sh = sheet(SHEET_TX);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  data.shift();
  return data
    .filter(function (r) { return r[1] !== ''; })
    .map(function (r) {
      return { id: r[10] || '', time: r[0], type: r[1], name: r[2], qty: r[3], bal: r[4],
        who: r[5], year: Number(r[6]) || '', fest: r[7], cat: r[8], target: r[9] };
    });
}

function readPlan() {
  const sh = sheet(SHEET_PLAN);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  data.shift();
  return data
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const o = {};
      PLAN_KEYS.forEach(function (k, i) { o[k] = r[i]; });
      return o;
    });
}

function readCompanies() {
  const sh = sheet(SHEET_CO);
  const out = {};
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  data.shift();
  data.forEach(function (r) {
    if (r[0] === '') return;
    var cats = {};
    try { cats = r[2] ? JSON.parse(r[2]) : {}; } catch (e) { cats = {}; }
    out[r[0]] = { name: r[1] || r[0], categories: cats };
  });
  return out;
}

/* 進/出料：更新庫存欄、寫一筆異動，回傳該筆異動 */
function applyTx(type, item, qty, who) {
  const newQty = type === '進料' ? item.qty + qty : item.qty - qty;
  sheet(SHEET_ITEMS).getRange(item.row, 8).setValue(newQty);
  item.qty = newQty;
  const id = newId();
  const time = nowStamp();
  const target = type === '出料' ? (who || '').trim() : '';
  sheet(SHEET_TX).appendRow([time, type, label(item), qty, newQty, who || '-',
    item.year, item.fest, item.cat, target, id]);
  return { id: id, time: time, type: type, name: label(item), qty: qty, bal: newQty,
    who: who || '-', year: item.year, fest: item.fest, cat: item.cat, target: target };
}

function newId() { return 't' + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36); }
function nowStamp() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM-dd HH:mm'); }
function label(i) { return i.name + '（' + i.year + i.fest + '）'; }

/* ======================================================
 *  模糊比對（與網頁原型相同）
 * ====================================================== */
function matchItems(items, key) {
  const norm = function (s) { return String(s).toLowerCase().replace(/\s+/g, ''); };
  const k = norm(key);
  if (!k) return [];
  const isSub = function (a, b) { var i = 0; for (var x = 0; x < b.length; x++) { if (b[x] === a[i]) i++; } return i === a.length; };
  return items.map(function (it) {
    const n = norm(it.name);
    var s = 0;
    if (n === k) s = 100; else if (n.indexOf(k) >= 0) s = 80; else if (isSub(k, n)) s = 60;
    return { it: it, s: s };
  }).filter(function (x) { return x.s > 0; }).sort(function (a, b) { return b.s - a.s; });
}
function resolveItem(items, key) {
  const m = matchItems(items, key);
  if (!m.length) return { status: 'none' };
  if (m.length === 1 || m[0].s > m[1].s) return { status: 'ok', item: m[0].it };
  return { status: 'ambiguous', candidates: m.filter(function (x) { return x.s === m[0].s; }).map(function (x) { return x.it; }) };
}
function ambigMsg(cands) {
  return '🔍 找到多個符合的禮盒，請輸入更完整的名稱：\n' +
    cands.map(function (i) { return '· ' + i.name + '（' + i.year + i.fest + '）'; }).join('\n');
}

/* ======================================================
 *  LINE 事件 + 指令解析
 * ====================================================== */
function handleEvent(ev) {
  if (ev.type !== 'message' || ev.message.type !== 'text') return;
  const text = ev.message.text;
  const t = text.trim();
  if (/^(總表|查總表|送禮總表|總覽)$/.test(t)) { replyTotalImage(ev.replyToken); return; }
  if (/^(未送|還沒送|誰沒送|沒送|未送達|送禮進度|誰還沒送)$/.test(t)) { replyToLine(ev.replyToken, buildPendingMsg() || '🎉 目前規劃中的對象都已送達！'); return; }
  const isWrite = /^(進料|出料)/.test(t);
  var lock;
  if (isWrite) { lock = LockService.getScriptLock(); lock.waitLock(15000); }
  try {
    replyToLine(ev.replyToken, parseCommand(text));
  } finally {
    if (lock) lock.releaseLock();
  }
}

function replyToLine(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
}

/* ======================================================
 *  送禮總表：即時截圖（thum.io，免註冊）
 * ====================================================== */
function replyTotalImage(replyToken) {
  const base = 'https://image.thum.io/get/width/1280/wait/9/noanimate/' + TOTAL_URL;
  const prev = 'https://image.thum.io/get/width/640/wait/9/noanimate/' + TOTAL_URL;
  // 先預熱：讓 thum.io 先產生最新截圖並快取，LINE 取圖才不會逾時
  try { UrlFetchApp.fetch(base, { muteHttpExceptions: true }); } catch (e) {}
  const link = TOTAL_URL.replace('&shot=1', '');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [
        { type: 'image', originalContentUrl: base, previewImageUrl: prev },
        { type: 'text', text: '🎁 送禮總表即時截圖\n看互動版：' + link }
      ]
    }),
    muteHttpExceptions: true
  });
}

/* ======================================================
 *  每小時推播「還有誰沒送」（廣播給所有好友）
 *  只在「有未送對象」時才發訊息（全部送完就安靜，省 LINE 額度）
 * ====================================================== */
/* 組「還沒送（未送達）」清單文字；全部送完回傳空字串 */
function buildPendingMsg() {
  const rows = readPlan();
  const pending = rows.filter(function (r) { return String(r.status) !== '已送達'; });
  if (!pending.length) return '';
  const coNames = readCompanies();
  const byCo = {};
  pending.forEach(function (r) { var k = r.co; (byCo[k] = byCo[k] || []).push(r); });
  var totBoxes = pending.reduce(function (s, r) { return s + (Number(r.plan) || 0); }, 0);
  var msg = '🎁 還沒送（未送達）\n共 ' + pending.length + ' 位，' + totBoxes + ' 盒\n';
  Object.keys(byCo).forEach(function (co) {
    var name = (coNames[co] && coNames[co].name) || co;
    msg += '\n【' + name + '】\n';
    byCo[co].slice(0, 25).forEach(function (r) {
      msg += '· ' + r.year + r.fest + ' ' + r.dept + '/' + r.target + '（' + r.plan + '盒·' + r.status + '）\n';
    });
    if (byCo[co].length > 25) msg += '…等共 ' + byCo[co].length + ' 位\n';
  });
  if (msg.length > 4900) msg = msg.slice(0, 4900) + '\n…(已截斷)';
  return msg.trim();
}

/* 每小時自動推播（廣播）：只在有未送對象時才發 */
function hourlyReminder() {
  const msg = buildPendingMsg();
  if (!msg) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ messages: [{ type: 'text', text: '【自動提醒】\n' + msg }] }),
    muteHttpExceptions: true
  });
}

/* 執行一次以啟用每小時自動推播；要停用就執行 removeHourlyReminder */
function setupHourlyReminder() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'hourlyReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('hourlyReminder').timeBased().everyHours(1).create();
}
function removeHourlyReminder() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'hourlyReminder') ScriptApp.deleteTrigger(t);
  });
}

function parseCommand(text) {
  const parts = text.trim().split(/\s+/);
  const cmd = parts[0];
  const items = readItems();

  if (/^(說明|help|\?|？)$/i.test(cmd)) {
    return '📖 指令說明\n進料 <禮盒名> <數量> [備註]\n出料 <禮盒名> <數量> [出貨對象]\n庫存 <禮盒名>　查單一禮盒\n庫存 <節日>　依過年/端午/中秋查\n庫存　　　　列出全部\n規格 <禮盒名>　查規格\n總表　　　　送禮總表即時截圖\n未送　　　　查誰還沒送到';
  }

  if (cmd === '庫存') {
    const arg = parts.slice(1).join(' ');
    if (arg && FESTS.indexOf(arg) >= 0) {
      const list = items.filter(function (i) { return i.fest === arg; });
      if (!list.length) return '查無「' + arg + '」禮盒';
      return '🎁 ' + arg + '禮盒庫存：\n' + list.map(function (i) {
        return '· [' + i.cat + '] ' + i.name + ' ' + i.qty + ' ' + i.unit + '（' + i.year + '）'; }).join('\n');
    }
    if (arg) {
      const res = resolveItem(items, arg);
      if (res.status === 'none') return '❓ 找不到禮盒「' + arg + '」';
      if (res.status === 'ambiguous') return ambigMsg(res.candidates);
      const it = res.item;
      return '🎁 ' + it.name + '\n分類：' + it.cat + ' ' + (CATS[it.cat] || it.cat) +
        '\n年分/節日：' + it.year + ' ' + it.fest + '\n規格：' + it.spec +
        '\n即時庫存：' + it.qty + ' ' + it.unit;
    }
    if (!items.length) return '目前無任何禮盒';
    return '🎁 全部禮盒庫存：\n' + items.map(function (i) {
      return '· [' + i.cat + '] ' + i.name + ' ' + i.qty + i.unit + '（' + i.year + i.fest + '）'; }).join('\n');
  }

  if (cmd === '規格') {
    const res = resolveItem(items, parts.slice(1).join(' '));
    if (res.status === 'none') return '❓ 找不到禮盒';
    if (res.status === 'ambiguous') return ambigMsg(res.candidates);
    const it = res.item;
    return '📐 ' + it.name + '\n規格：' + it.spec + '\n分類：' + it.cat + ' ' + (CATS[it.cat] || it.cat) + '\n單位：' + it.unit;
  }

  if (cmd === '進料' || cmd === '出料') {
    if (parts.length < 3) return '格式：' + cmd + ' <禮盒名> <數量> [' + (cmd === '出料' ? '出貨對象' : '備註') + ']';
    const qtyIdx = parts.findIndex(function (p, i) { return i > 0 && /^\d+$/.test(p); });
    if (qtyIdx < 0) return '❓ 請輸入數量，例如：' + cmd + ' 公益月餅禮盒 100';
    const name = parts.slice(1, qtyIdx).join(' ');
    const qty = parseInt(parts[qtyIdx], 10);
    const extra = parts.slice(qtyIdx + 1).join(' ');
    const res = resolveItem(items, name);
    if (res.status === 'none') return '❓ 找不到禮盒「' + name + '」，可先在試算表或網頁新增';
    if (res.status === 'ambiguous') return ambigMsg(res.candidates);
    const it = res.item;
    if (cmd === '出料' && qty > it.qty) return '❌ ' + it.name + ' 庫存不足，目前僅 ' + it.qty + ' ' + it.unit;
    applyTx(cmd, it, qty, extra);
    const sign = cmd === '進料' ? '+' : '−';
    const xlabel = cmd === '出料' ? '出貨對象' : '備註';
    return '✅ ' + cmd + '登記成功\n' + label(it) + ' ' + sign + qty + ' ' + it.unit +
      '\n目前庫存：' + it.qty + ' ' + it.unit + (extra ? '\n' + xlabel + '：' + extra : '');
  }

  return '🤔 看不懂指令，輸入「說明」查看用法';
}

/* ======================================================
 *  一鍵建立／補齊工作表（非破壞性：不清掉既有資料）
 * ====================================================== */
function initSheets() {
  const book = ss();

  // 禮盒：沒有才建，空表只建立表頭
  var shItems = book.getSheetByName(SHEET_ITEMS);
  if (!shItems) shItems = book.insertSheet(SHEET_ITEMS);
  if (shItems.getLastRow() === 0) {
    shItems.getRange(1, 1, 1, 8)
      .setValues([['id', '分類', '禮盒名稱', '規格', '年分', '節日', '單位', '庫存']]).setFontWeight('bold');
    shItems.setFrozenRows(1);
  }

  // 異動：確保表頭為 11 欄（含 id）；不動既有資料
  var shTx = book.getSheetByName(SHEET_TX);
  if (!shTx) shTx = book.insertSheet(SHEET_TX);
  shTx.getRange(1, 1, 1, 11)
    .setValues([['時間', '類型', '禮盒', '數量', '結存', '人員/備註', '年分', '節日', '分類', '出貨對象', 'id']])
    .setFontWeight('bold');
  shTx.setFrozenRows(1);

  // 送禮規劃：沒有才建，空表只建立表頭
  var shPlan = book.getSheetByName(SHEET_PLAN);
  if (!shPlan) shPlan = book.insertSheet(SHEET_PLAN);
  if (shPlan.getLastRow() === 0) {
    shPlan.getRange(1, 1, 1, PLAN_HEADERS.length).setValues([PLAN_HEADERS]).setFontWeight('bold');
    shPlan.setFrozenRows(1);
  }

  // 公司：沒有才建、空的才放範例（第3欄放各公司自訂禮盒類別的 JSON，預設空物件）
  var shCo = book.getSheetByName(SHEET_CO);
  if (!shCo) shCo = book.insertSheet(SHEET_CO);
  if (shCo.getLastRow() === 0) {
    shCo.getRange(1, 1, 1, 3).setValues([['key', '名稱', '自訂類別JSON']]).setFontWeight('bold');
    shCo.getRange(2, 1, 2, 3).setValues([['c1', '公司一', '{}'], ['c2', '公司二', '{}']]);
    shCo.setFrozenRows(1);
  }

  // 規劃歷史：版本快照
  var shHist = book.getSheetByName(SHEET_HIST);
  if (!shHist) shHist = book.insertSheet(SHEET_HIST);
  if (shHist.getLastRow() === 0) {
    shHist.getRange(1, 1, 1, 5).setValues([['版本時間', '編輯者', '對象數', '規劃JSON', '公司JSON']]).setFontWeight('bold');
    shHist.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
}

/*
 * 正式啟用前清空所有業務資料。
 * 此函式刻意不開放為公開 API，避免任何人透過網址誤刪資料；
 * 只能由試算表綁定的 Apps Script 編輯器手動執行。
 */
function clearAllDataForLaunch() {
  [SHEET_ITEMS, SHEET_TX, SHEET_PLAN, SHEET_CO, SHEET_HIST].forEach(function (name) {
    var sh = sheet(name);
    if (sh) sh.clearContents();
  });
  initSheets();
}
