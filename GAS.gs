/* 路由 */
function doGet(e){
  var p = (e && e.parameter) || {};
  var target = String(p.target || "");
  var payload = { status:"ok", msg:"get_disabled" };
  if (!target) return _json(payload);
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    payload.fileExists = true;
  } catch (_) {
    payload.fileExists = false;
    return _json(payload);
  }
  var found = ss.getSheets().some(function(sh){ return sh.getName() === sn_1; });
  if (found) payload.sheetExists = true;
  return _json(payload);
}


function doPost(e){
  if (!e || !e.postData) return _json({status:"error", msg:"no_post_data"});

  var p = {};
  try {
    if (e.postData.type === 'application/json') {
      p = JSON.parse(e.postData.contents || "{}");
    } else {
      p = e.parameter || {};
    }
  } catch (_){
    return _json({status:"error", msg:"bad_json"});
  }

  var action = String(p.action || "").toLowerCase();
  if (!action) return _json({ status: "error", msg: "unknown_action" });

  switch (action) {
    case "ping":
      return _json({ status: "ok" });

    case "list_recent":
      sheet = _sheet(sn_1);
      return _listRecent(sheet);

    case "list_recent2":
      sheet = _sheet(sn_2);
      return _listRecent2(sheet);

    case "auth_check":
      sheet = _sheet(sn_2);
      return _check(sheet, p);

    case "submit":
    case "upsert":
    case "lucky":
    case "soft_delete":

const mhp = getCache();
const daok = mhp && (String(mhp[p.uid] || '') === String(p.swd || ''));
if (!daok) return _json({ status: 'errorPW', msg: 'auth_failed' });

      return withLock(60000, () => {
        switch (action) {
          case "submit":
            return _submit(_sheet(sn_1), p);

          case "upsert":
            return _upsert(_sheet(sn_2), p);

          case "lucky": {
            const rankedIds = _buildLuckyRanks(_sheet(sn_2));
            return _drawLucky(_sheet(sn_1), p, rankedIds);
          }

          case "soft_delete":
            return _softDelete(_sheet(sn_2), p);

          default:
            return _json({ status: "error", msg: "unknown_action" });
        }
      });

    default:
      return _json({ status: "error", msg: "unknown_action" });
  }
}

/* 讀取：取底部 520 列， fields+values */
function _listRecent(sh){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return _json({status:"error", msg:"no_data"});
  }

  var lastCol  = sh.getLastColumn();
  var startRow = Math.max(2,lastRow - 519);
  var rows     = lastRow - startRow + 1;

  var values = sh.getRange(startRow, 1, rows, lastCol).getValues();

  const epoch = Date.UTC(1899,11,30);
  values.forEach(function(row, i){
    row[2] = _toSerialInt(row[2], epoch);
    //row.push(startRow + i);
  });

  var fields = [
"submittedAt","key","date","id","shift","dN","name","uid","deletedAt","lucky"];

  return _json({
    status: "ok", fields: fields, values: values
  });
}

function _submit(sh, p){
  var row = [
    nowTw(),     // 1
    p.key,       // 2
    p.date,      // 3
    p.id,        // 4
    p.shift,     // 5
    p.dN,        // 6
    p.uid        // 7
  ];

  var hitR = idxSync(sh, 2, p.key);
  if (hitR > 1){
    sh.getRange(hitR, 1, 1, 7).setValues([row]);
    sh.getRange(hitR, 3).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"更新"});
  } else {
    sh.appendRow(row);
    const last = sh.getLastRow();
    sh.getRange(last, 3).setNumberFormat('mm/dd');
    idxSync(sh, 2, p.key, last, 'upd');
    return _json({status:"ok", mode:"新增"});
  }
}

function idxSync(sh, col, key, val, mode) {
  const cache = CacheService.getScriptCache();
  const tag = 'IDX:' + sh.getName() + ':' + col;
  let j = cache.get(tag);
  let map = j ? JSON.parse(j) : null;

  if (mode === 'upd') {
    if (!map) return false;
    map[String(key)] = val;
    cache.put(tag, JSON.stringify(map), TTL);
    return true;
  }

  const TTL = 21600;
  if (!map) {
    const last = sh.getLastRow();
    if (last < 2) return 0;
    const vals = sh.getRange(2, col, last - 1, 1).getValues();
    map = {};
    for (let i = 0; i < vals.length; i++) {
      const k = String(vals[i][0] || '').trim();
      if (k && !map[k]) map[k] = i + 2;
    }
    cache.put(tag, JSON.stringify(map), TTL);
  } else {
    cache.put(tag, JSON.stringify(map), TTL);
  }
  const found = map[String(key)] || 0;
  return found;
}



/* 工具 - on */
function _sheet(sn) {
var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
var sh = ss ? ss.getSheetByName(sn) : null;
if (!sh) {
throw _json({ status: "error", msg: "sheet_not_found" });
}
return sh;
}

function _findRowByKey(sh, key, col){
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const arr = sh.getRange(2, col, last-1, 1).getValues().flat().map(v=>String(v));
  const i = arr.indexOf(String(key));
  return i >= 0 ? i+2 : 0;
}

/*** 日期序號轉換（yyyy-mm-dd）***/
function serialToYmd(n, epoch){
  if (!n || isNaN(n)) return '';
  const ms = n * 86400000 + epoch;
  const d  = new Date(ms);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/*** 日期序號轉換（Excel 基準）***/
function _toSerialInt(v, epoch){
  if (typeof v === 'number') return Math.floor(v);
  if (Object.prototype.toString.call(v) === '[object Date]') {
    v = Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd');
  }

  const s = String(v).trim();
  if (s === '') return 0;
  const y = parseInt(s.slice(0,4),10);
  const m = parseInt(s.slice(5,7),10) - 1;
  const d = parseInt(s.slice(8,10),10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return 0;
  return Math.floor((Date.UTC(y, m, d) - epoch) / 86400000);
}

function nowTw() {
  const raw = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
  const parts = raw.split(' ');
  const time = parts[1].split(':').map(v => v.padStart(2, '0')).join(':');
  return `${parts[0]} ${time}`;
}

function _json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
/* 工具 - off */

function _check(sh, p) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return _json({ status: "error", msg: "no_data" });

  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const found = values.find(r => r[0] === p.uid);
  if (!found) {
    return _json({ status: "error", msg: "用戶未登錄" });
  }

  if (found[6] === p.swd) {
    return _json({ status: "ok", mode: "秘鑰通過" });
  }

  return _json({ status: "error", msg: "密碼錯誤" });
}

function _listRecent2(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return _json({ status: "error", msg: "no_data" });
  }

  var lastCol = sh.getLastColumn();
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const epoch = Date.UTC(1899, 11, 30);
  const startRow = 2;

  values.forEach(function(row, i) {
    row[3] = _toSerialInt(row[3], epoch);
  });

  const fields = ["id", "name", "tier", "limit_date", "updatedAt", "uid", "pws"];

  return _json({
    status: "ok", fields: fields, values: values
  });
}


function _upsert(sh, p){
  const row = [
    p.id,          // A id
    p.name,        // B name
    p.tier,        // C tier
    p.limit_date,  // D limit_date
    nowTw(),       // E updatedAt
    p.uid          // F uid
  ];

  const hitRow = _findRowByKey(sh, String(p.id), 1);

  if (hitRow > 0){
    sh.getRange(hitRow, 1, 1, 6).setValues([row]);
    sh.getRange(hitRow, 4).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"更新"});
  } else {
    sh.appendRow([row]);
    const last = sh.getLastRow();
    sh.getRange(last, 4).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"新增"});
  }
}
