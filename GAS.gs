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
  // 解析：JSON 優先，否則 x-www-form-urlencoded（e.parameter）
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

  sheet = _sheet(sn_1);
  switch (action) {
    case "ping":
      return _json({ status: "ok" });

    case "list_recent":
      return _listRecent(sheet);

    case "list_recent2":
      sheet = _sheet(sn_2);
      return _listRecent2(sheet);

    case "submit":
    case "upsert":
    case "soft_delete":
      return withLock(60000, () => {
        switch (action) {
          case "submit":
            return _submit(sheet, p);
          case "soft_delete":
            return _softDelete(sheet, p);
          case "upsert":
            sheet = _sheet(sn_2);
            return _upsert(sheet, p);

          case "lucky": 
            sheet = _sheet(sn_2);
            const rankedIds = _buildLuckyRanks(sheet);
            sheet = _sheet(sn_1);
            const dateSerial = Number(p.dateSerial);
            return _drawLucky(sheet, dateSerial, rankedIds);
        }
      });

    default:
      return _json({ status: "error", msg: "unknown_action" });
  }
}
/* 寫入1..7：命中 key 覆寫 1..7 */
function _submit(sheet, p){
  var submittedAt = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
  var row = [
    submittedAt,   // 1
    p.key,         // 2
    p.date,        // 3
    p.id,          // 4
    p.shift,       // 5
    p.dN,          // 6
    'user'         // 7
  ];

  var hitRow = _findRowByKey(sheet, String(p.key), 2);
  if (hitRow > 0){
    sheet.getRange(hitRow, 1, 1, 7).setValues([row]); // 覆寫 1..7
    sheet.getRange(hitRow, 3).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"更新"});
  } else {
    sheet.appendRow(row);
    var last = sheet.getLastRow();
    sheet.getRange(last, 3).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"新增"});
  }
}


function _softDelete(sheet, p){
  var admin_id = String(p.admin_id || "");
  if (!admin_id) 
    return _json({status:"error", msg:"no_admin_id"});
  var lastRow = sheet.getLastRow();
  var targets = [];
  for (var k in p){
    if (/^row\d+$/.test(k)) {
      var r = Number(p[k]);
      if (r && r >= 2 && r <= lastRow) targets.push(r);
    }
  }
  if (targets.length === 0) 
    return _json({status:"error", msg:"not_found"});

  var deletedAt = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
  var rowValue  = ["DEL", admin_id, deletedAt]; // 1×3

  for (var i = 0; i < targets.length; i++){
    sheet.getRange(targets[i], 7, 1, 3).setValues([rowValue]);
  }
  return _json({status:"ok", count: targets.length});
}

/* 讀取：取底部 160 列，依第1欄 降冪 fields+values */
function _listRecent(sheet){
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _json({status:"error", msg:"no_data"});
  }

  var lastCol  = sheet.getLastColumn();
  var startRow = Math.max(2, lastRow - 160 + 1);
  var rows     = lastRow - startRow + 1;

  var values = sheet.getRange(startRow, 1, rows, lastCol).getValues();

  const epoch = Date.UTC(1899,11,30);
  values.forEach(function(row, i){
    row[2] = _toSerialInt(row[2], epoch);
    row.push(startRow + i);
  });

  values.sort((a,b)=> a[5].localeCompare(b[5]));
  values.sort((a,b)=>{
  return String(a[3]).localeCompare(String(b[3]), 'en', { numeric:true });
});
  values.sort((a,b)=> b[2] - a[2]);

  var fields = [
"submittedAt","key","date","id","shift","dN","admin_id","deletedAt","lucky","row"];

  return _json({
    status: "ok", fields: fields, values: values
  });
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

function _findRowByKey(sheet, key, ct){
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var rows = lastRow - 1;
  var keys  = sheet.getRange(2, ct, rows, 1).getValues();
  for (var i=0;i<rows;i++){
    if (String(keys[i][0]) === key) return i + 2;
  }
  return 0;
}

/*** 日期序號轉換（yyyy-mm-dd）***/
function serialToYmd(n, epoch){
  if (!n || isNaN(n)) return '';
  const ms = n * 86400000 + epoch;
  const d  = new Date(ms);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/*** 日期序號轉換（Excel 基準：1899-12-30） ***/
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

function _json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 工具 - off */

function _listRecent2(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _json({ status: "error", msg: "no_data" });
  }

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const epoch = Date.UTC(1899, 11, 30);
  const startRow = 2;

  values.forEach(function(row, i) {
    row[3] = _toSerialInt(row[3], epoch);
    row.push(startRow + i);
  });

  values.sort(function(a,b){ return b[4] - a[4]; });

  const fields = ["id", "name", "tier", "limit_date", "updatedAt", "admin_id", "row"];

  return _json({
    status: "ok", fields: fields, values: values
  });
}


function _upsert(sheet, p){
  const updatedAt = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');

  const row = [
    p.id,          // A id
    p.name,        // B name
    p.tier,        // C tier
    p.limit_date,  // D limit_date
    updatedAt,     // E updatedAt
    '18B16'        // F admin_id
  ];

  const hitRow = _findRowByKey(sheet, String(p.id), 1);

  if (hitRow > 0){
    sheet.getRange(hitRow, 1, 1, 6).setValues([row]);
    sheet.getRange(hitRow, 4).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"更新"});
  } else {
    sheet.appendRow(row);
    const last = sheet.getLastRow();
    sheet.getRange(last, 4).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"新增"});
  }
}