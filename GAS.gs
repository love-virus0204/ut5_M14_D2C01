/* 路由 */
function doGet(e){
  var p = (e && e.parameter) || {};
  var target = String(p.target || "");
  var payload = { status:"ok", msg:"get_disabled" }; // 只要能回
  if (!target) return _json(payload);
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    payload.fileExists = true;
  } catch (_) {
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

  var action = String(p.action||"").toLowerCase();
  if (!action) return _json({status:"error", msg:"unknown_action"});

  if (action === "ping") return _json({status:"ok"});
  var sheet = _sheet();
  if (!sheet) return _json({status:"error", msg:"sheet_not_found"});

// 讀表
  if (action === "list_recent") return _listRecent(sheet);

// 寫入/軟刪：加鎖
  if (action === "submit" || action === "upsert" || action === "soft_delete"){
    return withLock(60000, () => {
      if (action === "soft_delete") return _softDelete(sheet, p);
      return _submit(sheet, p);
    });
  }
    return _json({status:"error", msg:"unknown_action"});
  }

/* 寫入1..6：命中 key 覆寫 1..6 */
function _submit(sheet, p){
  var submittedAt = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
  var row = [
    submittedAt,   // 1
    p.key,         // 2
    p.date,        // 3
    p.ID,          // 4
    p.shift,       // 5
    p.dN           // 6
  ];

  var hitRow = _findRowByKey(sheet, String(p.key));
  if (hitRow > 0){
    sheet.getRange(hitRow, 1, 1, 6).setValues([row]); // 覆寫 1..6
    return _json({status:"ok", mode:"更新"});
  } else {
    sheet.appendRow([...row, "", ""]);
    var last = sheet.getLastRow();
    sheet.getRange(last, 3).setNumberFormat('mm/dd');
    return _json({status:"ok", mode:"新增"});
  }
}

/* 軟刪：逐筆覆寫 7..9 */
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

/* 讀取：取底部 150 列，依第1欄 降冪 fields+values */
function _listRecent(sheet){
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _json({status:"error", msg:"no_data"});
  }

  var lastCol  = sheet.getLastColumn() - 2;
  var startRow = Math.max(2, lastRow - 150 + 1);
  var rows     = lastRow - startRow + 1;

  var values = sheet.getRange(startRow, 1, rows, 7).getValues();

  const epoch = Date.UTC(1899,11,30);
  values.forEach(function(row, i){
    row[0] = _toSerialInt(row[0], epoch);
    row.push(startRow + i);
  });
  values.sort(function(a,b){ return b[0] - a[0]; });

  var fields = [
"submittedAt","key","date","ID","shift","dN",""okN,"row"];

  return _json({
    status: "ok", fields: fields, values: values
  });
}

/* 工具 */
function _sheet(){
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss ? ss.getSheetByName(sn_1) : null;
}

function _findRowByKey(sheet, key){
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var count = lastRow - 1;
  var keys  = sheet.getRange(2, 2, count, 1).getValues(); // 2
  for (var i=0;i<count;i++){
    if (String(keys[i][0]) === key) return i + 2;
  }
  return 0;
}

/*** 日期序號轉換（Excel 基準：1899-12-30） ***/
function _toSerialInt(v, epoch){
  if (typeof v === "number") return Math.floor(v);
  if (Object.prototype.toString.call(v) === "[object Date]"){
    return Math.floor((v.getTime() - epoch) / 86400000);
  }
  return 0;
}

function _json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
