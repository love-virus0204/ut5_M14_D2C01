/** 清理 sh_1 緩存*/
function Notify(){
  return withLock(60000, () => {
    const sh1   = _sheet(sn_1);
    const cache = CacheService.getScriptCache();
    cache.remove('IDX:' + sh1.getName() + ':2');
  });
}

/** 解禁 + 寄信 + 清理 sh_1（sn_2 無資料跳過）*/
function sweepExpiredAndNotify(){
  return withLock(60000, () => {
    const sh = _sheet(sn_2);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const values = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    const epoch = Date.UTC(1899, 11, 30);
    const now   = nowTw();
    const nowStr = now.slice(0,10).replace(/-/g,'/');
    const newTn  = _toSerialInt(nowStr, epoch);
    const ymd    = serialToYmd(newTn, epoch);
    const notify = [];
    for (let i = 0; i < values.length; i++) {
      const [id, name, tier, limit, flag, ad] = values[i];
      const limitSerial = _toSerialInt(limit, epoch);
      if (limitSerial && newTn > limitSerial) {
        const row = i + 2;
        const mt = [
          id,     // A id
          name,   // B name
          2,      // C tier
          '',     // D limit_date
          now,    // E updatedAt
          'sys'   // F admin_id
        ];
        sh.getRange(row, 1, 1, 6).setValues([mt]);
        notify.push(`ID：${id}　｜　姓名：${name}　｜　限制日：${serialToYmd(limitSerial, epoch)}`);
      }
    }

    if (notify.length) {
      try {
        MailApp.sendEmail(MAIL_TO, ymd + ' 解禁名單', notify.join('\n'));
      } catch (e) {
        Logger.log('寄信失敗: ' + e.message);
      }
    }

    // === sh_1 清理 ==
    if (newTn % 7 !== 3) return;
    const sh1   = _sheet(sn_1);
    const cache = CacheService.getScriptCache();
    cache.remove('IDX:' + sh1.getName() + ':2');
    const aLast = sh1.getLastRow();
    const upper = Math.min(520 , aLast - 1);
    if (upper < 2) return;
      const colC = sh1.getRange(2, 3, upper, 1).getValues();
      const delRows = [];
      for (let i = 0; i < colC.length; i++){
        const cSer = _toSerialInt(colC[i][0], epoch);
        if (cSer && (newTn - cSer) > 30) delRows.push(i + 2);
      }
      for (let k = delRows.length - 1; k >= 0; k--) sh1.deleteRow(delRows[k]);
  });
}