function sweepExpiredAndNotify() {
  return withLock(60000, () => {
    const sh = _sheet(sn_2);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const values = sh.getRange(2, 1, lastRow - 1, 5).getValues(); // A:E
    const epoch = Date.UTC(1899, 11, 30);
    const nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
    const newTn  = _toSerialInt(nowStr, epoch);

    const notify = [];

    for (let i = 0; i < values.length; i++) {
      const [id, name, tier, limit, flag] = values[i];
      const limitSerial = _toSerialInt(limit, epoch); // ← 加這行：統一轉序號
      if (limitSerial && newTn > limitSerial) {
        const row = i + 2;
        sh.getRange(row, 3).setValue(2);
        sh.getRange(row, 4).setValue('');
        sh.getRange(row, 5).setValue(nowStr);
        sh.getRange(row, 6).setValue('sys');
        notify.push(`ID:${id} 限制日:${serialToYmd(limitSerial)}`);
      }
    }

    if (notify.length) {
      const title = nowStr + ' 解禁名單';
      const body  = notify.join('\n');
      try {
        MailApp.sendEmail(MAIL_TO, title, body);
      } catch (e) {
        Logger.log('寄信失敗: ' + e.message);
      }
    }
  });
}