function sweepExpiredAndNotify() {
  return withLock(60000, () => {
    const sh = _sheet(sn_2);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const values = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    const epoch = Date.UTC(1899, 11, 30);
    const nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
    const newTn  = _toSerialInt(nowStr, epoch);

    const notify = [];

    for (let i = 0; i < values.length; i++) {
      const [id, name, tier, limit, flag, ad] = values[i];
      const limitSerial = _toSerialInt(limit, epoch);
      if (limitSerial && newTn > limitSerial) {
        const row = i + 2;
        var mt = [
          id,     // 1
          name,   // 2
          '2',    // 3
          '',     // 4
          nowStr, // 5
          'sys2'   // 6
        ];
        sh.getRange(row, 1, 1, 6).setValues(mt);

        // notify.push(`ID:${id} 限制日:${serialToYmd(limitSerial)}`);
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