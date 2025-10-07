function sweepExpiredAndNotify() {
  return withLock(60000, () => {
    sheet = _sheet(sn_2);
    const resp = _listRecent2(sheet);
    // if (!resp || resp.status !== 'ok') return;

    const F = resp.fields;
    const iId  = F.indexOf('id');
    const iLim = F.indexOf('limit_date');
    const iRow = F.indexOf('row');

    const nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
    const epoch = Date.UTC(1899,11,30);
    const newTn  = _toSerialInt(nowStr, epoch);

    const hits = [];
    const notify = [];

    resp.values.forEach(v => {
      const limit = v[iLim];
      if (limit && newTn > limit) {
        const row = v[iRow];
        hits.push(row);
        notify.push(`ID:${v[iId]} 限制日:${serialToYmd(limit)}`);
        sheet.getRange(row, 5).setValue('');
        sheet.getRange(row, 4).setValue(2);
      }
    });

    if (notify.length) {
      const title = nowStr + ' 解禁名單';
      const body  = notify.join('\n');
      MailApp.sendEmail(MAIL_TO, title, body);
    }
  });
}