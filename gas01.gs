function _drawLucky(sh, dateSerial, rankedIds) {
  const epoch = Date.UTC(1899,11,30);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');

  const last  = sh.getLastRow();
  const start = Math.max(2, last - 519);
  const count = last - start + 1;

  const big = sh.getRange(start, 1, count, 9).getValues();
  const sub = sh.getRange(start, 7, count, 3).getValues();

  const rankMap = Object.create(null);
  for (let i=0;i<rankedIds.length;i++) rankMap[ rankedIds[i] ] = i+1;

  const C_DATE=2, D_ID=3, E_FLAG=4, F_KIND=5, G_SUB=0, H_SUB=1, I_SUB=2, I_BIG=8;
  let updated=0;

  for (let i=0;i<count;i++){
    const serial = _toSerialInt(big[i][C_DATE], epoch);
    if (serial !== dateSerial) continue;

    const id   = String(big[i][D_ID]);
    const eFlg = String(big[i][E_FLAG]);
    const kind = String(big[i][F_KIND]);
    const r    = rankMap[id];

    let val = '9999';
    if (eFlg !== 'n' && r > 0) {
      const prefix = (kind === '假日') ? '1' : '2';
      val = prefix + pad3(r);
    }

    big[i][I_BIG] = val;
    sub[i][I_SUB] = val;
    sub[i][G_SUB] = 'lucky';
    sub[i][H_SUB] = now;
    updated++;
  }

  sh.getRange(start, 7, count, 3).setValues(sub);

  big.forEach((row, i) => {
    row[2] = _toSerialInt(row[2], epoch);
    row.push(start + i);
  });

  big.sort((a,b)=> a[8] - b[8]);

  const fields = ["submittedAt","key","date","id","shift","dN","admin_id","deletedAt","lucky","row"];

  return _json({ status:"ok", fields, values: big });

  function pad3(n){ return ('000'+n).slice(-3); }
}




function _buildLuckyRanks(sh){
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const pool = [];
  const weightMap = {};

  for (let i = 0; i < values.length; i++) {
    const [id, , tier] = values[i];
    const w = Number(tier);
    if (w <= 0) continue;
    weightMap[id] = w;
    for (let j = 0; j < w; j++) pool.push(id);
  }

  for (let i = pool.length - 1; i > 0; i--) {
    const j = (Math.random() * (i + 1)) | 0;
    [pool[i], pool[j]] = [pool[j], pool[i]];
  }

  const ranked = [...new Set(pool)];
  return ranked;
}