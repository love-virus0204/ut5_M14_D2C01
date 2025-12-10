    // === sh_1 清理 ==
  const sh1  = _sheet(sn_1);
  const last = sh1.getLastRow();
  const end  = Math.min(300, last);
  if (end < 2) return;

  const colC = sh1.getRange(1, 3, end, 1).getValues();
  const delRows = [];

  for (let r = 1; r < colC.length; r++) {
    const cSer = _toSerialInt(colC[r][0], epoch);
    if (cSer && cSer > 45940 && (newTn - cSer) > 30) {
      delRows.push(r + 1);
    }
  }

  for (let i = delRows.length - 1; i >= 0; i--) {
    sh1.deleteRow(delRows[i]);
  }
  });