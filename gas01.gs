function buildLuckyRanks(sh){
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