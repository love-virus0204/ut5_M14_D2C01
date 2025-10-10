const TZ             = 'Asia/Taipei';
const SPREADSHEET_ID = '1e6ShYFbRd9OL6pVHH_BtrPUaio7t6Kxk-5KG-PYcdfI';
const sn_1     = 'sh1';
const sn_2     = 'sh2';
var sheet;
const MAIL_TO = 'nak.visu@gmail.com,eric781230@gmail.com';

function createDailyTrigger() {
  // 先刪除舊的同名觸發器（可選）
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'sweepExpiredAndNotify') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // 建立每天 9:35 執行的觸發器
  ScriptApp.newTrigger('sweepExpiredAndNotify')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .nearMinute(10)
    .create();

  Logger.log('已建立每天 09:35 執行的排程。');
}