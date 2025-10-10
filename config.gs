const TZ             = 'Asia/Taipei';
const SPREADSHEET_ID = '1JxWgzTA6_Dsd3M4t7eT6UCj2fraaKtpN3SqJ4Z_ylP4';
const sn_1     = 'sh1';
const sn_2     = 'sh2';
var sheet;
const MAIL_TO = 'ut.ypd14@gmail.com,eric781230@gmail.com';

function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'sweepExpiredAndNotify') {
      ScriptApp.deleteTrigger(t);
    }
  } ScriptApp.newTrigger('sweepExpiredAndNotify')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .nearMinute(10)
    .create();
  Logger.log('已建立每天 10:10 執行的排程。');
}

function createPingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'ping') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('ping')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('✅ 已建立每 5 分鐘執行一次 ping() 的排程。');
}