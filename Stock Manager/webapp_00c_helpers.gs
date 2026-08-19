/**
 * ============================================================
 * webapp_00c_helpers.gs /CONFIG  Backup ไฟล์
 * ------------------------------------------------------------
 * สำรองไฟล์ Google Sheet (+ Apps Script ที่ผูกอยู่กับไฟล์นั้นโดยอัตโนมัติ —
 * เวลา DriveApp.makeCopy() ไฟล์ Sheet ที่มี container-bound script ตัว copy
 * จะได้โค้ด .gs ทั้งหมดติดไปด้วย ไม่ต้องใช้ Apps Script API ให้ยุ่งยาก)
 *
 * รันจากโปรเจกต์ไหนก็ได้ (ไม่ต้องเป็นโปรเจกต์เดียวกับไฟล์ที่จะสำรอง) ขอแค่
 * บัญชีที่รันมีสิทธิ์ดูไฟล์ต้นทางและเขียนใน Google Drive
 *
 * ⚠️ ต้องทำก่อนใช้งาน:
 *  1. ใส่ Spreadsheet ID ที่ต้องการสำรองใน BACKUP_TARGETS ด้านล่าง
 *  2. รัน setupDailyBackupTrigger() ครั้งเดียว (ตั้ง Trigger รายวัน)
 *  3. (แนะนำ) รัน testBackupNow() ก่อน 1 ครั้ง เช็คว่าสำรองได้จริงและ
 *     เจอโฟลเดอร์ปลายทางถูกไฟล์
 * ============================================================
 */

const BACKUP_TARGETS = [
  // ── ใส่ Spreadsheet ID จริงแทนตรงนี้ — label ใช้ตั้งชื่อไฟล์ backup เท่านั้น
  //    ไม่กระทบการทำงาน ตั้งให้จำง่ายพอ ──
  { id: '1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo', label: 'Stock_Database' },
  { id: '12rlj7SR-Xofj8tdyu3atA9kUTLC44Y6Ja9cNfd2PUvw', label: 'STOCK_PRICE_DATABAS' },
  // { id: 'PASTE_SPREADSHEET_ID_HERE_3', label: 'Watchlist' },
];

const BACKUP_FOLDER_NAME = 'Stock Manager Backups'; // สร้างอัตโนมัติถ้ายังไม่มีในไฟล์ My Drive
const BACKUP_RETENTION_DAYS = 30; // เก็บย้อนหลัง 30 วัน เกินกว่านี้ลบทิ้ง (ย้ายลงถังขยะ ไม่ลบถาวร)

/**
 * MAIN — รันทุกวันผ่าน Trigger (setupDailyBackupTrigger ตั้งให้อัตโนมัติ)
 * สำรองทุก target ที่ตั้งไว้ + เก็บกวาดไฟล์เก่าเกิน retention
 */

/**
 * หา/สร้างโฟลเดอร์ปลายทางใน My Drive ของบัญชีที่รันสคริปต์
 */
function _getOrCreateBackupFolder() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * ลบไฟล์ backup ที่เก่าเกิน BACKUP_RETENTION_DAYS (ย้ายลงถังขยะ ไม่ลบถาวร
 * กันเผลอลบของที่ยังต้องใช้ — กู้คืนได้ใน Google Drive Trash ภายใน 30 วัน)
 * คืนจำนวนไฟล์ที่ลบ
 */
function _cleanupOldBackups(folder) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_RETENTION_DAYS);

  let deleted = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() < cutoff) {
      f.setTrashed(true);
      deleted++;
    }
  }
  return deleted;
}

/**
 * แจ้งเตือนผ่าน Telegram ถ้าโปรเจกต์นี้มีฟังก์ชัน sendTelegram() อยู่แล้ว
 * (เผื่อรันคนละโปรเจกต์กับที่มี sendTelegram — ไม่ error ถ้าไม่มี แค่ข้าม)
 */


/**
 * ตั้ง Trigger รันสำรองอัตโนมัติทุกวัน — รันฟังก์ชันนี้ครั้งเดียวตอนติดตั้ง
 * เวลา 03:30 (เลี่ยงชนกับงานประจำวันอื่นที่มักตั้งไว้ตี 2)
 */
function setupDailyBackupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupAllSheets')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('backupAllSheets')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .nearMinute(30)
    .create();

  Logger.log('setupDailyBackupTrigger: ตั้ง Trigger รัน backupAllSheets() ทุกวัน ~03:30 แล้ว');
  try {
    SpreadsheetApp.getUi().alert('✅ ตั้ง Trigger สำรองข้อมูลรายวันแล้ว (รันทุกวัน ~03:30)');
  } catch (e) { /* รันจาก Editor ไม่มี UI ก็ไม่เป็นไร */ }
}

/**
 * ทดสอบสำรองด้วยมือ — รันใน Apps Script Editor แล้วดู Logger + เช็คโฟลเดอร์
 * "Stock Manager Backups" ใน Google Drive ว่ามีไฟล์ใหม่จริงไหม
 */
function testBackupNow() {
  backupAllSheets();
}


function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('💾 สำรองข้อมูล')
    .addItem('สำรองตอนนี้', 'manualBackupWithAlert')
    .addToUi();
}

function manualBackupWithAlert(){
  backupAllSheets();
  SpreadsheetApp.getUi().alert('✅ สำรองข้อมูลเสร็จแล้ว\n\nไฟล์ถูกเก็บไว้ในโฟลเดอร์ Drive: ' + BACKUP_FOLDER_NAME);
}










