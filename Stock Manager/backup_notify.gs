// ============================================================
// backup_notify.gs
// แจ้งเตือนผลการ Backup (สำเร็จ / ล้มเหลว) ผ่าน Telegram
// ------------------------------------------------------------
// ⚠️ ไฟล์นี้ "แทนที่" ฟังก์ชัน backupAllSheets() และ _backupNotify() เดิม
// ที่อยู่ใน webapp_00c_helpers.gs
//
// สาเหตุที่ต้องแก้: ฟังก์ชัน _backupNotify() เดิมเรียก sendTelegram()
// ซึ่งไม่มีฟังก์ชันนี้อยู่จริงในโปรเจกต์ (มีแต่ sendTelegramSafe() /
// sendTelegramError()) เลยไม่เคยส่งข้อความ Telegram ได้จริงเวลา backup
// เสร็จ — ไฟล์นี้แก้ให้ใช้ฟังก์ชันที่มีอยู่แล้วแทน และแยกแจ้งผลสำเร็จ/
// ล้มเหลวเป็นรายไฟล์ให้ชัดเจนขึ้น
//
// วิธีติดตั้ง:
//  1. เปิด webapp_00c_helpers.gs แล้ว "ลบ" ฟังก์ชัน backupAllSheets()
//     กับ _backupNotify() เดิมทิ้ง (ชื่อฟังก์ชันซ้ำกันจะรันไม่ได้)
//     — ค่าคงที่ BACKUP_TARGETS, BACKUP_FOLDER_NAME, BACKUP_RETENTION_DAYS
//     และฟังก์ชัน _getOrCreateBackupFolder(), _cleanupOldBackups(),
//     setupDailyBackupTrigger(), testBackupNow() "เก็บไว้เหมือนเดิม"
//  2. เพิ่มไฟล์นี้ (backup_notify.gs) เข้าไปในโปรเจกต์
//  3. ต้องเคยรัน setupProperties() ตั้งค่า BOT_TOKEN / CHAT_ID ไว้แล้ว
//  4. รัน testBackupNow() 1 ครั้งเพื่อทดสอบว่าได้รับข้อความ Telegram
// ============================================================

function backupAllSheets() {
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  try {
    const invalid = BACKUP_TARGETS.filter(t => !t.id || t.id.indexOf('PASTE_') === 0);
    if (invalid.length === BACKUP_TARGETS.length) {
      Logger.log('backupAllSheets: ยังไม่ได้ใส่ Spreadsheet ID จริงใน BACKUP_TARGETS เลย — ข้าม');
      sendTelegramSafe(
        '⚠️ Backup ' + todayStr + '\nยังไม่ได้ตั้งค่า BACKUP_TARGETS — ข้ามการสำรองข้อมูล'
      );
      return;
    }

    const folder = _getOrCreateBackupFolder();
    const results = [];

    BACKUP_TARGETS.forEach(target => {
      if (!target.id || target.id.indexOf('PASTE_') === 0) {
        results.push({ label: target.label, status: 'skipped_no_id' });
        return;
      }
      try {
        const backupName = target.label + ' - Backup ' + todayStr;

        // กันสำรองซ้ำถ้ารันมากกว่า 1 ครั้งในวันเดียวกัน
        const already = folder.getFilesByName(backupName);
        if (already.hasNext()) {
          Logger.log('backupAllSheets: "' + backupName + '" มีอยู่แล้ว — ข้าม');
          results.push({ label: target.label, status: 'skipped_duplicate' });
          return;
        }

        const sourceFile = DriveApp.getFileById(target.id);
        sourceFile.makeCopy(backupName, folder);
        Logger.log('backupAllSheets: สำรอง "' + target.label + '" สำเร็จ → ' + backupName);
        results.push({ label: target.label, status: 'ok' });
      } catch (e) {
        Logger.log('backupAllSheets: สำรอง "' + target.label + '" ล้มเหลว — ' + e.message);
        results.push({ label: target.label, status: 'error', error: e.message });
      }
    });

    const cleaned = _cleanupOldBackups(folder);
    _backupNotify(results, cleaned, todayStr);

  } catch (e) {
    // ครอบ error ระดับฟังก์ชันทั้งหมด (เช่น ไม่มีสิทธิ์เข้าถึง Drive)
    // แจ้งเตือนทันทีว่า backup ล้มเหลวทั้งกระบวนการ
    sendTelegramError('backupAllSheets', e);
  }
}

/**
 * สร้างและส่งข้อความสรุปผล backup รายไฟล์ (สำเร็จ / ล้มเหลว / ข้าม)
 * ผ่าน sendTelegramSafe() ซึ่งมีอยู่แล้วในโปรเจกต์ (helper.gs)
 */
function _backupNotify(results, cleaned, todayStr) {
  const okList   = results.filter(r => r.status === 'ok');
  const errList  = results.filter(r => r.status === 'error');
  const skipList = results.filter(r => r.status.indexOf('skipped') === 0);

  const allFailed   = results.length > 0 && okList.length === 0 && errList.length > 0;
  const headerIcon  = allFailed ? '🚨' : (errList.length ? '⚠️' : '✅');

  let msg = headerIcon + ' Backup ประจำวัน ' + todayStr + '\n';
  msg += '━━━━━━━━━━━━\n';

  if (okList.length) {
    msg += '✅ สำเร็จ (' + okList.length + '):\n';
    okList.forEach(r => { msg += '  • ' + r.label + '\n'; });
  }

  if (errList.length) {
    msg += '❌ ล้มเหลว (' + errList.length + '):\n';
    errList.forEach(r => { msg += '  • ' + r.label + ' — ' + r.error + '\n'; });
  }

  if (skipList.length) {
    msg += '⏭ ข้าม (' + skipList.length + '):\n';
    skipList.forEach(r => {
      const reason = r.status === 'skipped_duplicate'
        ? 'สำรองไปแล้ววันนี้'
        : 'ยังไม่ได้ใส่ Spreadsheet ID';
      msg += '  • ' + r.label + ' — ' + reason + '\n';
    });
  }

  if (cleaned) {
    msg += '🗑 ลบไฟล์เก่าเกิน ' + BACKUP_RETENTION_DAYS + ' วัน: ' + cleaned + ' ไฟล์\n';
  }

  msg += '🕐 ' + getNow();

  Logger.log(msg);
  sendTelegramSafe(msg);
}

