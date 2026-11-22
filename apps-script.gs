// ===== ลอยกระทงออนไลน์ · ตัวเก็บคำอธิษฐานลง Google Sheet =====
// 1) เปิด Google Sheet ใหม่ > ส่วนขยาย (Extensions) > Apps Script
// 2) ลบโค้ดเดิม วางโค้ดนี้ทั้งหมด กดบันทึก
// 3) Deploy > New deployment > ประเภท Web app
//      Execute as: Me   |   Who has access: Anyone
// 4) คัดลอก URL ที่ลงท้ายด้วย /exec ไปใส่ SAVE_URL ใน index.html
//
// คำอธิษฐานอ่านได้เฉพาะเจ้าของชีต — หน้าเว็บดึงไปแค่ "ชื่อ" กับ "แบบกระทง" เพื่อให้ลอยวนบนจอ

const SHEET_NAME = 'คำอธิษฐาน';
const TYPES = {leaf: 'ใบตอง', bread: 'ขนมปัง', hyacinth: 'ผักตบชวา', coconut: 'กะลา'};

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['เวลา', 'ชื่อ', 'คำอธิษฐาน', 'กระทง']);
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    // ตัดความยาว + กันสูตรที่ขึ้นต้นด้วย = + - @
    const clean = (v, n) => String(v || '').trim().slice(0, n).replace(/^[=+\-@]/, "'$&");
    const name = clean(d.name, 24), wish = clean(d.wish, 120);
    if (!name) return ContentService.createTextOutput('empty');
    sheet_().appendRow([new Date(), name, wish, TYPES[d.type] || '']);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// ส่งเฉพาะชื่อ + แบบกระทง 60 ใบล่าสุด (ไม่ส่งคำอธิษฐาน)
function doGet() {
  const sh = sheet_();
  const total = Math.max(0, sh.getLastRow() - 1);
  let items = [];
  if (total > 0) {
    const n = Math.min(60, total);
    const rows = sh.getRange(sh.getLastRow() - n + 1, 2, n, 3).getValues(); // ชื่อ, คำอธิษฐาน, กระทง
    const rev = {};
    Object.keys(TYPES).forEach(k => rev[TYPES[k]] = k);
    items = rows.reverse().map(r => ({name: String(r[0]).replace(/^'/, ''), type: rev[r[2]] || 'leaf'}));
  }
  return ContentService.createTextOutput(JSON.stringify({total: total, items: items}))
    .setMimeType(ContentService.MimeType.JSON);
}
