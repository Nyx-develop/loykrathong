// ลอยกระทงออนไลน์ · เก็บคำอธิษฐาน (สร้าง Google Sheet ให้อัตโนมัติใน Drive ของคุณ)
const TYPES = {leaf: 'ใบตอง', bread: 'ขนมปัง', hyacinth: 'ผักตบชวา', coconut: 'กะลา'};

function sheet_(create) {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    if (!create) return null;
    ss = SpreadsheetApp.create('ลอยกระทง · คำอธิษฐาน');
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['เวลา', 'ชื่อ', 'คำอธิษฐาน', 'กระทง']);
  }
  return ss.getSheets()[0];
}

// กด ▶ Run ฟังก์ชันนี้ครั้งเดียวก็ได้ จะได้ลิงก์ชีตใน Execution log
function setup() { Logger.log(sheet_(true).getParent().getUrl()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    const clean = (v, n) => String(v || '').trim().slice(0, n).replace(/^[=+\-@]/, "'$&");
    const name = clean(d.name, 24), wish = clean(d.wish, 120);
    if (!name) return ContentService.createTextOutput('empty');
    sheet_(true).appendRow([new Date(), name, wish, TYPES[d.type] || '']);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// ส่งให้หน้าเว็บแค่ชื่อ + แบบกระทง 60 ใบล่าสุด (คำอธิษฐานไม่ถูกส่งออกไป)
function doGet() {
  const sh = sheet_(false);
  const total = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  let items = [];
  if (total > 0) {
    const n = Math.min(60, total);
    const rev = {};
    Object.keys(TYPES).forEach(k => rev[TYPES[k]] = k);
    items = sh.getRange(sh.getLastRow() - n + 1, 2, n, 3).getValues().reverse()
      .map(r => ({name: String(r[0]).replace(/^'/, ''), type: rev[r[2]] || 'leaf'}));
  }
  return ContentService.createTextOutput(JSON.stringify({total: total, items: items}))
    .setMimeType(ContentService.MimeType.JSON);
}
