// เอามาจาก URL: https://docs.google.com/spreadsheets/d/<ตรงนี้คือ ID>/edit
const SPREADSHEET_ID = '';

// เอามาจาก URL: https://drive.google.com/drive/folders/<ตรงนี้คือ ID>
const IMAGE_FOLDER_ID = '';

const SHEET_MENU = 'Menu';
const SHEET_ORDERS = 'Orders';
const SHEET_USERS = 'Users';

function initAllsheet() {
  const ss = getSS();

  let menuSheet = ss.getSheetByName(SHEET_MENU);
  if (!menuSheet) {
    menuSheet = ss.insertSheet(SHEET_MENU);
    // เพิ่มคอลัมน์ DriveId
    menuSheet.appendRow(['ID', 'หมวดหมู่', 'ชื่อเมนู', 'ราคา', 'เปิดขาย', 'DriveId']);
    menuSheet.setFrozenRows(1);
  } else {
    // กรณีมีชีตเดิม ตรวจสอบว่ามีหัวคอลัมน์ DriveId หรือยัง
    const headers = menuSheet.getRange(1, 1, 1, menuSheet.getLastColumn()).getValues()[0];
    if (headers.length < 6) {
      menuSheet.getRange(1, 6).setValue('DriveId');
    }
  }

  let orderSheet = ss.getSheetByName(SHEET_ORDERS);
  if (!orderSheet) {
    orderSheet = ss.insertSheet(SHEET_ORDERS);
    orderSheet.appendRow(['เลขที่ออเดอร์', 'วันที่เวลา', 'โต๊ะ/ลูกค้า', 'รายการ (JSON)', 'ยอดรวม', 'ชำระโดย', 'สถานะ']);
    orderSheet.setFrozenRows(1);
  }

  let userSheet = ss.getSheetByName(SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USERS);
    userSheet.appendRow(['Username', 'Password', 'Role']);
    userSheet.appendRow(['admin', '1234', 'Admin']);
    userSheet.setFrozenRows(1);
  }

  return 'ตั้งค่าชีตเรียบร้อย';
}

// หน้าเว็บถูกโฮสต์แยกต่างหาก (index.html) สคริปต์นี้ทำหน้าที่เป็น API ผ่าน doPost เท่านั้น
function doGet() {
  initAllsheetOnce();
  return jsonOutput({ result: 'POS API พร้อมใช้งาน' });
}

// รายชื่อฟังก์ชันที่เปิดให้ api.js เรียกได้ (whitelist — กันการเรียกฟังก์ชันอื่นจากภายนอก)
const API_ACTIONS = {
  checkLogin: checkLogin,
  getMenu: getMenu,
  getAllMenuForAdmin: getAllMenuForAdmin,
  addMenuItem: addMenuItem,
  updateMenuItem: updateMenuItem,
  deleteMenuItem: deleteMenuItem,
  submitOrder: submitOrder,
  getOrdersByDate: getOrdersByDate
};

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API_ACTIONS[req.action];
    if (!fn) return jsonOutput({ error: 'ไม่รู้จัก action: ' + req.action });

    initAllsheetOnce();
    return jsonOutput({ result: fn.apply(null, req.args || []) });
  } catch (err) {
    return jsonOutput({ error: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// openById เป็นคำสั่งที่ช้าที่สุดในไฟล์นี้ คำขอหนึ่งครั้งเปิดครั้งเดียวพอ
let SS = null;

function getSS() {
  if (!SS) SS = SpreadsheetApp.openById(SPREADSHEET_ID);
  return SS;
}

const PROP_SHEETS_READY = 'SHEETS_READY';

// PropertiesService เป็นการเรียกข้ามเครือข่าย ถ้าถามทุกคำขอจะบวกเวลาให้ทุกปุ่มรวมถึงตอนล็อกอิน
// Apps Script ใช้ instance เดิมซ้ำกับคำขอที่มาติด ๆ กัน จำใส่ตัวแปรไว้จึงข้ามการถามไปได้เลย
let SHEETS_READY = false;

function initAllsheetOnce() {
  if (SHEETS_READY) return;

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_SHEETS_READY) !== '1') {
    initAllsheet();
    props.setProperty(PROP_SHEETS_READY, '1');
  }
  SHEETS_READY = true;
}

function resetInitFlag() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SHEETS_READY);
  SHEETS_READY = false;
  return 'ล้าง flag แล้ว ครั้งต่อไปที่เปิดหน้าเว็บจะ init ชีตใหม่';
}

/* รันฟังก์ชันนี้ในตัวแก้ไข Apps Script แล้วดูที่บันทึกการดำเนินการ
   เพื่อบอกว่า openById ล้มเหลวเพราะ ID ผิด หรือเพราะรันอยู่คนละบัญชีกับเจ้าของไฟล์ */
function checkConfig() {
  const lines = ['บัญชีที่รันสคริปต์: ' + (Session.getEffectiveUser().getEmail() || '(ไม่ทราบ)')];

  lines.push('SPREADSHEET_ID: ' + SPREADSHEET_ID + ' (' + SPREADSHEET_ID.length + ' ตัวอักษร)');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    lines.push('  เปิดได้: "' + ss.getName() + '"');
    lines.push('  ชีตที่มี: ' + ss.getSheets().map(s => s.getName()).join(', '));
  } catch (err) {
    lines.push('  เปิดไม่ได้: ' + err.message);
    lines.push('  → ตรวจว่า ID ตรงกับ URL ของไฟล์ และไฟล์อยู่ในบัญชีข้างบนหรือถูกแชร์ให้บัญชีนั้นแล้ว');
  }

  lines.push('IMAGE_FOLDER_ID: ' + IMAGE_FOLDER_ID + ' (' + IMAGE_FOLDER_ID.length + ' ตัวอักษร)');
  try {
    lines.push('  เปิดได้: "' + DriveApp.getFolderById(IMAGE_FOLDER_ID).getName() + '"');
  } catch (err) {
    lines.push('  เปิดไม่ได้: ' + err.message);
  }

  const report = lines.join('\n');
  Logger.log(report);
  return report;
}

function checkLogin(username, password) {
  const sheet = getSS().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: 'ไม่พบชีต Users กรุณารัน initAllsheet' };

  // กันการล็อกอินด้วยค่าว่างไปตรงกับแถวว่างที่ค้างอยู่ในชีต
  if (String(username || '').trim() === '' || String(password || '') === '') {
    return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
  }

  const data = sheet.getDataRange().getValues();
  data.shift();

  const user = data.find(r => String(r[0]).trim() === String(username).trim() && String(r[1]) === String(password));

  if (user) {
    return { success: true, username: user[0], role: user[2] || 'Staff' };
  } else {
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
}

// Helper: ดึง File ID ออกจากลิงก์ Drive ทุกรูปแบบ (หรือรับ ID ดิบก็ได้)
function extractDriveId(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const m = s.match(/\/d\/([-\w]{20,})/) || s.match(/[?&]id=([-\w]{20,})/);
  if (m) return m[1];
  return /^[-\w]{20,}$/.test(s) ? s : '';
}

// Helper: ลิงก์ Drive แบบปกติสำหรับเก็บลงชีต
function getDriveLink(fileId) {
  return fileId ? 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing' : '';
}

// Helper: แปลงค่าในคอลัมน์ DriveId เป็น URL สำหรับแสดงภาพ
function getImageUrl(value) {
  const id = extractDriveId(value);
  return id ? 'https://lh3.googleusercontent.com/d/' + id : '';
}

// คอลัมน์ "เปิดขาย" อาจเป็นช่องติ๊ก (boolean) หรือถูกพิมพ์เป็นข้อความ TRUE / ใช่ / 1
// ถ้าเทียบ === true อย่างเดียว เมนูที่กรอกมือจะหายไปจากหน้าขายทั้งหมด
function toBool(v) {
  if (v === true || v === 1) return true;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y'
    || s === 'ใช่' || s === 'เปิด' || s === 'เปิดขาย';
}

// ราคาที่พิมพ์มาพร้อม "฿" หรือคอมมา (เช่น "1,250") ต้องอ่านเป็นตัวเลขให้ได้ ไม่ใช่ NaN
function toNumber(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// แถวที่ยังไม่มี ID แต่มีชื่อเมนู ควรถูกอ่านด้วย ไม่ใช่ถูกตัดทิ้งเงียบ ๆ
function isMenuRow(r) {
  return String(r[0] || '').trim() !== '' || String(r[2] || '').trim() !== '';
}

// ---------- เมนู ----------

function getMenu() {
  const sheet = getSS().getSheetByName(SHEET_MENU);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data
    .filter(r => isMenuRow(r) && toBool(r[4]))
    .map(r => ({
      id: r[0],
      category: r[1],
      name: r[2],
      price: toNumber(r[3]),
      driveId: r[5] || '',
      fileId: extractDriveId(r[5]),
      imageUrl: getImageUrl(r[5])
    }));
}

function getAllMenuForAdmin() {
  const sheet = getSS().getSheetByName(SHEET_MENU);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data
    .filter(isMenuRow)
    .map(r => ({
      id: r[0],
      category: r[1],
      name: r[2],
      price: toNumber(r[3]),
      active: toBool(r[4]),
      driveId: r[5] || '',
      fileId: extractDriveId(r[5]),
      imageUrl: getImageUrl(r[5])
    }));
}

/**
 * ฟังก์ชันรับไฟล์ base64 แล้วอัปโหลดไปที่ Google Drive
 * คืนค่าลิงก์ Drive แบบปกติเพื่อนำไปบันทึกลงใน Sheet
 */
function uploadImageToDrive(base64Data, filename) {
  try {
    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, filename);
    
    // ยังไม่ได้ตั้ง IMAGE_FOLDER_ID ก็ยังอัปโหลดได้ แต่ไฟล์จะไปกองอยู่ที่ My Drive
    const file = IMAGE_FOLDER_ID
      ? DriveApp.getFolderById(IMAGE_FOLDER_ID).createFile(blob)
      : DriveApp.createFile(blob);

    // นโยบายบัญชี/องค์กรอาจบล็อกการแชร์สาธารณะ ถ้าตั้งไม่ได้ก็ปล่อยผ่าน
    // ไฟล์ยังอยู่ใน Drive และบันทึกลงชีตได้ (ให้ไปตั้งสิทธิ์ที่โฟลเดอร์แทน)
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e2) { /* ปล่อยให้สืบทอดสิทธิ์จากโฟลเดอร์ */ }
    }

    return getDriveLink(file.getId());
  } catch (err) {
    throw new Error('การอัปโหลดรูปภาพล้มเหลว: ' + err.message);
  }
}

function addMenuItem(item) {
  // รับได้ทั้งลิงก์ Drive และ ID ดิบ แล้วเก็บลงชีตเป็นลิงก์เสมอ
  let driveLink = getDriveLink(extractDriveId(item.driveId));

  // หากมีการแนบรูปภาพใหม่เข้ามาแบบ base64
  if (item.imageBase64) {
    driveLink = uploadImageToDrive(item.imageBase64, 'menu_' + Date.now());
  }

  // สองเครื่องกดเพิ่มเมนูพร้อมกันจะอ่าน maxId ตัวเดียวกันแล้วได้ ID ซ้ำ
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSS().getSheetByName(SHEET_MENU);
    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const n = Number(data[i][0]);
      if (n > maxId) maxId = n;
    }
    const newId = maxId + 1;
    sheet.appendRow([newId, item.category, item.name, toNumber(item.price), true, driveLink]);
    return { success: true, id: newId };
  } finally {
    lock.releaseLock();
  }
}

function updateMenuItem(item) {
  let driveLink = getDriveLink(extractDriveId(item.driveId));

  // หากมีการเลือกรูปใหม่เข้ามา
  if (item.imageBase64) {
    driveLink = uploadImageToDrive(item.imageBase64, 'menu_' + item.id + '_' + Date.now());
  }

  const sheet = getSS().getSheetByName(SHEET_MENU);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(item.id)) {
      sheet.getRange(i + 1, 2, 1, 5).setValues([[item.category, item.name, toNumber(item.price), toBool(item.active), driveLink]]);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบเมนูนี้' };
}

function deleteMenuItem(id) {
  const sheet = getSS().getSheetByName(SHEET_MENU);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบเมนูนี้' };
}

// ---------- ออเดอร์ ----------

function submitOrder(order) {
  if (!order || !order.items || order.items.length === 0) {
    return { success: false, message: 'ไม่มีรายการอาหารในออเดอร์' };
  }
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const timeStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss');

  let total = 0;
  order.items.forEach(it => { total += toNumber(it.price) * toNumber(it.qty); });

  // สองบิลที่ปิดในวินาทีเดียวกันจะได้เลขที่บิลซ้ำ ต่อท้ายด้วยมิลลิวินาทีกันชน
  const orderId = 'OD' + Utilities.formatDate(now, tz, 'yyMMdd-HHmmss')
    + '-' + ('00' + (now.getTime() % 1000)).slice(-3);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSS().getSheetByName(SHEET_ORDERS).appendRow([
      orderId,
      timeStr,
      order.tableName || '-',
      JSON.stringify(order.items),
      total,
      order.payment || 'เงินสด',
      'สำเร็จ'
    ]);
  } finally {
    lock.releaseLock();
  }

  return { success: true, orderId: orderId, time: timeStr, total: total };
}

// date เป็น 'yyyy-MM-dd' ตามที่ <input type="date"> ส่งมา ถ้าไม่ส่งมาจะถือว่าเป็นวันนี้
function getOrdersByDate(date) {
  const sheet = getSS().getSheetByName(SHEET_ORDERS);
  if (!sheet) return { orders: [], count: 0, totalSales: 0 };
  const data = sheet.getDataRange().getValues();
  data.shift();
  const tz = Session.getScriptTimeZone();
  const ymd = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(date || ''));
  const dayStr = ymd
    ? ymd[3] + '/' + ymd[2] + '/' + ymd[1]
    : Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

  // ชีตอาจแปลงข้อความวันที่ที่บันทึกไปเป็นชนิด Date เอง จึงต้องรองรับทั้งสองแบบ
  const timeStr = v => v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm:ss') : String(v);

  let totalSales = 0;
  const orders = data
    .filter(r => String(r[0] || '').trim() !== '' && timeStr(r[1]).indexOf(dayStr) === 0)
    .map(r => {
      totalSales += toNumber(r[4]);
      let items = [];
      try { items = JSON.parse(r[3] || '[]'); } catch (e) { items = []; }
      return {
        orderId: r[0],
        time: timeStr(r[1]),
        table: r[2],
        items: items,
        total: toNumber(r[4]),
        payment: r[5],
        status: r[6]
      };
    })
    .reverse();

  return { orders: orders, count: orders.length, totalSales: totalSales };
}