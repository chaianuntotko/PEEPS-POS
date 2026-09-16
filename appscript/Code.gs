// เอามาจาก URL: https://docs.google.com/spreadsheets/d/<ตรงนี้คือ ID>/edit
const SPREADSHEET_ID = '1A1ICU3qc6RNmhSeyhw1FcDnE_fwAZfyxKoHnhHGTEfA';

// เอามาจาก URL: https://drive.google.com/drive/folders/<ตรงนี้คือ ID>
const IMAGE_FOLDER_ID = '1nu2b0TAViaqJnk0UtATuzjiqNxmW1GzC';

const SHEET_MENU = 'Menu';
const SHEET_ORDERS = 'Orders';
const SHEET_USERS = 'Users';

function initAllsheet() {
  const ss = getSS();

  let menuSheet = ss.getSheetByName(SHEET_MENU);
  if (!menuSheet) {
    menuSheet = ss.insertSheet(SHEET_MENU);
    // เพิ่มคอลัมน์ DriveId, โปรโมชัน 3 ชิ้น 100 และสต๊อก
    menuSheet.appendRow(['ID', 'หมวดหมู่', 'ชื่อเมนู', 'ราคา', 'เปิดขาย', 'DriveId', 'โปรโมชัน 3 ชิ้น 100', 'สต๊อก']);
    menuSheet.setFrozenRows(1);
  } else {
    // กรณีมีชีตเดิม ตรวจสอบว่ามีหัวคอลัมน์ DriveId / โปรโมชัน / สต๊อก หรือยัง
    const headers = menuSheet.getRange(1, 1, 1, menuSheet.getLastColumn()).getValues()[0];
    if (headers.length < 6) {
      menuSheet.getRange(1, 6).setValue('DriveId');
    }
    if (headers.length < 7) {
      menuSheet.getRange(1, 7).setValue('โปรโมชัน 3 ชิ้น 100');
    }
    if (headers.length < 8) {
      menuSheet.getRange(1, 8).setValue('สต๊อก');
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

/* รันฟังก์ชันนี้ในตัวแก้ไข Apps Script แล้วดูที่บันทึกการดำเนินการ (Ctrl+Enter หรือ View > Logs)
   ใช้เช็คตอนที่กดบันทึกเมนูจากหน้าเว็บแล้ว "สต๊อก" หรือ "โปรโมชัน" ไม่ขึ้นในชีตจริง
   ถ้าเลขคอลัมน์ของ โปรโมชัน/สต๊อก ที่รายงานออกมาไม่ใช่ 7/8 ให้รัน initAllsheet() ก่อน
   ถ้าหัวคอลัมน์ถูกต้องแต่ยังไม่บันทึก แปลว่าโค้ดที่รันอยู่จริงใน Apps Script Editor (ที่ผูกกับ URL /exec)
   เป็นคนละเวอร์ชันกับไฟล์นี้ — ต้องคัดลอกไฟล์นี้ไปวางทับใน Editor แล้ว Deploy > Manage deployments
   > แก้ deployment เดิม > เลือก Version: New version ใหม่อีกครั้ง (แค่กด Save ในตัวแก้ไขไม่พอ
   เพราะ URL /exec จะยังชี้ไปที่เวอร์ชันเก่าที่ deploy ไว้ก่อนหน้าจนกว่าจะ deploy เวอร์ชันใหม่) */
function checkMenuSheet() {
  const sheet = getSS().getSheetByName(SHEET_MENU);
  if (!sheet) return 'ไม่พบชีต ' + SHEET_MENU + ' — รัน initAllsheet() ก่อน';

  const lines = [];
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  lines.push('จำนวนคอลัมน์ที่มีหัวตาราง: ' + lastCol);
  headers.forEach((h, i) => lines.push('  คอลัมน์ ' + (i + 1) + ': "' + h + '"'));

  const promoCol = headers.indexOf('โปรโมชัน 3 ชิ้น 100') + 1;
  const stockCol = headers.indexOf('สต๊อก') + 1;
  lines.push('คอลัมน์ "โปรโมชัน 3 ชิ้น 100" อยู่ที่: ' + (promoCol || 'ไม่พบ (ควรเป็น 7)'));
  lines.push('คอลัมน์ "สต๊อก" อยู่ที่: ' + (stockCol || 'ไม่พบ (ควรเป็น 8)'));

  if (lastRow > 1) {
    const lastRowData = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
    lines.push('แถวข้อมูลล่าสุด (แถวที่ ' + lastRow + '): ' + JSON.stringify(lastRowData));
  } else {
    lines.push('ยังไม่มีแถวข้อมูลเมนูเลย');
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

// ช่องสต๊อกว่าง = ไม่ได้ติดตามสต๊อก (ขายได้ไม่จำกัด) ต้องแยกจาก 0 = ของหมด
function toStockOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  return toNumber(v);
}

// ค่าที่จะเขียนกลับลงชีต: ไม่ได้ติดตามสต๊อกให้เก็บเป็นค่าว่าง ไม่ใช่ 0
function stockCell(v) {
  return (v === null || v === undefined || v === '') ? '' : toNumber(v);
}

// ---------- เมนู ----------

// ใช้ร่วมกันทั้งตอนอ่านทั้งชีต (getAllMenuForAdmin) และตอนคืนค่าแถวเดียวหลังเพิ่ม/แก้ไข
// (addMenuItem/updateMenuItem) เพื่อให้ฝั่งเว็บอัปเดตแถวนั้นในแคชได้เลยโดยไม่ต้องขอทั้งชีตซ้ำ
function buildAdminItem(id, category, name, price, active, driveId, promo, stock) {
  return {
    id: id,
    category: category,
    name: name,
    price: price,
    active: active,
    promo: promo,
    stock: toStockOrNull(stock),
    driveId: driveId || '',
    fileId: extractDriveId(driveId),
    imageUrl: getImageUrl(driveId)
  };
}

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
      promo: toBool(r[6]),
      stock: toStockOrNull(r[7]),
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
    .map(r => buildAdminItem(r[0], r[1], r[2], toNumber(r[3]), toBool(r[4]), r[5], toBool(r[6]), r[7]));
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
    const priceVal = toNumber(item.price);
    const promoVal = !!item.promo;
    const stockVal = stockCell(item.stock);
    sheet.appendRow([newId, item.category, item.name, priceVal, true, driveLink, promoVal, stockVal]);
    // ส่งแถวที่เพิ่งบันทึกกลับไปเลย ฝั่งเว็บจะได้อัปเดตแคชในตัวได้ทันที
    // ไม่ต้องขอ getAllMenuForAdmin() ทั้งชีตซ้ำอีกรอบ (ลดเวลารอบันทึกเมนูลงครึ่งหนึ่ง)
    return { success: true, id: newId, item: buildAdminItem(newId, item.category, item.name, priceVal, true, driveLink, promoVal, stockVal) };
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
      // สต๊อกอาจถูกหักไปแล้วจากการขายหลังจากเบราว์เซอร์โหลดเมนูมาแคชไว้ ถ้าคำขอนี้ไม่ได้ตั้งใจแก้ไข
      // ค่าสต๊อกจริง ๆ (เช่น กดแค่เปิด/ปิดขาย ไม่ได้ส่ง stock มา) ต้องไม่เขียนทับด้วยค่าเก่าที่ค้างอยู่ในแคช
      const stockVal = item.stock === undefined ? data[i][7] : stockCell(item.stock);
      const priceVal = toNumber(item.price);
      const activeVal = toBool(item.active);
      const promoVal = !!item.promo;
      sheet.getRange(i + 1, 2, 1, 7).setValues([[item.category, item.name, priceVal, activeVal, driveLink, promoVal, stockVal]]);
      // ส่งแถวที่เพิ่งบันทึกกลับไปเลย ฝั่งเว็บจะได้อัปเดตแคชในตัวได้ทันที ไม่ต้องขอทั้งชีตซ้ำ
      return { success: true, item: buildAdminItem(item.id, item.category, item.name, priceVal, activeVal, driveLink, promoVal, stockVal) };
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

// รวมจำนวนที่ต้องการหักต่อ id เมนู — เมนูปกติหักด้วย id ตรง ๆ ส่วนโปรโมชัน 3 ชิ้น 100
// ที่มัดรวมเป็นบรรทัดเดียวจะแนบ picks (id + จำนวนต่อชุด) มาให้ทีละตัว คูณด้วยจำนวนชุดที่ซื้อ
function neededStockById(items) {
  const need = {};
  (items || []).forEach(it => {
    if (it.picks && it.picks.length) {
      it.picks.forEach(p => {
        const id = String(p.id);
        need[id] = (need[id] || 0) + toNumber(p.qty) * toNumber(it.qty || 1);
      });
    } else if (it.id !== undefined && it.id !== null && it.id !== '') {
      const id = String(it.id);
      need[id] = (need[id] || 0) + toNumber(it.qty);
    }
  });
  return need;
}

// เช็คว่าสต๊อกพอก่อนบันทึกบิลจริง — คืนชื่อ+จำนวนของรายการแรกที่ไม่พอ หรือ null ถ้าพอทุกอย่าง
// ช่องสต๊อกว่าง (ไม่ได้ติดตาม) ถือว่าขายได้ไม่จำกัด ข้ามการเช็ค
function findStockShortage(items, data) {
  const need = neededStockById(items);
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    if (!(id in need)) continue;
    const cur = data[i][7];
    if (cur === '' || cur === null || cur === undefined) continue;
    if (toNumber(cur) < need[id]) {
      return { name: data[i][2], have: toNumber(cur), need: need[id] };
    }
  }
  return null;
}

// หักสต๊อกจริงตามที่เช็คผ่านแล้ว — เรียกหลัง findStockShortage คืน null เท่านั้น
// คืนค่าสต๊อกใหม่ของแต่ละ id ที่ถูกหัก ให้ฝั่งเว็บเอาไปอัปเดตตัวเลข "เหลือ" บนการ์ดได้ทันที
// โดยไม่ต้องขอเมนูทั้งชีตใหม่ทั้งหมดหลังชำระเงินทุกครั้ง
function deductStock(items, sheet, data) {
  const rowById = {};
  for (let i = 1; i < data.length; i++) rowById[String(data[i][0])] = i;

  const need = neededStockById(items);
  const updates = {};
  Object.keys(need).forEach(id => {
    const i = rowById[id];
    if (i == null) return;
    const cur = data[i][7];
    if (cur === '' || cur === null || cur === undefined) return;
    const newVal = toNumber(cur) - need[id];
    sheet.getRange(i + 1, 8).setValue(newVal);
    updates[id] = newVal;
  });
  return updates;
}

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

  let stockUpdates = {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const menuSheet = getSS().getSheetByName(SHEET_MENU);
    const menuData = menuSheet ? menuSheet.getDataRange().getValues() : [];

    // เช็คสต๊อกในล็อกเดียวกับการบันทึกบิล กันสองบิลพร้อมกันอ่านสต๊อกเดิมซ้ำแล้วขายเกินของจริง
    const shortage = menuSheet ? findStockShortage(order.items, menuData) : null;
    if (shortage) {
      return {
        success: false,
        message: '"' + shortage.name + '" เหลือไม่พอ (เหลือ ' + shortage.have + ', ต้องการ ' + shortage.need + ')'
      };
    }

    getSS().getSheetByName(SHEET_ORDERS).appendRow([
      orderId,
      timeStr,
      order.tableName || '-',
      JSON.stringify(order.items),
      total,
      order.payment || 'โอนเงิน',
      'สำเร็จ'
    ]);
    if (menuSheet) stockUpdates = deductStock(order.items, menuSheet, menuData);
  } finally {
    lock.releaseLock();
  }

  return { success: true, orderId: orderId, time: timeStr, total: total, stockUpdates: stockUpdates };
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