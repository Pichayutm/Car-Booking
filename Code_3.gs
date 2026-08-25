/**
 * ระบบลง-คืนรถบริษัท (Car Booking System)
 * Backend: Google Apps Script ผูกกับ Google Sheet นี้ (Container-bound script)
 *
 * วิธีติดตั้ง: ดูไฟล์ คู่มือติดตั้ง.md
 */

const SHEET_VEHICLES = 'Vehicles';
const SHEET_PROJECTS = 'Projects';
const SHEET_USERS = 'Users';
const SHEET_PURPOSES = 'Purposes';
const SHEET_TRANSACTIONS = 'Transactions';
const DRIVE_FOLDER_NAME = 'CarBooking_Photos';

// ============================================================
// SETUP — รันครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชัน setupSheets แล้วกด Run)
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---------- Vehicles ----------
  let vSheet = ss.getSheetByName(SHEET_VEHICLES);
  if (!vSheet) vSheet = ss.insertSheet(SHEET_VEHICLES);
  vSheet.clear();
  vSheet.appendRow(['VehicleID', 'ทะเบียนรถ', 'สถานะ', 'ผู้ใช้ปัจจุบัน', 'โปรเจคปัจจุบัน', 'วัตถุประสงค์ปัจจุบัน', 'สถานที่ปัจจุบัน', 'เวลาลงใช้', 'ไมล์ตอนลง', 'กำหนดคืน']);
  const vehicleSeed = [
    '4ขษ6204', '4ขษ6406', '4ขณ8380', '4ขก7909 (Service)',
    '1นก7916 (รถบริษัทฯ)', 'รถตู้เช่า', 'รถเช่าเพิ่มนอกเหนือ 5 คัน (Extra)', 'รถส่วนตัว'
  ];
  vehicleSeed.forEach((plate, i) => {
    vSheet.appendRow(['V' + String(i + 1).padStart(3, '0'), plate, 'ว่าง', '', '', '', '', '', '', '']);
  });
  vSheet.setFrozenRows(1);

  // ---------- Projects ----------
  let pSheet = ss.getSheetByName(SHEET_PROJECTS);
  if (!pSheet) pSheet = ss.insertSheet(SHEET_PROJECTS);
  pSheet.clear();
  pSheet.appendRow(['ProjectID', 'ชื่อโปรเจค', 'Active']);
  const projectSeed = [
    'NT09-NAVY3126', 'NT15-DELL19', 'NT16-CYBER17', 'NT18-CLS2', 'NT19-CLS3',
    'NT20-MADE', 'NT21-BBIP58', 'NT22-100GHW', 'NT23-DWDM24', 'MEA10-OFCS66',
    'MEA11-VDOWALL', 'MEA12-MADWDM68', 'MEA13-RMU2MEA', 'PEA03 Warranty Digital Radio',
    'PEA05 IP Access P2', 'PEA06 IP Access P3', 'PEA07 IP Access P4', 'PEA09 IP Access P5',
    'PEA12-DIGITAL', 'PEA13-IPACCP6', 'PEA16-IPACCP7', 'PEA17-MAOFC66S', 'PEA21-IPACCC67',
    'PEA22-TELEPea', 'PEA23-TELEForth', 'PEA24-MA P1', 'PEA26-IPACC69', 'PEA27-MAOFC6M',
    'OHEC16-OFCMA69', 'RTA28-Repair', 'RTA30-Repair69', 'RTARF03-Skill63', 'RTARF04-Drone65',
    'RTARF05-COM64', 'RTARF06-Cb-test', 'ส่วนกลาง'
  ];
  projectSeed.forEach((name, i) => {
    pSheet.appendRow(['PJ' + String(i + 1).padStart(3, '0'), name, 'Y']);
  });
  pSheet.setFrozenRows(1);

  // ---------- Users ----------
  let uSheet = ss.getSheetByName(SHEET_USERS);
  if (!uSheet) uSheet = ss.insertSheet(SHEET_USERS);
  uSheet.clear();
  uSheet.appendRow(['UserID', 'ชื่อ-นามสกุล', 'Active']);
  const userSeed = [
    'ศรายุทธ์ เทียนเสน', 'จักรพงษ์ อุดรไสว', 'พรสวรรค์ มีศรี', 'ภานุเมธ ประทีปภัทร',
    'สืบคุณ น้อยคล้าย', 'วุฒิชัย รัตนราช', 'เอกพงษ์ พลพุทธา', 'สืบสกุล เทียนมงคล',
    'ธีรภัทร์ ต่อสกุล', 'นิพนธ์ ดิษเจริญ', 'วิสูตร เจริญยศ', 'วัชระพล ศิริธนาคร',
    'ชาญวิทย์ มุสิกวัตร', 'ฌาณเกียรติ ติรกาญจนา', 'พิชณุตม์ มิตรล้วน', 'จักรกฤษ กฤษณายุธ',
    'พรทิพย์ สกุลมาลัยทอง', 'ดวงเดือน สิทธิบุรี', 'test'
  ];
  userSeed.forEach((name, i) => {
    uSheet.appendRow(['U' + String(i + 1).padStart(3, '0'), name, 'Y']);
  });
  uSheet.setFrozenRows(1);

  // ---------- Purposes ----------
  let puSheet = ss.getSheetByName(SHEET_PURPOSES);
  if (!puSheet) puSheet = ss.insertSheet(SHEET_PURPOSES);
  puSheet.clear();
  puSheet.appendRow(['PurposeID', 'ข้อความ', 'Active']);
  const purposeSeed = [
    'ไปทำงานต่างจังหวัด', 'รับกรรมการไปตรวจรับ', 'Onsite MA Service (SLA)',
    'ขนของ/ติดตั้ง', 'ซื้อของ', 'ส่งหนังสือ, ส่งเอกสาร', 'ประชุม'
  ];
  purposeSeed.forEach((text, i) => {
    puSheet.appendRow(['PU' + String(i + 1).padStart(3, '0'), text, 'Y']);
  });
  puSheet.setFrozenRows(1);

  // ---------- Transactions ----------
  let tSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!tSheet) tSheet = ss.insertSheet(SHEET_TRANSACTIONS);
  tSheet.clear();
  tSheet.appendRow(['TransactionID', 'VehicleID', 'ทะเบียนรถ', 'ชื่อผู้ใช้', 'ProjectID', 'ชื่อโปรเจค', 'วัตถุประสงค์', 'สถานที่', 'เวลาลงใช้', 'ไมล์เริ่ม', 'รูปไมล์เริ่ม', 'เวลาคืน', 'ไมล์คืน', 'รูปไมล์คืน', 'ระยะทาง(KM)', 'สถานะ', 'กำหนดคืน']);
  tSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  return 'Setup complete: สร้างชีต Vehicles, Projects, Users, Purposes, Transactions เรียบร้อย';
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getData') {
      result = {
        ok: true,
        vehicles: getVehicles(),
        projects: getProjects(),
        users: getUsers(),
        purposes: getPurposes()
      };
    } else {
      result = { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'checkout') result = checkoutVehicle(data);
    else if (action === 'checkin') result = checkinVehicle(data);
    else if (action === 'ocr') result = ocrMileage(data);
    else result = { ok: false, error: 'unknown action' };
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// READ
// ============================================================
function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
}

function getVehicles() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VEHICLES));
}

function getProjects() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROJECTS)).filter(p => p['Active'] === 'Y');
}

function getUsers() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS)).filter(u => u['Active'] === 'Y');
}

function getPurposes() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PURPOSES)).filter(p => p['Active'] === 'Y');
}

// ============================================================
// CHECKOUT — ลงใช้รถ
// ============================================================
function checkoutVehicle(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName(SHEET_VEHICLES);
    const vValues = vSheet.getDataRange().getValues();
    const vHeaders = vValues[0];
    const colStatus = vHeaders.indexOf('สถานะ');
    const colVehicleID = vHeaders.indexOf('VehicleID');
    const colPlate = vHeaders.indexOf('ทะเบียนรถ');

    let rowIndex = -1;
    for (let i = 1; i < vValues.length; i++) {
      if (vValues[i][colVehicleID] === data.vehicleId) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { ok: false, error: 'ไม่พบรถคันนี้ในระบบ' };
    if (vValues[rowIndex][colStatus] === 'ไม่ว่าง') {
      return { ok: false, error: 'รถคันนี้ถูกใช้งานอยู่แล้ว กรุณาเลือกคันอื่น หรือรอให้คืนก่อน' };
    }

    let photoUrl = '';
    if (data.photoBase64) {
      photoUrl = savePhotoToDrive(data.photoBase64, `checkout_${data.vehicleId}_${new Date().getTime()}.jpg`);
    }

    const now = new Date();
    const txId = 'TX' + now.getTime();
    const rowNum = rowIndex + 1; // 1-indexed sheet row

    vSheet.getRange(rowNum, colStatus + 1).setValue('ไม่ว่าง');
    vSheet.getRange(rowNum, vHeaders.indexOf('ผู้ใช้ปัจจุบัน') + 1).setValue(data.userName);
    vSheet.getRange(rowNum, vHeaders.indexOf('โปรเจคปัจจุบัน') + 1).setValue(data.projectName);
    vSheet.getRange(rowNum, vHeaders.indexOf('วัตถุประสงค์ปัจจุบัน') + 1).setValue(data.purpose);
    vSheet.getRange(rowNum, vHeaders.indexOf('สถานที่ปัจจุบัน') + 1).setValue(data.location);
    vSheet.getRange(rowNum, vHeaders.indexOf('เวลาลงใช้') + 1).setValue(now);
    vSheet.getRange(rowNum, vHeaders.indexOf('ไมล์ตอนลง') + 1).setValue(data.mileage);
    vSheet.getRange(rowNum, vHeaders.indexOf('กำหนดคืน') + 1).setValue(data.expectedReturn || '');

    const tSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
    tSheet.appendRow([
      txId,
      data.vehicleId,
      vValues[rowIndex][colPlate],
      data.userName,
      data.projectId,
      data.projectName,
      data.purpose,
      data.location,
      now,
      data.mileage,
      photoUrl,
      '',
      '',
      '',
      '',
      'กำลังใช้',
      data.expectedReturn || ''
    ]);

    return { ok: true, transactionId: txId };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// CHECKIN — คืนรถ
// ============================================================
function checkinVehicle(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
    const tValues = tSheet.getDataRange().getValues();
    const tHeaders = tValues[0];
    const colVehicleID = tHeaders.indexOf('VehicleID');
    const colStatus = tHeaders.indexOf('สถานะ');
    const colMileStart = tHeaders.indexOf('ไมล์เริ่ม');

    let rowIndex = -1;
    for (let i = tValues.length - 1; i >= 1; i--) {
      if (tValues[i][colVehicleID] === data.vehicleId && tValues[i][colStatus] === 'กำลังใช้') {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { ok: false, error: 'ไม่พบรายการลงใช้ที่ยังไม่คืนของรถคันนี้' };

    const mileageStart = Number(tValues[rowIndex][colMileStart]);
    const mileageEnd = Number(data.mileage);
    if (isNaN(mileageEnd)) return { ok: false, error: 'เลขไมล์ไม่ถูกต้อง' };
    if (mileageEnd < mileageStart) {
      return { ok: false, error: `เลขไมล์ตอนคืน (${mileageEnd}) น้อยกว่าตอนลงใช้ (${mileageStart}) กรุณาตรวจสอบรูปอีกครั้ง` };
    }
    const distance = mileageEnd - mileageStart;

    let photoUrl = '';
    if (data.photoBase64) {
      photoUrl = savePhotoToDrive(data.photoBase64, `checkin_${data.vehicleId}_${new Date().getTime()}.jpg`);
    }

    const now = new Date();
    const rowNum = rowIndex + 1;
    tSheet.getRange(rowNum, tHeaders.indexOf('เวลาคืน') + 1).setValue(now);
    tSheet.getRange(rowNum, tHeaders.indexOf('ไมล์คืน') + 1).setValue(mileageEnd);
    tSheet.getRange(rowNum, tHeaders.indexOf('รูปไมล์คืน') + 1).setValue(photoUrl);
    tSheet.getRange(rowNum, tHeaders.indexOf('ระยะทาง(KM)') + 1).setValue(distance);
    tSheet.getRange(rowNum, colStatus + 1).setValue('คืนแล้ว');

    // reset vehicle row back to available
    const vSheet = ss.getSheetByName(SHEET_VEHICLES);
    const vValues = vSheet.getDataRange().getValues();
    const vHeaders = vValues[0];
    const vColVehicleID = vHeaders.indexOf('VehicleID');
    let vRowIndex = -1;
    for (let i = 1; i < vValues.length; i++) {
      if (vValues[i][vColVehicleID] === data.vehicleId) {
        vRowIndex = i;
        break;
      }
    }
    if (vRowIndex > -1) {
      const vRowNum = vRowIndex + 1;
      vSheet.getRange(vRowNum, vHeaders.indexOf('สถานะ') + 1).setValue('ว่าง');
      vSheet.getRange(vRowNum, vHeaders.indexOf('ผู้ใช้ปัจจุบัน') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('โปรเจคปัจจุบัน') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('วัตถุประสงค์ปัจจุบัน') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('สถานที่ปัจจุบัน') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('เวลาลงใช้') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('ไมล์ตอนลง') + 1).setValue('');
      vSheet.getRange(vRowNum, vHeaders.indexOf('กำหนดคืน') + 1).setValue('');
    }

    return { ok: true, distance: distance, mileageStart: mileageStart, mileageEnd: mileageEnd };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// PHOTO STORAGE (Google Drive)
// ============================================================
function savePhotoToDrive(base64Data, filename) {
  const folder = getOrCreateFolder();
  const parts = base64Data.split(',');
  const contentTypeMatch = parts[0].match(/data:(.*);base64/);
  const contentType = contentTypeMatch ? contentTypeMatch[1] : 'image/jpeg';
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes, contentType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

// ============================================================
// OCR — อ่านเลขไมล์จากรูปด้วย Google Cloud Vision API
// ต้องตั้งค่า Script Property ชื่อ VISION_API_KEY ก่อนใช้งาน
// (Apps Script editor > Project Settings > Script Properties)
// ============================================================
function ocrMileage(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'ยังไม่ได้ตั้งค่า VISION_API_KEY ใน Script Properties ดูวิธีตั้งค่าในคู่มือติดตั้ง' };
  }
  if (!data.photoBase64) return { ok: false, error: 'ไม่พบรูปภาพ' };

  const base64 = data.photoBase64.split(',')[1];
  const payload = {
    requests: [
      {
        image: { content: base64 },
        features: [{ type: 'TEXT_DETECTION' }]
      }
    ]
  };

  const response = UrlFetchApp.fetch('https://vision.googleapis.com/v1/images:annotate?key=' + apiKey, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) return { ok: false, error: 'Vision API error: ' + json.error.message };

  const annotation = json.responses && json.responses[0] && json.responses[0].fullTextAnnotation;
  const text = annotation ? annotation.text : '';

  // ดึงชุดตัวเลขที่ยาวที่สุดในภาพ (โดยทั่วไปคือเลขไมล์บนหน้าปัด)
  const matches = text.match(/\d[\d.,]*\d|\d/g) || [];
  let bestNumber = '';
  matches.forEach(m => {
    const cleaned = m.replace(/[.,]/g, '');
    if (cleaned.length > bestNumber.length) bestNumber = cleaned;
  });

  return { ok: true, rawText: text, detectedMileage: bestNumber };
}
