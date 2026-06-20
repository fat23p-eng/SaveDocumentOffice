/**
 * ระบบจัดเก็บไฟล์เอกสารสำนักงาน — Apps Script Backend (v3)
 * ----------------------------------------------------------
 * เปลี่ยนจาก v2:
 * - กลุ่ม "ทั่วไป" ตอนนี้เห็นเอกสารได้ 3 ส่วน: ของตัวเอง + งานนายทะเบียนสหกรณ์ + งานสำนักงาน
 *   (ก่อนหน้านี้เห็นแค่ของตัวเองอย่างเดียว)
 * - "ดูทุกกลุ่ม" และ "ผู้ดูแลระบบ" ทำงานเหมือนเดิม (เห็นได้ทุกกลุ่ม)
 * - ลบเอกสารยังจำกัดแบบเดิม: ทั่วไป/ดูทุกกลุ่ม ลบได้เฉพาะของตัวเอง, ผู้ดูแลระบบลบได้ทุกอัน
 *
 * ถ้าต้องการเปลี่ยนชื่อ 2 กลุ่มที่ "ทุกคนเห็นได้" ให้แก้ค่าใน ALWAYS_VISIBLE_GROUPS ด้านล่างนี้
 * ให้ตรงกับชื่อกลุ่มในชีต users เป๊ะๆ (รวมเว้นวรรค/สะกด)
 *
 * วิธีติดตั้ง: แทนที่เนื้อหา Code.gs เดิมทั้งหมดด้วยไฟล์นี้
 */

const SHEET_DATA = 'ข้อมูลหลัก';
const SHEET_USERS = 'users';
const SHEET_DOCTYPES = 'ประเภทเอกสาร';
const ROOT_FOLDER_NAME = 'ระบบจัดเก็บไฟล์เอกสารสำนักงาน';
const FOLDER_PROP_KEY = 'ROOT_FOLDER_ID';

const ROLE_NORMAL = 'ทั่วไป';
const ROLE_VIEW_ALL = 'ดูทุกกลุ่ม';
const ROLE_ADMIN = 'ผู้ดูแลระบบ';

// กลุ่มที่ทุกคน (ไม่ว่าสิทธิ์ใด) มองเห็นเอกสารได้เสมอ นอกเหนือจากของตัวเอง
const ALWAYS_VISIBLE_GROUPS = ['งานนายทะเบียนสหกรณ์', 'งานสำนักงาน'];

// ---------- Web App Entry ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบจัดเก็บไฟล์เอกสารสำนักงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Helpers ----------
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function normalizeRole(raw) {
  const r = (raw || '').toString().trim();
  if (r === ROLE_VIEW_ALL || r === ROLE_ADMIN) return r;
  return ROLE_NORMAL;
}

function uniqueList(arr) {
  const seen = {};
  const out = [];
  arr.forEach(v => {
    if (v && !seen[v]) { seen[v] = true; out.push(v); }
  });
  return out;
}

function getOrCreateRootFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(FOLDER_PROP_KEY);
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // fall through and recreate
    }
  }
  const existing = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  let folder;
  if (existing.hasNext()) {
    folder = existing.next();
  } else {
    folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  }
  props.setProperty(FOLDER_PROP_KEY, folder.getId());
  return folder;
}

function getOrCreateSubFolder(parent, name) {
  const safeName = (name || 'อื่นๆ').toString().substring(0, 100);
  const existing = parent.getFoldersByName(safeName);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(safeName);
}

// ---------- Auth ----------
function login(groupName, password) {
  const sheet = getSS().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() === String(groupName).trim() &&
        String(row[1]).trim() === String(password).trim()) {
      return { success: true, groupName: row[0], role: normalizeRole(row[2]) };
    }
  }
  return { success: false, message: 'ชื่อกลุ่มหรือรหัสผ่านไม่ถูกต้อง' };
}

// ---------- Init data (groups + doc types) ----------
function getInitData() {
  const ss = getSS();
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0]) groups.push(usersData[i][0]);
  }

  const typeSheet = ss.getSheetByName(SHEET_DOCTYPES);
  const typeData = typeSheet.getDataRange().getValues();
  const docTypes = [];
  for (let i = 1; i < typeData.length; i++) {
    if (typeData[i][0]) docTypes.push(typeData[i][0]);
  }

  return { groups: groups, docTypes: docTypes, alwaysVisibleGroups: ALWAYS_VISIBLE_GROUPS };
}

// ---------- Upload ----------
function uploadFile(payload, requestorGroup, requestorRole) {
  try {
    const role = normalizeRole(requestorRole);
    const effectiveGroup = (role === ROLE_ADMIN && payload.groupName) ? payload.groupName : requestorGroup;

    const rootFolder = getOrCreateRootFolder();
    const groupFolder = getOrCreateSubFolder(rootFolder, effectiveGroup);
    const typeFolder = getOrCreateSubFolder(groupFolder, payload.docType);

    const decoded = Utilities.base64Decode(payload.base64Data);
    const blob = Utilities.newBlob(decoded, payload.mimeType, payload.fileName);
    const file = typeFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sheet = getSS().getSheetByName(SHEET_DATA);
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sheet.appendRow([
      dateStr,
      effectiveGroup,
      payload.docType,
      payload.note || '',
      file.getUrl(),
      file.getName()
    ]);

    return { success: true, fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (err) {
    return { success: false, message: 'อัพโหลดไม่สำเร็จ: ' + err.message };
  }
}

// ---------- Search / list files ----------
/**
 * - ทั่วไป : เห็นได้เฉพาะ [ของตัวเอง] + ALWAYS_VISIBLE_GROUPS เท่านั้น
 *            ถ้า filters.groupName ที่ client ส่งมาไม่อยู่ในรายการที่อนุญาต จะถูกเซ็ตกลับเป็นกลุ่มตัวเอง
 *            และมีการกรองซ้ำอีกชั้นในลูปด้านล่าง (กันกรณี client ส่งค่าแปลกๆ มา)
 * - ดูทุกกลุ่ม / ผู้ดูแลระบบ : เห็นได้ทุกกลุ่มตามเดิม
 */
function getFiles(filters, requestorGroup, requestorRole) {
  filters = filters || {};
  const role = normalizeRole(requestorRole);

  let allowedGroups = null; // null = ไม่จำกัด (เห็นได้ทุกกลุ่ม)
  if (role === ROLE_NORMAL) {
    allowedGroups = uniqueList([requestorGroup].concat(ALWAYS_VISIBLE_GROUPS));
    if (filters.groupName && filters.groupName !== 'ทั้งหมด' && allowedGroups.indexOf(filters.groupName) === -1) {
      filters.groupName = requestorGroup;
    }
  }

  const sheet = getSS().getSheetByName(SHEET_DATA);
  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[5] && !row[4]) continue;
    const date = row[0];
    const groupName = row[1];
    const docType = row[2];
    const note = row[3];
    const fileUrl = row[4];
    const fileName = row[5];

    // ชั้นป้องกันหลัก: ถ้าจำกัดกลุ่มไว้ (ทั่วไป) แถวนี้ต้องอยู่ในรายการที่อนุญาตเท่านั้น
    if (allowedGroups && allowedGroups.indexOf(groupName) === -1) continue;

    if (filters.groupName && filters.groupName !== 'ทั้งหมด' && groupName !== filters.groupName) continue;
    if (filters.docType && filters.docType !== 'ทั้งหมด' && docType !== filters.docType) continue;
    if (filters.keyword) {
      const kw = filters.keyword.toString().trim().toLowerCase();
      const haystack = ((fileName || '') + ' ' + (note || '')).toLowerCase();
      if (kw && haystack.indexOf(kw) === -1) continue;
    }

    let dateOut = date;
    if (date instanceof Date) {
      dateOut = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    }

    const canManage = (role === ROLE_ADMIN) || (groupName === requestorGroup);

    results.push({
      rowIndex: i + 1,
      date: dateOut,
      groupName: groupName,
      docType: docType,
      note: note,
      fileUrl: fileUrl,
      fileName: fileName,
      canManage: canManage,
      canEdit: role === ROLE_ADMIN
    });
  }

  results.sort((a, b) => b.rowIndex - a.rowIndex);
  return results;
}

// ---------- Delete ----------
function deleteFile(rowIndex, requestorGroup, requestorRole) {
  try {
    const role = normalizeRole(requestorRole);
    const sheet = getSS().getSheetByName(SHEET_DATA);
    const rowGroup = sheet.getRange(rowIndex, 2).getValue();

    const allowed = (role === ROLE_ADMIN) || (rowGroup === requestorGroup);
    if (!allowed) {
      return { success: false, message: 'ไม่มีสิทธิ์ลบเอกสารของกลุ่มอื่น' };
    }

    const fileUrl = sheet.getRange(rowIndex, 5).getValue();
    if (fileUrl) {
      const idMatch = fileUrl.toString().match(/[-\w]{25,}/);
      if (idMatch) {
        try {
          DriveApp.getFileById(idMatch[0]).setTrashed(true);
        } catch (e) {
          // ไฟล์อาจถูกลบไปแล้ว ไม่เป็นไร
        }
      }
    }
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (err) {
    return { success: false, message: 'ลบไม่สำเร็จ: ' + err.message };
  }
}

// ---------- Edit (เฉพาะผู้ดูแลระบบ) ----------
function editFile(rowIndex, updates, requestorGroup, requestorRole) {
  const role = normalizeRole(requestorRole);
  if (role !== ROLE_ADMIN) {
    return { success: false, message: 'เฉพาะผู้ดูแลระบบเท่านั้นที่แก้ไขรายการได้' };
  }
  try {
    const sheet = getSS().getSheetByName(SHEET_DATA);
    if (updates.groupName) sheet.getRange(rowIndex, 2).setValue(updates.groupName);
    if (updates.docType) sheet.getRange(rowIndex, 3).setValue(updates.docType);
    if (updates.note !== undefined) sheet.getRange(rowIndex, 4).setValue(updates.note);
    return { success: true };
  } catch (err) {
    return { success: false, message: 'แก้ไขไม่สำเร็จ: ' + err.message };
  }
}
