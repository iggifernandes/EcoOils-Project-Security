// CODE.GS VERSION: 1.3.0 — EcoOils Security backend (guard database, documents, photos)
/**
 * Security Timesheet — Apps Script backend.
 *
 * Deploy this as a Web App (Deploy > New deployment > Web app,
 * Execute as: Me, Who has access: Anyone). The deployment URL you get
 * is what the frontend calls instead of localStorage/window.storage.
 *
 * Data model:
 *  - A "KV" sheet tab in THIS spreadsheet acts as a simple key/value store,
 *    mirroring the app's existing loadShared/saveShared(key, value) calls:
 *      Column A: key       (e.g. "roster", "entries:<guardId>", "schedule", "settings")
 *      Column B: updatedAt (ISO timestamp, informational only)
 *      Column C onward: value (JSON string, split into 40k-char chunks)
 *  - Uploaded files (photos, SKCK scans, certificates) are saved as real
 *    files in the Drive folder set below, not stuffed into a cell — Sheets
 *    cells cap out around 50,000 characters, far too small for a 2MB file.
 *
 * SETUP — only two things to edit below:
 *   1. DRIVE_FOLDER_ID — the parent Drive folder files should be saved into.
 *      (Already set to the "Security Personnel Certificates — EcoOils
 *      Mentari" folder created earlier.)
 *   2. Nothing else needs editing — the KV sheet is created automatically
 *      on first use if it doesn't exist yet.
 */

// Leave empty: the script creates a Drive folder "EcoOils Security Documents" in the account
// that runs it (the account that owns this script) the first time a file is uploaded.
const CODE_VERSION = "1.3.0";
const DRIVE_FOLDER_ID = "";
const ROOT_FOLDER_NAME = "EcoOils Security Documents";

function getRootFolder_() {
  if (DRIVE_FOLDER_ID) return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("ROOT_FOLDER_ID");
  if (saved) {
    try { return DriveApp.getFolderById(saved); } catch (e) { /* deleted: create a new one */ }
  }
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty("ROOT_FOLDER_ID", folder.getId());
  return folder;
}
const KV_SHEET_NAME = "KV";
const CHUNK = 40000; // Sheets cells cap at 50,000 chars; values are split across columns C, D, E...

// Layout: A = key, B = updatedAt, C.. = value chunks (stored as plain text).

function getKvSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(KV_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(KV_SHEET_NAME);
    sheet.appendRow(["key", "updatedAt", "value (chunked)"]);
  }
  return sheet;
}

function findRow_(sheet, key) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) return i + 2;
  }
  return -1;
}

function rowToValue_(row) {
  let out = "";
  for (let c = 2; c < row.length; c++) {
    if (row[c] === "" || row[c] == null) break;
    out += String(row[c]);
  }
  return out;
}

function kvGet_(key) {
  const sheet = getKvSheet_();
  const r = findRow_(sheet, key);
  if (r === -1) return null;
  const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
  return rowToValue_(row) || null;
}

function kvGetPrefix_(prefix) {
  const sheet = getKvSheet_();
  const last = sheet.getLastRow();
  const items = {};
  if (last < 2) return items;
  const data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const k = String(data[i][0]);
    if (k.indexOf(prefix) === 0) items[k] = rowToValue_(data[i]);
  }
  return items;
}

function kvSet_(key, value) {
  const sheet = getKvSheet_();
  const str = String(value);
  const chunks = [];
  for (let i = 0; i < str.length; i += CHUNK) chunks.push(str.substring(i, i + CHUNK));
  if (chunks.length === 0) chunks.push("");
  let r = findRow_(sheet, key);
  if (r === -1) r = Math.max(sheet.getLastRow(), 1) + 1;
  const needCols = 2 + chunks.length;
  if (sheet.getMaxColumns() < needCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  const width = Math.max(sheet.getLastColumn(), needCols);
  const rowVals = [key, new Date().toISOString()].concat(chunks);
  while (rowVals.length < width) rowVals.push(""); // clear leftover old chunks
  const range = sheet.getRange(r, 1, 1, width);
  range.setNumberFormat("@"); // plain text: no number/date/formula coercion
  range.setValues([rowVals]);
  return true;
}

function kvDelete_(key) {
  const sheet = getKvSheet_();
  const r = findRow_(sheet, key);
  if (r !== -1) sheet.deleteRow(r);
  return true;
}

function getGuardFolder_(name) {
  const parent = getRootFolder_();
  const clean = String(name || "").replace(/[\\\/:*?"<>|]/g, " ").trim();
  if (!clean) return parent;
  const it = parent.getFoldersByName(clean);
  return it.hasNext() ? it.next() : parent.createFolder(clean);
}

function uploadFile_(fileName, mimeType, base64Data, folderName) {
  const folder = getGuardFolder_(folderName);
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName);
  const file = folder.createFile(blob);
  // "Anyone with the link can view" so guards' phones can load/download it
  // without needing to sign in to Google themselves.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* sharing blocked by account policy: the file is still saved */ }
  return {
    id: file.getId(),
    url: file.getUrl(),
    viewUrl: "https://drive.google.com/uc?export=view&id=" + file.getId(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId(),
  };
}

function deleteFile_(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (e) {
    return false;
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: "bad-request" });
  }

  const action = body.action;

  // One write at a time: prevents two phones corrupting the same row.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return jsonResponse_({ ok: false, error: "busy" });
  }

  try {
    if (action === "getPrefix") {
      return jsonResponse_({ ok: true, items: kvGetPrefix_(String(body.prefix || "")) });
    }
    if (action === "get") {
      const value = kvGet_(body.key);
      return jsonResponse_({ ok: true, value: value });
    }
    if (action === "set") {
      kvSet_(body.key, body.value);
      return jsonResponse_({ ok: true });
    }
    if (action === "delete") {
      kvDelete_(body.key);
      return jsonResponse_({ ok: true });
    }
    if (action === "uploadFile") {
      const result = uploadFile_(body.fileName, body.mimeType, body.base64Data, body.folderName);
      return jsonResponse_({ ok: true, file: result });
    }
    if (action === "deleteFile") {
      const ok = deleteFile_(body.fileId);
      return jsonResponse_({ ok: ok });
    }
    return jsonResponse_({ ok: false, error: "unknown-action" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return jsonResponse_({ ok: true, message: "EcoOils Security backend is running.", version: CODE_VERSION });
}
