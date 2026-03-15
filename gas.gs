/**
 * charm+ 帳簿アプリ - Google Apps Script
 * Version: 2.0.0 (Phase 3: マスタシート・タグ・領収書対応)
 *
 * スプレッドシート構成:
 *   取引シート: id / type / date / amount / category /
 *              description / paymentMethod / createdAt /
 *              deletedAt / tags / receiptId
 *   設定シート: key / value
 *   マスタシート: key / value（JSON文字列で保存）
 */

const SHEET_TRANSACTIONS = '取引';
const SHEET_SETTINGS     = '設定';
const SHEET_MASTER       = 'マスタ';
const SHEET_CUSTOMERS    = '顧客';

// ─────────────────────────────────────────
//  エントリポイント
// ─────────────────────────────────────────

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  try {
    let result;
    if (action === 'getAll')         result = getAll();
    else if (action === 'getMaster')    result = getMaster();
    else if (action === 'getCustomers') result = getCustomers();
    else result = { error: '不明なアクション: ' + action };
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === 'addTransaction')    result = addTransaction(body.data);
    else if (action === 'updateTransaction') result = updateTransaction(body.data);
    else if (action === 'deleteTransaction') result = deleteTransaction(body.id);
    else if (action === 'saveSettings')      result = saveSettings(body.data);
    else if (action === 'saveMaster')        result = saveMaster(body.data);
    else if (action === 'uploadReceipt')     result = uploadReceipt(body);
    else if (action === 'addCustomer')       result = addCustomer(body.data);
    else if (action === 'updateCustomer')    result = updateCustomer(body.data);
    else if (action === 'deleteCustomer')    result = deleteCustomer(body.id);
    else result = { error: '不明なアクション: ' + action };

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────
//  シート取得ユーティリティ
// ─────────────────────────────────────────

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_TRANSACTIONS) {
      sheet.appendRow([
        'id','type','date','amount','category',
        'description','paymentMethod','createdAt',
        'deletedAt','tags','receiptId'
      ]);
    } else if (name === SHEET_SETTINGS) {
      sheet.appendRow(['key','value']);
    } else if (name === SHEET_MASTER) {
      sheet.appendRow(['key','value']);
    } else if (name === SHEET_CUSTOMERS) {
      sheet.appendRow([
        'id','createdAt','name','nameKana','phone','email',
        'birthday','allergyNotes','memo','lastVisit','visitCount','totalSpend','tags'
      ]);
    }
  }
  return sheet;
}

// ─────────────────────────────────────────
//  全データ取得
// ─────────────────────────────────────────

function getAll() {
  const txSheet  = getOrCreateSheet(SHEET_TRANSACTIONS);
  const setSheet = getOrCreateSheet(SHEET_SETTINGS);

  const transactions = readTransactions(txSheet);
  const settings     = readSettings(setSheet);

  return { transactions, settings };
}

/** Sheets の Date オブジェクトを YYYY-MM-DD 文字列に変換 */
function formatCellDate(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val === undefined || val === null ? '' : val);
}

function readTransactions(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(String);
  const result  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const t   = {};
    headers.forEach((h, j) => {
      // date 列は必ず YYYY-MM-DD 形式で返す（Sheets の自動日付変換対策）
      if (h === 'date' || h === 'createdAt' || h === 'deletedAt') {
        t[h] = row[j] instanceof Date ? formatCellDate(row[j]) : String(row[j] ?? '');
      } else {
        t[h] = row[j] === undefined || row[j] === null ? '' : String(row[j]);
      }
    });

    // 論理削除済みは除外
    if (t.deletedAt && t.deletedAt !== '') continue;

    // tags を配列に変換
    if (t.tags && t.tags !== '') {
      t.tags = t.tags.split(',').map(s => s.trim()).filter(Boolean);
    } else {
      t.tags = [];
    }

    // amount を数値に
    t.amount = Number(t.amount) || 0;

    result.push(t);
  }
  return result;
}

function readSettings(sheet) {
  const data   = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]);
    const val = data[i][1];
    if (key) {
      if (key === 'initialCash' || key === 'initialBank') {
        result[key] = Number(val) || 0;
      } else {
        result[key] = String(val);
      }
    }
  }
  return result;
}

// ─────────────────────────────────────────
//  取引の追加
// ─────────────────────────────────────────

function addTransaction(data) {
  const sheet  = getOrCreateSheet(SHEET_TRANSACTIONS);
  const tagsStr = Array.isArray(data.tags) ? data.tags.join(',') : (data.tags || '');
  sheet.appendRow([
    data.id          || '',
    data.type        || '',
    data.date        || '',
    Number(data.amount) || 0,
    data.category    || '',
    data.description || '',
    data.paymentMethod || '',
    data.createdAt   || new Date().toISOString(),
    '',                // deletedAt
    tagsStr,
    data.receiptId   || '',
  ]);
  return { success: true };
}

// ─────────────────────────────────────────
//  取引の更新
// ─────────────────────────────────────────

function updateTransaction(data) {
  const sheet   = getOrCreateSheet(SHEET_TRANSACTIONS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(data.id)) {
      const colMap = {};
      headers.forEach((h, j) => { colMap[h] = j + 1; }); // 1-indexed for setCell

      const tagsStr = Array.isArray(data.tags) ? data.tags.join(',') : (data.tags || '');

      if (colMap.type)          sheet.getRange(i+1, colMap.type).setValue(data.type || '');
      if (colMap.date)          sheet.getRange(i+1, colMap.date).setValue(data.date || '');
      if (colMap.amount)        sheet.getRange(i+1, colMap.amount).setValue(Number(data.amount) || 0);
      if (colMap.category)      sheet.getRange(i+1, colMap.category).setValue(data.category || '');
      if (colMap.description)   sheet.getRange(i+1, colMap.description).setValue(data.description || '');
      if (colMap.paymentMethod) sheet.getRange(i+1, colMap.paymentMethod).setValue(data.paymentMethod || '');
      if (colMap.tags)          sheet.getRange(i+1, colMap.tags).setValue(tagsStr);
      if (colMap.receiptId && data.receiptId !== undefined)
        sheet.getRange(i+1, colMap.receiptId).setValue(data.receiptId || '');

      return { success: true };
    }
  }
  return { error: '取引が見つかりません: ' + data.id };
}

// ─────────────────────────────────────────
//  取引の削除（論理削除）
// ─────────────────────────────────────────

function deleteTransaction(id) {
  const sheet   = getOrCreateSheet(SHEET_TRANSACTIONS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');
  const delIdx  = headers.indexOf('deletedAt');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(id)) {
      sheet.getRange(i + 1, delIdx + 1).setValue(new Date().toISOString());
      return { success: true };
    }
  }
  return { error: '取引が見つかりません: ' + id };
}

// ─────────────────────────────────────────
//  設定の保存
// ─────────────────────────────────────────

function saveSettings(data) {
  const sheet = getOrCreateSheet(SHEET_SETTINGS);
  const rows  = sheet.getDataRange().getValues();
  const keys  = rows.slice(1).map(r => String(r[0]));

  Object.entries(data).forEach(([key, val]) => {
    const idx = keys.indexOf(key);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(String(val));
    } else {
      sheet.appendRow([key, String(val)]);
      keys.push(key);
    }
  });
  return { success: true };
}

// ─────────────────────────────────────────
//  マスタデータの取得・保存
// ─────────────────────────────────────────

function getMaster() {
  const sheet = getOrCreateSheet(SHEET_MASTER);
  const rows  = sheet.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]);
    const val = String(rows[i][1]);
    if (!key) continue;
    try {
      result[key] = JSON.parse(val);
    } catch {
      result[key] = val;
    }
  }
  return result;
}

function saveMaster(data) {
  const sheet = getOrCreateSheet(SHEET_MASTER);
  const rows  = sheet.getDataRange().getValues();
  const keys  = rows.slice(1).map(r => String(r[0]));

  Object.entries(data).forEach(([key, val]) => {
    const jsonVal = JSON.stringify(val);
    const idx     = keys.indexOf(key);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(jsonVal);
    } else {
      sheet.appendRow([key, jsonVal]);
      keys.push(key);
    }
  });
  return { success: true };
}

// ─────────────────────────────────────────
//  領収書アップロード（Googleドライブ）
// ─────────────────────────────────────────

function uploadReceipt(body) {
  const { fileName, mimeType, base64Data } = body;
  if (!base64Data) return { error: 'base64Dataが空です' };

  // 保存先フォルダ（「charm+ 領収書」フォルダを自動作成）
  const folderName = 'charm+ 領収書';
  const folders    = DriveApp.getFoldersByName(folderName);
  const folder     = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(folderName);

  const bytes = Utilities.base64Decode(base64Data);
  const blob  = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || 'receipt.jpg');
  const file  = folder.createFile(blob);

  // 閲覧リンクを「リンクを知っている人」に共有
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { success: true, fileId: file.getId() };
}

// ─────────────────────────────────────────
//  顧客管理
// ─────────────────────────────────────────

function getCustomers() {
  const sheet = getOrCreateSheet(SHEET_CUSTOMERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { customers: [] };

  const headers = data[0].map(String);
  const result  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const c   = {};
    headers.forEach((h, j) => {
      if (h === 'visitCount' || h === 'totalSpend') {
        c[h] = Number(row[j]) || 0;
      } else if (h === 'createdAt' || h === 'lastVisit') {
        c[h] = row[j] instanceof Date ? formatCellDate(row[j]) : String(row[j] ?? '');
      } else {
        c[h] = row[j] === undefined || row[j] === null ? '' : String(row[j]);
      }
    });
    if (c.id) result.push(c);
  }
  return { customers: result };
}

function addCustomer(data) {
  const sheet = getOrCreateSheet(SHEET_CUSTOMERS);
  sheet.appendRow([
    data.id          || '',
    data.createdAt   || new Date().toISOString(),
    data.name        || '',
    data.nameKana    || '',
    data.phone       || '',
    data.email       || '',
    data.birthday    || '',
    data.allergyNotes || '',
    data.memo        || '',
    data.lastVisit   || '',
    Number(data.visitCount)  || 0,
    Number(data.totalSpend)  || 0,
    data.tags        || '',
  ]);
  return { success: true };
}

function updateCustomer(data) {
  const sheet   = getOrCreateSheet(SHEET_CUSTOMERS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(data.id)) {
      const colMap = {};
      headers.forEach((h, j) => { colMap[h] = j + 1; });

      const fields = ['name','nameKana','phone','email','birthday','allergyNotes','memo','lastVisit','visitCount','totalSpend','tags'];
      fields.forEach(f => {
        if (colMap[f] && data[f] !== undefined) {
          sheet.getRange(i + 1, colMap[f]).setValue(
            (f === 'visitCount' || f === 'totalSpend') ? (Number(data[f]) || 0) : (data[f] || '')
          );
        }
      });
      return { success: true };
    }
  }
  return { error: '顧客が見つかりません: ' + data.id };
}

function deleteCustomer(id) {
  const sheet   = getOrCreateSheet(SHEET_CUSTOMERS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: '顧客が見つかりません: ' + id };
}
