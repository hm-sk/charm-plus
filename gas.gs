/**
 * charm+ 帳簿アプリ - Google Apps Script
 * Version: 2.1.0 (Phase 3+: データ保護・バックアップ対応)
 *
 * スプレッドシート構成:
 *   取引シート: id / type / date / amount / category /
 *              description / paymentMethod / createdAt /
 *              deletedAt / tags / receiptId
 *   設定シート: key / value
 *   マスタシート: key / value（JSON文字列で保存）
 *
 * データ保護:
 *   - シート保護: スクリプト以外からの編集をロック
 *   - 自動バックアップ: 日次でスプレッドシートをDriveにコピー
 *   - 手動バックアップ: UIから即座にバックアップ可能
 */

const SHEET_TRANSACTIONS = '取引';
const SHEET_SETTINGS     = '設定';
const SHEET_MASTER       = 'マスタ';
const SHEET_CUSTOMERS    = '顧客';
const SHEET_APPOINTMENTS = '予約';

// ─────────────────────────────────────────
//  エントリポイント
// ─────────────────────────────────────────

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  try {
    let result;
    if (action === 'getAll')              result = getAll();
    else if (action === 'getMaster')         result = getMaster();
    else if (action === 'getCustomers')      result = getCustomers();
    else if (action === 'getBookingInfo')    result = getBookingInfo();
    else if (action === 'getAvailableSlots') result = getAvailableSlots(e.parameter.menuId, e.parameter.date);
    else if (action === 'getAppointments')   result = getAppointments(e.parameter.from, e.parameter.to);
    else if (action === 'getBackupStatus')   result = getBackupStatus();
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
    else if (action === 'addCustomer')            result = addCustomer(body.data);
    else if (action === 'updateCustomer')         result = updateCustomer(body.data);
    else if (action === 'deleteCustomer')         result = deleteCustomer(body.id);
    else if (action === 'createAppointment')      result = createAppointment(body.data);
    else if (action === 'updateAppointmentStatus') result = updateAppointmentStatus(body.id, body.status, body.staffNote);
    else if (action === 'cancelAppointment')      result = cancelAppointment(body.id);
    else if (action === 'setupProtection')        result = setupSheetProtection();
    else if (action === 'createBackup')           result = createManualBackup();
    else if (action === 'saveBackupSettings')     result = saveBackupSettings(body.data);
    else if (action === 'setupBackupTrigger')     result = setupBackupTrigger();
    else if (action === 'removeBackupTrigger')    result = removeBackupTrigger();
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
    } else if (name === SHEET_APPOINTMENTS) {
      sheet.appendRow([
        'id','createdAt','customerName','phone','email',
        'menuId','menuName','price','dateTime','duration',
        'status','notes','staffNote','reminderSent','customerId','transactionId'
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

// ─────────────────────────────────────────
//  予約管理
// ─────────────────────────────────────────

/** 予約フォーム用の初期データ（メニュー一覧 + 営業時間） */
function getBookingInfo() {
  const master = getMaster();
  const menus  = master.menus || [];
  const activeMenus = menus.filter(m => m.isActive !== false);
  return {
    menus:             activeMenus,
    businessHours:     master.businessHours     || null,
    bookingWindowDays: master.bookingWindowDays  || 60,
  };
}

/** 指定日の空き時間スロットを返す */
function getAvailableSlots(menuId, date) {
  if (!menuId || !date) return { slots: [] };

  const master  = getMaster();
  const menus   = master.menus || [];
  const menu    = menus.find(m => m.id === menuId);
  if (!menu) return { error: 'メニューが見つかりません', slots: [] };

  const duration   = Number(menu.duration) || 60; // 施術時間（分）
  const bh         = master.businessHours || {};
  const dayNames   = ['sun','mon','tue','wed','thu','fri','sat'];
  const dayOfWeek  = dayNames[new Date(date + 'T00:00:00').getDay()];
  const dayHours   = bh[dayOfWeek];

  if (!dayHours) return { slots: [] }; // 定休日

  // スロット生成（開始〜終了まで duration 刻み）
  const slots = [];
  const [openH, openM]   = dayHours.open.split(':').map(Number);
  const [closeH, closeM] = dayHours.close.split(':').map(Number);
  const openMin  = openH  * 60 + openM;
  const closeMin = closeH * 60 + closeM;

  for (let t = openMin; t + duration <= closeMin; t += duration) {
    const hh = String(Math.floor(t / 60)).padStart(2, '0');
    const mm = String(t % 60).padStart(2, '0');
    slots.push(`${hh}:${mm}`);
  }

  // 既存予約と重複するスロットを除外
  const sheet = getOrCreateSheet(SHEET_APPOINTMENTS);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length > 1) {
    const headers = rows[0].map(String);
    const dtIdx   = headers.indexOf('dateTime');
    const durIdx  = headers.indexOf('duration');
    const stIdx   = headers.indexOf('status');

    rows.slice(1).forEach(row => {
      const status = String(row[stIdx] || '');
      if (status === 'cancelled' || status === 'noshow') return;
      const apptDt = String(row[dtIdx] || '');
      if (!apptDt.startsWith(date)) return;
      const apptTime    = apptDt.split('T')[1]?.substring(0, 5) || apptDt.split(' ')[1]?.substring(0, 5) || '';
      const apptDur     = Number(row[durIdx]) || 60;
      const [ah, am]    = apptTime.split(':').map(Number);
      const apptStart   = ah * 60 + am;
      const apptEnd     = apptStart + apptDur;

      // このスロットが既存予約と重複するか
      for (let i = slots.length - 1; i >= 0; i--) {
        const [sh, sm] = slots[i].split(':').map(Number);
        const slotStart = sh * 60 + sm;
        const slotEnd   = slotStart + duration;
        if (slotStart < apptEnd && slotEnd > apptStart) {
          slots.splice(i, 1);
        }
      }
    });
  }

  // 今日の場合、現在時刻以前のスロットを除外
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (date === today) {
    const now    = new Date();
    const nowMin = now.getHours() * 60 + now.getMinutes() + 30; // 30分バッファ
    return { slots: slots.filter(s => {
      const [sh, sm] = s.split(':').map(Number);
      return sh * 60 + sm >= nowMin;
    })};
  }

  return { slots };
}

/** 予約一覧取得（管理用） */
function getAppointments(from, to) {
  const sheet = getOrCreateSheet(SHEET_APPOINTMENTS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { appointments: [] };

  const headers = data[0].map(String);
  const result  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const a   = {};
    headers.forEach((h, j) => {
      if (h === 'price' || h === 'duration') {
        a[h] = Number(row[j]) || 0;
      } else if (h === 'reminderSent') {
        a[h] = row[j] === true || row[j] === 'TRUE';
      } else if (h === 'createdAt') {
        a[h] = row[j] instanceof Date ? formatCellDate(row[j]) : String(row[j] ?? '');
      } else {
        a[h] = row[j] === undefined || row[j] === null ? '' : String(row[j]);
      }
    });
    if (!a.id) continue;
    // 日付範囲フィルタ
    const apptDate = a.dateTime ? a.dateTime.substring(0, 10) : '';
    if (from && apptDate < from) continue;
    if (to   && apptDate > to)   continue;
    result.push(a);
  }
  return { appointments: result };
}

/** 予約作成（重複チェック付き） */
function createAppointment(data) {
  // 空き確認（再チェック）
  const slotsResult = getAvailableSlots(data.menuId, data.dateTime?.substring(0, 10));
  const requestedTime = data.dateTime?.split('T')[1]?.substring(0, 5) || data.dateTime?.split(' ')[1]?.substring(0, 5) || '';
  if (!slotsResult.slots.includes(requestedTime)) {
    return { error: 'ご希望の時間は既に埋まっています。別の時間をお選びください。' };
  }

  const sheet = getOrCreateSheet(SHEET_APPOINTMENTS);
  const id    = 'appt_' + new Date().getTime();

  sheet.appendRow([
    id,
    new Date().toISOString(),
    data.customerName  || '',
    data.phone         || '',
    data.email         || '',
    data.menuId        || '',
    data.menuName      || '',
    Number(data.price) || 0,
    data.dateTime      || '',
    Number(data.duration) || 60,
    'pending',
    data.notes         || '',
    '',  // staffNote
    false,
    data.customerId    || '',
    '',  // transactionId
  ]);

  // 顧客に確認メール
  if (data.email) {
    try {
      const subject = `【ご予約受付】${data.menuName} ${data.dateTime}`;
      const body    = `${data.customerName} 様\n\nご予約を受け付けました。\n\n`
                    + `メニュー: ${data.menuName}\n`
                    + `日時: ${data.dateTime}\n`
                    + `金額: ¥${Number(data.price).toLocaleString()}\n\n`
                    + `確定メールをお待ちください。\n\n`;
      MailApp.sendEmail(data.email, subject, body);
    } catch(e) {
      Logger.log('メール送信失敗: ' + e.message);
    }
  }

  // オーナーに通知（設定から取得）
  try {
    const ownerEmail = Session.getActiveUser().getEmail();
    if (ownerEmail) {
      MailApp.sendEmail(
        ownerEmail,
        `【新規予約】${data.customerName} 様 ${data.menuName}`,
        `新しい予約が入りました。\n\nお名前: ${data.customerName}\n電話: ${data.phone}\nメニュー: ${data.menuName}\n日時: ${data.dateTime}\n備考: ${data.notes || 'なし'}\n`
      );
    }
  } catch(e) {
    Logger.log('オーナー通知失敗: ' + e.message);
  }

  return { success: true, id };
}

/** ステータス更新（confirmed / completed / cancelled / noshow） */
function updateAppointmentStatus(id, status, staffNote) {
  const sheet   = getOrCreateSheet(SHEET_APPOINTMENTS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');
  const stIdx   = headers.indexOf('status');
  const snIdx   = headers.indexOf('staffNote');
  const emailIdx = headers.indexOf('email');
  const nameIdx  = headers.indexOf('customerName');
  const menuIdx  = headers.indexOf('menuName');
  const dtIdx   = headers.indexOf('dateTime');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) !== String(id)) continue;

    sheet.getRange(i + 1, stIdx + 1).setValue(status);
    if (staffNote !== undefined && snIdx >= 0) {
      sheet.getRange(i + 1, snIdx + 1).setValue(staffNote || '');
    }

    // confirmed 時に確定メール送信
    if (status === 'confirmed') {
      const email = String(rows[i][emailIdx] || '');
      if (email) {
        try {
          MailApp.sendEmail(
            email,
            `【予約確定】${rows[i][menuIdx]} ${rows[i][dtIdx]}`,
            `${rows[i][nameIdx]} 様\n\nご予約が確定しました。\n\nメニュー: ${rows[i][menuIdx]}\n日時: ${rows[i][dtIdx]}\n\nご来店をお待ちしております。\n`
          );
        } catch(e) { Logger.log('確定メール失敗: ' + e.message); }
      }
    }

    return { success: true };
  }
  return { error: '予約が見つかりません: ' + id };
}

function cancelAppointment(id) {
  return updateAppointmentStatus(id, 'cancelled');
}

// ─────────────────────────────────────────
//  データ保護：シート保護
// ─────────────────────────────────────────

/**
 * 全データシートに保護を適用
 * スクリプトオーナーのみ編集可能にし、他ユーザーの直接編集を禁止
 */
function setupSheetProtection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [SHEET_TRANSACTIONS, SHEET_SETTINGS, SHEET_MASTER, SHEET_CUSTOMERS, SHEET_APPOINTMENTS];
  const me = Session.getEffectiveUser();
  const results = [];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { results.push({ sheet: name, status: 'not_found' }); return; }

    // 既存の保護を確認・削除してから再設定
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existing.forEach(p => p.remove());

    const protection = sheet.protect().setDescription('charm+ データ保護: ' + name);
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== me.getEmail()));

    // 警告表示（完全ロックできない場合のフォールバック）
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    results.push({ sheet: name, status: 'protected' });
  });

  return { success: true, results };
}

// ─────────────────────────────────────────
//  データ保護：自動バックアップ
// ─────────────────────────────────────────

/** バックアップ設定をマスタシートから取得 */
function _getBackupConfig() {
  const master = getMaster();
  return master.backupConfig || {
    enabled: false,
    folderPath: 'charm+ バックアップ',
    maxKeep: 30,         // 保持する世代数
    lastBackup: null,
    lastBackupId: null,
  };
}

/** バックアップ設定を保存 */
function saveBackupSettings(data) {
  const current = _getBackupConfig();
  const updated = {
    enabled:    data.enabled !== undefined ? data.enabled : current.enabled,
    folderPath: data.folderPath || current.folderPath,
    maxKeep:    Number(data.maxKeep) || current.maxKeep,
    lastBackup: current.lastBackup,
    lastBackupId: current.lastBackupId,
  };
  return saveMaster({ backupConfig: updated });
}

/** バックアップ先フォルダを取得または作成 */
function _getOrCreateBackupFolder(folderPath) {
  const parts = folderPath.split('/').filter(Boolean);
  let parent = DriveApp.getRootFolder();

  for (const part of parts) {
    const folders = parent.getFoldersByName(part);
    if (folders.hasNext()) {
      parent = folders.next();
    } else {
      parent = parent.createFolder(part);
    }
  }
  return parent;
}

/** バックアップ実行（共通処理） */
function _executeBackup(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = _getOrCreateBackupFolder(config.folderPath);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const backupName = ss.getName() + '_backup_' + timestamp;

  // スプレッドシートをコピー
  const copy = ss.copy(backupName);
  const file = DriveApp.getFileById(copy.getId());

  // バックアップフォルダに移動
  file.moveTo(folder);

  // 古いバックアップを世代管理で削除
  const maxKeep = config.maxKeep || 30;
  const files = folder.getFiles();
  const backups = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().includes('_backup_')) {
      backups.push({ file: f, date: f.getDateCreated() });
    }
  }
  backups.sort((a, b) => b.date - a.date);
  if (backups.length > maxKeep) {
    backups.slice(maxKeep).forEach(b => b.file.setTrashed(true));
  }

  // 最終バックアップ情報を更新
  config.lastBackup = new Date().toISOString();
  config.lastBackupId = copy.getId();
  saveMaster({ backupConfig: config });

  return {
    success: true,
    backupName,
    fileId: copy.getId(),
    folderPath: config.folderPath,
    timestamp: config.lastBackup,
  };
}

/** 手動バックアップ（UIから実行） */
function createManualBackup() {
  const config = _getBackupConfig();
  return _executeBackup(config);
}

/** 自動バックアップ（トリガーから実行） */
function runScheduledBackup() {
  const config = _getBackupConfig();
  if (!config.enabled) return;
  _executeBackup(config);
}

/** バックアップ状態を取得 */
function getBackupStatus() {
  const config = _getBackupConfig();

  // トリガーの有無を確認
  const triggers = ScriptApp.getProjectTriggers();
  const hasBackupTrigger = triggers.some(t => t.getHandlerFunction() === 'runScheduledBackup');

  // バックアップフォルダの情報
  let folderUrl = null;
  let backupCount = 0;
  try {
    const folder = _getOrCreateBackupFolder(config.folderPath);
    folderUrl = folder.getUrl();
    const files = folder.getFiles();
    while (files.hasNext()) {
      if (files.next().getName().includes('_backup_')) backupCount++;
    }
  } catch (e) { /* フォルダ未作成 */ }

  // シート保護状態
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [SHEET_TRANSACTIONS, SHEET_SETTINGS, SHEET_MASTER, SHEET_CUSTOMERS, SHEET_APPOINTMENTS];
  const protectionStatus = {};
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protectionStatus[name] = protections.length > 0;
    }
  });

  return {
    backup: {
      enabled:       config.enabled,
      folderPath:    config.folderPath,
      maxKeep:       config.maxKeep,
      lastBackup:    config.lastBackup,
      lastBackupId:  config.lastBackupId,
      triggerActive: hasBackupTrigger,
      folderUrl,
      backupCount,
    },
    protection: protectionStatus,
  };
}

/** 日次バックアップトリガーをセットアップ */
function setupBackupTrigger() {
  // 既存のバックアップトリガーを削除
  removeBackupTrigger();

  // 毎日午前3時に実行
  ScriptApp.newTrigger('runScheduledBackup')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();

  // 設定を有効化
  const config = _getBackupConfig();
  config.enabled = true;
  saveMaster({ backupConfig: config });

  return { success: true, message: '日次バックアップトリガーを設定しました（毎日3:00）' };
}

/** バックアップトリガーを削除 */
function removeBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runScheduledBackup') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const config = _getBackupConfig();
  config.enabled = false;
  saveMaster({ backupConfig: config });

  return { success: true, message: 'バックアップトリガーを解除しました' };
}

// ─────────────────────────────────────────
//  予約リマインダー
// ─────────────────────────────────────────

/** 前日リマインダー送信（GASトリガーで毎朝実行） */
function sendDailyReminders() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const result = getAppointments(tomorrowStr, tomorrowStr);
  const sheet  = getOrCreateSheet(SHEET_APPOINTMENTS);
  const rows   = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const idIdx   = headers.indexOf('id');
  const rrIdx   = headers.indexOf('reminderSent');

  result.appointments.forEach(a => {
    if (a.status !== 'confirmed') return;
    if (a.reminderSent === true)  return;
    if (!a.email) return;

    try {
      MailApp.sendEmail(
        a.email,
        `【明日のご予約】${a.menuName}`,
        `${a.customerName} 様\n\n明日のご予約のご確認です。\n\nメニュー: ${a.menuName}\n日時: ${a.dateTime}\n\nご不明な点はご連絡ください。\n`
      );
      // reminderSent = true にマーク
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][idIdx]) === a.id) {
          sheet.getRange(i + 1, rrIdx + 1).setValue(true);
          break;
        }
      }
    } catch(e) { Logger.log('リマインダー失敗: ' + e.message); }
  });
}
