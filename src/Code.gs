/**
 * シフト表作成Webアプリ (GAS) Ver1.0
 * メインエントリーポイント
 */

// スプレッドシートID（デプロイ時に設定）
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';

// シート名定数
const SHEET_NAMES = {
  USERS: 'Users',
  MONTH_DEFINITION: 'MonthDefinition',
  REQUIRED_HOLIDAYS: 'RequiredHolidays',
  HOLIDAYS: 'Holidays',
  SHIFT_OPTIONS: 'ShiftOptions',
  SHIFT_REQUESTS: 'ShiftRequests',
  SHIFT_FINAL: 'ShiftFinal',
  LOCKS: 'Locks'
};

/**
 * Webアプリのエントリーポイント（GET）
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('シフト表管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLファイルのインクルード用
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シートを取得（なければ作成）
 */
function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 初期セットアップ：全シートとヘッダーを作成
 */
function setupSpreadsheet() {
  // Users シート
  const usersSheet = getOrCreateSheet(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() === 0) {
    usersSheet.appendRow([
      'employeeId', 'name', 'role', 'category', 'password',
      'paidLeaveHoursPerDay', 'includePaidLeaveInDaysOff', 'active'
    ]);
    // デフォルト管理者を作成（社員番号: admin, パスワード: admin）
    const defaultHash = hashPassword('admin');
    usersSheet.appendRow([
      'admin', '管理者', 'admin', '社員', defaultHash, 8, false, true
    ]);
  }

  // MonthDefinition シート
  const monthDefSheet = getOrCreateSheet(SHEET_NAMES.MONTH_DEFINITION);
  if (monthDefSheet.getLastRow() === 0) {
    monthDefSheet.appendRow([
      'yearMonth', 'startDate', 'endDate', 'effectiveFrom'
    ]);
  }

  // RequiredHolidays シート
  const reqHolSheet = getOrCreateSheet(SHEET_NAMES.REQUIRED_HOLIDAYS);
  if (reqHolSheet.getLastRow() === 0) {
    reqHolSheet.appendRow(['yearMonth', 'requiredDaysOff']);
  }

  // Holidays シート
  const holSheet = getOrCreateSheet(SHEET_NAMES.HOLIDAYS);
  if (holSheet.getLastRow() === 0) {
    holSheet.appendRow(['date']);
  }

  // ShiftOptions シート
  const optSheet = getOrCreateSheet(SHEET_NAMES.SHIFT_OPTIONS);
  if (optSheet.getLastRow() === 0) {
    optSheet.appendRow(['optionId', 'label', 'startTime', 'endTime', 'sortOrder', 'active']);
    // デフォルト選択肢
    const defaults = [
      ['opt1', '9:00-18:00', '09:00', '18:00', 1, true],
      ['opt2', '8:00-17:00', '08:00', '17:00', 2, true],
      ['opt3', '9:00-14:00', '09:00', '14:00', 3, true],
      ['opt4', '公休', '', '', 10, true],
      ['opt5', '有休', '', '', 11, true]
    ];
    defaults.forEach(row => optSheet.appendRow(row));
  }

  // ShiftRequests シート
  const reqSheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  if (reqSheet.getLastRow() === 0) {
    reqSheet.appendRow(['yearMonth', 'employeeId', 'date', 'shiftOptionId', 'updatedAt']);
  }

  // ShiftFinal シート
  const finalSheet = getOrCreateSheet(SHEET_NAMES.SHIFT_FINAL);
  if (finalSheet.getLastRow() === 0) {
    finalSheet.appendRow(['yearMonth', 'employeeId', 'date', 'shiftOptionId', 'updatedAt']);
  }

  // Locks シート
  const locksSheet = getOrCreateSheet(SHEET_NAMES.LOCKS);
  if (locksSheet.getLastRow() === 0) {
    locksSheet.appendRow(['yearMonth', 'locked', 'lockedAt', 'lockedBy']);
  }

  return { success: true, message: '初期セットアップが完了しました' };
}

/**
 * パスワードハッシュ化
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + 'shiftapp_salt_v1',
    Utilities.Charset.UTF_8
  );
  return rawHash.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/**
 * 日付をyyyy-MM-dd文字列に変換
 */
function formatDate(date) {
  if (typeof date === 'string') return date;
  const d = new Date(date);
  const year = d.getFullYear();
  const month = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

/**
 * yyyy-MM形式の年月文字列を生成
 */
function formatYearMonth(year, month) {
  return year + '-' + ('0' + month).slice(-2);
}
