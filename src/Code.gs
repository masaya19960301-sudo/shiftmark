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
  LOCKS: 'Locks',
  USER_DEFAULTS: 'UserDefaults',
  ANNOUNCEMENTS: 'Announcements',
  SHIFT_SUBMISSIONS: 'ShiftSubmissions'
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
      'paidLeaveHoursPerDay', 'includePaidLeaveInDaysOff', 'active',
      'displayOrder', 'paidLeaveStartTime', 'paidLeaveEndTime'
    ]);
    // デフォルト管理者を作成（社員番号: admin, パスワード: admin）
    const defaultHash = hashPassword('admin');
    usersSheet.appendRow([
      'admin', '管理者', 'admin', '社員', defaultHash, 8, false, true, 0, '', ''
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
    optSheet.appendRow(['optionId', 'label', 'startTime', 'endTime', 'sortOrder', 'active', 'code']);
    // デフォルト選択肢
    const defaults = [
      ['opt1', '9:00-18:00', '09:00', '18:00', 1, true, '441Z'],
      ['opt2', '8:00-17:00', '08:00', '17:00', 2, true, '441X'],
      ['opt3', '9:00-14:00', '09:00', '14:00', 3, true, '441Z5'],
      ['opt4', '公休', '', '', 10, true, '公休'],
      ['opt5', '有休', '', '', 11, true, '有休']
    ];
    defaults.forEach(row => optSheet.appendRow(row));
    // 時刻列をテキスト形式に設定
    const optLastRow = optSheet.getLastRow();
    if (optLastRow > 1) {
      optSheet.getRange(2, 3, optLastRow - 1, 2).setNumberFormat('@');
    }
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

  // UserDefaults シート（ユーザーごとの曜日別デフォルトシフト）
  const defaultsSheet = getOrCreateSheet(SHEET_NAMES.USER_DEFAULTS);
  if (defaultsSheet.getLastRow() === 0) {
    defaultsSheet.appendRow(['employeeId', 'dayOfWeek', 'shiftOptionId']);
  }

  // Announcements シート
  const annoSheet = getOrCreateSheet(SHEET_NAMES.ANNOUNCEMENTS);
  if (annoSheet.getLastRow() === 0) {
    annoSheet.appendRow(['yearMonth', 'message']);
  }

  return { success: true, message: '初期セットアップが完了しました' };
}

/**
 * パスワードハッシュ化（引数は必ず文字列化）
 */
function hashPassword(password) {
  const pw = String(password);
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pw + 'shiftapp_salt_v1',
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
  if (!date) return '';
  if (typeof date === 'string') return date;
  if (date instanceof Date) {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const day = ('0' + date.getDate()).slice(-2);
    return year + '-' + month + '-' + day;
  }
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
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

/**
 * 時刻値を"HH:MM"文字列に正規化
 * Google Sheetsが"09:00"をDate型に自動変換する問題に対応
 */
function normalizeTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return ('0' + val.getHours()).slice(-2) + ':' + ('0' + val.getMinutes()).slice(-2);
  }
  var s = String(val);
  // "1899-12-30T09:00:00.000Z" のようなISO文字列対応
  if (s.includes('T') && s.includes(':')) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
    }
  }
  // 既にHH:MM形式ならそのまま
  var m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    return ('0' + m[1]).slice(-2) + ':' + m[2];
  }
  return s;
}

/**
 * yearMonth値を正規化（Date/数値/文字列→"YYYY-MM"文字列）
 * Google Sheetsが"2025-03"を日付型に自動変換する問題に対応
 */
function normalizeYearMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2);
  }
  // 文字列の場合そのまま（"2025-03"形式を想定）
  return String(val);
}

/**
 * デフォルト反映済みフラグを設定
 */
function markDefaultsGeneratedForMonth(yearMonth) {
  PropertiesService.getScriptProperties().setProperty('defaults_gen_' + normalizeYearMonth(yearMonth), 'true');
}

/**
 * デフォルト反映済みかどうかを確認
 */
function isDefaultsGeneratedForMonth(yearMonth) {
  return PropertiesService.getScriptProperties().getProperty('defaults_gen_' + normalizeYearMonth(yearMonth)) === 'true';
}
