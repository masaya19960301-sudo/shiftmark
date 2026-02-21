/**
 * データアクセス層
 * スプレッドシートのCRUD操作
 */

// ========== Users ==========

/**
 * 社員番号でユーザーを検索
 */
function findUserByEmployeeId(employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  const targetId = String(employeeId);
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === targetId) {
      const user = {};
      headers.forEach((h, idx) => { user[h] = data[i][idx]; });
      // employeeIdは必ず文字列として保持
      user.employeeId = String(user.employeeId);
      user._row = i + 1;
      return user;
    }
  }
  return null;
}

/**
 * 全ユーザー取得（アクティブのみ）
 */
function getAllActiveUsers() {
  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  ensureUsersHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const user = {};
    headers.forEach((h, idx) => { user[h] = data[i][idx]; });
    if (user.active === true || user.active === 'TRUE' || user.active === 'true') {
      delete user.password;
      user.employeeId = String(user.employeeId);
      users.push(user);
    }
  }
  users.sort((a, b) => ((a.displayOrder || 0) - (b.displayOrder || 0)));
  return users;
}

/**
 * 全ユーザー取得（管理者用、非アクティブ含む）
 */
function getAllUsers() {
  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  ensureUsersHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const user = {};
    headers.forEach((h, idx) => { user[h] = data[i][idx]; });
    delete user.password;
    user.employeeId = String(user.employeeId);
    user._row = i + 1;
    users.push(user);
  }
  return users;
}

/**
 * ユーザー追加
 */
function addUser(userData) {
  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  const empId = String(userData.employeeId);
  const existing = findUserByEmployeeId(empId);
  if (existing) {
    return { success: false, message: 'この社員番号は既に登録されています' };
  }

  const passwordHash = hashPassword(String(userData.password || 'password'));
  sheet.appendRow([
    empId,
    userData.name,
    userData.role || 'user',
    userData.category || '社員',
    passwordHash,
    userData.paidLeaveHoursPerDay || 8,
    userData.includePaidLeaveInDaysOff || false,
    true,
    userData.displayOrder || 0,
    userData.paidLeaveStartTime || '',
    userData.paidLeaveEndTime || ''
  ]);
  // 社員番号のセルをテキスト形式に設定し、値を再セット（先頭ゼロ保持）
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('@').setValue(empId);
  sheet.getRange(lastRow, 10).setNumberFormat('@');
  sheet.getRange(lastRow, 11).setNumberFormat('@');

  return { success: true, message: 'ユーザーを登録しました' };
}

/**
 * ユーザー削除
 */
function deleteUser(employeeId) {
  const user = findUserByEmployeeId(employeeId);
  if (!user) {
    return { success: false, message: 'ユーザーが見つかりません' };
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  sheet.deleteRow(user._row);
  SpreadsheetApp.flush();
  return { success: true, message: 'ユーザーを削除しました' };
}

/**
 * ユーザー更新
 */
function updateUser(employeeId, userData) {
  const user = findUserByEmployeeId(employeeId);
  if (!user) {
    return { success: false, message: 'ユーザーが見つかりません' };
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  const row = user._row;
  sheet.getRange(row, 2).setValue(userData.name);
  sheet.getRange(row, 3).setValue(userData.role);
  sheet.getRange(row, 4).setValue(userData.category);
  sheet.getRange(row, 6).setValue(userData.paidLeaveHoursPerDay);
  sheet.getRange(row, 7).setValue(userData.includePaidLeaveInDaysOff);
  sheet.getRange(row, 8).setValue(userData.active !== false);
  sheet.getRange(row, 9).setValue(userData.displayOrder || 0);
  sheet.getRange(row, 10).setNumberFormat('@').setValue(userData.paidLeaveStartTime || '');
  sheet.getRange(row, 11).setNumberFormat('@').setValue(userData.paidLeaveEndTime || '');

  return { success: true, message: 'ユーザー情報を更新しました' };
}

/**
 * ユーザーパスワード更新
 */
function updateUserPassword(employeeId, newPasswordHash) {
  const user = findUserByEmployeeId(employeeId);
  if (!user) return false;

  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
  sheet.getRange(user._row, 5).setValue(newPasswordHash);
  return true;
}

/**
 * 管理者によるパスワードリセット
 */
function resetUserPassword(employeeId, newPassword) {
  const newHash = hashPassword(String(newPassword));
  return updateUserPassword(employeeId, newHash);
}

// ========== MonthDefinition ==========

/**
 * 月度定義を取得
 */
function getMonthDefinitions() {
  const sheet = getOrCreateSheet(SHEET_NAMES.MONTH_DEFINITION);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const defs = [];
  for (let i = 1; i < data.length; i++) {
    const def = {};
    headers.forEach((h, idx) => {
      if (h === 'startDate' || h === 'endDate') {
        def[h] = formatDate(data[i][idx]);
      } else if (h === 'yearMonth' || h === 'effectiveFrom') {
        def[h] = normalizeYearMonth(data[i][idx]);
      } else {
        def[h] = data[i][idx];
      }
    });
    def._row = i + 1;
    defs.push(def);
  }
  return defs;
}

/**
 * 特定年月の月度定義を取得
 * 完全一致がない場合、effectiveFromが設定された定義から日付パターンを継承
 */
function getMonthDefinition(yearMonth) {
  var target = normalizeYearMonth(yearMonth);
  var defs = getMonthDefinitions();

  // 完全一致を優先
  var exact = defs.find(function(d) { return d.yearMonth === target; });
  if (exact) return exact;

  // effectiveFromが設定されていて対象月以前の定義を検索
  var applicable = defs.filter(function(d) {
    return d.effectiveFrom && d.effectiveFrom <= target;
  });

  if (applicable.length === 0) return null;

  // 最も新しいeffectiveFromの定義をテンプレートとして使用
  applicable.sort(function(a, b) {
    return b.effectiveFrom.localeCompare(a.effectiveFrom);
  });
  var template = applicable[0];

  // テンプレートの月度から対象月までの月数差分を計算
  var tParts = template.yearMonth.split('-');
  var tYear = parseInt(tParts[0]);
  var tMonth = parseInt(tParts[1]);

  var targetParts = target.split('-');
  var targetYear = parseInt(targetParts[0]);
  var targetMonth = parseInt(targetParts[1]);

  var monthDiff = (targetYear * 12 + targetMonth) - (tYear * 12 + tMonth);

  // テンプレートの開始日・終了日を月数分ずらして適用
  var tStart = new Date(template.startDate);
  var tEnd = new Date(template.endDate);

  var newStart = new Date(tStart.getFullYear(), tStart.getMonth() + monthDiff, tStart.getDate());
  var newEnd = new Date(tEnd.getFullYear(), tEnd.getMonth() + monthDiff, tEnd.getDate());

  return {
    yearMonth: target,
    startDate: formatDate(newStart),
    endDate: formatDate(newEnd),
    effectiveFrom: template.effectiveFrom,
    _row: template._row
  };
}

/**
 * 月度定義を保存/更新
 */
function saveMonthDefinition(yearMonth, startDate, endDate, effectiveFrom) {
  const sheet = getOrCreateSheet(SHEET_NAMES.MONTH_DEFINITION);
  const ymStr = String(yearMonth);
  const existing = getMonthDefinition(ymStr);

  if (existing) {
    sheet.getRange(existing._row, 2).setValue(startDate);
    sheet.getRange(existing._row, 3).setValue(endDate);
    sheet.getRange(existing._row, 4).setValue(effectiveFrom || '');
    // yearMonthと effectiveFromをテキスト形式に
    sheet.getRange(existing._row, 1).setNumberFormat('@');
    sheet.getRange(existing._row, 4).setNumberFormat('@');
  } else {
    sheet.appendRow([ymStr, startDate, endDate, effectiveFrom || '']);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
    sheet.getRange(lastRow, 4).setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * 年月から日付範囲を取得（月度定義がなければカレンダー月を使用）
 */
function getDateRangeForYearMonth(yearMonth) {
  const def = getMonthDefinition(yearMonth);
  if (def) {
    return { startDate: def.startDate, endDate: def.endDate };
  }

  // デフォルト：カレンダー月
  const ymStr = String(yearMonth);
  const parts = ymStr.split('-');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);
  const startDate = formatDate(new Date(year, month - 1, 1));
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = formatDate(new Date(year, month - 1, lastDay));
  return { startDate: startDate, endDate: endDate };
}

// ========== RequiredHolidays ==========

/**
 * 必要公休日数を取得
 */
function getRequiredHolidays(yearMonth) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REQUIRED_HOLIDAYS);
  const data = sheet.getDataRange().getValues();
  const target = normalizeYearMonth(yearMonth);
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === target) {
      return { yearMonth: target, requiredDaysOff: data[i][1], _row: i + 1 };
    }
  }
  return null;
}

/**
 * 全ての必要公休日設定を取得
 */
function getAllRequiredHolidays() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REQUIRED_HOLIDAYS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      yearMonth: normalizeYearMonth(data[i][0]),
      requiredDaysOff: data[i][1]
    });
  }
  return result;
}

/**
 * 必要公休日数を保存
 */
function saveRequiredHolidays(yearMonth, requiredDaysOff) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REQUIRED_HOLIDAYS);
  const ymStr = String(yearMonth);
  const existing = getRequiredHolidays(ymStr);
  if (existing) {
    sheet.getRange(existing._row, 2).setValue(requiredDaysOff);
  } else {
    sheet.appendRow([ymStr, requiredDaysOff]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// ========== Holidays ==========

/**
 * 祝日一覧取得
 */
function getHolidays() {
  const sheet = getOrCreateSheet(SHEET_NAMES.HOLIDAYS);
  const data = sheet.getDataRange().getValues();
  const holidays = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      holidays.push(formatDate(data[i][0]));
    }
  }
  return holidays;
}

/**
 * 祝日登録（一括）
 */
function registerHolidays(dateList) {
  const sheet = getOrCreateSheet(SHEET_NAMES.HOLIDAYS);
  // 既存データをクリアしてヘッダーのみ残す
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).clearContent();
  }
  // 新しいデータを追加
  dateList.forEach(date => {
    sheet.appendRow([date]);
  });
  SpreadsheetApp.flush();
  return { success: true, count: dateList.length };
}

/**
 * 祝日追加（追記）
 */
function addHolidays(dateList) {
  const sheet = getOrCreateSheet(SHEET_NAMES.HOLIDAYS);
  const existing = getHolidays();
  const added = [];
  dateList.forEach(date => {
    const formatted = formatDate(date);
    if (!existing.includes(formatted)) {
      sheet.appendRow([formatted]);
      added.push(formatted);
    }
  });
  SpreadsheetApp.flush();
  return { success: true, addedCount: added.length };
}

// ========== ShiftOptions ==========

/**
 * ShiftOptionsシートのヘッダーに不足カラムがあれば追加
 */
function ensureShiftOptionsHeaders_(sheet) {
  if (sheet.getLastRow() === 0) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('code') === -1) {
    const nextCol = headers.length + 1;
    sheet.getRange(1, nextCol).setValue('code');
  }
}

/**
 * Usersシートのヘッダーに不足カラムがあれば追加
 */
function ensureUsersHeaders_(sheet) {
  if (sheet.getLastRow() === 0) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = ['displayOrder', 'paidLeaveStartTime', 'paidLeaveEndTime'];
  needed.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      headers.push(col);
    }
  });
}

/**
 * 勤務時間候補取得（アクティブのみ）
 */
function getActiveShiftOptions() {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_OPTIONS);
  ensureShiftOptionsHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const options = [];
  for (let i = 1; i < data.length; i++) {
    const opt = {};
    headers.forEach((h, idx) => { opt[h] = data[i][idx]; });
    opt.optionId = String(opt.optionId);
    opt.startTime = normalizeTime(opt.startTime);
    opt.endTime = normalizeTime(opt.endTime);
    if (opt.active === true || opt.active === 'TRUE' || opt.active === 'true') {
      options.push(opt);
    }
  }
  options.sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
  return options;
}

/**
 * 全勤務時間候補取得（管理者用）
 */
function getAllShiftOptions() {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_OPTIONS);
  ensureShiftOptionsHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const options = [];
  for (let i = 1; i < data.length; i++) {
    const opt = {};
    headers.forEach((h, idx) => { opt[h] = data[i][idx]; });
    opt.optionId = String(opt.optionId);
    opt.startTime = normalizeTime(opt.startTime);
    opt.endTime = normalizeTime(opt.endTime);
    opt._row = i + 1;
    options.push(opt);
  }
  options.sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
  return options;
}

/**
 * 勤務時間候補を保存
 */
function saveShiftOption(optionData) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_OPTIONS);
  const all = getAllShiftOptions();
  const existing = all.find(o => o.optionId === String(optionData.optionId));
  const st = normalizeTime(optionData.startTime);
  const et = normalizeTime(optionData.endTime);

  if (existing) {
    const row = existing._row;
    sheet.getRange(row, 2).setValue(optionData.label);
    sheet.getRange(row, 3).setNumberFormat('@').setValue(st);
    sheet.getRange(row, 4).setNumberFormat('@').setValue(et);
    sheet.getRange(row, 5).setValue(optionData.sortOrder || 0);
    sheet.getRange(row, 6).setValue(optionData.active !== false);
    sheet.getRange(row, 7).setValue(optionData.code || '');
  } else {
    sheet.appendRow([
      optionData.optionId || Utilities.getUuid(),
      optionData.label,
      st,
      et,
      optionData.sortOrder || 0,
      optionData.active !== false,
      optionData.code || ''
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat('@');
    sheet.getRange(lastRow, 4).setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// ========== UserDefaults（ユーザー別デフォルトシフト） ==========

/**
 * ユーザーのデフォルトシフトを取得
 * 戻り値: { 0: optionId, 1: optionId, ... 6: optionId } (0=日, 6=土)
 */
function getUserDefaults(employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.USER_DEFAULTS);
  const data = sheet.getDataRange().getValues();
  const targetEmp = String(employeeId);
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === targetEmp) {
      result[String(data[i][1])] = String(data[i][2]);
    }
  }
  return result;
}

/**
 * 全ユーザーのデフォルトシフトを取得
 */
function getAllUserDefaults() {
  const sheet = getOrCreateSheet(SHEET_NAMES.USER_DEFAULTS);
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const emp = String(data[i][0]);
    if (!result[emp]) result[emp] = {};
    result[emp][String(data[i][1])] = String(data[i][2]);
  }
  return result;
}

/**
 * ユーザーのデフォルトシフトを保存（全曜日一括）
 * defaults: { "0": optionId, "1": optionId, ... "6": optionId }
 */
function saveUserDefaults(employeeId, defaults) {
  const sheet = getOrCreateSheet(SHEET_NAMES.USER_DEFAULTS);
  const data = sheet.getDataRange().getValues();
  const targetEmp = String(employeeId);

  // 既存レコードを削除（下から削除して行番号ズレ防止）
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === targetEmp) {
      sheet.deleteRow(i + 1);
    }
  }

  // 新しいデフォルトを追加
  for (let day = 0; day <= 6; day++) {
    const optId = defaults[String(day)] || '';
    if (optId) {
      sheet.appendRow([targetEmp, day, optId]);
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 1).setNumberFormat('@').setValue(targetEmp);
    }
  }
  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * デフォルトシフトから希望シフトを自動生成
 * バッチ処理で高速化
 */
function generateShiftFromDefaults(yearMonth, employeeId) {
  const defaults = getUserDefaults(employeeId);
  if (Object.keys(defaults).length === 0) {
    return { success: false, message: 'デフォルトシフトが設定されていません' };
  }

  const dateRange = getDateRangeForYearMonth(yearMonth);
  const dates = generateDateList(dateRange.startDate, dateRange.endDate);
  const empId = String(employeeId);

  const entries = [];
  dates.forEach(function(dateStr) {
    const dow = new Date(dateStr).getDay();
    const optId = defaults[String(dow)] || '';
    if (optId) {
      entries.push({ employeeId: empId, date: dateStr, shiftOptionId: optId });
    }
  });

  if (entries.length > 0) {
    batchSaveShiftRequests(yearMonth, entries);
  }

  return { success: true, count: entries.length };
}

/**
 * デフォルトシフトから希望シフトを全ユーザー一括自動生成（管理者用）
 * バッチ処理で高速化: シートの読み書きを1回にまとめる
 */
function generateAllRequestsFromDefaults(yearMonth) {
  const users = getAllActiveUsers();
  const allDefaults = getAllUserDefaults();
  const dateRange = getDateRangeForYearMonth(yearMonth);
  const dates = generateDateList(dateRange.startDate, dateRange.endDate);

  const entries = [];
  users.forEach(function(user) {
    const empId = String(user.employeeId);
    const defaults = allDefaults[empId];
    if (!defaults || Object.keys(defaults).length === 0) return;

    dates.forEach(function(dateStr) {
      const dow = new Date(dateStr).getDay();
      const optId = defaults[String(dow)] || '';
      if (optId) {
        entries.push({ employeeId: empId, date: dateStr, shiftOptionId: optId });
      }
    });
  });

  if (entries.length > 0) {
    batchSaveShiftRequests(yearMonth, entries);
  }

  markDefaultsGeneratedForMonth(yearMonth);
  return { success: true, count: entries.length };
}

/**
 * 既存の希望シフトがデフォルトと異なるかチェック
 */
function checkRequestsVsDefaults(yearMonth) {
  const allRequests = getShiftRequests(yearMonth);
  if (allRequests.length === 0) {
    return { hasModified: false, requestCount: 0, modifiedCount: 0 };
  }

  const allDefaults = getAllUserDefaults();

  let modifiedCount = 0;
  allRequests.forEach(req => {
    const defaults = allDefaults[req.employeeId];
    if (!defaults) { modifiedCount++; return; }

    const dt = new Date(req.date);
    const dow = dt.getDay();
    const expectedOptId = defaults[String(dow)] || '';

    if (req.shiftOptionId !== expectedOptId) {
      modifiedCount++;
    }
  });

  return { hasModified: modifiedCount > 0, modifiedCount: modifiedCount, requestCount: allRequests.length };
}

/**
 * デフォルトシフトから確定シフトを自動生成（管理者用：全ユーザー一括）
 * バッチ処理で高速化
 */
function generateFinalFromDefaults(yearMonth) {
  const users = getAllActiveUsers();
  const allDefaults = getAllUserDefaults();
  const dateRange = getDateRangeForYearMonth(yearMonth);
  const dates = generateDateList(dateRange.startDate, dateRange.endDate);

  const entries = [];
  users.forEach(function(user) {
    const empId = String(user.employeeId);
    const defaults = allDefaults[empId];
    if (!defaults || Object.keys(defaults).length === 0) return;

    dates.forEach(function(dateStr) {
      const dow = new Date(dateStr).getDay();
      const optId = defaults[String(dow)] || '';
      if (optId) {
        entries.push({ employeeId: empId, date: dateStr, shiftOptionId: optId });
      }
    });
  });

  if (entries.length > 0) {
    batchSaveShiftFinal(yearMonth, entries);
  }

  return { success: true, count: entries.length };
}

// ========== バッチ書き込み ==========

/**
 * 希望シフトを一括保存（バッチ処理）
 * entries: [{ employeeId, date, shiftOptionId }, ...]
 * シート全体を1回だけ読み書きするため高速
 */
function batchSaveShiftRequests(yearMonth, entries) {
  var sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  var rawData = sheet.getDataRange().getValues();
  var targetYM = normalizeYearMonth(yearMonth);
  var now = new Date().toISOString();

  // GASがDate型に自動変換するため、全データを文字列に正規化
  var allData = [];
  for (var i = 0; i < rawData.length; i++) {
    if (i === 0) {
      allData.push(rawData[i]); // ヘッダーはそのまま
    } else {
      allData.push([
        normalizeYearMonth(rawData[i][0]),
        String(rawData[i][1]),
        formatDate(rawData[i][2]),
        String(rawData[i][3]),
        rawData[i][4] ? String(rawData[i][4]) : ''
      ]);
    }
  }

  // 既存レコードのインデックスを構築（employeeId|date -> 行index）
  var indexMap = {};
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === targetYM) {
      var key = allData[i][1] + '|' + allData[i][2];
      indexMap[key] = i;
    }
  }

  // エントリを処理：既存は上書き、新規は追記リストに追加
  var appendRows = [];
  entries.forEach(function(entry) {
    var empId = String(entry.employeeId);
    var dateStr = formatDate(entry.date);
    var optId = String(entry.shiftOptionId);
    var k = empId + '|' + dateStr;

    if (indexMap[k] !== undefined) {
      var idx = indexMap[k];
      allData[idx][3] = optId;
      allData[idx][4] = now;
    } else {
      appendRows.push([String(yearMonth), empId, dateStr, optId, now]);
    }
  });

  // 既存データを一括書き戻し（正規化済みの全データ）
  if (allData.length > 0) {
    // yearMonthとemployeeId列をテキスト形式に設定（先頭ゼロ落ち防止のためsetValuesより先に実行）
    sheet.getRange(1, 1, allData.length, 1).setNumberFormat('@');
    sheet.getRange(1, 2, allData.length, 1).setNumberFormat('@');
    sheet.getRange(1, 1, allData.length, 5).setValues(allData);
  }

  // 新規行を一括追記
  if (appendRows.length > 0) {
    var startRow = allData.length + 1;
    sheet.getRange(startRow, 1, appendRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 2, appendRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, appendRows.length, 5).setValues(appendRows);
  }

  SpreadsheetApp.flush();
}

/**
 * 確定シフトを一括保存（バッチ処理）
 * entries: [{ employeeId, date, shiftOptionId }, ...]
 */
function batchSaveShiftFinal(yearMonth, entries) {
  var sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_FINAL);
  var rawData = sheet.getDataRange().getValues();
  var targetYM = normalizeYearMonth(yearMonth);
  var now = new Date().toISOString();

  // GASがDate型に自動変換するため、全データを文字列に正規化
  var allData = [];
  for (var i = 0; i < rawData.length; i++) {
    if (i === 0) {
      allData.push(rawData[i]);
    } else {
      allData.push([
        normalizeYearMonth(rawData[i][0]),
        String(rawData[i][1]),
        formatDate(rawData[i][2]),
        String(rawData[i][3]),
        rawData[i][4] ? String(rawData[i][4]) : ''
      ]);
    }
  }

  var indexMap = {};
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === targetYM) {
      var key = allData[i][1] + '|' + allData[i][2];
      indexMap[key] = i;
    }
  }

  var appendRows = [];
  entries.forEach(function(entry) {
    var empId = String(entry.employeeId);
    var dateStr = formatDate(entry.date);
    var optId = String(entry.shiftOptionId);
    var k = empId + '|' + dateStr;

    if (indexMap[k] !== undefined) {
      var idx = indexMap[k];
      allData[idx][3] = optId;
      allData[idx][4] = now;
    } else {
      appendRows.push([String(yearMonth), empId, dateStr, optId, now]);
    }
  });

  if (allData.length > 0) {
    // yearMonthとemployeeId列をテキスト形式に設定（先頭ゼロ落ち防止のためsetValuesより先に実行）
    sheet.getRange(1, 1, allData.length, 1).setNumberFormat('@');
    sheet.getRange(1, 2, allData.length, 1).setNumberFormat('@');
    sheet.getRange(1, 1, allData.length, 5).setValues(allData);
  }

  if (appendRows.length > 0) {
    var startRow = allData.length + 1;
    sheet.getRange(startRow, 1, appendRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 2, appendRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, appendRows.length, 5).setValues(appendRows);
  }

  SpreadsheetApp.flush();
}

// ========== ShiftRequests ==========

/**
 * 希望シフトを取得（年月×社員）
 * 同一社員×日付の重複エントリがある場合は最後の1件のみ保持
 */
function getShiftRequests(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = employeeId ? String(employeeId) : null;
  const resultMap = {};
  for (let i = 1; i < data.length; i++) {
    const rowYM = normalizeYearMonth(data[i][0]);
    const rowEmp = String(data[i][1]);
    if (rowYM === targetYM && (!targetEmp || rowEmp === targetEmp)) {
      const dateStr = formatDate(data[i][2]);
      const key = rowEmp + '_' + dateStr;
      resultMap[key] = {
        yearMonth: rowYM,
        employeeId: rowEmp,
        date: dateStr,
        shiftOptionId: String(data[i][3]),
        updatedAt: data[i][4]
      };
    }
  }
  return Object.values(resultMap);
}

/**
 * 希望シフトを保存（1セル分）
 */
function saveShiftRequest(yearMonth, employeeId, date, shiftOptionId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const formattedDate = formatDate(date);
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = String(employeeId);

  // 既存レコードを探す
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === targetYM &&
        String(data[i][1]) === targetEmp &&
        formatDate(data[i][2]) === formattedDate) {
      if (shiftOptionId === '' || shiftOptionId === null || shiftOptionId === undefined) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(String(shiftOptionId));
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      }
      SpreadsheetApp.flush();
      return { success: true };
    }
  }

  // 新規追加
  if (shiftOptionId && shiftOptionId !== '') {
    sheet.appendRow([String(yearMonth), targetEmp, formattedDate, String(shiftOptionId), new Date().toISOString()]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
    sheet.getRange(lastRow, 2).setNumberFormat('@').setValue(targetEmp);
    SpreadsheetApp.flush();
  }
  return { success: true };
}

// ========== ShiftFinal ==========

/**
 * 確定シフトを取得
 * 同一社員×日付の重複エントリがある場合は最後の1件のみ保持
 */
function getShiftFinal(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_FINAL);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = employeeId ? String(employeeId) : null;
  const resultMap = {};
  for (let i = 1; i < data.length; i++) {
    const rowYM = normalizeYearMonth(data[i][0]);
    const rowEmp = String(data[i][1]);
    if (rowYM === targetYM && (!targetEmp || rowEmp === targetEmp)) {
      const dateStr = formatDate(data[i][2]);
      const key = rowEmp + '_' + dateStr;
      resultMap[key] = {
        yearMonth: rowYM,
        employeeId: rowEmp,
        date: dateStr,
        shiftOptionId: String(data[i][3]),
        updatedAt: data[i][4]
      };
    }
  }
  return Object.values(resultMap);
}

/**
 * 確定シフトを保存（1セル分）
 */
function saveShiftFinal(yearMonth, employeeId, date, shiftOptionId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_FINAL);
  const data = sheet.getDataRange().getValues();
  const formattedDate = formatDate(date);
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = String(employeeId);

  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === targetYM &&
        String(data[i][1]) === targetEmp &&
        formatDate(data[i][2]) === formattedDate) {
      if (shiftOptionId === '' || shiftOptionId === null || shiftOptionId === undefined) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(String(shiftOptionId));
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      }
      SpreadsheetApp.flush();
      return { success: true };
    }
  }

  if (shiftOptionId && shiftOptionId !== '') {
    sheet.appendRow([String(yearMonth), targetEmp, formattedDate, String(shiftOptionId), new Date().toISOString()]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
    sheet.getRange(lastRow, 2).setNumberFormat('@').setValue(targetEmp);
    SpreadsheetApp.flush();
  }
  return { success: true };
}

/**
 * 希望シフトを確定シフトにコピー（一括）
 * バッチ処理で高速化
 */
function copyRequestsToFinal(yearMonth) {
  const requests = getShiftRequests(yearMonth);

  const entries = requests.map(function(req) {
    return { employeeId: req.employeeId, date: req.date, shiftOptionId: req.shiftOptionId };
  });

  if (entries.length > 0) {
    batchSaveShiftFinal(yearMonth, entries);
  }

  return { success: true, count: requests.length };
}

// ========== Locks ==========

/**
 * 締切状態を取得
 */
function getLockStatus(yearMonth) {
  const sheet = getOrCreateSheet(SHEET_NAMES.LOCKS);
  const data = sheet.getDataRange().getValues();
  const target = normalizeYearMonth(yearMonth);
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === target) {
      return {
        yearMonth: target,
        locked: data[i][1] === true || data[i][1] === 'TRUE',
        lockedAt: data[i][2],
        lockedBy: String(data[i][3]),
        _row: i + 1
      };
    }
  }
  return { yearMonth: target, locked: false };
}

/**
 * 全ての締切状態を取得
 */
function getAllLocks() {
  const sheet = getOrCreateSheet(SHEET_NAMES.LOCKS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      yearMonth: normalizeYearMonth(data[i][0]),
      locked: data[i][1] === true || data[i][1] === 'TRUE',
      lockedAt: data[i][2],
      lockedBy: String(data[i][3])
    });
  }
  return result;
}

/**
 * 締切設定/解除
 */
function setLock(yearMonth, locked, adminEmployeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.LOCKS);
  const ymStr = String(yearMonth);
  const status = getLockStatus(ymStr);

  if (status._row) {
    sheet.getRange(status._row, 2).setValue(locked);
    sheet.getRange(status._row, 3).setValue(new Date().toISOString());
    sheet.getRange(status._row, 4).setValue(String(adminEmployeeId));
  } else {
    sheet.appendRow([ymStr, locked, new Date().toISOString(), String(adminEmployeeId)]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  return { success: true, locked: locked };
}

// ========== Announcements ==========

function getAnnouncement(yearMonth) {
  const sheet = getOrCreateSheet(SHEET_NAMES.ANNOUNCEMENTS);
  const data = sheet.getDataRange().getValues();
  const target = normalizeYearMonth(yearMonth);
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === target) {
      return { yearMonth: target, message: data[i][1] || '', _row: i + 1 };
    }
  }
  return null;
}

function getAllAnnouncements() {
  const sheet = getOrCreateSheet(SHEET_NAMES.ANNOUNCEMENTS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ yearMonth: normalizeYearMonth(data[i][0]), message: data[i][1] || '' });
  }
  return result;
}

function saveAnnouncement(yearMonth, message) {
  const sheet = getOrCreateSheet(SHEET_NAMES.ANNOUNCEMENTS);
  const ymStr = String(yearMonth);
  const existing = getAnnouncement(ymStr);
  if (existing) {
    sheet.getRange(existing._row, 2).setValue(message);
  } else {
    sheet.appendRow([ymStr, message]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// ========== ShiftSubmissions（希望提出状態） ==========

/**
 * 希望提出を記録
 */
function submitShiftRequest(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_SUBMISSIONS);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = String(employeeId);

  // 既存レコードを更新
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === targetYM && String(data[i][1]) === targetEmp) {
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return { success: true };
    }
  }

  // 新規追加
  sheet.appendRow([targetYM, targetEmp, new Date().toISOString()]);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('@');
  sheet.getRange(lastRow, 2).setNumberFormat('@').setValue(targetEmp);
  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * 希望提出状態を取得（月度単位、全ユーザー）
 */
function getShiftSubmissions(yearMonth) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_SUBMISSIONS);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (normalizeYearMonth(data[i][0]) === targetYM) {
      result[String(data[i][1])] = data[i][2]; // employeeId -> submittedAt
    }
  }
  return result;
}

/**
 * 希望提出を取消（シフト変更時に自動リセット）
 */
function clearShiftSubmission(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_SUBMISSIONS);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = String(employeeId);
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizeYearMonth(data[i][0]) === targetYM && String(data[i][1]) === targetEmp) {
      sheet.deleteRow(i + 1);
    }
  }
  SpreadsheetApp.flush();
}

// ========== データクリーンアップ ==========

/**
 * 1年以上前のシフトデータを削除
 * ShiftRequests, ShiftFinal, Locks, Announcements の yearMonth列を対象
 * GASのトリガーで月1回実行する想定
 * （スクリプトエディタ → トリガー → cleanupOldData を月次タイマーに設定）
 */
function cleanupOldData() {
  const now = new Date();
  const cutoff = new Date(now.getFullYear() - 1, now.getMonth(), 1);
  const cutoffYM = Utilities.formatDate(cutoff, 'Asia/Tokyo', 'yyyy-MM');

  const targets = [
    SHEET_NAMES.SHIFT_REQUESTS,
    SHEET_NAMES.SHIFT_FINAL,
    SHEET_NAMES.LOCKS,
    SHEET_NAMES.ANNOUNCEMENTS,
    SHEET_NAMES.SHIFT_SUBMISSIONS
  ];

  let totalDeleted = 0;
  targets.forEach(sheetName => {
    const sheet = getOrCreateSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    // 下の行から削除（行番号ズレ防止）
    for (let i = data.length - 1; i >= 1; i--) {
      const ym = normalizeYearMonth(data[i][0]);
      if (ym && ym < cutoffYM) {
        sheet.deleteRow(i + 1);
        totalDeleted++;
      }
    }
  });

  Logger.log('クリーンアップ完了: ' + totalDeleted + '行削除（基準: ' + cutoffYM + ' より前）');
  return { success: true, deleted: totalDeleted, cutoff: cutoffYM };
}
