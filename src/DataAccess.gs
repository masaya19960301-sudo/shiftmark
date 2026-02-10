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
  return users;
}

/**
 * 全ユーザー取得（管理者用、非アクティブ含む）
 */
function getAllUsers() {
  const sheet = getOrCreateSheet(SHEET_NAMES.USERS);
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
    true
  ]);
  // 社員番号のセルをテキスト形式に設定（数値変換防止）
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('@');

  return { success: true, message: 'ユーザーを登録しました' };
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
 */
function getMonthDefinition(yearMonth) {
  const target = normalizeYearMonth(yearMonth);
  const defs = getMonthDefinitions();
  return defs.find(d => d.yearMonth === target) || null;
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
 * 勤務時間候補取得（アクティブのみ）
 */
function getActiveShiftOptions() {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_OPTIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const options = [];
  for (let i = 1; i < data.length; i++) {
    const opt = {};
    headers.forEach((h, idx) => { opt[h] = data[i][idx]; });
    opt.optionId = String(opt.optionId);
    opt.startTime = String(opt.startTime || '');
    opt.endTime = String(opt.endTime || '');
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
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const options = [];
  for (let i = 1; i < data.length; i++) {
    const opt = {};
    headers.forEach((h, idx) => { opt[h] = data[i][idx]; });
    opt.optionId = String(opt.optionId);
    opt.startTime = String(opt.startTime || '');
    opt.endTime = String(opt.endTime || '');
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

  if (existing) {
    const row = existing._row;
    sheet.getRange(row, 2).setValue(optionData.label);
    sheet.getRange(row, 3).setValue(optionData.startTime || '');
    sheet.getRange(row, 4).setValue(optionData.endTime || '');
    sheet.getRange(row, 5).setValue(optionData.sortOrder || 0);
    sheet.getRange(row, 6).setValue(optionData.active !== false);
  } else {
    sheet.appendRow([
      optionData.optionId || Utilities.getUuid(),
      optionData.label,
      optionData.startTime || '',
      optionData.endTime || '',
      optionData.sortOrder || 0,
      optionData.active !== false
    ]);
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// ========== ShiftRequests ==========

/**
 * 希望シフトを取得（年月×社員）
 */
function getShiftRequests(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = employeeId ? String(employeeId) : null;
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const rowYM = normalizeYearMonth(data[i][0]);
    const rowEmp = String(data[i][1]);
    if (rowYM === targetYM && (!targetEmp || rowEmp === targetEmp)) {
      result.push({
        yearMonth: rowYM,
        employeeId: rowEmp,
        date: formatDate(data[i][2]),
        shiftOptionId: String(data[i][3]),
        updatedAt: data[i][4]
      });
    }
  }
  return result;
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
    sheet.getRange(lastRow, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
  }
  return { success: true };
}

// ========== ShiftFinal ==========

/**
 * 確定シフトを取得
 */
function getShiftFinal(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_FINAL);
  const data = sheet.getDataRange().getValues();
  const targetYM = normalizeYearMonth(yearMonth);
  const targetEmp = employeeId ? String(employeeId) : null;
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const rowYM = normalizeYearMonth(data[i][0]);
    const rowEmp = String(data[i][1]);
    if (rowYM === targetYM && (!targetEmp || rowEmp === targetEmp)) {
      result.push({
        yearMonth: rowYM,
        employeeId: rowEmp,
        date: formatDate(data[i][2]),
        shiftOptionId: String(data[i][3]),
        updatedAt: data[i][4]
      });
    }
  }
  return result;
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
    sheet.getRange(lastRow, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
  }
  return { success: true };
}

/**
 * 希望シフトを確定シフトにコピー（一括）
 */
function copyRequestsToFinal(yearMonth) {
  const requests = getShiftRequests(yearMonth);
  requests.forEach(req => {
    saveShiftFinal(yearMonth, req.employeeId, req.date, req.shiftOptionId);
  });
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
