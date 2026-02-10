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

  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(employeeId)) {
      const user = {};
      headers.forEach((h, idx) => { user[h] = data[i][idx]; });
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
  const existing = findUserByEmployeeId(userData.employeeId);
  if (existing) {
    return { success: false, message: 'この社員番号は既に登録されています' };
  }

  const passwordHash = hashPassword(userData.password || 'password');
  sheet.appendRow([
    userData.employeeId,
    userData.name,
    userData.role || 'user',
    userData.category || '社員',
    passwordHash,
    userData.paidLeaveHoursPerDay || 8,
    userData.includePaidLeaveInDaysOff || false,
    true
  ]);
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
  const newHash = hashPassword(newPassword);
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
      def[h] = (h === 'startDate' || h === 'endDate') ? formatDate(data[i][idx]) : data[i][idx];
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
  const defs = getMonthDefinitions();
  return defs.find(d => d.yearMonth === yearMonth) || null;
}

/**
 * 月度定義を保存/更新
 */
function saveMonthDefinition(yearMonth, startDate, endDate, effectiveFrom) {
  const sheet = getOrCreateSheet(SHEET_NAMES.MONTH_DEFINITION);
  const existing = getMonthDefinition(yearMonth);

  if (existing) {
    sheet.getRange(existing._row, 2).setValue(startDate);
    sheet.getRange(existing._row, 3).setValue(endDate);
    sheet.getRange(existing._row, 4).setValue(effectiveFrom);
  } else {
    sheet.appendRow([yearMonth, startDate, endDate, effectiveFrom]);
  }
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
  const parts = yearMonth.split('-');
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth) {
      return { yearMonth: data[i][0], requiredDaysOff: data[i][1], _row: i + 1 };
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
    result.push({ yearMonth: data[i][0], requiredDaysOff: data[i][1] });
  }
  return result;
}

/**
 * 必要公休日数を保存
 */
function saveRequiredHolidays(yearMonth, requiredDaysOff) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REQUIRED_HOLIDAYS);
  const existing = getRequiredHolidays(yearMonth);
  if (existing) {
    sheet.getRange(existing._row, 2).setValue(requiredDaysOff);
  } else {
    sheet.appendRow([yearMonth, requiredDaysOff]);
  }
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
  const existing = all.find(o => o.optionId === optionData.optionId);

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
  return { success: true };
}

// ========== ShiftRequests ==========

/**
 * 希望シフトを取得（年月×社員）
 */
function getShiftRequests(yearMonth, employeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SHIFT_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth && (!employeeId || String(data[i][1]) === String(employeeId))) {
      result.push({
        yearMonth: data[i][0],
        employeeId: String(data[i][1]),
        date: formatDate(data[i][2]),
        shiftOptionId: data[i][3],
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

  // 既存レコードを探す
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth &&
        String(data[i][1]) === String(employeeId) &&
        formatDate(data[i][2]) === formattedDate) {
      if (shiftOptionId === '' || shiftOptionId === null) {
        // 空の場合は行を削除
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(shiftOptionId);
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      }
      return { success: true };
    }
  }

  // 新規追加
  if (shiftOptionId && shiftOptionId !== '') {
    sheet.appendRow([yearMonth, employeeId, formattedDate, shiftOptionId, new Date().toISOString()]);
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
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth && (!employeeId || String(data[i][1]) === String(employeeId))) {
      result.push({
        yearMonth: data[i][0],
        employeeId: String(data[i][1]),
        date: formatDate(data[i][2]),
        shiftOptionId: data[i][3],
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

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth &&
        String(data[i][1]) === String(employeeId) &&
        formatDate(data[i][2]) === formattedDate) {
      if (shiftOptionId === '' || shiftOptionId === null) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(shiftOptionId);
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      }
      return { success: true };
    }
  }

  if (shiftOptionId && shiftOptionId !== '') {
    sheet.appendRow([yearMonth, employeeId, formattedDate, shiftOptionId, new Date().toISOString()]);
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yearMonth) {
      return {
        yearMonth: data[i][0],
        locked: data[i][1] === true || data[i][1] === 'TRUE',
        lockedAt: data[i][2],
        lockedBy: data[i][3],
        _row: i + 1
      };
    }
  }
  return { yearMonth: yearMonth, locked: false };
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
      yearMonth: data[i][0],
      locked: data[i][1] === true || data[i][1] === 'TRUE',
      lockedAt: data[i][2],
      lockedBy: data[i][3]
    });
  }
  return result;
}

/**
 * 締切設定/解除
 */
function setLock(yearMonth, locked, adminEmployeeId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.LOCKS);
  const status = getLockStatus(yearMonth);

  if (status._row) {
    sheet.getRange(status._row, 2).setValue(locked);
    sheet.getRange(status._row, 3).setValue(new Date().toISOString());
    sheet.getRange(status._row, 4).setValue(adminEmployeeId);
  } else {
    sheet.appendRow([yearMonth, locked, new Date().toISOString(), adminEmployeeId]);
  }
  return { success: true, locked: locked };
}
