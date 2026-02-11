/**
 * 管理者用API
 */

// ========== ユーザー管理 ==========

/**
 * ユーザー一覧取得（管理者）
 */
function apiAdminGetUsers(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, users: getAllUsers() };
}

/**
 * ユーザー登録（管理者）
 */
function apiAdminAddUser(token, userData) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return addUser(userData);
}

/**
 * ユーザー更新（管理者）
 */
function apiAdminUpdateUser(token, employeeId, userData) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return updateUser(employeeId, userData);
}

/**
 * パスワードリセット（管理者）
 */
function apiAdminResetPassword(token, employeeId, newPassword) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const result = resetUserPassword(employeeId, newPassword);
  if (result) {
    return { success: true, message: 'パスワードをリセットしました' };
  }
  return { success: false, message: 'パスワードリセットに失敗しました' };
}

/**
 * ユーザー削除（管理者）
 */
function apiAdminDeleteUser(token, employeeId) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  if (String(auth.session.employeeId) === String(employeeId)) {
    return { success: false, message: '自分自身は削除できません' };
  }

  return deleteUser(employeeId);
}

// ========== 月度設定 ==========

/**
 * 月度定義一覧取得
 */
function apiAdminGetMonthDefinitions(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, definitions: getMonthDefinitions() };
}

/**
 * 月度定義保存
 */
function apiAdminSaveMonthDefinition(token, yearMonth, startDate, endDate, effectiveFrom) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  saveMonthDefinition(yearMonth, startDate, endDate, effectiveFrom);
  return { success: true, message: '月度設定を保存しました' };
}

// ========== 必要公休日 ==========

/**
 * 必要公休日一覧取得
 */
function apiAdminGetRequiredHolidays(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, holidays: getAllRequiredHolidays() };
}

/**
 * 必要公休日保存
 */
function apiAdminSaveRequiredHolidays(token, yearMonth, requiredDaysOff) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  saveRequiredHolidays(yearMonth, requiredDaysOff);
  return { success: true, message: '必要公休日数を保存しました' };
}

// ========== 祝日管理 ==========

/**
 * 祝日一覧取得
 */
function apiAdminGetHolidays(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, holidays: getHolidays() };
}

/**
 * 祝日登録（CSVデータから）
 */
function apiAdminRegisterHolidays(token, csvData) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  // CSVパース
  const dates = parseDateListFromCSV(csvData);
  if (dates.length === 0) {
    return { success: false, message: '有効な日付データが見つかりません' };
  }

  const result = addHolidays(dates);
  return {
    success: true,
    message: result.addedCount + '件の祝日を登録しました',
    addedCount: result.addedCount
  };
}

/**
 * CSVから日付リストをパース
 */
function parseDateListFromCSV(csvData) {
  const lines = csvData.split(/[\r\n]+/);
  const dates = [];
  lines.forEach(line => {
    const cells = line.split(/[,\t]/);
    cells.forEach(cell => {
      const trimmed = cell.trim().replace(/"/g, '');
      if (trimmed) {
        // 様々な日付形式に対応
        const date = tryParseDate(trimmed);
        if (date) {
          dates.push(formatDate(date));
        }
      }
    });
  });
  return dates;
}

/**
 * 日付文字列のパース試行
 */
function tryParseDate(str) {
  // yyyy/mm/dd or yyyy-mm-dd
  let match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }
  // mm/dd/yyyy
  match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    return new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
  }
  return null;
}

// ========== 勤務時間候補 ==========

/**
 * 勤務時間候補一覧取得（管理者用、全件）
 */
function apiAdminGetShiftOptions(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, options: getAllShiftOptions() };
}

/**
 * 勤務時間候補保存
 */
function apiAdminSaveShiftOption(token, optionData) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  saveShiftOption(optionData);
  return { success: true, message: '勤務時間候補を保存しました' };
}

// ========== 締切管理 ==========

/**
 * 締切状態一覧取得
 */
function apiAdminGetLocks(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, locks: getAllLocks() };
}

/**
 * 締切設定/解除
 */
function apiAdminSetLock(token, yearMonth, locked) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const result = setLock(yearMonth, locked, auth.session.employeeId);
  const action = locked ? '締切' : '締切解除';
  return { success: true, message: yearMonth + 'の希望を' + action + 'しました' };
}

// ========== 確定シフト編集 ==========

/**
 * 確定シフト取得
 */
function apiAdminGetShiftFinal(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const shifts = getShiftFinal(yearMonth);
  const users = getAllActiveUsers();
  const options = getActiveShiftOptions();
  const dateRange = getDateRangeForYearMonth(yearMonth);
  const holidays = getHolidays();
  const lockStatus = getLockStatus(yearMonth);
  const reqHolidays = getRequiredHolidays(yearMonth);

  return {
    success: true,
    shifts: shifts,
    users: users,
    options: options,
    dateRange: dateRange,
    holidays: holidays,
    locked: lockStatus.locked,
    requiredDaysOff: reqHolidays ? reqHolidays.requiredDaysOff : null
  };
}

/**
 * 確定シフト保存（1セル）
 */
function apiAdminSaveShiftFinal(token, yearMonth, employeeId, date, shiftOptionId) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return saveShiftFinal(yearMonth, employeeId, date, shiftOptionId);
}

/**
 * 希望シフトを確定シフトに一括コピー
 */
function apiAdminCopyRequestsToFinal(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return copyRequestsToFinal(yearMonth);
}

// ========== デフォルトシフト管理 ==========

/**
 * 全ユーザーのデフォルトシフト取得
 */
function apiAdminGetAllUserDefaults(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const users = getAllActiveUsers();
  const allDefaults = getAllUserDefaults();
  const options = getActiveShiftOptions();

  return {
    success: true,
    users: users,
    allDefaults: allDefaults,
    options: options
  };
}

/**
 * デフォルトから希望シフトを全ユーザー一括生成（管理者）
 */
function apiAdminGenerateAllRequestsFromDefaults(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  return generateAllRequestsFromDefaults(yearMonth);
}

/**
 * 既存の希望シフトがデフォルトと異なるかチェック（管理者）
 */
function apiAdminCheckRequestsVsDefaults(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const result = checkRequestsVsDefaults(yearMonth);
  return { success: true, hasModified: result.hasModified, modifiedCount: result.modifiedCount, requestCount: result.requestCount };
}

/**
 * ユーザーのデフォルトシフト保存
 */
function apiAdminSaveUserDefaults(token, employeeId, defaults) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  saveUserDefaults(employeeId, defaults);
  return { success: true, message: 'デフォルトシフトを保存しました' };
}

// ========== 周知事項 ==========

function apiAdminGetAnnouncements(token) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  return { success: true, announcements: getAllAnnouncements() };
}

function apiAdminSaveAnnouncement(token, yearMonth, message) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };
  saveAnnouncement(yearMonth, message);
  return { success: true, message: '周知事項を保存しました' };
}
