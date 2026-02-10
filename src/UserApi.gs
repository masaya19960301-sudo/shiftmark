/**
 * 一般ユーザー用API
 */

/**
 * 希望シフト入力画面用データ取得
 */
function apiGetShiftInputData(token, yearMonth) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  const dateRange = getDateRangeForYearMonth(yearMonth);
  const options = getActiveShiftOptions();
  const lockStatus = getLockStatus(yearMonth);
  const holidays = getHolidays();
  const users = getAllActiveUsers();

  // 全ユーザーの希望シフト取得
  const allRequests = getShiftRequests(yearMonth);

  // 自分の希望を分離
  const myRequests = allRequests.filter(r => String(r.employeeId) === String(session.employeeId));
  const otherRequests = allRequests.filter(r => String(r.employeeId) !== String(session.employeeId));

  // 必要公休日数
  const reqHolidays = getRequiredHolidays(yearMonth);

  // 自分のユーザー情報
  const myUser = findUserByEmployeeId(session.employeeId);
  const myInfo = myUser ? {
    category: myUser.category,
    paidLeaveHoursPerDay: myUser.paidLeaveHoursPerDay,
    includePaidLeaveInDaysOff: myUser.includePaidLeaveInDaysOff
  } : {};

  return {
    success: true,
    yearMonth: yearMonth,
    dateRange: dateRange,
    options: options,
    locked: lockStatus.locked,
    isAdmin: session.role === 'admin',
    holidays: holidays,
    myRequests: myRequests,
    otherRequests: otherRequests,
    users: users,
    requiredDaysOff: reqHolidays ? reqHolidays.requiredDaysOff : null,
    myInfo: myInfo
  };
}

/**
 * 希望シフト保存（1セル）
 */
function apiSaveShiftRequest(token, yearMonth, date, shiftOptionId) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  // 締切チェック（管理者は締切後も編集可）
  if (session.role !== 'admin') {
    const lockStatus = getLockStatus(yearMonth);
    if (lockStatus.locked) {
      return { success: false, message: 'この月度は締切済みのため編集できません' };
    }
  }

  return saveShiftRequest(yearMonth, session.employeeId, date, shiftOptionId);
}

/**
 * シフト閲覧画面用データ取得（確定シフト）
 */
function apiGetShiftView(token, yearMonth) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  const dateRange = getDateRangeForYearMonth(yearMonth);
  const options = getActiveShiftOptions();
  const holidays = getHolidays();
  const users = getAllActiveUsers();

  // 確定シフトがあればそれを表示、なければ希望シフト
  let shifts = getShiftFinal(yearMonth);
  let isFinal = true;
  if (shifts.length === 0) {
    shifts = getShiftRequests(yearMonth);
    isFinal = false;
  }

  return {
    success: true,
    yearMonth: yearMonth,
    dateRange: dateRange,
    options: options,
    holidays: holidays,
    users: users,
    shifts: shifts,
    isFinal: isFinal
  };
}

/**
 * 利用可能な年月一覧を取得
 */
function apiGetAvailableYearMonths(token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  // 月度定義があればそれを使う
  const defs = getMonthDefinitions();
  const locks = getAllLocks();

  // 現在の年月を基準に前後の月を生成
  const now = new Date();
  const months = [];

  for (let i = -3; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const ym = formatYearMonth(d.getFullYear(), d.getMonth() + 1);
    const def = defs.find(dd => dd.yearMonth === ym);
    const lock = locks.find(l => l.yearMonth === ym);
    months.push({
      yearMonth: ym,
      label: d.getFullYear() + '年' + (d.getMonth() + 1) + '月度',
      hasDef: !!def,
      locked: lock ? lock.locked : false
    });
  }

  return {
    success: true,
    months: months,
    isAdmin: session.role === 'admin'
  };
}

/**
 * 公休日数チェック（希望入力時のリアルタイム検証）
 */
function apiCheckDaysOff(token, yearMonth) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  const myRequests = getShiftRequests(yearMonth, session.employeeId);
  const options = getActiveShiftOptions();
  const reqHolidays = getRequiredHolidays(yearMonth);
  const user = findUserByEmployeeId(session.employeeId);

  // 公休としてカウントするオプションを特定
  const daysOffOptions = options.filter(o =>
    o.label === '公休' || o.label.includes('公休')
  ).map(o => o.optionId);

  const paidLeaveOptions = options.filter(o =>
    o.label === '有休' || o.label.includes('有休')
  ).map(o => o.optionId);

  let daysOffCount = 0;
  let paidLeaveCount = 0;

  myRequests.forEach(req => {
    if (daysOffOptions.includes(req.shiftOptionId)) {
      daysOffCount++;
    }
    if (paidLeaveOptions.includes(req.shiftOptionId)) {
      paidLeaveCount++;
    }
  });

  // 有休を公休に含めるかの設定
  const includePL = user && (user.includePaidLeaveInDaysOff === true ||
                             user.includePaidLeaveInDaysOff === 'TRUE');
  const totalDaysOff = includePL ? daysOffCount + paidLeaveCount : daysOffCount;

  const required = reqHolidays ? reqHolidays.requiredDaysOff : null;
  let warning = null;

  if (required !== null) {
    const diff = totalDaysOff - required;
    if (diff < 0) {
      warning = '公休が' + Math.abs(diff) + '日不足しています（必要: ' + required + '日, 現在: ' + totalDaysOff + '日）';
    } else if (diff > 0) {
      warning = '公休が' + diff + '日超過しています（必要: ' + required + '日, 現在: ' + totalDaysOff + '日）';
    }
  }

  return {
    success: true,
    daysOffCount: daysOffCount,
    paidLeaveCount: paidLeaveCount,
    totalDaysOff: totalDaysOff,
    required: required,
    warning: warning,
    includePaidLeave: includePL
  };
}

/**
 * デフォルトシフトから自分の希望シフトを自動生成
 */
function apiGenerateMyShiftFromDefaults(token, yearMonth) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'セッションが無効です' };

  // 締切チェック（管理者は締切後も編集可）
  if (session.role !== 'admin') {
    const lockStatus = getLockStatus(yearMonth);
    if (lockStatus.locked) {
      return { success: false, message: 'この月度は締切済みのため編集できません' };
    }
  }

  return generateShiftFromDefaults(yearMonth, session.employeeId);
}
