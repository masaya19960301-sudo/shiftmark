/**
 * Excel出力モジュール
 * 勤務コード/公休形式でのExcel出力
 */

/**
 * Excel出力API（管理者用）
 */
function apiAdminExportExcel(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const dateRange = getDateRangeForYearMonth(yearMonth);
  const users = getAllActiveUsers();
  const options = getActiveShiftOptions();
  const holidays = getHolidays();

  // 確定シフト優先、なければ希望シフト
  let shifts = getShiftFinal(yearMonth);
  let source = '確定';
  if (shifts.length === 0) {
    shifts = getShiftRequests(yearMonth);
    source = '希望';
  }

  // オプションIDからラベルへのマップ
  const optionMap = {};
  options.forEach(o => { optionMap[o.optionId] = o; });

  // 日付リスト生成
  const dates = generateDateList(dateRange.startDate, dateRange.endDate);

  // ========== 時間表記シート用データ ==========
  const timeData = [];
  // ヘッダー行
  const headerRow = ['人員'];
  dates.forEach(d => {
    const dt = new Date(d);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    headerRow.push(dayNames[dt.getDay()] + '\n' + (dt.getMonth() + 1) + '/' + dt.getDate());
  });
  timeData.push(headerRow);

  // データ行
  users.forEach(user => {
    const row = [user.name];
    dates.forEach(date => {
      const shift = shifts.find(s =>
        String(s.employeeId) === String(user.employeeId) && s.date === date
      );
      if (shift && optionMap[shift.shiftOptionId]) {
        const opt = optionMap[shift.shiftOptionId];
        if (opt.startTime && opt.endTime) {
          row.push(opt.startTime + '-\n' + opt.endTime);
        } else {
          row.push(opt.label);
        }
      } else {
        row.push('');
      }
    });
    timeData.push(row);
  });

  // ========== 勤務コード/公休シート用データ ==========
  const codeData = [];
  const codeHeader = ['人員', '社員番号'];
  dates.forEach(d => {
    const dt = new Date(d);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    codeHeader.push(dayNames[dt.getDay()] + '\n' + (dt.getMonth() + 1) + '/' + dt.getDate());
  });
  codeData.push(codeHeader);

  users.forEach(user => {
    const row = [user.name, String(user.employeeId)];
    dates.forEach(date => {
      const shift = shifts.find(s =>
        String(s.employeeId) === String(user.employeeId) && s.date === date
      );
      if (shift && optionMap[shift.shiftOptionId]) {
        const opt = optionMap[shift.shiftOptionId];
        row.push(getWorkCode(opt, user, options));
      } else {
        row.push('');
      }
    });
    codeData.push(row);
  });

  return {
    success: true,
    timeData: timeData,
    codeData: codeData,
    yearMonth: yearMonth,
    source: source,
    dates: dates,
    holidays: holidays
  };
}

/**
 * 勤務コードを取得
 * 有休の場合はユーザーの有休時間帯に対応する勤務コードを使用
 */
function getWorkCode(option, user, allOptions) {
  const label = option.label;

  // 有休の場合：ユーザーの有休時間帯から対応する勤務コードを導出
  if (label === '有休' || label.includes('有休')) {
    if (user.paidLeaveStartTime && user.paidLeaveEndTime && allOptions) {
      const matchOpt = allOptions.find(function(o) {
        return o.startTime === user.paidLeaveStartTime && o.endTime === user.paidLeaveEndTime;
      });
      if (matchOpt && matchOpt.code) return String(matchOpt.code);
    }
    return '有休';
  }

  // codeフィールドがあればそれを使用
  if (option.code) {
    return String(option.code);
  }

  if (label === '公休' || label.includes('公休')) return '公休';
  if (option.startTime && option.endTime) {
    return option.startTime + '-' + option.endTime;
  }
  return label;
}

/**
 * 日付リスト生成
 */
function generateDateList(startDate, endDate) {
  const dates = [];
  const start = new Date(startDate);
  const end = new Date(endDate);

  const current = new Date(start);
  while (current <= end) {
    dates.push(formatDate(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

/**
 * Excel出力用の.xlsxファイル作成（勤務コードシートのみ）
 * base64でクライアントに返してブラウザダウンロードさせる
 */
function apiAdminCreateExcelFile(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const exportData = apiAdminExportExcel(token, yearMonth);
  if (!exportData.success) return exportData;

  // 一時スプレッドシートを作成
  const tmpName = '_tmp_shift_export_' + Date.now();
  const ss = SpreadsheetApp.create(tmpName);

  try {
    // 勤務コードシートにデータを書き込み
    const sheet = ss.getActiveSheet();
    sheet.setName('勤務コード');
    if (exportData.codeData.length > 0) {
      const numRows = exportData.codeData.length;
      const numCols = exportData.codeData[0].length;
      sheet.getRange(1, 1, numRows, numCols).setValues(exportData.codeData);

      // 社員番号列をテキスト形式に設定
      sheet.getRange(2, 2, numRows - 1, 1).setNumberFormat('@');

      formatExportSheet(sheet, exportData.dates, exportData.holidays, numRows, numCols);
    }
    SpreadsheetApp.flush();

    // .xlsx形式でエクスポート
    const ssId = ss.getId();
    const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    const blob = response.getBlob();
    const fileName = 'シフト表_' + yearMonth + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.xlsx';

    // base64エンコードしてクライアントに返す
    const base64 = Utilities.base64Encode(blob.getBytes());

    return {
      success: true,
      message: 'Excelファイルを作成しました（' + exportData.source + 'シフトベース）',
      fileName: fileName,
      base64: base64
    };
  } finally {
    // 一時スプレッドシートを削除
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  }
}

/**
 * 出力シートの書式設定
 */
function formatExportSheet(sheet, dates, holidays, numRows, numCols) {
  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#4472C4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(true);

  // 全体のボーダー
  sheet.getRange(1, 1, numRows, numCols).setBorder(
    true, true, true, true, true, true,
    '#000000', SpreadsheetApp.BorderStyle.SOLID
  );

  // 日曜・祝日の列を色分け（人員+社員番号の後に日付列が続く）
  const dateStartCol = numCols - dates.length + 1;
  dates.forEach((date, idx) => {
    const dt = new Date(date);
    const col = dateStartCol + idx;

    if (dt.getDay() === 0 || holidays.includes(date)) {
      // 日曜・祝日: ピンク
      sheet.getRange(1, col, numRows, 1).setBackground('#FFE0E0');
    } else if (dt.getDay() === 6) {
      // 土曜: 薄い青
      sheet.getRange(1, col, numRows, 1).setBackground('#E0E8FF');
    }
  });

  // 列幅調整
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 80);
  for (let c = dateStartCol; c <= numCols; c++) {
    sheet.setColumnWidth(c, 70);
  }

  // セルの中央揃え
  if (numRows > 1) {
    sheet.getRange(2, 1, numRows - 1, numCols).setHorizontalAlignment('center');
    sheet.getRange(2, 1, numRows - 1, 1).setHorizontalAlignment('left');
  }
}
