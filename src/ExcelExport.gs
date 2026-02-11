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
        row.push(getWorkCode(opt, user));
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
 * Ver1.0では簡易的な変換を行う（勤務コードマスタは後工程）
 */
function getWorkCode(option, user) {
  // codeフィールドがあればそれを使用
  if (option.code) {
    return String(option.code);
  }

  const label = option.label;
  if (label === '公休' || label.includes('公休')) return '公休';
  if (label === '有休' || label.includes('有休')) return '有休';
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
 * Excel出力用のスプレッドシート作成
 */
function apiAdminCreateExcelFile(token, yearMonth) {
  const auth = requireAdmin(token);
  if (!auth.valid) return { success: false, message: auth.message };

  const exportData = apiAdminExportExcel(token, yearMonth);
  if (!exportData.success) return exportData;

  // 新しいスプレッドシートを作成
  const fileName = 'シフト表_' + yearMonth + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const ss = SpreadsheetApp.create(fileName);

  // 時間表記シート
  const timeSheet = ss.getActiveSheet();
  timeSheet.setName('時間表記');
  if (exportData.timeData.length > 0) {
    const numRows = exportData.timeData.length;
    const numCols = exportData.timeData[0].length;
    timeSheet.getRange(1, 1, numRows, numCols).setValues(exportData.timeData);

    // 書式設定
    formatExportSheet(timeSheet, exportData.dates, exportData.holidays, numRows, numCols);
  }

  // 勤務コード/公休シート
  const codeSheet = ss.insertSheet('勤務コード');
  if (exportData.codeData.length > 0) {
    const numRows = exportData.codeData.length;
    const numCols = exportData.codeData[0].length;
    codeSheet.getRange(1, 1, numRows, numCols).setValues(exportData.codeData);

    formatExportSheet(codeSheet, exportData.dates, exportData.holidays, numRows, numCols);
  }

  const fileUrl = ss.getUrl();
  return {
    success: true,
    message: 'Excelファイルを作成しました（' + exportData.source + 'シフトベース）',
    fileUrl: fileUrl,
    fileName: fileName
  };
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

  // 日曜・祝日の列を色分け
  dates.forEach((date, idx) => {
    const dt = new Date(date);
    const col = idx + 2; // 1列目は名前

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
  for (let c = 2; c <= numCols; c++) {
    sheet.setColumnWidth(c, 70);
  }

  // セルの中央揃え
  if (numRows > 1) {
    sheet.getRange(2, 1, numRows - 1, numCols).setHorizontalAlignment('center');
    sheet.getRange(2, 1, numRows - 1, 1).setHorizontalAlignment('left');
  }
}
