// =============================================
// תצורה כללית
// =============================================

const MONTHS_HE = [
  'ינואר','פברואר','מרץ','אפריל','מאי','יוני',
  'יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'
];
const DAYS_HE = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];

// עמודות (1-based)
const COL = { DATE:1, DAY:2, SHIFT:3, OWED:4, CASH:5, CHECK:6, TOTAL:7, DIFF:8, NOTES:9 };
const NUM_COLS = 9;
const HEADERS = ['תאריך','יום','משמרת','סכום מגיע','מזומן',"צ'ק",'סה"כ התקבל','הפרש','הערות'];

const DATA_START_ROW = 3;   // שורה 1 = כותרת, שורה 2 = headers, משורה 3 = נתונים
const DATA_MAX_ROWS   = 33; // מקסימום 33 שורות נתונים (מספיק לכל חודש + זליגות)
const SUMMARY_START   = DATA_START_ROW + DATA_MAX_ROWS + 2; // שורה 38+

// צבעים
const C_TITLE_BG   = '#1a237e';
const C_HEADER_BG  = '#1565c0';
const C_ODD_ROW    = '#f5f7ff';
const C_EVEN_ROW   = '#ffffff';
const C_SUMMARY_BG = '#e8eaf6';
const C_GREEN_BG   = '#e8f5e9';
const C_GREEN_FG   = '#1b5e20';
const C_RED_BG     = '#ffebee';
const C_RED_FG     = '#b71c1c';
const C_WARN_BG    = '#fff8e1';
const C_WARN_FG    = '#e65100';

// =============================================
// תפריט
// =============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 ניהול הכנסות')
    .addItem('➕  הוסף יום חדש',        'showAddEntryDialog')
    .addSeparator()
    .addItem('📅  צור גליון לחודש זה', 'createCurrentMonthSheetMenu')
    .addItem('🔄  רענן סיכומים',        'refreshCurrentSheetSummary')
    .addSeparator()
    .addItem('📊  עדכן לוח מחוונים',   'updateDashboard')
    .toUi();
}

// =============================================
// ניהול גליונות
// =============================================

function sheetName(year, month) {
  return MONTHS_HE[month] + ' ' + year;
}

function getOrCreateMonthSheet(year, month) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheetName(year, month);
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initMonthSheet(sheet, year, month);
  }
  return sheet;
}

function createCurrentMonthSheetMenu() {
  const now   = new Date();
  const sheet = getOrCreateMonthSheet(now.getFullYear(), now.getMonth());
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ גליון "' + sheet.getName() + '" מוכן!');
}

function initMonthSheet(sheet, year, month) {
  sheet.clear();
  sheet.clearFormats();

  const name = MONTHS_HE[month] + ' ' + year;

  // שורה 1 - כותרת ראשית
  const titleRange = sheet.getRange(1, 1, 1, NUM_COLS);
  titleRange.merge()
    .setValue('💰  הכנסות ברמן  –  ' + name)
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C_TITLE_BG).setFontColor('#ffffff');
  sheet.setRowHeight(1, 52);

  // שורה 2 - כותרות עמודות
  sheet.getRange(2, 1, 1, NUM_COLS)
    .setValues([HEADERS])
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C_HEADER_BG).setFontColor('#ffffff')
    .setWrap(true);
  sheet.setRowHeight(2, 38);

  // רוחב עמודות
  sheet.setColumnWidth(COL.DATE,  105);
  sheet.setColumnWidth(COL.DAY,    78);
  sheet.setColumnWidth(COL.SHIFT, 108);
  sheet.setColumnWidth(COL.OWED,  108);
  sheet.setColumnWidth(COL.CASH,   95);
  sheet.setColumnWidth(COL.CHECK,  95);
  sheet.setColumnWidth(COL.TOTAL, 118);
  sheet.setColumnWidth(COL.DIFF,   95);
  sheet.setColumnWidth(COL.NOTES, 165);

  sheet.setFrozenRows(2);

  // כתיבת אזור הסיכום הריק (יתמלא כשיהיו נתונים)
  writeSummarySection(sheet, 0, 0, 0, 0, 0, 0);
}

// =============================================
// הוספת רשומה
// =============================================

function showAddEntryDialog() {
  const now = new Date();
  getOrCreateMonthSheet(now.getFullYear(), now.getMonth()); // ודא שהגליון קיים
  const html = HtmlService.createHtmlOutputFromFile('AddEntry')
    .setWidth(440).setHeight(570)
    .setTitle('הוסף יום חדש');
  SpreadsheetApp.getUi().showModalDialog(html, '➕  הוסף יום חדש');
}

function addEntry(data) {
  const dateStr = data.date;                           // "YYYY-MM-DD"
  const parts   = dateStr.split('-');
  const year    = parseInt(parts[0]);
  const month   = parseInt(parts[1]) - 1;             // 0-based
  const day     = parseInt(parts[2]);
  const dateObj = new Date(year, month, day);

  const sheet = getOrCreateMonthSheet(year, month);

  const owed     = parseFloat(data.owed)  || 0;
  const cash     = parseFloat(data.cash)  || 0;
  const chk      = parseFloat(data.check) || 0;
  const total    = cash + chk;
  const diff     = total - owed;                       // חיובי = קיבל יותר, שלילי = חסר

  const dayName     = DAYS_HE[dateObj.getDay()];
  const formattedDate = Utilities.formatDate(dateObj, 'Asia/Jerusalem', 'dd/MM/yyyy');

  const targetRow = findNextEmptyDataRow(sheet);
  if (!targetRow) {
    return { success: false, message: 'הגליון מלא! צור גליון חדש.' };
  }

  sheet.getRange(targetRow, 1, 1, NUM_COLS).setValues([[
    formattedDate, dayName, data.shift || '',
    owed, cash, chk, total, diff, data.notes || ''
  ]]);

  applyRowFormat(sheet, targetRow, diff);
  sortDataArea(sheet);
  recolorAllRows(sheet);
  refreshSummaryFromSheet(sheet);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return { success: true, message: 'נשמר בהצלחה: ' + formattedDate };
}

function findNextEmptyDataRow(sheet) {
  const lastPossible = DATA_START_ROW + DATA_MAX_ROWS - 1;
  for (let r = DATA_START_ROW; r <= lastPossible; r++) {
    const val = sheet.getRange(r, COL.DATE).getValue();
    if (val === '' || val === null) return r;
  }
  return null;
}

// =============================================
// עיצוב שורות
// =============================================

function applyRowFormat(sheet, row, diff) {
  const bg = (row % 2 === 0) ? C_EVEN_ROW : C_ODD_ROW;
  const range = sheet.getRange(row, 1, 1, NUM_COLS);
  range.setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);

  // עיצוב עמודות מטבע
  [COL.OWED, COL.CASH, COL.CHECK, COL.TOTAL, COL.DIFF].forEach(col => {
    sheet.getRange(row, col).setNumberFormat('₪#,##0.00');
  });

  // צבע עמודת הפרש
  const diffCell = sheet.getRange(row, COL.DIFF);
  if (diff >= 0) {
    diffCell.setBackground(C_GREEN_BG).setFontColor(C_GREEN_FG).setFontWeight('bold');
  } else {
    diffCell.setBackground(C_RED_BG).setFontColor(C_RED_FG).setFontWeight('bold');
  }

  sheet.getRange(row, COL.NOTES).setHorizontalAlignment('right');
}

function recolorAllRows(sheet) {
  const lastData = getLastDataRow(sheet);
  for (let r = DATA_START_ROW; r <= lastData; r++) {
    const diff = parseFloat(sheet.getRange(r, COL.DIFF).getValue()) || 0;
    applyRowFormat(sheet, r, diff);
  }
}

function getLastDataRow(sheet) {
  let last = DATA_START_ROW - 1;
  const maxRow = DATA_START_ROW + DATA_MAX_ROWS - 1;
  for (let r = DATA_START_ROW; r <= maxRow; r++) {
    const val = sheet.getRange(r, COL.DATE).getValue();
    if (val !== '' && val !== null) last = r;
  }
  return last;
}

// =============================================
// מיון לפי תאריך
// =============================================

function sortDataArea(sheet) {
  const lastData = getLastDataRow(sheet);
  if (lastData < DATA_START_ROW) return;
  const numRows = lastData - DATA_START_ROW + 1;
  sheet.getRange(DATA_START_ROW, 1, numRows, NUM_COLS)
    .sort({ column: COL.DATE, ascending: true });
}

// =============================================
// סיכום
// =============================================

function refreshCurrentSheetSummary() {
  refreshSummaryFromSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
  SpreadsheetApp.getUi().alert('🔄 הסיכום עודכן!');
}

function refreshSummaryFromSheet(sheet) {
  const lastData = getLastDataRow(sheet);
  if (lastData < DATA_START_ROW) {
    writeSummarySection(sheet, 0, 0, 0, 0, 0, 0);
    return;
  }

  const numRows = lastData - DATA_START_ROW + 1;
  const data    = sheet.getRange(DATA_START_ROW, 1, numRows, NUM_COLS).getValues();

  let days = 0, totalOwed = 0, totalCash = 0, totalCheck = 0, totalReceived = 0, totalDiff = 0;
  data.forEach(row => {
    if (row[COL.DATE - 1] !== '') {
      days++;
      totalOwed     += parseFloat(row[COL.OWED  - 1]) || 0;
      totalCash     += parseFloat(row[COL.CASH  - 1]) || 0;
      totalCheck    += parseFloat(row[COL.CHECK - 1]) || 0;
      totalReceived += parseFloat(row[COL.TOTAL - 1]) || 0;
      totalDiff     += parseFloat(row[COL.DIFF  - 1]) || 0;
    }
  });

  writeSummarySection(sheet, days, totalOwed, totalCash, totalCheck, totalReceived, totalDiff);
}

function writeSummarySection(sheet, days, owed, cash, chk, received, diff) {
  // נקה את אזור הסיכום הישן
  const clearRows = 15;
  sheet.getRange(SUMMARY_START - 1, 1, clearRows, NUM_COLS).clearContent().clearFormat();

  // קו הפרדה
  sheet.getRange(SUMMARY_START - 1, 1, 1, NUM_COLS).merge()
    .setValue('').setBackground(C_HEADER_BG);
  sheet.setRowHeight(SUMMARY_START - 1, 6);

  // כותרת סיכום
  sheet.getRange(SUMMARY_START, 1, 1, NUM_COLS).merge()
    .setValue('📊  סיכום חודשי')
    .setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');
  sheet.setRowHeight(SUMMARY_START, 38);

  const fmt = '₪#,##0.00';
  const r   = SUMMARY_START + 1;

  // שורה 1: ימי עבודה + סכום מגיע
  writeSummaryRow(sheet, r,
    'ימי עבודה', days + ' ימים', false,
    'סכום מגיע', owed, fmt, false);

  // שורה 2: מזומן + צ'קים
  writeSummaryRow(sheet, r + 1,
    'מזומן שהתקבל', cash, fmt, false,
    "צ'קים", chk, fmt, false);

  // שורה 3: סה"כ + הפרש
  writeSummaryRow(sheet, r + 2,
    'סה"כ התקבל', received, fmt, true,
    'הפרש כולל', diff, fmt, true, diff >= 0);

  sheet.setRowHeight(r, 34);
  sheet.setRowHeight(r + 1, 34);
  sheet.setRowHeight(r + 2, 34);

  // הערת תלוש שכר
  const noteRow = r + 4;
  sheet.getRange(noteRow, 1, 1, NUM_COLS).merge()
    .setValue('⚠️  תלוש שכר מתקבל ב-10 לחודש — ייתכן שסכומים מתחילת החודש הבא ייכנסו לתלוש הנוכחי')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center').setWrap(true)
    .setBackground(C_WARN_BG).setFontColor(C_WARN_FG);
  sheet.setRowHeight(noteRow, 32);
}

function writeSummaryRow(sheet, row, label1, val1, fmt1, bold1, label2, val2, fmt2, bold2, isPositive) {
  // תא label 1
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue(label1).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');

  // תא value 1
  const c1 = sheet.getRange(row, 3);
  c1.setValue(val1).setHorizontalAlignment('center')
    .setBackground('#ffffff').setFontColor('#1565c0');
  if (fmt1)  c1.setNumberFormat(fmt1);
  if (bold1) c1.setFontWeight('bold').setFontSize(12);

  // רווח
  sheet.getRange(row, 4).setValue('').setBackground('#ffffff');

  // תא label 2
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue(label2).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');

  // תא value 2
  const c2 = sheet.getRange(row, 7);
  c2.setValue(val2).setHorizontalAlignment('center');
  if (fmt2)  c2.setNumberFormat(fmt2);
  if (bold2) c2.setFontWeight('bold').setFontSize(12);

  // צבע הפרש
  if (bold2 && label2 === 'הפרש כולל') {
    if (isPositive) {
      c2.setBackground(C_GREEN_BG).setFontColor(C_GREEN_FG);
    } else {
      c2.setBackground(C_RED_BG).setFontColor(C_RED_FG);
    }
  } else {
    c2.setBackground('#ffffff').setFontColor('#1565c0');
  }

  // ניקוי שאר העמודות
  sheet.getRange(row, 8, 1, NUM_COLS - 7).clearContent().clearFormat()
    .setBackground('#ffffff');
}

// =============================================
// לוח מחוונים שנתי
// =============================================

function updateDashboard() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const currentYear = new Date().getFullYear();

  let dash = ss.getSheetByName('📊 לוח מחוונים');
  if (!dash) {
    dash = ss.insertSheet('📊 לוח מחוונים', 0);
  }

  dash.clear();
  dash.clearFormats();

  // כותרת
  dash.getRange(1, 1, 1, 8).merge()
    .setValue('📊  סיכום שנתי  –  ' + currentYear)
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_TITLE_BG).setFontColor('#ffffff');
  dash.setRowHeight(1, 55);

  // headers
  const hdr = ['חודש','ימי עבודה','סכום מגיע','מזומן',"צ'קים",'סה"כ התקבל','הפרש'];
  dash.getRange(2, 1, 1, 7).setValues([hdr])
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center')
    .setBackground(C_HEADER_BG).setFontColor('#ffffff');
  dash.setRowHeight(2, 38);
  dash.setFrozenRows(2);

  let yDays = 0, yOwed = 0, yCash = 0, yChk = 0, yRec = 0, yDiff = 0;

  MONTHS_HE.forEach((mName, mIdx) => {
    const sName = sheetName(currentYear, mIdx);
    const mSheet = ss.getSheetByName(sName);
    const r = 3 + mIdx;

    let days = 0, owed = 0, cash = 0, chk = 0, rec = 0, dif = 0;

    if (mSheet) {
      const lastData = getLastDataRow(mSheet);
      if (lastData >= DATA_START_ROW) {
        const numRows = lastData - DATA_START_ROW + 1;
        const rows = mSheet.getRange(DATA_START_ROW, 1, numRows, NUM_COLS).getValues();
        rows.forEach(row => {
          if (row[COL.DATE - 1] !== '') {
            days++;
            owed += parseFloat(row[COL.OWED  - 1]) || 0;
            cash += parseFloat(row[COL.CASH  - 1]) || 0;
            chk  += parseFloat(row[COL.CHECK - 1]) || 0;
            rec  += parseFloat(row[COL.TOTAL - 1]) || 0;
            dif  += parseFloat(row[COL.DIFF  - 1]) || 0;
          }
        });
      }
    }

    const bg = (mIdx % 2 === 0) ? C_ODD_ROW : C_EVEN_ROW;
    const fmt = '₪#,##0.00';

    // קישור לגליון החודשי
    dash.getRange(r, 1)
      .setValue(mName).setFontWeight('bold')
      .setBackground(bg).setHorizontalAlignment('center').setFontColor('#1a237e');

    if (mSheet) {
      dash.getRange(r, 1).setFormula(
        '=HYPERLINK("#gid=' + mSheet.getSheetId() + '","' + mName + '")'
      );
    }

    dash.getRange(r, 2).setValue(days).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 3).setValue(owed).setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 4).setValue(cash).setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 5).setValue(chk) .setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 6).setValue(rec) .setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center')
      .setFontWeight('bold');

    const diffCell = dash.getRange(r, 7);
    diffCell.setValue(dif).setNumberFormat(fmt).setHorizontalAlignment('center').setFontWeight('bold');
    diffCell.setBackground(dif >= 0 ? C_GREEN_BG : C_RED_BG)
      .setFontColor(dif >= 0 ? C_GREEN_FG : C_RED_FG);

    dash.setRowHeight(r, 30);

    yDays += days; yOwed += owed; yCash += cash;
    yChk  += chk;  yRec  += rec;  yDiff += dif;
  });

  // שורת סיכום שנתי
  const totRow = 16;
  dash.getRange(totRow, 1, 1, 7).setBackground(C_TITLE_BG).setFontColor('#ffffff').setFontWeight('bold');
  [' סה"כ שנתי', yDays, yOwed, yCash, yChk, yRec, yDiff].forEach((v, i) => {
    const cell = dash.getRange(totRow, i + 1);
    cell.setValue(v).setHorizontalAlignment('center').setFontSize(12);
    if (i >= 2) cell.setNumberFormat('₪#,##0.00');
  });
  dash.setRowHeight(totRow, 42);

  // רוחב עמודות
  [100, 85, 115, 105, 105, 125, 105].forEach((w, i) => dash.setColumnWidth(i + 1, w));

  ss.setActiveSheet(dash);
  SpreadsheetApp.getUi().alert('✅ לוח המחוונים עודכן!');
}
