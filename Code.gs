// =============================================
// תצורה כללית
// =============================================

const MONTHS_HE = [
  'ינואר','פברואר','מרץ','אפריל','מאי','יוני',
  'יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'
];
const DAYS_HE = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];

// עמודות (1-based)
const COL = { DATE:1, DAY:2, SHIFT:3, HOURS:4, OWED:5, CASH:6, CHECK:7, TOTAL:8, DIFF:9, NOTES:10 };
const NUM_COLS = 10;
const HEADERS = ['תאריך','יום','משמרת','שעות','סכום מגיע','מזומן',"צ'ק",'סה"כ התקבל','הפרש','הערות'];

const DATA_START_ROW = 3;
const DATA_MAX_ROWS  = 33;
const DATA_END_ROW   = DATA_START_ROW + DATA_MAX_ROWS - 1; // 35
const SUMMARY_START  = DATA_END_ROW + 3;                   // 38

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

// המרת מספר עמודה לאות (1→A, 2→B …)
function colLetter(n) { return String.fromCharCode(64 + n); }

// =============================================
// WEB APP
// =============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('💰 הכנסות ברמן')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCurrentMonthSummary() {
  const now   = new Date();
  const sheet = getOrCreateMonthSheet(now.getFullYear(), now.getMonth());
  const lastData = getLastDataRow(sheet);
  const monthLabel = MONTHS_HE[now.getMonth()] + ' ' + now.getFullYear();

  if (lastData < DATA_START_ROW) {
    return { days:0, owed:0, cash:0, check:0, received:0, diff:0, month: monthLabel };
  }

  const numRows = lastData - DATA_START_ROW + 1;
  const data    = sheet.getRange(DATA_START_ROW, 1, numRows, NUM_COLS).getValues();
  let days=0, owed=0, cash=0, chk=0, received=0, diff=0;
  data.forEach(row => {
    if (row[COL.DATE-1] !== '') {
      days++;
      owed     += parseFloat(row[COL.OWED -1]) || 0;
      cash     += parseFloat(row[COL.CASH -1]) || 0;
      chk      += parseFloat(row[COL.CHECK-1]) || 0;
      received += parseFloat(row[COL.TOTAL-1]) || 0;
      diff     += parseFloat(row[COL.DIFF -1]) || 0;
    }
  });
  return { days, owed, cash, check:chk, received, diff, month: monthLabel };
}

function getRecentEntries() {
  const now   = new Date();
  const sheet = getOrCreateMonthSheet(now.getFullYear(), now.getMonth());
  const lastData = getLastDataRow(sheet);
  if (lastData < DATA_START_ROW) return [];

  const numRows = lastData - DATA_START_ROW + 1;
  const data    = sheet.getRange(DATA_START_ROW, 1, numRows, NUM_COLS).getValues();
  return data
    .filter(r => r[COL.DATE-1] !== '')
    .slice(-5)
    .reverse()
    .map(r => ({
      date:     r[COL.DATE -1],
      day:      r[COL.DAY  -1],
      shift:    r[COL.SHIFT-1],
      hours:    r[COL.HOURS-1],
      owed:     r[COL.OWED -1],
      cash:     r[COL.CASH -1],
      check:    r[COL.CHECK-1],
      received: r[COL.TOTAL-1],
      diff:     r[COL.DIFF -1]
    }));
}

// =============================================
// ריענון כל הגליונות
// =============================================

function fixAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const CL = colLetter;
  let totalFixed = 0;
  const report = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    // עבד רק גליונות שמתאימים לפורמט "חודש שנה"
    const isMonthSheet = MONTHS_HE.some(m => name.startsWith(m));
    if (!isMonthSheet) return;

    const lastData = getLastDataRow(sheet);
    if (lastData < DATA_START_ROW) {
      report.push(name + ': אין נתונים');
      return;
    }

    let fixed = 0;
    for (let r = DATA_START_ROW; r <= lastData; r++) {
      if (sheet.getRange(r, COL.DATE).getValue() === '') continue;
      sheet.getRange(r, COL.TOTAL).setFormula(`=${CL(COL.CASH)}${r}+${CL(COL.CHECK)}${r}`);
      sheet.getRange(r, COL.DIFF) .setFormula(`=${CL(COL.TOTAL)}${r}-${CL(COL.OWED)}${r}`);
      applyRowFormat(sheet, r);
      fixed++;
    }

    writeSummarySection(sheet);
    addConditionalFormatting(sheet);

    report.push(name + ': ' + fixed + ' שורות');
    totalFixed += fixed;
  });

  Logger.log('=== תיקון גליונות ===');
  report.forEach(line => Logger.log(line));
  Logger.log('סה"כ שורות שתוקנו: ' + totalFixed);
  return { fixed: totalFixed, report: report };
}

// =============================================
// תפריט
// =============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 ניהול הכנסות')
    .addItem('➕  הוסף יום חדש',          'showAddEntryDialog')
    .addSeparator()
    .addItem('📅  צור גליון לחודש זה',         'createCurrentMonthSheetMenu')
    .addItem('🔄  רענן נוסחאות ועיצוב',        'refreshCurrentSheetSummary')
    .addItem('🛠️  תקן גליון ישן (המר לנוסחאות)', 'fixOldSheet')
    .addSeparator()
    .addItem('📊  עדכן לוח מחוונים',     'updateDashboard')
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
  sheet.clearConditionalFormatRules();

  const name = MONTHS_HE[month] + ' ' + year;

  // שורה 1 – כותרת
  sheet.getRange(1, 1, 1, NUM_COLS).merge()
    .setValue('💰  הכנסות ברמן  –  ' + name)
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C_TITLE_BG).setFontColor('#ffffff');
  sheet.setRowHeight(1, 52);

  // שורה 2 – headers
  sheet.getRange(2, 1, 1, NUM_COLS)
    .setValues([HEADERS])
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C_HEADER_BG).setFontColor('#ffffff')
    .setWrap(true);
  sheet.setRowHeight(2, 38);

  // רוחב עמודות
  sheet.setColumnWidth(COL.DATE,   105);
  sheet.setColumnWidth(COL.DAY,     78);
  sheet.setColumnWidth(COL.SHIFT,  108);
  sheet.setColumnWidth(COL.HOURS,   70);
  sheet.setColumnWidth(COL.OWED,   108);
  sheet.setColumnWidth(COL.CASH,    95);
  sheet.setColumnWidth(COL.CHECK,   95);
  sheet.setColumnWidth(COL.TOTAL,  118);
  sheet.setColumnWidth(COL.DIFF,    95);
  sheet.setColumnWidth(COL.NOTES,  165);

  sheet.setFrozenRows(2);

  // נוסחאות סיכום + עיצוב תנאי
  writeSummarySection(sheet);
  addConditionalFormatting(sheet);
}

// =============================================
// הוספת רשומה
// =============================================

function showAddEntryDialog() {
  const now = new Date();
  getOrCreateMonthSheet(now.getFullYear(), now.getMonth());
  const html = HtmlService.createHtmlOutputFromFile('AddEntry')
    .setWidth(440).setHeight(580)
    .setTitle('הוסף יום חדש');
  SpreadsheetApp.getUi().showModalDialog(html, '➕  הוסף יום חדש');
}

function addEntry(data) {
  const dateStr = data.date;
  const parts   = dateStr.split('-');
  const year    = parseInt(parts[0]);
  const month   = parseInt(parts[1]) - 1;
  const day     = parseInt(parts[2]);
  const dateObj = new Date(year, month, day);

  const sheet = getOrCreateMonthSheet(year, month);

  const hours = parseFloat(data.hours) || '';
  const owed  = parseFloat(data.owed)  || 0;
  const cash  = parseFloat(data.cash)  || 0;
  const chk   = parseFloat(data.check) || 0;

  const dayName       = DAYS_HE[dateObj.getDay()];
  const formattedDate = Utilities.formatDate(dateObj, 'Asia/Jerusalem', 'dd/MM/yyyy');

  const targetRow = findNextEmptyDataRow(sheet);
  if (!targetRow) {
    return { success: false, message: 'הגליון מלא! צור גליון חדש.' };
  }

  // כתיבת ערכים שניתן לערוך ידנית
  sheet.getRange(targetRow, COL.DATE ).setValue(formattedDate);
  sheet.getRange(targetRow, COL.DAY  ).setValue(dayName);
  sheet.getRange(targetRow, COL.SHIFT).setValue(data.shift || '');
  sheet.getRange(targetRow, COL.HOURS).setValue(hours);
  sheet.getRange(targetRow, COL.OWED ).setValue(owed);
  sheet.getRange(targetRow, COL.CASH ).setValue(cash);
  sheet.getRange(targetRow, COL.CHECK).setValue(chk);
  sheet.getRange(targetRow, COL.NOTES).setValue(data.notes || '');

  // כתיבת נוסחאות מחושבות (מתעדכנות אוטומטית עם עריכה ידנית)
  const CL = colLetter;
  sheet.getRange(targetRow, COL.TOTAL)
    .setFormula(`=${CL(COL.CASH)}${targetRow}+${CL(COL.CHECK)}${targetRow}`);
  sheet.getRange(targetRow, COL.DIFF)
    .setFormula(`=${CL(COL.TOTAL)}${targetRow}-${CL(COL.OWED)}${targetRow}`);

  applyRowFormat(sheet, targetRow);
  sortDataArea(sheet);
  recolorAllRows(sheet);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return { success: true, message: 'נשמר בהצלחה: ' + formattedDate };
}

function findNextEmptyDataRow(sheet) {
  for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
    const val = sheet.getRange(r, COL.DATE).getValue();
    if (val === '' || val === null) return r;
  }
  return null;
}

// =============================================
// עיצוב שורות
// =============================================

function applyRowFormat(sheet, row) {
  const bg = (row % 2 === 0) ? C_EVEN_ROW : C_ODD_ROW;

  // רקע לכל השורה
  sheet.getRange(row, 1, 1, NUM_COLS)
    .setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);

  // נקה רקע של תא הפרש — יטופל ע"י Conditional Formatting
  sheet.getRange(row, COL.DIFF).setBackground(null).setFontColor(null).setFontWeight('bold');

  // פורמט מטבע
  [COL.OWED, COL.CASH, COL.CHECK, COL.TOTAL, COL.DIFF].forEach(col =>
    sheet.getRange(row, col).setNumberFormat('₪#,##0.00')
  );

  // פורמט שעות
  if (sheet.getRange(row, COL.HOURS).getValue() !== '') {
    sheet.getRange(row, COL.HOURS).setNumberFormat('0.0');
  }

  sheet.getRange(row, COL.NOTES).setHorizontalAlignment('right');
}

function recolorAllRows(sheet) {
  const lastData = getLastDataRow(sheet);
  for (let r = DATA_START_ROW; r <= lastData; r++) {
    applyRowFormat(sheet, r);
  }
}

function getLastDataRow(sheet) {
  let last = DATA_START_ROW - 1;
  for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
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
  sheet.getRange(DATA_START_ROW, 1, lastData - DATA_START_ROW + 1, NUM_COLS)
    .sort({ column: COL.DATE, ascending: true });
}

// =============================================
// סיכום עם נוסחאות (מתעדכן אוטומטית!)
// =============================================

function refreshCurrentSheetSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  writeSummarySection(sheet);
  addConditionalFormatting(sheet);
  SpreadsheetApp.getUi().alert('🔄 הנוסחאות והעיצוב עודכנו!');
}

// ממיר גליון ישן (ערכים קשיחים) לנוסחאות אוטומטיות
function fixOldSheet() {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastData = getLastDataRow(sheet);

  if (lastData < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('אין נתונים לתיקון בגליון זה.');
    return;
  }

  const CL = colLetter;
  let fixed = 0;

  for (let r = DATA_START_ROW; r <= lastData; r++) {
    // דלג על שורות ריקות
    if (sheet.getRange(r, COL.DATE).getValue() === '') continue;

    // החלף ערכים קשיחים בנוסחאות
    sheet.getRange(r, COL.TOTAL)
      .setFormula(`=${CL(COL.CASH)}${r}+${CL(COL.CHECK)}${r}`);
    sheet.getRange(r, COL.DIFF)
      .setFormula(`=${CL(COL.TOTAL)}${r}-${CL(COL.OWED)}${r}`);

    applyRowFormat(sheet, r);
    fixed++;
  }

  // עדכן סיכום ועיצוב תנאי
  writeSummarySection(sheet);
  addConditionalFormatting(sheet);

  SpreadsheetApp.getUi().alert(`✅ תוקנו ${fixed} שורות — כל עמודות סה"כ והפרש הן עכשיו נוסחאות!`);
}

function writeSummarySection(sheet) {
  const CL  = colLetter;
  const fmt = '₪#,##0.00';

  // ניקוי
  sheet.getRange(SUMMARY_START - 1, 1, 20, NUM_COLS).clearContent().clearFormat();

  // קו הפרדה
  sheet.getRange(SUMMARY_START - 1, 1, 1, NUM_COLS).merge()
    .setValue('').setBackground(C_HEADER_BG);
  sheet.setRowHeight(SUMMARY_START - 1, 6);

  // כותרת
  sheet.getRange(SUMMARY_START, 1, 1, NUM_COLS).merge()
    .setValue('📊  סיכום חודשי  (מתעדכן אוטומטית)')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');
  sheet.setRowHeight(SUMMARY_START, 36);

  const r = SUMMARY_START + 1;

  // שורה 1: ימי עבודה | סכום מגיע
  writeSummaryFormulaRow(sheet, r,
    'ימי עבודה',
    `=COUNTA(A${DATA_START_ROW}:A${DATA_END_ROW})`,
    '0 "ימים"', false,
    'סכום מגיע',
    `=SUM(${CL(COL.OWED)}${DATA_START_ROW}:${CL(COL.OWED)}${DATA_END_ROW})`,
    fmt, false);

  // שורה 2: מזומן | צ'קים
  writeSummaryFormulaRow(sheet, r + 1,
    'מזומן שהתקבל',
    `=SUM(${CL(COL.CASH)}${DATA_START_ROW}:${CL(COL.CASH)}${DATA_END_ROW})`,
    fmt, false,
    "צ'קים",
    `=SUM(${CL(COL.CHECK)}${DATA_START_ROW}:${CL(COL.CHECK)}${DATA_END_ROW})`,
    fmt, false);

  // שורה 3: סה"כ התקבל | הפרש כולל
  writeSummaryFormulaRow(sheet, r + 2,
    'סה"כ התקבל',
    `=SUM(${CL(COL.TOTAL)}${DATA_START_ROW}:${CL(COL.TOTAL)}${DATA_END_ROW})`,
    fmt, true,
    'הפרש כולל',
    `=SUM(${CL(COL.DIFF)}${DATA_START_ROW}:${CL(COL.DIFF)}${DATA_END_ROW})`,
    fmt, true);

  // כותרת ממוצעים
  sheet.getRange(r + 3, 1, 1, NUM_COLS).merge()
    .setValue('📈  ממוצעים')
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e3f2fd').setFontColor('#0d47a1');
  sheet.setRowHeight(r + 3, 28);

  // שורה 4: ממוצע ליום | ממוצע לשעה
  writeSummaryFormulaRow(sheet, r + 4,
    'ממוצע ליום',
    `=IFERROR(SUM(${CL(COL.TOTAL)}${DATA_START_ROW}:${CL(COL.TOTAL)}${DATA_END_ROW})/COUNTA(A${DATA_START_ROW}:A${DATA_END_ROW}),0)`,
    fmt, false,
    'ממוצע לשעה',
    `=IFERROR(SUM(${CL(COL.TOTAL)}${DATA_START_ROW}:${CL(COL.TOTAL)}${DATA_END_ROW})/SUM(${CL(COL.HOURS)}${DATA_START_ROW}:${CL(COL.HOURS)}${DATA_END_ROW}),0)`,
    fmt, false);

  // שורה 5: ממוצע שעות ליום
  writeSummaryFormulaRow(sheet, r + 5,
    'ממוצע שעות ליום',
    `=IFERROR(SUM(${CL(COL.HOURS)}${DATA_START_ROW}:${CL(COL.HOURS)}${DATA_END_ROW})/COUNTA(A${DATA_START_ROW}:A${DATA_END_ROW}),0)`,
    '0.0 "שע\'"', false,
    null, null, null, false);

  [r, r + 1, r + 2, r + 4, r + 5].forEach(row => sheet.setRowHeight(row, 34));

  // הערת תלוש שכר
  const noteRow = r + 7;
  sheet.getRange(noteRow, 1, 1, NUM_COLS).merge()
    .setValue('⚠️  תלוש שכר מתקבל ב-10 לחודש — ייתכן שסכומים מתחילת החודש הבא ייכנסו לתלוש הנוכחי')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center').setWrap(true)
    .setBackground(C_WARN_BG).setFontColor(C_WARN_FG);
  sheet.setRowHeight(noteRow, 32);
}

function writeSummaryFormulaRow(sheet, row, label1, formula1, fmt1, bold1, label2, formula2, fmt2, bold2) {
  // תווית 1
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue(label1).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');

  // ערך 1 (נוסחה)
  const c1 = sheet.getRange(row, 3);
  c1.setFormula(formula1).setHorizontalAlignment('center')
    .setBackground('#ffffff').setFontColor('#1565c0');
  if (fmt1)  c1.setNumberFormat(fmt1);
  if (bold1) c1.setFontWeight('bold').setFontSize(12);

  // אם אין צד ימין — נקה ואל תכתוב
  if (!label2) {
    sheet.getRange(row, 4, 1, NUM_COLS - 3).clearContent().clearFormat()
      .setBackground('#ffffff');
    return;
  }

  // רווח
  sheet.getRange(row, 4).setValue('').setBackground('#ffffff');

  // תווית 2
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue(label2).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_SUMMARY_BG).setFontColor('#1a237e');

  // ערך 2 (נוסחה) – הפרש מקבל צבע מ-Conditional Formatting
  const c2 = sheet.getRange(row, 7);
  c2.setFormula(formula2).setHorizontalAlignment('center')
    .setBackground(null).setFontColor(null);
  if (fmt2)  c2.setNumberFormat(fmt2);
  if (bold2) c2.setFontWeight('bold').setFontSize(12);

  // ניקוי שאר
  sheet.getRange(row, 8, 1, NUM_COLS - 7).clearContent().clearFormat()
    .setBackground('#ffffff');
}

// =============================================
// עיצוב תנאי – הפרש (ירוק / אדום אוטומטי)
// =============================================

function addConditionalFormatting(sheet) {
  // תא הפרש בנתונים
  const diffData    = sheet.getRange(DATA_START_ROW, COL.DIFF, DATA_MAX_ROWS, 1);
  // תא הפרש בסיכום (שורה r+2, עמודה 7)
  const diffSummary = sheet.getRange(SUMMARY_START + 3, 7);

  const posRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setBackground(C_GREEN_BG).setFontColor(C_GREEN_FG).setBold(true)
    .setRanges([diffData, diffSummary])
    .build();

  const negRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground(C_RED_BG).setFontColor(C_RED_FG).setBold(true)
    .setRanges([diffData, diffSummary])
    .build();

  sheet.setConditionalFormatRules([posRule, negRule]);
}

// =============================================
// לוח מחוונים שנתי
// =============================================

function updateDashboard() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const currentYear = new Date().getFullYear();

  let dash = ss.getSheetByName('📊 לוח מחוונים');
  if (!dash) dash = ss.insertSheet('📊 לוח מחוונים', 0);

  dash.clear();
  dash.clearFormats();

  dash.getRange(1, 1, 1, 8).merge()
    .setValue('📊  סיכום שנתי  –  ' + currentYear)
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(C_TITLE_BG).setFontColor('#ffffff');
  dash.setRowHeight(1, 55);

  const hdr = ['חודש','ימי עבודה','סכום מגיע','מזומן',"צ'קים",'סה"כ התקבל','הפרש'];
  dash.getRange(2, 1, 1, 7).setValues([hdr])
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center')
    .setBackground(C_HEADER_BG).setFontColor('#ffffff');
  dash.setRowHeight(2, 38);
  dash.setFrozenRows(2);

  let yDays=0, yOwed=0, yCash=0, yChk=0, yRec=0, yDiff=0;

  MONTHS_HE.forEach((mName, mIdx) => {
    const mSheet = ss.getSheetByName(sheetName(currentYear, mIdx));
    const r      = 3 + mIdx;
    let days=0, owed=0, cash=0, chk=0, rec=0, dif=0;

    if (mSheet) {
      const lastData = getLastDataRow(mSheet);
      if (lastData >= DATA_START_ROW) {
        const rows = mSheet.getRange(DATA_START_ROW, 1, lastData - DATA_START_ROW + 1, NUM_COLS).getValues();
        rows.forEach(row => {
          if (row[COL.DATE-1] !== '') {
            days++;
            owed += parseFloat(row[COL.OWED -1]) || 0;
            cash += parseFloat(row[COL.CASH -1]) || 0;
            chk  += parseFloat(row[COL.CHECK-1]) || 0;
            rec  += parseFloat(row[COL.TOTAL-1]) || 0;
            dif  += parseFloat(row[COL.DIFF -1]) || 0;
          }
        });
      }
    }

    const bg  = (mIdx % 2 === 0) ? C_ODD_ROW : C_EVEN_ROW;
    const fmt = '₪#,##0.00';

    if (mSheet) {
      dash.getRange(r, 1).setFormula('=HYPERLINK("#gid=' + mSheet.getSheetId() + '","' + mName + '")');
    } else {
      dash.getRange(r, 1).setValue(mName);
    }
    dash.getRange(r, 1).setFontWeight('bold').setBackground(bg).setHorizontalAlignment('center').setFontColor('#1a237e');
    dash.getRange(r, 2).setValue(days).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 3).setValue(owed).setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 4).setValue(cash).setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 5).setValue(chk) .setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center');
    dash.getRange(r, 6).setValue(rec) .setNumberFormat(fmt).setBackground(bg).setHorizontalAlignment('center').setFontWeight('bold');

    const dc = dash.getRange(r, 7);
    dc.setValue(dif).setNumberFormat(fmt).setHorizontalAlignment('center').setFontWeight('bold')
      .setBackground(dif >= 0 ? C_GREEN_BG : C_RED_BG)
      .setFontColor(dif >= 0 ? C_GREEN_FG : C_RED_FG);

    dash.setRowHeight(r, 30);
    yDays+=days; yOwed+=owed; yCash+=cash; yChk+=chk; yRec+=rec; yDiff+=dif;
  });

  const totRow = 16;
  dash.getRange(totRow, 1, 1, 7).setBackground(C_TITLE_BG).setFontColor('#ffffff').setFontWeight('bold');
  [' סה"כ שנתי', yDays, yOwed, yCash, yChk, yRec, yDiff].forEach((v, i) => {
    const cell = dash.getRange(totRow, i + 1);
    cell.setValue(v).setHorizontalAlignment('center').setFontSize(12);
    if (i >= 2) cell.setNumberFormat('₪#,##0.00');
  });
  dash.setRowHeight(totRow, 42);

  [100, 85, 115, 105, 105, 125, 105].forEach((w, i) => dash.setColumnWidth(i + 1, w));

  ss.setActiveSheet(dash);
  SpreadsheetApp.getUi().alert('✅ לוח המחוונים עודכן!');
}
