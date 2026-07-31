/**
 * IIEC — membership interest form backend.
 *
 * POST from IIEC_membership_v1.html  ->  formatted Google Sheet  ->  .xlsx in Drive.
 *
 * FIRST RUN
 *  1. script.google.com -> your project -> replace the file with this one -> Save.
 *  2. Run `setup` from the editor and approve the authorisation prompt
 *     (Sheets + Drive + external requests). This builds the header row, banding,
 *     filter, column formats, dropdowns, named range and the Summary tab, then
 *     writes the first .xlsx.
 *  3. Deploy -> Manage deployments -> pencil -> Version: NEW VERSION -> Deploy.
 *     The /exec URL serves the last PUBLISHED version, so repeat this step every
 *     time you edit the code. Editing and saving alone changes nothing.
 *
 * HELPERS you can run from the editor at any time
 *    setup()           rebuild formatting + Summary, then export
 *    reformat()        re-apply table formatting over existing rows
 *    refreshXlsxNow()  regenerate the .xlsx without adding a row
 *    testSubmit()      insert one fake row end-to-end
 *    deleteTestRows()  delete rows whose email ends in @example.edu
 */

/* ------------------------------ CONFIG ------------------------------ */
var CONFIG = {
  SHEET_ID:       '1lG_RAayRJlLjd1wHhhXUZQ8_RHHoaA_MK_4hew-JLGc',
  SHEET_NAME:     'Interest List',
  SUMMARY_NAME:   'Summary',
  XLSX_FOLDER_ID: '1TH93VmoOlydJkjpY1p8jO4XWxwY1jSCc',
  XLSX_NAME:      'IIEC Interest List.xlsx',
  NAMED_RANGE:    'InterestData',
  SHARED_SECRET:  '',      // '' = no token check
  NOTIFY_EMAIL:   ''       // '' = no email alerts; else 'you@gmail.com'
};

/* Table layout. Keep HEADERS, COL and WIDTHS in step if you add a field. */
var HEADERS = ['No.', 'Timestamp', 'Name', 'Email', 'Phone', 'Branch',
               'Branch (label)', 'Year of study', 'Page', 'User agent'];
var COL = { NO:1, TS:2, NAME:3, EMAIL:4, PHONE:5, BRANCH:6, LABEL:7, YEAR:8, PAGE:9, UA:10 };
var WIDTHS = [50, 155, 190, 235, 130, 92, 205, 108, 210, 250];

/* Must mirror the <select> options in the form. */
var BRANCHES = ['CSE','IT','AIML','DS','ECE','EEE','MECH','CIVIL','CHEM','BIOTECH','MBA','OTHER'];
var YEARS    = ['1','2','3','4','other'];

var THEME = { header:'#1D4ED8', headerText:'#FFFFFF', line:'#D1D5DB', band:'#F3F6FC' };

/* ------------------------------ ROUTES ------------------------------ */

function doGet() {
  return json({ ok:true, service:'IIEC interest form', time:new Date().toISOString() });
}

function doPost(e) {
  try {
    var data = parseBody(e);

    if (CONFIG.SHARED_SECRET && data.token !== CONFIG.SHARED_SECRET) {
      return json({ ok:false, error:'Unauthorised' });
    }

    var name   = trim(data.fullName);
    var email  = trim(data.email);
    var phone  = trim(data.phone);
    var digits = phone.replace(/\D/g, '');

    if (!name) return json({ ok:false, error:'Name is required' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json({ ok:false, error:'Invalid email' });
    if (digits.length < 10 || digits.length > 13) return json({ ok:false, error:'Invalid phone number' });
    if (!trim(data.branch)) return json({ ok:false, error:'Branch is required' });

    var rowNumber;
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);                     // serialise concurrent submissions
    try {
      var sheet = getSheet();

      if (findRow(sheet, email, digits) > 0) {
        return json({ ok:true, duplicate:true, message:'Already on the list' });
      }

      rowNumber = sheet.getLastRow() + 1;
      sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([[
        rowNumber - 1,                                        // No.
        data.submittedAt ? new Date(data.submittedAt) : new Date(),
        titleCase(name),
        email.toLowerCase(),
        phone,
        trim(data.branch).toUpperCase(),
        trim(data.branchLabel),
        trim(data.studyYear),
        trim(data.page),
        trim(data.userAgent)
      ]]);

      styleTable(sheet);                      // keeps banding, filter and borders in step
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    // Neither of these may break the submission.
    var xlsxOk = false;
    try { exportXlsx(); xlsxOk = true; }
    catch (err) { console.warn('xlsx export failed: ' + err); }

    if (CONFIG.NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, 'New IIEC interest signup',
          [name, email, phone, trim(data.branchLabel), trim(data.studyYear)].join('\n'));
      } catch (err) { console.warn('notify failed: ' + err); }
    }

    return json({ ok:true, row:rowNumber, xlsx:xlsxOk });
  } catch (err) {
    console.error(err);
    return json({ ok:false, error:String(err && err.message || err) });
  }
}

/* ------------------------------ SHEET / TABLE ------------------------------ */

function getSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME, 0);
    styleTable(sheet);
  } else if (sheet.getLastRow() === 0) {
    styleTable(sheet);
  }
  return sheet;
}

/**
 * Applies the whole table treatment: header row, frozen pane, alternating row
 * bands, filter, borders, column widths, number formats and dropdowns.
 * Safe to run repeatedly — it clears what it owns before re-applying.
 */
function styleTable(sheet) {
  var cols = HEADERS.length;
  var last = Math.max(sheet.getLastRow(), 1);
  var rows = last - 1;                                  // data rows, excluding header

  // trim stray columns so the exported table has no empty tail
  if (sheet.getMaxColumns() > cols) sheet.deleteColumns(cols + 1, sheet.getMaxColumns() - cols);

  // header
  sheet.getRange(1, 1, 1, cols)
    .setValues([HEADERS])
    .setBackground(THEME.header)
    .setFontColor(THEME.headerText)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(1, 30);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  WIDTHS.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  // banding — the closest equivalent to an Excel table style that survives export
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(1, 1, Math.max(last, 2), cols)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(THEME.header)
    .setFirstRowColor('#FFFFFF')
    .setSecondRowColor(THEME.band);

  // filter over the whole table
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, Math.max(last, 2), cols).createFilter();

  // formats
  sheet.getRange(2, COL.NO, Math.max(rows, 1), 1).setNumberFormat('0').setHorizontalAlignment('center');
  sheet.getRange(2, COL.TS, Math.max(rows, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, COL.PHONE, Math.max(rows, 1), 1).setNumberFormat('@');   // keeps +91 / leading zeros
  sheet.getRange(2, COL.BRANCH, Math.max(rows, 1), 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL.YEAR, Math.max(rows, 1), 1).setHorizontalAlignment('center');

  if (rows > 0) {
    var body = sheet.getRange(2, 1, rows, cols);
    body.setFontSize(10).setVerticalAlignment('middle').setWrap(false);
    body.setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);

    // dropdowns so manual edits stay consistent with the form
    sheet.getRange(2, COL.BRANCH, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(BRANCHES, true).setAllowInvalid(true).build());
    sheet.getRange(2, COL.YEAR, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(YEARS, true).setAllowInvalid(true).build());

    // renumber the No. column so it always reads 1..n after deletions
    var seq = [];
    for (var i = 0; i < rows; i++) seq.push([i + 1]);
    sheet.getRange(2, COL.NO, rows, 1).setValues(seq);
  }

  // named range: use =InterestData in formulas, exports to Excel as a defined name
  var ss = sheet.getParent();
  ss.getNamedRanges().forEach(function (nr) {
    if (nr.getName() === CONFIG.NAMED_RANGE) nr.remove();
  });
  ss.setNamedRange(CONFIG.NAMED_RANGE, sheet.getRange(1, 1, Math.max(last, 2), cols));
}

/** Returns the row index of a matching email or phone, or 0. */
function findRow(sheet, email, digits) {
  var rows = sheet.getLastRow() - 1;
  if (rows < 1) return 0;
  var values = sheet.getRange(2, COL.EMAIL, rows, 2).getValues();   // email + phone
  var mail = String(email).trim().toLowerCase();
  for (var i = 0; i < values.length; i++) {
    var rowMail  = String(values[i][0]).trim().toLowerCase();
    var rowPhone = String(values[i][1]).replace(/\D/g, '');
    if (rowMail === mail || (digits && rowPhone === digits)) return i + 2;
  }
  return 0;
}

/* ------------------------------ SUMMARY TAB ------------------------------ */
/**
 * Live counts written as formulas, so they stay correct in the Sheet and in the
 * exported .xlsx without the script having to recalculate anything.
 */
function buildSummary() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SUMMARY_NAME) || ss.insertSheet(CONFIG.SUMMARY_NAME);
  var src = "'" + CONFIG.SHEET_NAME + "'!";

  sheet.clear();
  sheet.getBandings().forEach(function (b) { b.remove(); });

  sheet.getRange('A1').setValue('IIEC — Interest List Summary')
    .setFontSize(14).setFontWeight('bold').setFontColor(THEME.header);
  sheet.getRange('A2').setFormula('="Updated "&TEXT(NOW(),"yyyy-mm-dd hh:mm")')
    .setFontColor('#6B7280').setFontSize(9);

  sheet.getRange('A4').setValue('Total signups').setFontWeight('bold');
  sheet.getRange('B4').setFormula('=COUNTA(' + src + '$D$2:$D)');
  sheet.getRange('A5').setValue('Latest submission').setFontWeight('bold');
  sheet.getRange('B5').setFormula('=IFERROR(TEXT(MAX(' + src + '$B$2:$B),"yyyy-mm-dd hh:mm"),"—")');
  sheet.getRange('A6').setValue('Signups today').setFontWeight('bold');
  sheet.getRange('B6').setFormula('=COUNTIFS(' + src + '$B$2:$B,">="&TODAY(),' + src + '$B$2:$B,"<"&TODAY()+1)');

  // by branch
  sheet.getRange('A8:B8').setValues([['Branch', 'Count']])
    .setFontWeight('bold').setBackground(THEME.header).setFontColor(THEME.headerText);
  BRANCHES.forEach(function (code, i) {
    var r = 9 + i;
    sheet.getRange(r, 1).setValue(code);
    sheet.getRange(r, 2).setFormula('=COUNTIF(' + src + '$F$2:$F,A' + r + ')');
  });
  var branchEnd = 8 + BRANCHES.length;
  sheet.getRange(branchEnd + 1, 1).setValue('Total').setFontWeight('bold');
  sheet.getRange(branchEnd + 1, 2).setFormula('=SUM(B9:B' + branchEnd + ')').setFontWeight('bold');

  // by year
  sheet.getRange('D8:E8').setValues([['Year of study', 'Count']])
    .setFontWeight('bold').setBackground(THEME.header).setFontColor(THEME.headerText);
  var yearLabels = ['1st year', '2nd year', '3rd year', '4th year', 'Other'];
  YEARS.forEach(function (code, i) {
    var r = 9 + i;
    sheet.getRange(r, 4).setValue(yearLabels[i]);
    sheet.getRange(r, 5).setFormula('=COUNTIF(' + src + '$H$2:$H,"' + code + '")');
  });
  sheet.getRange(9 + YEARS.length, 4).setValue('Not answered').setFontWeight('bold');
  sheet.getRange(9 + YEARS.length, 5)
    .setFormula('=COUNTA(' + src + '$D$2:$D)-SUM(E9:E' + (8 + YEARS.length) + ')').setFontWeight('bold');

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);
  sheet.getRange(8, 1, BRANCHES.length + 2, 2)
    .setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(8, 4, YEARS.length + 2, 2)
    .setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
}

/* ------------------------------ XLSX ------------------------------ */
/**
 * Exports the whole workbook as .xlsx into the Drive folder, updating the file
 * in place so its ID and share link never change.
 */
function exportXlsx() {
  var token = ScriptApp.getOAuthToken();

  var blob = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + token } }
  ).getBlob().setName(CONFIG.XLSX_NAME);

  var folder = DriveApp.getFolderById(CONFIG.XLSX_FOLDER_ID);
  var existing = folder.getFilesByName(CONFIG.XLSX_NAME);

  if (existing.hasNext()) {
    var fileId = existing.next().getId();
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files/' + fileId + '?uploadType=media&supportsAllDrives=true',
      {
        method: 'patch',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        payload: blob.getBytes(),
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (res.getResponseCode() >= 300) {
      throw new Error('Drive update ' + res.getResponseCode() + ': ' + res.getContentText());
    }
    return fileId;
  }
  return folder.createFile(blob).getId();
}

/* ------------------------------ HELPERS ------------------------------ */

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { /* fall through */ }
  }
  return (e && e.parameter) || {};
}

function trim(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function titleCase(s) {
  return s.replace(/\s+/g, ' ').replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ MAINTENANCE ------------------------------ */

/** Run once after pasting this file: builds everything and grants permissions. */
function setup() {
  var sheet = getSheet();
  styleTable(sheet);
  buildSummary();
  console.log('sheet ready, rows: ' + Math.max(sheet.getLastRow() - 1, 0));
  console.log('xlsx file id: ' + exportXlsx());
}

/** Re-apply table formatting over whatever rows exist now. */
function reformat() {
  styleTable(getSheet());
  buildSummary();
  exportXlsx();
  console.log('reformatted and exported');
}

function refreshXlsxNow() {
  console.log('xlsx file id: ' + exportXlsx());
}

function testSubmit() {
  var out = doPost({ postData: { contents: JSON.stringify({
    token: CONFIG.SHARED_SECRET,
    fullName: 'test student',
    email: 'test.student@example.edu',
    phone: '9876543210',
    branch: 'CSE',
    branchLabel: 'Computer Science (CSE)',
    studyYear: '2',
    submittedAt: new Date().toISOString(),
    page: 'manual test'
  }) } });
  console.log(out.getContent());
}

/** Removes rows whose email ends in @example.edu (the test entries). */
function deleteTestRows() {
  var sheet = getSheet();
  var rows = sheet.getLastRow() - 1;
  if (rows < 1) return console.log('nothing to delete');
  var emails = sheet.getRange(2, COL.EMAIL, rows, 1).getValues();
  var removed = 0;
  for (var i = emails.length - 1; i >= 0; i--) {
    if (/@example\.edu$/i.test(String(emails[i][0]).trim())) {
      sheet.deleteRow(i + 2);
      removed++;
    }
  }
  styleTable(sheet);
  buildSummary();
  exportXlsx();
  console.log('deleted ' + removed + ' test row(s)');
}
