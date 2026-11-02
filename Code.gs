/**
 * Daily & Hourly Tracker — Comprehensive Google Apps Script Automation (`Code.gs`)
 *
 * Automates the creation and management of the 7-tab productivity workbook:
 * 1. Master_Log (Backend Database)
 * 2. Dashboard & Summary (Executive Analytics)
 * 3. Date-Wise View (Daily Review)
 * 4. Task-Wise View (Lifecycle History)
 * 5. Stakeholder-Wise View (Collaboration & Accountability)
 * 6. Master Progress View (Panoramic Audit Grid)
 * 7. Timeline View (Execution Gantt & Date Filtering)
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Calendar Follow-Up')
    .addItem('Schedule Selected Task Follow-Up', 'scheduleFollowUpCalendarEvent')
    .addSeparator()
    .addItem('Initialize / Reset Workbook', 'setupMasterLogTracker')
    .addToUi();
}

/**
 * Handles Web App GET requests from the web dashboard to fetch tasks or auto-initialize sheet.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Master_Log');

  // If Master_Log doesn't exist, auto-initialize the workbook structure
  if (!logSheet) {
    setupMasterLogTracker();
    ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.getSheetByName('Master_Log');
  }

  var rows = logSheet.getDataRange().getValues();
  var headers = rows.shift(); // remove header

  var data = rows.map(function(r, index) {
    return {
      id: index + 1,
      date: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      taskName: r[1] || '',
      category: r[2] || '',
      timeBlock: r[3] || '',
      priority: r[4] || '',
      planned: r[5] || 1,
      actual: r[6] || 1,
      stakeholder: r[7] || '',
      status: r[8] || 'In Progress',
      progress: r[9] || 0,
      nextUpdate: r[10] ? Utilities.formatDate(new Date(r[10]), Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      followUp: r[11] || 'No',
      timeSlot: r[12] || '',
      notes: r[13] || ''
    };
  });

  return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles Web App POST requests to save new task entries directly into Master_Log.
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Master_Log');
    if (!logSheet) {
      setupMasterLogTracker();
      logSheet = ss.getSheetByName('Master_Log');
    }

    var payload = JSON.parse(e.postData.contents);

    logSheet.appendRow([
      payload.date || new Date(),
      payload.taskName || '',
      payload.category || 'Deep Work',
      payload.timeBlock || '09:00 AM - 10:00 AM',
      payload.priority || 'High',
      payload.planned || 1.0,
      payload.actual !== undefined ? payload.actual : '=IF(ISBLANK(D' + (logSheet.getLastRow() + 1) + '), 0, 1)',
      payload.stakeholder || 'Internal',
      payload.status || 'In Progress',
      payload.progress || 0,
      payload.nextUpdate || '',
      payload.followUp || 'No',
      payload.timeSlot || '10:00 AM',
      payload.notes || ''
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Row appended successfully' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Workbook Generator & Initializer (`setupMasterLogTracker`)
 * Creates all 7 sheets, configures professional header styling, alternating row colors,
 * column widths, frozen rows, conditional formatting, data validations, and formulas.
 */
function setupMasterLogTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetNames = [
    'Master_Log',
    'Dashboard & Summary',
    'Date-Wise View',
    'Task-Wise View',
    'Stakeholder-Wise View',
    'Master Progress View',
    'Timeline View'
  ];

  sheetNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // ----------------------------------------------------
  // TAB 1: Master_Log Setup (Backend SSOT)
  // ----------------------------------------------------
  var logSheet = ss.getSheetByName('Master_Log');
  logSheet.clear();

  var logHeaders = [
    'Date', 'Task Name', 'Category', 'Time Block', 'Priority',
    'Planned (Hrs)', 'Actual (Hrs - Auto)', 'Stakeholder', 'Status',
    'Progress (%)', 'Next Update Expected', 'Follow-Up Required?',
    'Follow-Up Date & Time', 'Daily Update / Notes'
  ];

  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  formatHeaderRow(logSheet, logHeaders.length);

  // Sample Data Entry
  var sampleRow = [
    new Date(), 'Architecture Blueprint Review', 'Deep Work', '09:00 AM - 10:00 AM',
    'High', 2.0, '=IF(ISBLANK(D2), 0, 1)', 'Internal Team', 'In Progress',
    50, new Date(new Date().getTime() + 86400000), 'Yes', '02:00 PM', 'Initial draft complete and reviewed.'
  ];
  logSheet.getRange(2, 1, 1, sampleRow.length).setValues([sampleRow]);
  logSheet.getRange('G2').setFormula('=IF(ISBLANK(D2), 0, 1)');

  logSheet.getRange('A2:A100').setNumberFormat('yyyy-mm-dd');
  logSheet.getRange('F2:G100').setNumberFormat('0.0');
  logSheet.getRange('K2:K100').setNumberFormat('yyyy-mm-dd');

  applyMasterLogValidations(logSheet);
  applyConditionalFormatting(logSheet);

  // ----------------------------------------------------
  // TAB 2: Dashboard & Summary Setup
  // ----------------------------------------------------
  var dashSheet = ss.getSheetByName('Dashboard & Summary');
  dashSheet.clear();

  dashSheet.getRange('A1').setValue('DAILY & HOURLY TRACKER — EXECUTIVE DASHBOARD');
  dashSheet.getRange('A1').setFontWeight('bold').setFontSize(16).setFontColor('#0f172a');
  dashSheet.getRange('A2').setValue('Single Source of Truth (SSOT) Executive Analytics & Performance Indicators');
  dashSheet.getRange('A2').setFontStyle('italic').setFontColor('#64748b').setFontSize(11);

  dashSheet.getRange('A4').setValue('Start Date:');
  dashSheet.getRange('A4').setFontWeight('bold');
  dashSheet.getRange('B4').setValue('');
  dashSheet.getRange('A5').setValue('End Date:');
  dashSheet.getRange('A5').setFontWeight('bold');
  dashSheet.getRange('B5').setValue('');

  var kpiLabels = ['TOTAL TASKS', 'COMPLETED TASKS', 'COMPLETION RATE', 'DEEP WORK HOURS', 'TOTAL LOGGED HOURS'];
  dashSheet.getRange('A7:E7').setValues([kpiLabels]);
  formatHeaderRow(dashSheet, 5);
  dashSheet.getRange('A7:E7').setFontSize(11);

  dashSheet.getRange('A8').setFormula('=COUNTA(Master_Log!A2:A200)');
  dashSheet.getRange('B8').setFormula('=COUNTIF(Master_Log!I2:I200, "Completed")');
  dashSheet.getRange('C8').setFormula('=IFERROR(B8/A8, 0)');
  dashSheet.getRange('C8').setNumberFormat('0.0%');
  dashSheet.getRange('D8').setFormula('=SUMIFS(Master_Log!G2:G200, Master_Log!C2:C200, "Deep Work", Master_Log!I2:I200, "Completed")');
  dashSheet.getRange('E8').setFormula('=SUM(Master_Log!G2:G200)');
  dashSheet.getRange('A8:E8').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');

  dashSheet.getRange('A11').setValue('Category Breakdown & Performance');
  dashSheet.getRange('A11').setFontWeight('bold').setFontSize(13);
  var catHeaders = ['Category', 'Task Count', 'Planned Hours', 'Actual Hours', 'Variance (P - A)', 'Share of Total'];
  dashSheet.getRange('A12:F12').setValues([catHeaders]);
  formatHeaderRow(dashSheet, 6);

  var categories = ['Deep Work', 'Admin', 'Meetings', 'Health & Fitness', 'Personal', 'Learning'];
  for (var i = 0; i < categories.length; i++) {
    var r = 13 + i;
    dashSheet.getRange(r, 1).setValue(categories[i]);
    dashSheet.getRange(r, 2).setFormula('=COUNTIF(Master_Log!$C$2:$C$200, A' + r + ')');
    dashSheet.getRange(r, 3).setFormula('=SUMIF(Master_Log!$C$2:$C$200, A' + r + ', Master_Log!$F$2:$F$200)');
    dashSheet.getRange(r, 4).setFormula('=SUMIF(Master_Log!$C$2:$C$200, A' + r + ', Master_Log!$G$2:$G$200)');
    dashSheet.getRange(r, 5).setFormula('=C' + r + ' - D' + r);
    dashSheet.getRange(r, 6).setFormula('=IFERROR(D' + r + '/SUM($D$13:$D$18), 0)');
    dashSheet.getRange(r, 3, 1, 3).setNumberFormat('0.0');
    dashSheet.getRange(r, 6).setNumberFormat('0.0%');
  }

  // ----------------------------------------------------
  // TAB 3: Date-Wise View Setup
  // ----------------------------------------------------
  var dateSheet = ss.getSheetByName('Date-Wise View');
  dateSheet.clear();
  dateSheet.getRange('A1').setValue('DAILY REVIEW — SINGLE DATE STANDUP');
  dateSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  dateSheet.getRange('A2').setValue('Select Date:');
  dateSheet.getRange('A2').setFontWeight('bold');
  dateSheet.getRange('B2').setValue(new Date());
  dateSheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  dateSheet.getRange('B2').setNumberFormat('yyyy-mm-dd');
  dateSheet.getRange('A4').setFormula('=IFERROR(FILTER(Master_Log!A1:N200, Master_Log!A1:A200 = B2), "No entries found for selected date.")');

  // ----------------------------------------------------
  // TAB 4: Task-Wise View Setup
  // ----------------------------------------------------
  var taskSheet = ss.getSheetByName('Task-Wise View');
  taskSheet.clear();
  taskSheet.getRange('A1').setValue('TASK LIFECYCLE HISTORY');
  taskSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  taskSheet.getRange('A2').setValue('Select Task:');
  taskSheet.getRange('A2').setFontWeight('bold');
  taskSheet.getRange('B2').setValue('Architecture Blueprint Review');
  var taskVal = SpreadsheetApp.newDataValidation().requireValueInRange(logSheet.getRange('B2:B50')).build();
  taskSheet.getRange('B2').setDataValidation(taskVal);
  taskSheet.getRange('A4').setFormula('=IFERROR(FILTER(Master_Log!A1:N200, Master_Log!B1:B200 = B2), "No history found for selected task.")');

  // ----------------------------------------------------
  // TAB 5: Stakeholder-Wise View Setup
  // ----------------------------------------------------
  var stakeSheet = ss.getSheetByName('Stakeholder-Wise View');
  stakeSheet.clear();
  stakeSheet.getRange('A1').setValue('STAKEHOLDER COLLABORATION & DELIVERABLES');
  stakeSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  stakeSheet.getRange('A2').setValue('Select Stakeholder:');
  stakeSheet.getRange('A2').setFontWeight('bold');
  stakeSheet.getRange('B2').setValue('Internal Team');
  var stakeVal = SpreadsheetApp.newDataValidation().requireValueInRange(logSheet.getRange('H2:H50')).build();
  stakeSheet.getRange('B2').setDataValidation(stakeVal);
  stakeSheet.getRange('A4').setFormula('=IFERROR(FILTER(Master_Log!A1:N200, Master_Log!H1:H200 = B2), "No deliverables found for stakeholder.")');

  // ----------------------------------------------------
  // TAB 6: Master Progress View Setup
  // ----------------------------------------------------
  var progSheet = ss.getSheetByName('Master Progress View');
  progSheet.clear();
  progSheet.getRange('A1').setValue('MASTER PROGRESS VIEW (PANORAMIC AUDIT GRID)');
  progSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  progSheet.getRange('A3').setFormula('=ARRAYFORMULA(Master_Log!A1:N200)');

  // ----------------------------------------------------
  // TAB 7: Timeline View Setup
  // ----------------------------------------------------
  var timeSheet = ss.getSheetByName('Timeline View');
  timeSheet.clear();
  timeSheet.getRange('A1').setValue('TIMELINE & DATE RANGE GANTT VIEW');
  timeSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  timeSheet.getRange('A2').setValue('Start Date:');
  timeSheet.getRange('A2').setFontWeight('bold');
  timeSheet.getRange('B2').setValue(new Date());
  timeSheet.getRange('B2').setNumberFormat('yyyy-mm-dd');
  timeSheet.getRange('A3').setValue('End Date:');
  timeSheet.getRange('A3').setFontWeight('bold');
  timeSheet.getRange('B3').setValue(new Date(new Date().getTime() + 7*86400000));
  timeSheet.getRange('B3').setNumberFormat('yyyy-mm-dd');
  timeSheet.getRange('A5').setFormula('=IFERROR(FILTER(Master_Log!A1:N200, Master_Log!A1:A200 >= B2, Master_Log!A1:A200 <= B3), "No tasks found in selected date range.")');

  sheetNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    try {
      sh.autoResizeColumns(1, 15);
    } catch(e) {}
  });

  SpreadsheetApp.getUi().alert('Daily & Hourly Tracker initialized with professional formatting across all 7 tabs!');
}

function applyMasterLogValidations(sheet) {
  var categories = ['Deep Work', 'Admin', 'Meetings', 'Health & Fitness', 'Personal', 'Learning'];
  sheet.getRange('C2:C200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(categories, true).build());

  var timeBlocks = [
    '06:00 AM - 07:00 AM', '07:00 AM - 08:00 AM', '08:00 AM - 09:00 AM', '09:00 AM - 10:00 AM',
    '10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '12:00 PM - 01:00 PM', '01:00 PM - 02:00 PM',
    '02:00 PM - 03:00 PM', '03:00 PM - 04:00 PM', '04:00 PM - 05:00 PM', '05:00 PM - 06:00 PM',
    '06:00 PM - 07:00 PM', '07:00 PM - 08:00 PM', '08:00 PM - 09:00 PM', '09:00 PM - 10:00 PM'
  ];
  sheet.getRange('D2:D200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(timeBlocks, true).build());

  var priorities = ['High', 'Medium', 'Low'];
  sheet.getRange('E2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(priorities, true).build());

  var statuses = ['Not Started', 'In Progress', 'Completed', 'Deferred'];
  sheet.getRange('I2:I200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).build());

  var progress = ['0', '25', '50', '75', '100'];
  sheet.getRange('J2:J200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(progress, true).build());

  var yesNo = ['Yes', 'No'];
  sheet.getRange('L2:L200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(yesNo, true).build());

  var slots = [
    '09:00 AM', '10:00 AM', '11:00 AM', '12:00 PM',
    '01:00 PM', '02:00 PM', '03:00 PM', '04:00 PM', '05:00 PM', '06:00 PM'
  ];
  sheet.getRange('M2:M200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(slots, true).build());
}

function applyConditionalFormatting(sheet) {
  var range = sheet.getRange('I2:I200');
  var rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Completed').setBackground('#d1fae5').setFontColor('#065f46').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('In Progress').setBackground('#fef3c7').setFontColor('#92400e').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Not Started').setBackground('#fee2e2').setFontColor('#991b1b').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Deferred').setBackground('#f1f5f9').setFontColor('#475569').setRanges([range]).build());

  sheet.setConditionalFormatRules(rules);
}

function formatHeaderRow(sheet, numCols) {
  var range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground('#0f172a');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function scheduleFollowUpCalendarEvent() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== 'Master_Log') {
    SpreadsheetApp.getUi().alert('Please select a task row in the Master_Log sheet.');
    return;
  }

  var row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a valid data row.');
    return;
  }

  var date = sheet.getRange(row, 1).getValue();
  var taskName = sheet.getRange(row, 2).getValue();
  var stakeholder = sheet.getRange(row, 8).getValue();
  var followUpReq = sheet.getRange(row, 12).getValue();
  var followUpTimeStr = sheet.getRange(row, 13).getValue();
  var notes = sheet.getRange(row, 14).getValue();

  if (followUpReq !== 'Yes') {
    SpreadsheetApp.getUi().alert('Follow-Up Required? is not set to "Yes" for this task.');
    return;
  }

  if (!date || !taskName) {
    SpreadsheetApp.getUi().alert('Task Date and Task Name are required.');
    return;
  }

  var eventDate = new Date(date);
  if (followUpTimeStr) {
    var parts = followUpTimeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (parts) {
      var hours = parseInt(parts[1], 10);
      var minutes = parseInt(parts[2], 10);
      var ampm = parts[3].toUpperCase();
      if (ampm === 'PM' && hours < 12) hours += 12;
      if (ampm === 'AM' && hours === 12) hours = 0;
      eventDate.setHours(hours, minutes, 0, 0);
    }
  } else {
    eventDate.setHours(10, 0, 0, 0);
  }

  var endDate = new Date(eventDate.getTime() + 3600000);
  var title = 'Follow-Up: ' + taskName + (stakeholder ? ' (' + stakeholder + ')' : '');
  var description = 'Daily & Hourly Tracker Follow-Up\n\nTask: ' + taskName + '\nStakeholder: ' + stakeholder + '\nNotes: ' + notes;

  CalendarApp.getDefaultCalendar().createEvent(title, eventDate, endDate, {
    description: description
  });

  SpreadsheetApp.getUi().alert('Successfully scheduled calendar follow-up for ' + taskName + ' on ' + eventDate.toLocaleString() + '!');
}
