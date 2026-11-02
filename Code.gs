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
 * Helper utility to return JSON responses.
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles Web App GET requests from the web dashboard to fetch tasks or auto-initialize sheet.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Master_Log');

  if (!logSheet) {
    setupMasterLogTracker();
    ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.getSheetByName('Master_Log');
  }

  var rows = logSheet.getDataRange().getValues();
  rows.shift(); // remove header

  var timeZone = Session.getScriptTimeZone();
  var data = rows.map(function(r, index) {
    return {
      id: index + 1,
      rowIndex: index + 2,
      date: r[0] ? Utilities.formatDate(new Date(r[0]), timeZone, "yyyy-MM-dd") : '',
      taskName: r[1] || '',
      category: r[2] || '',
      timeBlock: r[3] || '',
      priority: r[4] || '',
      planned: r[5] || 1,
      actual: r[6] || 1,
      stakeholder: r[7] || '',
      status: r[8] || 'In Progress',
      progress: r[9] || 0,
      nextUpdate: r[10] ? Utilities.formatDate(new Date(r[10]), timeZone, "yyyy-MM-dd") : '',
      followUp: r[11] || 'No',
      timeSlot: r[12] || '',
      notes: r[13] || ''
    };
  });

  return jsonResponse({ status: 'success', data: data });
}

/**
 * Handles Web App POST requests to save new task entries or delete entries in Master_Log.
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

    if (payload.action === 'delete') {
      var rowIndex = payload.rowIndex;
      var lastRow = logSheet.getLastRow();

      if (rowIndex && rowIndex >= 2 && rowIndex <= lastRow) {
        logSheet.deleteRow(rowIndex);
        return jsonResponse({ status: 'success', message: 'Row deleted successfully' });
      }

      var targetTask = payload.taskName;
      var targetDate = payload.date;
      var dataRange = logSheet.getDataRange().getValues();
      var timeZone = Session.getScriptTimeZone();

      for (var i = 1; i < dataRange.length; i++) {
        var rowDate = dataRange[i][0] ? Utilities.formatDate(new Date(dataRange[i][0]), timeZone, "yyyy-MM-dd") : '';
        var rowTask = dataRange[i][1] || '';
        if (rowTask === targetTask && (!targetDate || rowDate === targetDate)) {
          logSheet.deleteRow(i + 1);
          return jsonResponse({ status: 'success', message: 'Row deleted by matching criteria' });
        }
      }

      return jsonResponse({ status: 'error', message: 'Row not found for deletion' });
    }

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

    return jsonResponse({ status: 'success', message: 'Row appended successfully' });
  } catch(err) {
    return jsonResponse({ status: 'error', message: err.toString() });
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
    dashSheet.getRange(r, 6).setFormula('=IFERROR(D' + r + '/E$8, 0)');
    dashSheet.getRange(r, 6).setNumberFormat('0.0%');
  }
  dashSheet.getRange('C13:E18').setNumberFormat('0.0');

  dashSheet.getRange('A20').setValue('Priority Breakdown');
  dashSheet.getRange('A20').setFontWeight('bold').setFontSize(13);
  var priHeaders = ['Priority', 'Task Count', 'Completed Count', 'Completion Rate'];
  dashSheet.getRange('A21:D21').setValues([priHeaders]);
  formatHeaderRow(dashSheet, 4);

  var priorities = ['High', 'Medium', 'Low'];
  for (var p = 0; p < priorities.length; p++) {
    var pr = 22 + p;
    dashSheet.getRange(pr, 1).setValue(priorities[p]);
    dashSheet.getRange(pr, 2).setFormula('=COUNTIF(Master_Log!$E$2:$E$200, A' + pr + ')');
    dashSheet.getRange(pr, 3).setFormula('=COUNTIFS(Master_Log!$E$2:$E$200, A' + pr + ', Master_Log!$I$2:$I$200, "Completed")');
    dashSheet.getRange(pr, 4).setFormula('=IFERROR(C' + pr + '/B' + pr + ', 0)');
    dashSheet.getRange(pr, 4).setNumberFormat('0.0%');
  }

  dashSheet.getRange('A27').setValue('Execution Status Breakdown');
  dashSheet.getRange('A27').setFontWeight('bold').setFontSize(13);
  var statHeaders = ['Status', 'Task Count', 'Share of Total'];
  dashSheet.getRange('A28:C28').setValues([statHeaders]);
  formatHeaderRow(dashSheet, 3);

  var statuses = ['In Progress', 'Not Started', 'Completed', 'Deferred'];
  for (var s = 0; s < statuses.length; s++) {
    var sr = 29 + s;
    dashSheet.getRange(sr, 1).setValue(statuses[s]);
    dashSheet.getRange(sr, 2).setFormula('=COUNTIF(Master_Log!$I$2:$I$200, A' + sr + ')');
    dashSheet.getRange(sr, 3).setFormula('=IFERROR(B' + sr + '/$A$8, 0)');
    dashSheet.getRange(sr, 3).setNumberFormat('0.0%');
  }

  dashSheet.setColumnWidth(1, 220);
  dashSheet.setColumnWidth(2, 140);
  dashSheet.setColumnWidth(3, 140);
  dashSheet.setColumnWidth(4, 140);
  dashSheet.setColumnWidth(5, 140);
  dashSheet.setColumnWidth(6, 140);
  dashSheet.setFrozenRows(8);

  // ----------------------------------------------------
  // TAB 3: Date-Wise View Setup
  // ----------------------------------------------------
  setupViewTab(ss, 'Date-Wise View', 'Select Date:', '=IFERROR(FILTER(Master_Log!A2:N200, Master_Log!A2:A200 = B2), "No tasks found for selected date.")', 'Date');

  // ----------------------------------------------------
  // TAB 4: Task-Wise View Setup
  // ----------------------------------------------------
  setupViewTab(ss, 'Task-Wise View', 'Select Task:', '=IFERROR(FILTER(Master_Log!A2:N200, Master_Log!B2:B200 = B2), "No history found for selected task.")', 'Task Name');

  // ----------------------------------------------------
  // TAB 5: Stakeholder-Wise View Setup
  // ----------------------------------------------------
  setupViewTab(ss, 'Stakeholder-Wise View', 'Select Stakeholder:', '=IFERROR(FILTER(Master_Log!A2:N200, Master_Log!H2:H200 = B2), "No records found for selected stakeholder.")', 'Stakeholder');

  // ----------------------------------------------------
  // TAB 6: Master Progress View Setup
  // ----------------------------------------------------
  var progSheet = ss.getSheetByName('Master Progress View');
  progSheet.clear();
  progSheet.getRange('A1').setValue('MASTER PROGRESS AUDIT GRID');
  progSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setFontColor('#0f172a');
  progSheet.getRange('A3').setFormula('=ARRAYFORMULA(Master_Log!A1:N200)');
  progSheet.setFrozenRows(3);
  formatHeaderRow(progSheet, 14);

  // ----------------------------------------------------
  // TAB 7: Timeline View Setup
  // ----------------------------------------------------
  var timeSheet = ss.getSheetByName('Timeline View');
  timeSheet.clear();
  timeSheet.getRange('A1').setValue('TIMELINE & DATE-RANGE EXECUTION VIEW');
  timeSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setFontColor('#0f172a');
  timeSheet.getRange('A3').setValue('Start Date:');
  timeSheet.getRange('A3').setFontWeight('bold');
  timeSheet.getRange('B3').setValue('');
  timeSheet.getRange('D3').setValue('End Date:');
  timeSheet.getRange('D3').setFontWeight('bold');
  timeSheet.getRange('E3').setValue('');
  timeSheet.getRange('A5').setFormula('=IFERROR(FILTER(Master_Log!A1:N200, Master_Log!A2:A200 >= B3, Master_Log!A2:A200 <= E3), "Select start and end dates to filter execution timeline.")');
  timeSheet.setFrozenRows(5);
  formatHeaderRow(timeSheet, 14);
}

function setupViewTab(ss, sheetName, labelText, formulaText, entityType) {
  var sheet = ss.getSheetByName(sheetName);
  sheet.clear();
  sheet.getRange('A1').setValue(sheetName.toUpperCase() + ' — FILTER VIEW');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14).setFontColor('#0f172a');
  sheet.getRange('A3').setValue(labelText);
  sheet.getRange('A3').setFontWeight('bold');
  sheet.getRange('B3').setValue('');
  sheet.getRange('A5').setFormula(formulaText);
  sheet.setFrozenRows(5);
  formatHeaderRow(sheet, 14);
}

function formatHeaderRow(sheet, numCols) {
  var range = sheet.getRange(1, 1, 1, numCols);
  if (sheet.getName() !== 'Master_Log' && sheet.getName() !== 'Master Progress View' && sheet.getName() !== 'Timeline View' && !sheet.getName().includes('Dashboard')) {
    // For view tabs where headers are at row 5
    range = sheet.getRange(5, 1, 1, numCols);
  } else if (sheet.getName() === 'Dashboard & Summary') {
    return; // Handled individually
  }
  range.setBackground('#0f172a');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
}

function applyMasterLogValidations(sheet) {
  var categories = ['Deep Work', 'Admin', 'Meetings', 'Health & Fitness', 'Personal', 'Learning'];
  var timeBlocks = ['06:00 AM - 07:00 AM', '07:00 AM - 08:00 AM', '08:00 AM - 09:00 AM', '09:00 AM - 10:00 AM', '10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '12:00 PM - 01:00 PM', '01:00 PM - 02:00 PM', '02:00 PM - 03:00 PM', '03:00 PM - 04:00 PM', '04:00 PM - 05:00 PM', '05:00 PM - 06:00 PM', '06:00 PM - 07:00 PM', '07:00 PM - 08:00 PM', '08:00 PM - 09:00 PM', '09:00 PM - 10:00 PM'];
  var priorities = ['High', 'Medium', 'Low'];
  var statuses = ['Not Started', 'In Progress', 'Completed', 'Deferred'];
  var progresses = ['0', '25', '50', '75', '100'];
  var yesNo = ['Yes', 'No'];
  var timeSlots = ['09:00 AM', '10:00 AM', '11:00 AM', '12:00 PM', '01:00 PM', '02:00 PM', '03:00 PM', '04:00 PM', '05:00 PM', '06:00 PM'];

  sheet.getRange('C2:C100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(categories).build());
  sheet.getRange('D2:D100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(timeBlocks).build());
  sheet.getRange('E2:E100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(priorities).build());
  sheet.getRange('I2:I100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(statuses).build());
  sheet.getRange('J2:J100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(progresses).build());
  sheet.getRange('L2:L100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(yesNo).build());
  sheet.getRange('M2:M100').setDataValidation(SpreadsheetApp.newDataValidation().explicitValueInList(timeSlots).build());
}

function applyConditionalFormatting(sheet) {
  var range = sheet.getRange('I2:I100');
  var rules = sheet.getConditionalFormatRules();

  var completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#d1fae5')
    .setFontColor('#065f46')
    .setRanges([range])
    .build();

  var inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground('#fef3c7')
    .setFontColor('#92400e')
    .setRanges([range])
    .build();

  var notStartedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not Started')
    .setBackground('#fee2e2')
    .setFontColor('#991b1b')
    .setRanges([range])
    .build();

  var deferredRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Deferred')
    .setBackground('#f1f5f9')
    .setFontColor('#475569')
    .setRanges([range])
    .build();

  rules.push(completedRule);
  rules.push(inProgressRule);
  rules.push(notStartedRule);
  rules.push(deferredRule);

  sheet.setConditionalFormatRules(rules);
}

function scheduleFollowUpCalendarEvent() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a valid task row in Master_Log.');
    return;
  }

  var data = sheet.getRange(row, 1, 1, 14).getValues()[0];
  var taskName = data[1];
  var followUpReq = data[11];
  var followUpTime = data[12] || '10:00 AM';
  var dateVal = data[0];
  var stakeholder = data[7];
  var notes = data[13];

  if (followUpReq !== 'Yes') {
    SpreadsheetApp.getUi().alert('Follow-Up Required? is set to "No" for this task.');
    return;
  }

  var eventDate = dateVal ? new Date(dateVal) : new Date();
  var timeParts = followUpTime.split(':');
  var hour = parseInt(timeParts[0], 10);
  var minute = parseInt(timeParts[1], 10);
  if (followUpTime.includes('PM') && hour < 12) hour += 12;
  if (followUpTime.includes('AM') && hour === 12) hour = 0;

  eventDate.setHours(hour, minute, 0, 0);
  var endDate = new Date(eventDate.getTime() + 3600000); // 1 hour duration

  var calendar = CalendarApp.getDefaultCalendar();
  calendar.createEvent('Follow-Up: ' + taskName + ' (' + stakeholder + ')', eventDate, endDate, {
    description: 'Daily & Hourly Tracker Follow-Up\n\nTask: ' + taskName + '\nStakeholder: ' + stakeholder + '\nNotes: ' + notes,
    location: 'Virtual / Meeting'
  });

  SpreadsheetApp.getUi().alert('Success! Calendar event scheduled for ' + eventDate.toLocaleString());
}
