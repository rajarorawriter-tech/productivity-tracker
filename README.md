# Daily & Hourly Tracker — Productivity Workbook

A comprehensive, professional Single Source of Truth (SSOT) productivity tracking system designed for Google Sheets with Google Apps Script automation and Google Calendar integration.

## ✨ Professional Design & Usability Features

1. **Automatic Database Provisioning**: When connecting a new or empty personal Google Sheet via the Apps Script Web App URL, the script automatically provisions and formats all 7 required tabs (`Master_Log`, `Dashboard & Summary`, `Date-Wise View`, `Task-Wise View`, `Stakeholder-Wise View`, `Master Progress View`, `Timeline View`), injects formulas, and applies data validation rules and conditional formatting.
2. **Executive Styling & Palette**: Styled with a professional deep slate navy (`#0f172a`) header theme, clean typography, and formatted columns (dates, decimals, percentages) for crystal-clear readability.
3. **Automated Conditional Formatting**: 
   - 🟢 **Completed**: Soft green fill with dark emerald text (`#d1fae5` / `#065f46`).
   - 🟡 **In Progress**: Soft amber fill with warm brown text (`#fef3c7` / `#92400e`).
   - 🔴 **Not Started**: Soft red fill with dark crimson text (`#fee2e2` / `#991b1b`).
   - ⚪ **Deferred**: Neutral grey tone (`#f1f5f9` / `#475569`).
4. **Master_Log (Backend SSOT)**: Centralized database capturing Date, Task Name, Category, Time Block, Priority, Planned vs. Actual Hours, Stakeholder, Status, Progress %, Next Update Expected, Follow-Up details, and Daily Notes.
5. **Dashboard & Summary**: Executive analytics with global date range filtering, KPI summary row (`Total Tasks`, `Completed Tasks`, `Completion Rate`, `Deep Work Hours`, `Total Logged Hours`), and category breakdown performance tables.
6. **Dedicated Views**:
   - `Date-Wise View`: Daily review and standup snapshot.
   - `Task-Wise View`: Rolling progress and time investment history per task.
   - `Stakeholder-Wise View`: Deliverables and follow-ups tied to specific clients or team members.
   - `Master Progress View`: Panoramic audit grid with full sorting and filtering.
   - `Timeline View`: Execution Gantt and date-range filtering.
7. **Google Apps Script Automations (`Code.gs`)**:
   - Automated workbook generator (`setupMasterLogTracker()`).
   - Google Calendar integration (`scheduleFollowUpCalendarEvent()`).
   - Web App API endpoints (`doGet` / `doPost`) for bi-directional live sync with web dashboards.

## Installation & Setup

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Click on **Extensions > Apps Script**.
3. Paste the contents of `Code.gs` into the script editor.
4. Deploy as a Web App (Execute as: **Me**, Who has access: **Anyone**).
5. Copy the Web App URL and paste it into the web dashboard ("Connect Google Sheets") to automatically provision and sync your personal Google Sheet!
