# Daily & Hourly Tracker — Productivity Workbook

A comprehensive, professional Single Source of Truth (SSOT) productivity tracking system designed for Google Sheets with Google Apps Script automation and Google Calendar integration.
# Installation & Setup

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Click on **Extensions > Apps Script**.
3. Paste the contents of `Code.gs` into the script editor.
 A remove all data in apps script first
 B copy and paste everything from code.gs
 C click on save (he icon that looks like a floppy disk)
 D click on onOpen and choose setupMasterLogTracker
 E Click on run (The play button next to the floppy disk button)
 F click on review permission on the popup
 F on the page that says "Google hasn’t verified this app" click on advanced then click on Go to Untitled project (unsafe) - because this is an untitle project we just created it has not been verified so this warning is normal and expected.
 G Select what Untitled project can access - click on select all and then continue. - script will run and create the google sheet. 
When you see execution completed go to step 4.
4. Deploy as a Web App (Execute as: **Me**, Who has access: **Anyone**).
 A Click on deploy  > New deployment - then a window will open Click on the button next to "select type" and then click on web app.
 B Description : Productivity tracker , (Execute as: **Me**, Who has access: **Anyone**). Click on Deploy - click on authorise access - again select all and then continue. Copy web app url also the Deployment id.
Then open tracker-app.html - click on connect google sheets 
5. Copy the Web App URL and paste it into the web dashboard ("Connect Google Sheets") to automatically provision and sync your personal Google Sheet!

After this close google sheets and use the tracker-app.html file to update details that you need. 
We made google sheets as your database only. Nothing needs to be done in google sheets.


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

#