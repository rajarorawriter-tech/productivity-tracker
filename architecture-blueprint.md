# Daily & Hourly Tracker — Comprehensive Architecture & Implementation Blueprint

This architectural blueprint outlines the complete design, data flow, schema, formula logic, and interactive features of the **Daily & Hourly Tracker** Google Sheets & Google Apps Script productivity workbook.

---

## 1. Master Architecture Overview

The workbook is structured around a **Single Source of Truth (SSOT)** backend (`Master_Log`) paired with dedicated executive analytics, dynamic filter views, and calendar automation. 

```
┌────────────────────────────────────────────────────────────────────────┐
│                        MASTER_LOG (Backend SSOT)                       │
│  Captures all raw entries: Date, Task, Category, Time Block, Priority, │
│  Planned/Actual Hours, Stakeholder, Status, Progress %, Follow-Ups     │
└──────────────────┬──────────────────┬──────────────────┬───────────────┘
                   │                  │                  │
        ┌──────────▼────────┐ ┌───────▼────────┐ ┌──────▼────────┐
        │ Dashboard &       │ │ Date-Wise View │ │ Task-Wise     │
        │ Summary           │ │ (Daily Review) │ │ View (History)│
        │ (KPIs + Charts +  │ └────────────────┘ └───────────────┘
        │ Date Filtering)   │ ┌────────────────┐ ┌───────────────┐
        └───────────────────│ │Stakeholder View│ │ Timeline View │
                            │ └────────────────┘ │ (Gantt-style) │
                            └────────────────────┴───────────────┘
```

---

## 2. Workbook Tab Breakdown

### Tab 1: `Master_Log` (Backend Database)
The core operational data store where all productivity logs are recorded. It features rigorous data validation and automated formulas to minimize manual data entry.

| Column | Header Name | Data Type / Validation | Description / Formula |
| :--- | :--- | :--- | :--- |
| **A** | `Date` | Date Dropdown / Picker | Native calendar date picker for task scheduling/execution. |
| **B** | `Task Name` | Free Text / Dropdown Source | Unique title of the task or project activity. |
| **C** | `Category` | Dropdown (`requireValueInList`) | `Deep Work`, `Admin`, `Meetings`, `Health & Fitness`, `Personal`, `Learning`. |
| **D** | `Time Block` | Dropdown (`requireValueInList`) | 1-hour intervals (`06:00 AM - 07:00 AM` to `09:00 PM - 10:00 PM`). |
| **E** | `Priority` | Dropdown (`requireValueInList`) | `High`, `Medium`, `Low`. |
| **F** | `Planned (Hrs)` | Number (Decimal) | Estimated hours allocated for the task. |
| **G** | `Actual (Hrs - Auto)`| **Auto Formula** | `=IF(ISBLANK(D...), 0, 1)` (Automatically counts 1 hour per active time block). |
| **H** | `Stakeholder` | Free Text / Entity Name | Client, team member, or department associated with the task. |
| **I** | `Status` | Dropdown (`requireValueInList`) | `Not Started`, `In Progress`, `Completed`, `Deferred`. |
| **J** | `Progress (%)` | Dropdown (`requireValueInList`) | `0`, `25`, `50`, `75`, `100`. |
| **K** | `Next Update Expected` | Date Picker | Scheduled date for the next status check-in. |
| **L** | `Follow-Up Required?` | Dropdown (`Yes`, `No`) | Triggers follow-up and calendar integration workflow. |
| **M** | `Follow-Up Date & Time`| Dropdown (`requireValueInList`) | Predefined time slots (`09:00 AM` to `06:00 PM`) for follow-up scheduling. |
| **N** | `Daily Update / Notes` | Free Text (Wrapped) | Detailed notes, blockers, or milestone accomplishments. |

---

### Tab 2: `Dashboard & Summary` (Executive Analytics)
Designed for high-level oversight and real-time performance tracking.
- **Global Date Range Filter (`B3` & `B4`)**: Users can specify an optional **Start Date** (`B3`) and **End Date** (`B4`). When left blank, metrics compute all-encompassing data. When populated, all dashboard cards and tables instantly filter to the specified date range.
- **KPI Summary Row**:
  - `TOTAL TASKS`: Count of all logged tasks.
  - `COMPLETED TASKS`: Count of successfully finished tasks.
  - `COMPLETION RATE`: Percentage of completed tasks against total logged.
  - `DEEP WORK HOURS`: Total actual hours spent on completed `Deep Work`.
  - `TOTAL LOGGED HOURS`: Sum of all actual hours logged.
- **Category Breakdown Table**: Aggregates Task Count, Planned Hours, Actual Hours, Variance, and Share of Total across all categories.
- **Priority & Execution Status Tables**: Tracks task distribution by priority level (`High`, `Medium`, `Low`) and workflow status (`Not Started`, `In Progress`, `Completed`, `Deferred`).
- **Embedded Visual Charts**:
  1. *Actual Time Allocation by Category* (Donut/Pie Chart).
  2. *Time Allocation Task-Wise* (Donut/Pie Chart).
  3. *Planned vs. Actual Hours Task-Wise* (Grouped Column Chart).

---

### Tab 3: `Date-Wise View` (Daily Review)
- **Cell B2**: Date Picker dropdown linked to dates present in `Master_Log`.
- **Formula**: `=IFERROR(FILTER('Master_Log'!A2:N201, 'Master_Log'!A2:A201 = B2), "")`
- **Purpose**: Instantly pulls all tasks scheduled or executed on a selected single day for daily standups or evening reviews.

---

### Tab 4: `Task-Wise View` (Lifecycle History)
- **Cell B2**: Task Name dropdown populated dynamically from `Master_Log`.
- **Formula**: `=IFERROR(FILTER('Master_Log'!A2:N201, 'Master_Log'!B2:B201 = B2), "")`
- **Purpose**: Tracks a single task's rolling progress, status updates, and hour investments across multiple days.

---

### Tab 5: `Stakeholder-Wise View` (Collaboration & Accountability)
- **Cell B2**: Stakeholder dropdown (e.g., *Client A*, *Internal Team*).
- **Formula**: `=IFERROR(FILTER('Master_Log'!A2:N201, 'Master_Log'!H2:H201 = B2), "")`
- **Purpose**: Isolates all deliverables, commitments, and follow-ups tied to a specific stakeholder.

---

### Tab 6: `Master Progress View` (Panoramic Audit Grid)
- **Formula**: `=ARRAYFORMULA('Master_Log'!A2:N201)`
- **Purpose**: A full chronological snapshot of all entries across the entire project. Fully compatible with native Google Sheets column sorting and filter views.

---

### Tab 7: `Timeline View` (Execution Gantt & Date Filtering)
- **Cells B2 & B3**: Start Date and End Date filters.
- **Formula**: `=IFERROR(FILTER('Master_Log'!A2:N201, 'Master_Log'!A2:A201 >= B2, 'Master_Log'!A2:A201 <= B3), "No tasks found in selected date range.")`
- **Embedded Chart**: Column chart visualizing task execution progress percentages across the selected date range.

---

## 3. Google Apps Script Automations (`Code.gs`)

1. **Workbook Generator (`setupMasterLogTracker()`)**:
   - Creates and initializes all 7 sheets, configures header styling, column widths, frozen rows, and number formats.
   - Injects dynamic array formulas and builds data validation dropdowns in bulk.
2. **Conditional Formatting**:
   - Automatically applies soft color coding for task statuses (`Not Started` = Red, `In Progress` = Orange, `Completed` = Green, `Deferred` = Amber).
   - Applies conditional color scales for progress percentages (`≥75%` Green, `50-74%` Orange, `<50%` Light Gray).
3. **Google Calendar Integration (`scheduleFollowUpCalendarEvent()`)**:
   - Custom spreadsheet menu (`📅 Calendar Follow-Up`).
   - Reads the active task row, checks if `Follow-Up Required?` is set to `Yes`, and automatically schedules a 1-hour appointment in your default Google Calendar using the selected time slot in Column M.
