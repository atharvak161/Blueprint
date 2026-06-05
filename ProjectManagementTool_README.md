# Project Management Tool — User Guide
## Version 1.0 | Excel Macro-Enabled Workbook

---

## Overview

This workbook is a comprehensive project management tool built for Excel.
It provides:

- **12 collapseable project phases** with Tasks and Subtasks
- **Auto-updating Task IDs** via VBA macro
- **UK Holiday-aware scheduling** (2025, 2026, 2027)
- **Resource management** with dropdown assignment
- **Status tracking** with colour-coded RAG indicators
- **Three one-click HTML report exports**: Gantt Chart, Project Dashboard, RAG Status Report

---

## Files Included

| File | Purpose |
|------|---------|
| `ProjectManagementTool.xlsx` | Main workbook — open this in Excel |
| `ProjectManagementTool_VBA.bas` | VBA macro code — import once into Excel |
| `README.md` | This guide |

---

## Setup Instructions (One-Time)

### Step 1 — Open the Workbook
Open `ProjectManagementTool.xlsx` in Microsoft Excel (2016 or later recommended).

### Step 2 — Save as Macro-Enabled (.xlsm)
The workbook must be saved as `.xlsm` to support VBA macros:
1. Go to **File → Save As**
2. In the "Save as type" dropdown, select **Excel Macro-Enabled Workbook (*.xlsm)**
3. Click **Save**

### Step 3 — Import the VBA Module
1. Press **Alt + F11** to open the Visual Basic Editor
2. In the menu, go to **File → Import File...**
3. Browse to and select `ProjectManagementTool_VBA.bas`
4. Click **Open** — the module will appear in the Project pane
5. Close the VBA Editor (Alt + F4 or close the window)

### Step 4 — Enable Macros
When prompted by Excel, click **Enable Content** or **Enable Macros**.
The tool will not function without macros enabled.

### Step 5 — Run Initial Setup
1. Go to **Developer → Macros** (or press Alt + F8)
2. Select **SetupProjectUI** and click **Run**
3. This applies conditional formatting and validates the sheet structure

---

## Sheet Reference

### Project Plan Sheet
The main working sheet. Contains:

| Column | Name | Description |
|--------|------|-------------|
| A | Task ID | Auto-assigned sequential number. Never edit manually. |
| B | Task Type | Dropdown: **Task**, **Subtask**, or **Phase** (section header) |
| C | Phase | The phase this task belongs to (auto-populated from section header) |
| D | Description | Task or subtask name/description |
| E | Planned Start | Start date in DD/MM/YYYY format |
| F | Planned End | End date in DD/MM/YYYY format |
| G | Resource | Assigned team member — dropdown from Resources sheet |
| H | % Done | Progress as decimal: 0 = 0%, 0.5 = 50%, 1 = 100% |
| I | Status | Dropdown: Not Started / In Progress / Completed / Overdue / Blocked |
| J | Notes / Blockers | Free text for notes, issues, dependencies |

**Project Summary (rows 1–9):**
- Rows 1–7: Auto-calculated KPIs (total tasks, completed, in progress, overdue, blocked, % complete)
- Row 9: Export buttons — click to generate HTML reports

### Resources Sheet
Add team members here BEFORE assigning them to tasks.

| Column | Description |
|--------|-------------|
| Name | Full name — must match exactly when assigning in Project Plan |
| Department | Team or department name |
| Role | Job title or role |
| Email / Contact | Contact information |

### Holidays Sheet
Pre-populated with UK public holidays for 2025, 2026, and 2027.

| Column | Description |
|--------|-------------|
| Holiday Name | Name of the public holiday or leave |
| Start Date | First date of the holiday (DD/MM/YYYY) |
| End Date | Last date of the holiday (same as start for single-day events) |
| Resource | Leave blank for public holidays (affects everyone). Enter a name for individual leave. |

> **Adding Individual Leave:** Enter the resource's name in column D.
> The scheduling macro will skip that date only for that person.

### Stakeholders Sheet
Used for email distribution when exporting reports.

| Column | Description |
|--------|-------------|
| Name | Stakeholder display name |
| Email | Full email address |
| Type | To, CC, or BCC |
| Organisation | Company or team |

---

## Working with Phases (Collapseable Sections)

Each **Phase** row is a section header. The tasks beneath it are grouped and can be collapsed:

- Click the **minus (–)** button on the left row numbers to collapse a phase
- Click the **plus (+)** button to expand it
- Use the **1** and **2** buttons at the top-left to collapse/expand all groups

**To add a new Phase:**
1. Insert a row at the top of where you want the new phase
2. Set **Task Type = Phase** in column B
3. Enter the phase name in column D
4. The VBA macro will automatically group the rows beneath it

---

## Task ID Management

Task IDs are automatically managed by the **RenumberIDs** VBA macro.

**Inserting a task:**
1. Right-click the row where you want to insert above/below
2. Select **Insert** from the context menu
3. Run **RenumberIDs** macro (Alt + F8 → RenumberIDs → Run)
4. All IDs will renumber sequentially

**Deleting a task:**
1. Select the row and press the Delete key (or right-click → Delete)
2. Run **RenumberIDs** to close the gap

> **Never manually edit column A** — it is managed entirely by the VBA macro.

---

## Status & Colour Coding

| Status | Colour | Meaning |
|--------|--------|---------|
| Not Started | Grey | Task has not begun |
| In Progress | Amber/Gold | Work is actively underway |
| Completed | Green | Task is finished |
| Overdue | Red | Past end date and not completed |
| Blocked | Purple | Cannot proceed — dependency or issue |

The % Done column should be entered as a decimal:
- `0` = 0%
- `0.25` = 25%
- `0.5` = 50%
- `0.75` = 75%
- `1` = 100%

---

## Holiday-Aware Date Scheduling

The workbook includes UK public holidays for 2025–2027 in the Holidays sheet.

**Auto-adjust dates:**
1. Press **Alt + F8**
2. Select **AdjustDatesForHolidays**
3. Click **Run**

This scans every task's start and end date and moves any that fall on a weekend or
UK public holiday to the next available working day.

**Adding custom holidays or leave:**
- Add a row to the Holidays sheet with the Name, Start Date, End Date
- Leave the Resource column blank for company-wide holidays
- Enter a resource name for individual leave (e.g. "David Patel")

---

## Exporting HTML Reports

The three export buttons on row 9 of the Project Plan sheet generate self-contained
HTML files. These open automatically in your default browser and can be shared
with stakeholders who do not have Excel.

### How to Assign Macros to Buttons

The export button labels are styled cells in row 9. To make them clickable:

1. Go to **Insert → Shapes → Rounded Rectangle**
2. Draw a shape over the button label area
3. Right-click the shape → **Format Shape**
   - Fill: Solid colour `#1F4E79` (navy)
   - Line: `#2E75B6` (blue)
4. Right-click → **Edit Text**, type the button label
5. Set font to Segoe UI, Bold, 11pt, White
6. Right-click → **Assign Macro**
7. Select the macro from the list and click OK

| Button | Macro to Assign |
|--------|----------------|
| 📊 EXPORT GANTT CHART | `ExportGanttChart` |
| 📋 EXPORT DASHBOARD | `ExportDashboard` |
| 🚦 EXPORT RAG REPORT | `ExportRAGReport` |

Alternatively, run any macro directly via **Alt + F8 → Select Macro → Run**.

---

## Macro Reference

| Macro | Purpose |
|-------|---------|
| `SetupProjectUI` | Applies conditional formatting and dropdowns. Run once after setup. |
| `RenumberIDs` | Renumbers all Task IDs sequentially. Run after inserting/deleting rows. |
| `AdjustDatesForHolidays` | Moves dates falling on weekends/holidays to next working day. |
| `ExportGanttChart` | Generates `ProjectGanttChart.html` — interactive Gantt with search & filters |
| `ExportDashboard` | Generates `ProjectDashboard.html` — KPI cards, charts, resource workload |
| `ExportRAGReport` | Generates `ProjectRAGReport.html` — phase-by-phase RAG status, printable |

---

## HTML Report Features

### Gantt Chart (`ProjectGanttChart.html`)
- Interactive timeline with coloured task bars
- Search by task name
- Filter by Status or Resource
- % Complete shown as fill within each bar
- Hover tooltip with full task details
- Today marker (red vertical line)
- Phase group headers always visible

### Project Dashboard (`ProjectDashboard.html`)
- KPI summary cards (total, completed, in progress, overdue, blocked, % complete)
- Status breakdown donut chart (Chart.js)
- Resource workload bar chart (top 10 resources)
- Phase completion bar chart
- Phase-by-phase progress bars
- Full task table with status badges and progress bars

### RAG Status Report (`ProjectRAGReport.html`)
- Overall project RAG (Red / Amber / Green) prominently displayed
- RAG count summary banner
- Phase-by-phase breakdown with collapseable sections
- Each task shows RAG rating, progress bar, and notes/blockers
- Print-friendly layout (File → Print or Ctrl+P)
- "Print / Save PDF" button in bottom-right corner

---

## Tips & Best Practices

1. **Always set Task Type first** when adding a new row — set Phase/Task/Subtask before filling other columns
2. **Add resources to the Resources sheet** before assigning them to tasks
3. **Save before exporting** — reports save to the same folder as the workbook
4. **Run AdjustDatesForHolidays** after importing this workbook or changing holidays
5. **Run RenumberIDs** after inserting or deleting task rows
6. **Do not edit column A (ID)** — it is managed by VBA
7. **Enter % Done as decimals** (0 to 1), not percentages (0% to 100%)
8. **Use Notes/Blockers column** to log issues, dependencies, and risks

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Macros not working | Re-enable macros via File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros |
| Resource dropdown empty | Add resources to the Resources sheet first |
| HTML report not opening | Check your browser is set as default; or open the .html file manually from the save folder |
| IDs are out of sequence | Run the `RenumberIDs` macro (Alt + F8 → RenumberIDs → Run) |
| Dates fall on weekends/holidays | Run `AdjustDatesForHolidays` macro |
| Report saves to wrong location | Save the workbook first — reports save to the same directory as the .xlsx/.xlsm file |
| Collapsed sections won't expand | Click the + button on the left row margin, or use the 1/2 outline level buttons in the top-left |

---

## Support & Customisation

**To add more phases:** Insert rows, set Task Type = Phase for the header row, then group the task rows beneath (Data → Group → Rows).

**To add more resources:** Add rows to the Resources sheet. The Resource dropdown in Project Plan will update automatically.

**To add holidays:** Add rows to the Holidays sheet. Run AdjustDatesForHolidays to re-check all dates.

**To customise report colours:** Edit the VBA code in the `BuildGanttHTML`, `BuildDashboardHTML`, and `BuildRAGHTML` functions. Colours are CSS hex values embedded in the HTML strings.

---

*Project Management Tool — Version 1.0*
*Designed for Microsoft Excel 2016 and later*
