# Capacity Planning Tool

A Python tool for visualising team workload and capacity. Reads work requests from an Excel file, calculates weekly and monthly capacity, and outputs professional presentation-ready charts as PNGs.

## What It Generates

### 1. Gantt Chart + Weekly Capacity (`capacity_gantt.png`)
- **Top panel:** Tasks grouped by workstream, sorted by priority (P1 first). Colour-coded rounded bars with hatch patterns per team member. Priority badges on workstream headers. Status symbols, estimation drift labels, planned-vs-actual ghost bars for completed tasks.
- **Bottom panel:** Weekly stacked bar chart with utilisation percentages. Over-capacity weeks highlighted in red.

### 2. Monthly Capacity Overview (`capacity_monthly.png`)
- Side-by-side bars per team member showing days allocated per month.
- Available capacity line (adjusts for actual working days per month).
- Utilisation percentage labels. Over-capacity months highlighted.

### 3. Strategic Roadmap (`roadmap.png`)
- Swim-lane view with one bar per workstream, sorted by priority (P1 at top).
- Priority badges on y-axis labels. Activity density shading.
- Diamond markers at each task start. Blocked task warning markers. Quarter boundary lines.
- Designed for exec-level conversations.

### Console Output
- **Executive summary**: task counts, utilisation per person, over-capacity weeks, priority breakdown, estimation drift totals
- **Schedule suggestions**: early finishers, overdue tasks, blocked duration, spare capacity gaps

```
============================================================
  EXECUTIVE SUMMARY
============================================================
  Tasks:         10 total (10 active, 0 complete, 0 on hold)
  Timeline:      14 weeks
  Utilisation:   73% overall
    Team Lead: 64.5 / 70 days (92%)
    Analyst: 38.0 / 70 days (54%)
  Over-capacity: 2 of 14 weeks

  By priority:
    P1: 4 tasks (53 days)
    P2: 3 tasks (19.5 days)
    P3: 2 tasks (18 days)
    P4: 1 task (12 days)

  Estimation drift: +2.5 days (+21%) across 2 tasks - scope increase
============================================================

SCHEDULE SUGGESTIONS:
  Analyst has spare capacity in w/c 16 Feb (0.0 days allocated, 5.0 days free)
  ...
```

## Setup

### Requirements
- Python 3.10+
- matplotlib, openpyxl, pandas

### Install Dependencies
```bash
pip install matplotlib openpyxl pandas
```

## Usage

### 1. Generate the Excel Template
```bash
python capacity_planner.py --template
```
Creates `capacity_data.xlsx` with three sheets and example data:
- **Team** — team members and available days per week
- **Workstreams** — workstream names, display colours, and priorities (P1-P4)
- **Tasks** — tasks with workstream, assignee, start date, original/current estimates, priority, status, actual end, blocked by, notes

The template includes:
- **Dropdown validations** for Status, Priority, Workstream, and Assigned To
- **Conditional formatting** for status colours, priority emphasis, and scope drift highlighting
- **Frozen headers** on all sheets

### 2. Edit the Excel File
Open `capacity_data.xlsx` and update with your actual tasks and team info.

**Status values:** `Planned`, `In Progress`, `Complete`, `On Hold`

**Priority values:** `P1`, `P2`, `P3`, `P4`

**Total Days:** working days the task takes (supports fractions like 2.5). Distributed across weekdays automatically, skipping weekends.

**Original Days:** the initial estimate — set once, used to calculate estimation drift.

**Actual End:** (optional) date a Complete task actually finished. Shows planned-vs-actual comparison on the Gantt.

**Blocked By:** (optional) free text for On Hold tasks explaining the blocker.

### 3. Generate Charts
```bash
# Generate all charts (default)
python capacity_planner.py

# Generate specific charts only
python capacity_planner.py --charts gantt
python capacity_planner.py --charts monthly roadmap

# Custom input/output paths
python capacity_planner.py --input path/to/file.xlsx --output path/to/chart.png
```

## Features

### Priority System (P1-P4)
- Both workstreams and tasks have priorities
- Gantt and roadmap sort by workstream priority, then task priority within each group
- Visual weight varies: P1 = bold/thick/full opacity, P4 = thin/faded
- Priority badges on workstream headers

### Estimation Drift Tracking
- Original Days vs Total Days shows scope changes
- Drift labels on Gantt bars: "(was 10d, now 15d +50%)"
- Executive summary totals across all drifted tasks

### Planned vs Actual
- Complete tasks with an Actual End date show two overlapping bars on the Gantt:
  - Ghost/dashed bar at planned position
  - Solid bar at actual position
  - Variance label: "+3d late" (red) or "-2d early" (green)

### Blocked/On Hold Tracking
- On Hold tasks shown in neutral blue-grey with cross-hatch pattern and dashed border — visually distinct from active work
- Planned tasks rendered at reduced opacity vs In Progress for clear status hierarchy
- Blocked duration calculated and displayed
- Blocked By reason shown if provided
- Warning markers on roadmap for workstreams with blocked tasks

### Schedule Suggestions
- Early finishers: recommends pulling forward subsequent tasks
- Overdue warnings for In Progress tasks past their planned end
- Blocked duration tracking
- Spare capacity alerts for future weeks

### Fractional Days
- Total Days and Original Days support decimals (0.5, 1.5, 2.5, etc.)
- Capacity calculations sum fractional allocations correctly

## Data Validation

The tool validates your Excel data before generating charts:
- Checks for missing task names, invalid dates, non-positive durations
- Detects workstream name mismatches with fuzzy suggestions
- Validates assigned team members against the Team sheet
- Warns on unrecognised status or priority values
- Warns if task priority is higher than its workstream priority
- Validates fractional days are positive

## File Structure
```
capacity_planner.py    # Single script - all logic (~1600 lines)
capacity_data.xlsx     # Excel input (generated via --template, then user-maintained)
output/
  capacity_gantt.png   # Gantt chart + weekly capacity
  capacity_monthly.png # Monthly capacity utilisation
  roadmap.png          # Strategic roadmap swim lanes
```
