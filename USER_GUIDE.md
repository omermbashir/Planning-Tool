# User Guide — Capacity Planning Tool

## Quick Start

### Step 1: Generate the Excel template
```bash
python capacity_planner.py --template
```
This creates `capacity_data.xlsx` with three sheets pre-filled with example data, dropdown validations, and conditional formatting.

### Step 2: Edit the Excel file
Open `capacity_data.xlsx` in Excel and replace the example data with your real team, workstreams, and tasks. The dropdowns and formatting rules will guide you.

### Step 3: Generate charts
```bash
python capacity_planner.py
```
This reads your Excel data and produces three PNG files in the `output/` folder, plus an executive summary and schedule suggestions in the console.

---

## Excel Sheets — Column Reference

### Team Sheet

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Name** | Yes | Team member name (used for task assignment) | `Team Lead` |
| **Role** | No | Descriptive role (not used in calculations) | `Senior Analyst` |
| **Days Per Week** | Yes | Available working days per week. Supports decimals for part-time. | `5`, `3`, `2.5` |

### Workstreams Sheet

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Workstream** | Yes | Name of the workstream. Must match exactly what's used in the Tasks sheet. | `Platform Migration` |
| **Color** | Yes | Hex colour code for chart rendering. | `#2196F3` |
| **Priority** | Yes | P1 (highest) to P4 (lowest). Controls sort order and visual emphasis. | `P1` |

### Tasks Sheet

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Task** | Yes | Name of the task. | `Requirements Gathering` |
| **Workstream** | Yes | Must match a name from the Workstreams sheet exactly. | `Platform Migration` |
| **Assigned To** | Yes | Must match a name from the Team sheet exactly. | `Team Lead` |
| **Start Date** | Yes | When the task starts. Format: date (Excel auto-formats). | `2026-02-17` |
| **Original Days** | Yes | Initial estimate in working days. Set once, never change. Used for drift tracking. | `10` |
| **Total Days** | Yes | Current best estimate in working days. Update as scope changes. Supports decimals. | `12.5` |
| **Priority** | Yes | P1-P4. Usually matches the workstream's priority but can differ. | `P2` |
| **Status** | Yes | One of: `Planned`, `In Progress`, `Complete`, `On Hold` | `In Progress` |
| **Actual End** | No | Date the task actually finished. Only relevant for Complete tasks. | `2026-03-14` |
| **Blocked By** | No | Free text explaining what's blocking an On Hold task. | `Waiting on Data Eng` |
| **Notes** | No | Any additional context. | `Scope increased after review` |

---

## Priority System (P1-P4)

Both workstreams and individual tasks have priorities. **Workstream priority always takes precedence** over task priority for ordering on charts.

### How ordering works
1. Workstreams are sorted P1 first, P4 last
2. Within each workstream, tasks are sorted by task priority (P1 first), then by start date

### Visual emphasis
| Priority | Bar thickness | Opacity | Font weight | Intent |
|----------|--------------|---------|-------------|--------|
| **P1** | Thickest (2.2) | Full (1.0) | Bold | Critical — demands attention |
| **P2** | Medium (1.5) | Near-full (0.95) | Medium | Important — clearly visible |
| **P3** | Standard (1.0) | Reduced (0.80) | Normal | Routine — present but subdued |
| **P4** | Thin (0.8) | Faded (0.65) | Normal | Low — background items |

### Tip: task priority vs workstream priority
A P1 task inside a P3 workstream is unusual but allowed. The tool will warn you in the console during validation. The workstream still appears in P3 position, but the task within it will be styled as P1.

---

## Status Colours on the Gantt Chart

| Status | Appearance | Description |
|--------|-----------|-------------|
| **In Progress** | Vivid workstream colour, full priority alpha | Active work — stands out clearly |
| **Planned** | Workstream colour at 70% of priority alpha | Future work — visibly muted compared to In Progress |
| **On Hold** | Blue-grey fill, cross-hatch pattern, dashed border | Blocked — neutral "parked" look, instantly recognisable |
| **Complete** | Workstream colour at 35% alpha, grey edge | Done — faded into the background |

---

## Estimation Drift

Track how estimates change over time by maintaining both **Original Days** and **Total Days**.

### Setup
1. When you first create a task, set both **Original Days** and **Total Days** to the same value
2. As scope changes, update **Total Days** only — leave **Original Days** untouched

### What you'll see
- **Gantt chart**: Drifted tasks show a label like `(was 10d, now 15d +50%)`
- **Executive summary**: Total drift across all tasks, e.g. `Estimation drift: +5.5 days (+18%) across 3 tasks`

### Tip
If Original Days and Total Days are equal, no drift is shown. This is the normal case for tasks where the estimate hasn't changed.

---

## Planned vs Actual Tracking

For Complete tasks, you can record when they actually finished to compare planned vs actual.

### Setup
1. Set Status to `Complete`
2. Fill in the **Actual End** date column with the date the task was actually finished

### What you'll see
- **Gantt chart**: Two overlapping bars:
  - Ghost/dashed outline at the planned position
  - Solid bar at the actual position
  - Variance label: `+3d late` (red) or `-2d early` (green)

### Tip
If you leave Actual End blank for a Complete task, it just shows a single faded bar at the planned position — no comparison is drawn.

---

## On Hold / Blocked Tasks

### How to block a task
1. Change Status to `On Hold` in Excel
2. Optionally fill **Blocked By** with the reason (e.g. "Waiting on Data Eng", "Depends on Platform team")

### How to unblock
1. Change Status back to `In Progress` (or `Planned` if work hasn't started yet)
2. Adjust the Start Date if the block has pushed it back

### What you'll see
- **Gantt chart**: Blue-grey cross-hatched bar with dashed border. Blocked duration and reason shown in the label.
- **Roadmap**: Warning marker at the blocked task's position within the workstream swim lane.
- **Console**: `{task} has been on hold for N working days` with the blocker reason.

---

## Fractional Days

Both Original Days and Total Days support decimal values like `0.5`, `1.5`, `2.5`.

### How it works
- For 2.5 days starting Monday: Mon (1.0 day), Tue (1.0 day), Wed (0.5 day)
- Weekends are always skipped
- Capacity calculations sum fractional allocations correctly

### Tip
Useful for small tasks (half-day meetings, 1.5-day reviews) or when splitting time across multiple workstreams.

---

## Reading the Charts

### Gantt Chart + Weekly Capacity (`capacity_gantt.png`)

**Top panel — Gantt view:**
- Rows grouped by workstream (grey header bars with priority badges)
- Each task is a coloured bar spanning its working days
- Hatch patterns distinguish team members
- Status symbols: `>` In Progress, `-` Planned, checkmark Complete, `||` On Hold
- Today line (red dashed vertical line)

**Bottom panel — Weekly capacity:**
- Stacked bars showing days allocated per person per week
- Red dashed line = team capacity
- Percentage labels on each bar
- Over-capacity weeks highlighted in red

### Monthly Capacity Overview (`capacity_monthly.png`)
- Grouped bars: one group per month, one bar per team member
- Green dashed line = available capacity (adjusts for actual working days in each month)
- Percentage labels showing utilisation
- Over-capacity months highlighted in red

### Strategic Roadmap (`roadmap.png`)
- One swim lane per workstream, sorted by priority (P1 at top)
- Priority badges on y-axis labels
- Diamond markers at each task's start date
- Activity density shading (darker = more tasks active)
- Quarter boundary lines
- Warning markers for workstreams with blocked tasks

---

## Console Output

### Executive Summary
Printed every time you generate charts. Shows:
- Total tasks (active, complete, on hold)
- Timeline span in weeks
- Overall utilisation percentage
- Per-person utilisation (days allocated / days available)
- Over-capacity week count
- Priority breakdown (tasks and days per priority level)
- Estimation drift total (if any tasks have drifted)

### Schedule Suggestions
Printed after the summary if any suggestions are detected:
- **Early finishers**: recommends pulling forward subsequent tasks
- **Overdue warnings**: flags In Progress tasks past their planned end
- **Blocked duration**: shows how long On Hold tasks have been blocked
- **Spare capacity**: highlights future weeks where a person has 3+ free days (capped at 5 entries)

---

## Command-Line Options

```bash
# Generate the Excel template (overwrites existing capacity_data.xlsx)
python capacity_planner.py --template

# Generate all charts (default)
python capacity_planner.py

# Generate specific charts only
python capacity_planner.py --charts gantt
python capacity_planner.py --charts monthly
python capacity_planner.py --charts roadmap
python capacity_planner.py --charts gantt roadmap

# Custom input file
python capacity_planner.py --input path/to/my_data.xlsx

# Custom output directory
python capacity_planner.py --output path/to/output_dir
```

---

## Tips

- **Re-run after every Excel update** — the tool reads fresh data each time
- **Keep Original Days constant** — only update Total Days when scope changes
- **Use P1 sparingly** — if everything is P1, nothing stands out
- **Check console output** — the schedule suggestions catch things the charts don't show (overdue tasks, spare capacity)
- **Fractional days for small tasks** — 0.5 for a half-day review, 1.5 for a day-and-a-half workshop
- **Workstream names must match exactly** — if a task references a workstream that doesn't exist on the Workstreams sheet, validation will flag it with a fuzzy match suggestion
