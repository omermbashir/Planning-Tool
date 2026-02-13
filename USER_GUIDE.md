# User Guide — Capacity Planning Tool

## Quick Start

### Step 1: Generate the Excel template
```bash
python capacity_planner.py --template
```
This creates `capacity_data.xlsx` with five sheets pre-filled with example data, dropdown validations, and conditional formatting.

### Step 2: Edit the Excel file
Open `capacity_data.xlsx` in Excel and replace the example data with your real team, workstreams, and tasks. The dropdowns and formatting rules will guide you.

**Only 6 mandatory fields per task** — the rest is optional or auto-filled. See the Tasks sheet reference below.

### Step 3: Generate charts
```bash
python capacity_planner.py
```
This reads your Excel data and produces four PNG files in the `output/` folder, plus an executive summary and schedule suggestions in the console.

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
| **Color** | Yes | Hex colour code for chart rendering. Must be a valid 6-digit hex code. | `#2196F3` |
| **Priority** | Yes | P1 (highest) to P4 (lowest). Controls sort order and visual emphasis. | `P1` |

### Tasks Sheet

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Task** | Yes | Name of the task. | `Requirements Gathering` |
| **Workstream** | Yes | Must match a name from the Workstreams sheet exactly. | `Platform Migration` |
| **Assigned To** | Yes | Must match a name from the Team sheet exactly. | `Team Lead` |
| **Start Date** | Yes | When the task starts. Format: date (Excel auto-formats). | `2026-02-17` |
| **Original Days** | No | Initial estimate in working days. **Auto-fills from Total Days if left blank.** Set once, never change. Used for drift tracking. | `10` |
| **Total Days** | Yes | Current best estimate in working days. Update as scope changes. Supports decimals. | `12.5` |
| **Priority** | No | P1-P4. **Inherits from workstream if left blank.** Only fill when a task differs from its workstream. | `P2` |
| **Status** | Yes | One of: `Planned`, `In Progress`, `Complete`, `On Hold`. Defaults to `Planned` if blank. | `In Progress` |
| **Actual End** | No | Date the task actually finished. Only relevant for Complete tasks. | `2026-03-14` |
| **Blocked By** | No | Free text explaining what's blocking an On Hold task. | `Waiting on Data Eng` |
| **Deadline** | No | Hard delivery date. Shows red diamond on Gantt, console warning if task overshoots. | `2026-03-31` |
| **Confidence** | No | Estimate quality: `High`, `Medium`, or `Low`. Shows coloured dot on Gantt. | `Medium` |
| **Notes** | No | Any additional context. | `Scope increased after review` |

**Smart defaults**: Only 6 fields are mandatory when adding a new task (Task, Workstream, Assigned To, Start Date, Total Days, Status). Priority inherits from the workstream, Original Days copies from Total Days, and everything else is optional.

### Public Holidays Sheet (optional)

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Date** | Yes | The public holiday date. Format: date. | `2026-04-03` |
| **Name** | No | Description of the holiday. | `Good Friday` |

Public holidays affect **all** team members — dates are treated like weekends (skipped in scheduling).

### Leave Sheet (optional)

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Person** | Yes | Must match a name from the Team sheet exactly. | `Team Lead` |
| **Start Date** | Yes | First day of leave. Format: date. | `2026-04-06` |
| **End Date** | Yes | Last day of leave. Format: date. | `2026-04-10` |
| **Type** | Yes | One of: `Annual Leave`, `Sick`, `Training`, `Conference`, `Other` | `Annual Leave` |
| **Notes** | No | Any additional context. | `Easter week` |

Leave affects only the named person — their tasks automatically extend around leave days.

Both the Public Holidays and Leave sheets are **optional**. If they don't exist in your Excel file, the tool works identically to before (no holidays, no leave adjustments). This means existing Excel files continue to work without changes.

---

## Smart Defaults

The tool is designed for minimal Excel maintenance. When adding a new task, you only need to fill 6 fields:

| Field | What to enter |
|-------|---------------|
| **Task** | Name of the task |
| **Workstream** | Pick from the dropdown |
| **Assigned To** | Pick from the dropdown |
| **Start Date** | When work begins |
| **Total Days** | How many working days it will take |
| **Status** | Pick from the dropdown (or leave blank for "Planned") |

Everything else fills itself or is optional:
- **Priority** → inherits from the workstream. Only override when a task is unusually urgent or low-priority compared to its workstream.
- **Original Days** → copies from Total Days. Only matters later if you update Total Days (drift tracking activates automatically).
- **Deadline, Confidence, Actual End, Blocked By, Notes** → leave blank until they're relevant.

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

## Deadline & Confidence (Communication Tools)

Both columns are optional. They exist to help you communicate risk and expectations to management. Leave them blank for most tasks — fill only the 2-3 where it matters.

### Deadline

A hard delivery date. Use it when your boss or stakeholder has said "this must be done by X date."

**Setup**: Fill the **Deadline** column with a date in the Tasks sheet.

**What you'll see**:
- **Gantt chart**: Red diamond marker (◆) at the deadline position on the task row. If the task bar extends past it, the overshoot gets a red tint.
- **Console**: `WARNING: 'Platform Migration' ends 3 days after deadline (deadline: 15 Mar)`
- **Executive summary**: Count of tasks at risk of missing their deadline.

**Tip**: If you don't fill Deadline, nothing changes — no markers, no warnings.

### Confidence

An estimate quality rating. Use it for tasks where the estimate might be wrong — new technology, external dependencies, vague requirements.

**Setup**: Fill the **Confidence** column with `High`, `Medium`, or `Low`.

**What you'll see**:
- **Gantt chart**: Small coloured dot next to the task label — green (High), amber (Medium), red (Low)
- **Console**: Low-confidence tasks are listed in the executive summary so you're aware of where surprises might come from.

**Tip**: High confidence = "done this before, estimate is solid". Low confidence = "first time, could easily double". This sets expectations early — when a low-confidence task overruns, it's not a surprise.

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
1. When you first create a task, just fill **Total Days** — Original Days auto-fills to the same value
2. As scope changes, update **Total Days** only — Original Days stays untouched automatically

### What you'll see
- **Gantt chart**: Drifted tasks show a label like `(was 10d, now 15d +50%)` — coloured amber for scope increase, green for scope decrease
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
- **Gantt chart**: Blue-grey cross-hatched bar with dashed border. Blocked duration and reason shown in the label. No artificial cap — a task blocked for 15 days on a 5-day estimate shows the full 15 days.
- **Roadmap**: Warning marker at the blocked task's position within the workstream swim lane.
- **Console**: `{task} has been on hold for N working days` with the blocker reason.

---

## Fractional Days

Both Original Days and Total Days support decimal values like `0.5`, `1.5`, `2.5`.

### How it works
- For 2.5 days starting Monday: Mon (1.0 day), Tue (1.0 day), Wed (0.5 day)
- Weekends, public holidays, and leave days are always skipped
- Capacity calculations sum fractional allocations correctly

### Tip
Useful for small tasks (half-day meetings, 1.5-day reviews) or when splitting time across multiple workstreams.

---

## Public Holidays & Leave

Account for time when team members aren't available. The tool automatically adjusts scheduling and capacity calculations.

### Adding public holidays
1. Go to the **Public Holidays** sheet
2. Add one row per holiday: Date and Name
3. These dates are treated like weekends for **all** team members

### Adding leave
1. Go to the **Leave** sheet
2. Add one row per leave period: Person, Start Date, End Date, Type, Notes
3. The date range is expanded into individual weekdays at load time
4. Only the named person is affected — their tasks extend around leave days

### How it affects scheduling
- A 5-day task starting Monday, with Wednesday being a public holiday: Mon, Tue, Thu, Fri, next Mon
- A 5-day task for someone with 2 leave days during the period: the task extends by 2 working days
- Start dates that fall on a non-working day (holiday or leave) snap to the next working day

### How it affects capacity
- **Weekly chart**: per-person capacity lines dip during leave/holiday weeks
- **Monthly chart**: available capacity line adjusts for holidays + leave
- **Executive summary**: utilisation uses leave-adjusted available days; leave summary shows per-person leave with types
- **Schedule suggestions**: spare capacity uses leave-adjusted availability; warns when leave overlaps active tasks

### What you'll see on charts
- **Gantt**: dotted purple lines for public holidays, small triangle markers (▼) for leave days on task rows
- **Weekly**: purple shading for holiday weeks, "NL" markers for leave days, capacity lines that dip
- **Monthly**: leave day annotations below months with 3+ leave days

### Tip
Both sheets are optional. If you don't have any holidays or leave to track, simply delete the example rows or remove the sheets entirely.

---

## Reading the Charts

### Gantt Chart (`capacity_gantt.png`)

- Rows grouped by workstream (grey header bars with priority badges)
- Each task is a coloured bar spanning its working days, labelled with duration in working days (e.g. "12.5 wd")
- Weekend shading (light grey vertical bands) for visual rhythm
- Hatch patterns distinguish team members
- Status symbols: `>` In Progress, `-` Planned, checkmark Complete, `||` On Hold
- Today line (red dashed vertical line)
- Public holiday dotted lines (purple)
- Leave day triangle markers (▼) above task rows
- Deadline diamond markers (red ◆) — red overshoot tint if task extends past deadline
- Confidence dots: green (High), amber (Medium), red (Low) next to task labels
- Long task names truncated at ~40 characters with "..."

### Weekly Capacity (`capacity_weekly.png`)

- Per-person side-by-side bars showing days allocated per week
- Individual dashed capacity lines per person (dip during leave/holiday weeks)
- Per-person utilisation percentage labels
- Over-capacity bars highlighted in red per person, with "+Nd" overshoot annotation
- Public holiday week shading (purple) with "N hol" label
- Leave markers ("NL") per person for weeks with leave

### Monthly Capacity Overview (`capacity_monthly.png`)
- Grouped bars: one group per month, one bar per team member
- Dashed line = available capacity (adjusts for working days, public holidays, and leave)
- Percentage labels showing utilisation
- Per-person over-capacity highlighting in red (only the individual over-capacity, not the whole group)
- Leave day annotations below months with 3+ leave days

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
- Public holiday count within the timeline period
- Overall utilisation percentage (leave-adjusted)
- Per-person utilisation (days allocated / leave-adjusted days available)
- Over-capacity week count with per-person detail (which person, which weeks, how much over)
- Leave summary with types per person
- Priority breakdown (tasks and days per priority level)
- Estimation drift total (if any tasks have drifted)
- Deadline warnings (tasks at risk of missing deadlines)
- Low-confidence task flags
- Concurrent task notes (when a person has 3+ tasks overlapping in a week)

### Schedule Suggestions
Printed after the summary if any suggestions are detected:
- **Early finishers**: recommends pulling forward subsequent tasks
- **Overdue warnings**: flags In Progress tasks past their planned end
- **Blocked duration**: shows how long On Hold tasks have been blocked (no artificial cap)
- **Leave overlaps**: warns when a person has leave during an active task
- **Spare capacity**: highlights future weeks where a person has 3+ free days (capped at 5 entries, uses leave-adjusted availability)

### Output File List
After chart generation, prints which files were written:
```
  Output:
    output/capacity_gantt.png
    output/capacity_weekly.png
    output/capacity_monthly.png
    output/roadmap.png
```

---

## Command-Line Options

```bash
# Generate the Excel template (overwrites existing capacity_data.xlsx)
python capacity_planner.py --template

# Generate all charts (default — 4 PNGs)
python capacity_planner.py

# Generate specific charts only
python capacity_planner.py --charts gantt
python capacity_planner.py --charts weekly
python capacity_planner.py --charts monthly
python capacity_planner.py --charts roadmap
python capacity_planner.py --charts gantt weekly
python capacity_planner.py --charts gantt roadmap

# Custom input file
python capacity_planner.py --input path/to/my_data.xlsx

# Custom output directory (affects all 4 PNGs)
python capacity_planner.py --outdir path/to/reports/

# Date window filter (only tasks overlapping with the range)
python capacity_planner.py --from 2026-04-01 --to 2026-06-30

# Combine options
python capacity_planner.py --charts gantt weekly --outdir reports/ --from 2026-04-01 --to 2026-06-30
```

---

## Tips

- **Re-run after every Excel update** — the tool reads fresh data each time
- **Only 6 fields to add a task** — Task, Workstream, Assigned To, Start Date, Total Days, Status. The rest is optional.
- **Leave Priority blank** — it inherits from the workstream. Only fill it when a task is unusually urgent or low-priority.
- **Leave Original Days blank** — it auto-fills from Total Days. Only matters later if you update Total Days (drift tracking).
- **Use P1 sparingly** — if everything is P1, nothing stands out
- **Deadline and Confidence for key tasks only** — don't fill these for routine tasks; save them for the 2-3 where you need to set expectations with your boss
- **Check console output** — the schedule suggestions catch things the charts don't show (overdue tasks, spare capacity, deadline warnings)
- **Fractional days for small tasks** — 0.5 for a half-day review, 1.5 for a day-and-a-half workshop
- **Workstream names must match exactly** — if a task references a workstream that doesn't exist on the Workstreams sheet, validation will flag it with a fuzzy match suggestion
- **Use `--from`/`--to` for focused views** — generate a "next quarter" chart without removing tasks from your Excel
- **Use `--outdir` for reports** — save charts to a specific folder for sharing
