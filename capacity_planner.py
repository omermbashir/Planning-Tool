"""
Capacity Planning Tool
Reads work requests from Excel, calculates weekly capacity for the team,
and outputs a professional Gantt chart PNG with capacity utilisation.
"""

import argparse
import os
import sys
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ── Constants ────────────────────────────────────────────────────────────────

DEFAULT_INPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "capacity_data.xlsx")
DEFAULT_OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "capacity_gantt.png")

STREAM_COLORS = {
    "2027 Product Strategy": "#2196F3",
    "People Champion": "#4CAF50",
    "Line Management": "#FF9800",
    "Leeds Social Committee": "#9C27B0",
    "Build New Analytics Platform": "#F44336",
    "Maintain Data During Transformation": "#00BCD4",
    "Deliver Ongoing Product Data (BAU)": "#FF5722",
    "Migrate to Intermediate Platform": "#795548",
    "Migrate to AWS Databricks": "#607D8B",
}

STATUS_VALUES = ["Planned", "In Progress", "Complete", "On Hold"]

PERSON_HATCHES = {
    0: "",       # First person: solid fill
    1: "//",     # Second person: diagonal hatch
    2: "\\\\",   # Third (if ever): back-diagonal
    3: "xx",     # Fourth: cross-hatch
}


# ── Template Generation ─────────────────────────────────────────────────────

def generate_template(output_path):
    """Create an Excel template with 3 sheets and example data."""
    wb = Workbook()

    # ── Styling ──
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2E3B4E", end_color="2E3B4E", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def style_header(ws, row=1):
        for cell in ws[row]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

    def style_data_rows(ws, start_row=2):
        for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

    # ── Sheet 1: Team ──
    ws_team = wb.active
    ws_team.title = "Team"
    ws_team.append(["Name", "Role", "Days Per Week"])
    ws_team.append(["Omer", "Lead", 5])
    ws_team.append(["Analyst", "Analyst", 5])
    ws_team.column_dimensions["A"].width = 20
    ws_team.column_dimensions["B"].width = 15
    ws_team.column_dimensions["C"].width = 16
    style_header(ws_team)
    style_data_rows(ws_team)

    # ── Sheet 2: Work Streams ──
    ws_streams = wb.create_sheet("Work Streams")
    ws_streams.append(["Stream", "Color"])
    for stream, color in STREAM_COLORS.items():
        ws_streams.append([stream, color])
    ws_streams.column_dimensions["A"].width = 45
    ws_streams.column_dimensions["B"].width = 12
    style_header(ws_streams)
    style_data_rows(ws_streams)
    # Color-fill the Color column cells to preview
    for row_idx in range(2, ws_streams.max_row + 1):
        color_cell = ws_streams.cell(row=row_idx, column=2)
        hex_color = color_cell.value.lstrip("#") if color_cell.value else "FFFFFF"
        color_cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    # ── Sheet 3: Tasks ──
    ws_tasks = wb.create_sheet("Tasks")
    ws_tasks.append(["Task", "Stream", "Assigned To", "Start Date", "Total Days", "Status", "Notes"])

    example_tasks = [
        ["Recommender AB Test Analysis", "Deliver Ongoing Product Data (BAU)", "Omer", "2026-02-16", 10, "In Progress", "Priority from product team"],
        ["Navigation AB Test Analysis", "Deliver Ongoing Product Data (BAU)", "Analyst", "2026-03-02", 8, "Planned", ""],
        ["Poker Strategy Research", "2027 Product Strategy", "Omer", "2026-03-02", 15, "Planned", "Stakeholder interviews + market review"],
        ["Team 1:1s & Reviews", "Line Management", "Omer", "2026-02-16", 2, "In Progress", "Recurring weekly - 2 days/month"],
        ["Q1 People Survey Actions", "People Champion", "Omer", "2026-03-16", 5, "Planned", ""],
        ["Social Event Planning", "Leeds Social Committee", "Analyst", "2026-02-23", 3, "Planned", "March team event"],
        ["Analytics Platform Requirements", "Build New Analytics Platform", "Omer", "2026-04-01", 20, "Planned", "Discovery & requirements gathering"],
        ["Legacy Dashboard Migration", "Migrate to Intermediate Platform", "Analyst", "2026-03-16", 15, "Planned", "Phase 1 dashboards"],
        ["Transformation Data Handover", "Maintain Data During Transformation", "Omer", "2026-03-09", 10, "Planned", "Document current pipelines"],
        ["AWS Databricks POC", "Migrate to AWS Databricks", "Analyst", "2026-05-01", 12, "Planned", "Initial proof of concept"],
    ]
    for task in example_tasks:
        ws_tasks.append(task)

    ws_tasks.column_dimensions["A"].width = 35
    ws_tasks.column_dimensions["B"].width = 42
    ws_tasks.column_dimensions["C"].width = 14
    ws_tasks.column_dimensions["D"].width = 14
    ws_tasks.column_dimensions["E"].width = 12
    ws_tasks.column_dimensions["F"].width = 12
    ws_tasks.column_dimensions["G"].width = 35
    style_header(ws_tasks)
    style_data_rows(ws_tasks)

    # Date format for start date column
    for row_idx in range(2, ws_tasks.max_row + 1):
        cell = ws_tasks.cell(row=row_idx, column=4)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(output_path)
    print(f"Template created: {output_path}")
    print("  - Sheet 'Team': define your team members and available days")
    print("  - Sheet 'Work Streams': work streams with display colors")
    print("  - Sheet 'Tasks': add your tasks with start dates and durations")
    print(f"\nEdit the file, then run again without --template to generate the chart.")


# ── Data Loading ─────────────────────────────────────────────────────────────

def load_team(filepath):
    """Load team members from the 'Team' sheet."""
    df = pd.read_excel(filepath, sheet_name="Team")
    df.columns = df.columns.str.strip()
    team = {}
    for _, row in df.iterrows():
        name = str(row["Name"]).strip()
        days = float(row["Days Per Week"])
        team[name] = days
    return team


def load_streams(filepath):
    """Load work streams and colors from the 'Work Streams' sheet."""
    df = pd.read_excel(filepath, sheet_name="Work Streams")
    df.columns = df.columns.str.strip()
    streams = {}
    for _, row in df.iterrows():
        stream = str(row["Stream"]).strip()
        color = str(row["Color"]).strip()
        streams[stream] = color
    return streams


def load_tasks(filepath):
    """Load tasks from the 'Tasks' sheet."""
    df = pd.read_excel(filepath, sheet_name="Tasks")
    df.columns = df.columns.str.strip()
    tasks = []
    for _, row in df.iterrows():
        start = row["Start Date"]
        if isinstance(start, str):
            start = datetime.strptime(start, "%Y-%m-%d")
        elif isinstance(start, pd.Timestamp):
            start = start.to_pydatetime()

        tasks.append({
            "task": str(row["Task"]).strip(),
            "stream": str(row["Stream"]).strip(),
            "assigned_to": str(row["Assigned To"]).strip(),
            "start_date": start,
            "total_days": int(row["Total Days"]),
            "status": str(row["Status"]).strip(),
            "notes": str(row.get("Notes", "")).strip(),
        })
    return tasks


def load_data(filepath):
    """Load all data from the Excel file."""
    team = load_team(filepath)
    streams = load_streams(filepath)
    tasks = load_tasks(filepath)
    return team, streams, tasks


# ── Schedule Calculation ─────────────────────────────────────────────────────

def get_end_date(start_date, total_working_days):
    """Calculate end date by adding working days (skipping weekends)."""
    current = start_date
    days_added = 0
    while days_added < total_working_days:
        if current.weekday() < 5:  # Mon-Fri
            days_added += 1
            if days_added == total_working_days:
                return current
        current += timedelta(days=1)
    return current


def calculate_schedule(tasks):
    """For each task, compute start/end dates and list of working days."""
    for task in tasks:
        start = task["start_date"]
        # Ensure start is a weekday
        while start.weekday() >= 5:
            start += timedelta(days=1)
        task["start_date"] = start
        task["end_date"] = get_end_date(start, task["total_days"])

        # Build list of actual working days
        working_days = []
        current = start
        while len(working_days) < task["total_days"]:
            if current.weekday() < 5:
                working_days.append(current)
            current += timedelta(days=1)
        task["working_days"] = working_days
    return tasks


# ── Capacity Calculation ─────────────────────────────────────────────────────

def get_week_start(date):
    """Get the Monday of the week containing the given date."""
    return date - timedelta(days=date.weekday())


def calculate_capacity(tasks, team):
    """Calculate per-person per-week allocation."""
    if not tasks:
        return {}, []

    # Find date range
    all_dates = []
    for t in tasks:
        all_dates.extend(t.get("working_days", []))
    if not all_dates:
        return {}, []

    min_date = get_week_start(min(all_dates))
    max_date = max(all_dates)
    max_week_start = get_week_start(max_date)

    # Build list of week-start Mondays
    weeks = []
    current = min_date
    while current <= max_week_start:
        weeks.append(current)
        current += timedelta(days=7)

    # Allocate: {week_start: {person: days_allocated}}
    allocation = {w: {name: 0 for name in team} for w in weeks}

    for task in tasks:
        person = task["assigned_to"]
        if person not in team:
            continue
        for day in task.get("working_days", []):
            ws = get_week_start(day)
            if ws in allocation:
                allocation[ws][person] += 1

    return allocation, weeks


# ── Chart Rendering ──────────────────────────────────────────────────────────

def render_chart(tasks, team, streams, allocation, weeks, output_path):
    """Render the two-panel chart and save as PNG."""
    # Filter out Complete tasks from active display (show them faded)
    active_tasks = [t for t in tasks if t["status"] != "Complete"]
    complete_tasks = [t for t in tasks if t["status"] == "Complete"]
    all_display_tasks = tasks  # show all, but style differently

    # Group tasks by stream
    stream_order = list(streams.keys())
    grouped = {}
    for stream in stream_order:
        stream_tasks = [t for t in all_display_tasks if t["stream"] == stream]
        if stream_tasks:
            grouped[stream] = stream_tasks

    if not grouped:
        print("No tasks to display. Check your Excel data.")
        return

    # Count total rows for Gantt (tasks + stream headers)
    total_rows = sum(len(ts) + 1 for ts in grouped.values())  # +1 per stream header

    # Person ordering for hatch patterns
    person_list = list(team.keys())
    person_hatch = {name: PERSON_HATCHES.get(i, "") for i, name in enumerate(person_list)}

    # ── Figure Setup ──
    fig_height = max(10, total_rows * 0.45 + 4)
    fig = plt.figure(figsize=(18, fig_height), facecolor="white")

    # GridSpec: top panel = Gantt (3/4), bottom panel = capacity (1/4)
    gs = fig.add_gridspec(2, 1, height_ratios=[3, 1], hspace=0.25)
    ax_gantt = fig.add_subplot(gs[0])
    ax_cap = fig.add_subplot(gs[1])

    # ── Date range ──
    if not weeks:
        print("No weeks to display.")
        return

    date_min = weeks[0] - timedelta(days=2)
    date_max = weeks[-1] + timedelta(days=9)

    # ── Render Gantt Panel ──
    y_pos = total_rows - 1
    y_ticks = []
    y_labels = []
    y_colors = []

    for stream in stream_order:
        if stream not in grouped:
            continue
        stream_color = streams.get(stream, "#888888")
        stream_tasks = grouped[stream]

        # Stream header row
        ax_gantt.barh(
            y_pos, (date_max - date_min).days,
            left=mdates.date2num(date_min),
            height=0.8, color=stream_color, alpha=0.12,
            edgecolor="none"
        )
        y_ticks.append(y_pos)
        y_labels.append(stream)
        y_colors.append(stream_color)
        y_pos -= 1

        # Task rows
        for task in stream_tasks:
            start_num = mdates.date2num(task["start_date"])
            end_num = mdates.date2num(task["end_date"])
            duration = end_num - start_num + 1

            # Determine style based on status
            alpha = 1.0
            edgecolor = stream_color
            linewidth = 1.5
            hatch = person_hatch.get(task["assigned_to"], "")

            if task["status"] == "Complete":
                alpha = 0.3
                edgecolor = "#AAAAAA"
            elif task["status"] == "On Hold":
                alpha = 0.5
                linewidth = 2.0
                edgecolor = "#333333"

            bar = ax_gantt.barh(
                y_pos, duration,
                left=start_num,
                height=0.6, color=stream_color, alpha=alpha,
                edgecolor=edgecolor, linewidth=linewidth,
                hatch=hatch
            )

            # Task label
            label_color = "#333333" if task["status"] != "Complete" else "#999999"
            label_x = start_num + duration / 2
            ax_gantt.text(
                label_x, y_pos,
                f"  {task['task']} ({task['assigned_to']})",
                va="center", ha="center",
                fontsize=7.5, color=label_color,
                fontweight="medium",
                clip_on=True
            )

            y_ticks.append(y_pos)
            y_labels.append("")
            y_colors.append(stream_color)
            y_pos -= 1

    # Gantt formatting
    ax_gantt.set_yticks(y_ticks)
    ax_gantt.set_yticklabels(y_labels, fontsize=8)
    for tick_label, color in zip(ax_gantt.get_yticklabels(), y_colors):
        if tick_label.get_text():  # Stream headers
            tick_label.set_fontweight("bold")
            tick_label.set_fontsize(9)
            tick_label.set_color(color)

    ax_gantt.set_xlim(mdates.date2num(date_min), mdates.date2num(date_max))
    ax_gantt.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=0))
    ax_gantt.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
    ax_gantt.xaxis.set_minor_locator(mdates.DayLocator())
    plt.setp(ax_gantt.xaxis.get_majorticklabels(), rotation=45, ha="right", fontsize=8)

    # Today line
    today = datetime.now()
    today_num = mdates.date2num(today)
    if mdates.date2num(date_min) <= today_num <= mdates.date2num(date_max):
        ax_gantt.axvline(today_num, color="#E53935", linewidth=1.5, linestyle="--", alpha=0.8, zorder=10)
        ax_gantt.text(
            today_num, total_rows - 0.5, "  Today",
            fontsize=8, color="#E53935", fontweight="bold", va="bottom"
        )

    ax_gantt.set_title("Team Capacity Plan — Gantt View", fontsize=14, fontweight="bold", pad=15)
    ax_gantt.grid(axis="x", alpha=0.2, linewidth=0.5)
    ax_gantt.set_axisbelow(True)
    ax_gantt.spines["top"].set_visible(False)
    ax_gantt.spines["right"].set_visible(False)

    # ── Render Capacity Panel ──
    total_capacity = sum(team.values())
    week_labels = [w.strftime("%d %b") for w in weeks]
    x_positions = np.arange(len(weeks))
    bar_width = 0.6

    # Stack bars per person
    bottom = np.zeros(len(weeks))
    person_bars = {}

    for person_idx, person in enumerate(person_list):
        values = [allocation[w].get(person, 0) for w in weeks]
        person_bars[person] = values

    # Determine over-capacity weeks
    total_per_week = np.zeros(len(weeks))
    for person in person_list:
        total_per_week += np.array(person_bars[person])

    for person_idx, person in enumerate(person_list):
        values = np.array(person_bars[person])
        colors = []
        for i, w in enumerate(weeks):
            if total_per_week[i] > total_capacity:
                colors.append("#E53935")  # Red for over-capacity
            else:
                colors.append("#4CAF50" if person_idx == 0 else "#2196F3")

        ax_cap.bar(
            x_positions, values, bar_width,
            bottom=bottom, color=colors, alpha=0.85,
            edgecolor="white", linewidth=0.5,
            hatch=PERSON_HATCHES.get(person_idx, ""),
            label=person
        )
        bottom += values

    # Capacity line
    ax_cap.axhline(
        total_capacity, color="#333333", linewidth=2,
        linestyle="--", alpha=0.8, label=f"Team capacity ({int(total_capacity)} days/wk)"
    )

    # Individual capacity lines
    cumulative = 0
    for person in person_list:
        cumulative += team[person]
        ax_cap.axhline(
            cumulative, color="#999999", linewidth=1,
            linestyle=":", alpha=0.5
        )

    ax_cap.set_xticks(x_positions)
    ax_cap.set_xticklabels(week_labels, rotation=45, ha="right", fontsize=8)
    ax_cap.set_ylabel("Days Allocated", fontsize=10)
    ax_cap.set_title("Weekly Capacity Utilisation", fontsize=12, fontweight="bold", pad=10)
    ax_cap.legend(loc="upper right", fontsize=8, framealpha=0.9)
    ax_cap.grid(axis="y", alpha=0.2, linewidth=0.5)
    ax_cap.spines["top"].set_visible(False)
    ax_cap.spines["right"].set_visible(False)
    ax_cap.set_axisbelow(True)

    # Set y-axis limit with headroom
    max_alloc = max(total_per_week) if len(total_per_week) > 0 else total_capacity
    ax_cap.set_ylim(0, max(total_capacity, max_alloc) * 1.15)

    # ── Legend for streams + person patterns ──
    legend_handles = []
    for stream, color in streams.items():
        if stream in grouped:
            legend_handles.append(mpatches.Patch(facecolor=color, edgecolor=color, label=stream))

    for person_idx, person in enumerate(person_list):
        legend_handles.append(mpatches.Patch(
            facecolor="#CCCCCC",
            edgecolor="#333333",
            hatch=PERSON_HATCHES.get(person_idx, ""),
            label=f"{person} (assigned)"
        ))

    ax_gantt.legend(
        handles=legend_handles, loc="upper left",
        bbox_to_anchor=(0, -0.02), ncol=3, fontsize=7.5,
        framealpha=0.9, edgecolor="#CCCCCC"
    )

    # ── Title & footer ──
    if weeks:
        date_range = f"{weeks[0].strftime('%d %b %Y')} — {weeks[-1].strftime('%d %b %Y')}"
    else:
        date_range = "No data"

    fig.suptitle(
        f"Capacity Plan: {date_range}",
        fontsize=16, fontweight="bold", y=0.98
    )
    fig.text(
        0.99, 0.01, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}",
        ha="right", fontsize=7, color="#999999"
    )

    # ── Save ──
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"Chart saved: {output_path}")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Capacity Planning Tool — generate Gantt charts from Excel data"
    )
    parser.add_argument(
        "--template", action="store_true",
        help="Generate a blank Excel template with example data"
    )
    parser.add_argument(
        "--input", default=DEFAULT_INPUT,
        help=f"Path to Excel input file (default: {DEFAULT_INPUT})"
    )
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT,
        help=f"Path for output PNG (default: {DEFAULT_OUTPUT})"
    )
    args = parser.parse_args()

    if args.template:
        generate_template(args.input)
        return

    # Load data
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        print("Run with --template first to create a template.")
        sys.exit(1)

    print(f"Loading data from: {args.input}")
    team, streams, tasks = load_data(args.input)
    print(f"  Team: {', '.join(f'{n} ({d}d/wk)' for n, d in team.items())}")
    print(f"  Streams: {len(streams)}")
    print(f"  Tasks: {len(tasks)}")

    # Calculate schedule
    tasks = calculate_schedule(tasks)

    # Calculate capacity
    allocation, weeks = calculate_capacity(tasks, team)
    print(f"  Weeks covered: {len(weeks)}")

    # Render chart
    render_chart(tasks, team, streams, allocation, weeks, args.output)


if __name__ == "__main__":
    main()
