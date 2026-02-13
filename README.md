# Capacity Planning Tool

A Python tool for visualising team workload and capacity. Reads work requests from an Excel file, calculates weekly capacity, and outputs a professional two-panel Gantt chart PNG showing where the team is under/over capacity.

## What It Does

- **Gantt Chart (top panel):** Tasks grouped by work stream, colour-coded, with hatch patterns per team member. Shows a "Today" marker line. Complete tasks are faded, On Hold tasks have a distinct outline.
- **Capacity View (bottom panel):** Weekly stacked bar chart of days allocated per person vs team capacity. Weeks exceeding capacity are highlighted in red.

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
This creates `capacity_data.xlsx` with three sheets and example data:
- **Team** — team members and available days per week
- **Work Streams** — work stream names and display colours
- **Tasks** — tasks with stream, assignee, start date, duration, and status

### 2. Edit the Excel File
Open `capacity_data.xlsx` and update with your actual tasks and team info.

**Status values:** `Planned`, `In Progress`, `Complete`, `On Hold`

**Total Days:** working days the task takes (the tool distributes across weekdays automatically, skipping weekends).

### 3. Generate the Chart
```bash
python capacity_planner.py
```
Outputs `output/capacity_gantt.png` — a high-DPI PNG suitable for presentations.

### Custom Paths
```bash
python capacity_planner.py --input path/to/file.xlsx --output path/to/chart.png
```

## Example Output

Running with the template example data produces a chart covering ~14 weeks with 10 tasks across 9 work streams, showing capacity utilisation for a 2-person team.

## File Structure
```
capacity_planner.py    # Single script — all logic
capacity_data.xlsx     # Excel input (generated via --template, then user-maintained)
output/
  capacity_gantt.png   # Generated chart output
```
