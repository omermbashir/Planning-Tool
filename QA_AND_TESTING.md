# QA & Testing

I ran 20 rounds of code review with an external reviewer. 42 bugs were found and fixed, 7 false positives were rejected, and I built a 157-test suite across 6 tiers. Three clean rounds in a row before shipping.

## Code Review Process

The reviewer audited the full codebase each round, reported findings with root causes, and I fixed each confirmed bug before the next round.

Over time, the bugs got less severe:

- **Early rounds** turned up core math errors - NaN propagation through capacity calculations, off-by-one chart alignment, time-of-day sensitivity in date comparisons.
- **Middle rounds** shifted to semantic bugs - status fields not being filtered properly in capacity math, planned dates used where actual dates should apply, validation gaps that silently distorted output.
- **Late rounds** were mostly cross-output contradictions - one chart showing a task as late while another showed the person as free, or a fix applied to the text summary but not to the Gantt.
- **Final rounds** only turned up false positives and minor edge cases. After 3 clean rounds, the tool shipped.

42 bugs fixed total. 7 findings were rejected as false positives (e.g., the reviewer flagged a missing `os.makedirs` call but missed the 4 existing calls inside the renderers).

## Bug Families

The 42 bugs weren't random - they clustered into 5 recurring families. Once I noticed this, I could grep for all instances of a family and fix them in one pass instead of waiting for each one to get caught.

| Family | Pattern | Count | Example |
|--------|---------|-------|---------|
| Planned vs Actual | Complete tasks using planned dates instead of actual dates for comparisons, capacity, or filtering | 7 | A task that finished early still consumed capacity until its planned end date |
| On Hold leaks | On Hold tasks incorrectly included in capacity calculations, concurrent counts, priority totals, or visual indicators | 6 | On Hold tasks inflated utilisation percentages and triggered false over-capacity warnings |
| Incomplete propagation | A fix applied to one output channel but not all others (Gantt, weekly, monthly, roadmap, summary, suggestions) | 5 | Blocked-task marker added to the roadmap but not to the Gantt |
| Fix side effects | A data structure modification that fixed one consumer but broke another | 1 | Trimming `working_days` for capacity correctness broke the variance label that relied on the original length |
| Asymmetric fixes | A fix applied in one direction but not the opposite | 1 | Early-finish capacity trimming was implemented, but late-finish extension was not |

After identifying each family, every fix was followed by a grep audit across the codebase to catch any remaining instances of that pattern.

## Review Patterns

The reviews led to a set of patterns I now use to catch bugs before they ship. Each one exists because a specific class of bug kept getting through without it.

- **Cross-function data flow tracing** - trace a value from where it's set through every function that reads it, not just the one being fixed.
- **Planned vs actual data selection** - after any change involving `end_date`, grep every usage and check whether `actual_end_date` should be used for Complete tasks.
- **Fix propagation scope** - after fixing a bug in one output channel, check the same logic across all other channels.
- **Status semantics enforcement** - maintain a matrix of which statuses participate in each calculation (capacity, concurrency, priority totals, warnings, visuals) and audit against it.
- **Pre-review grep audits** - before each review round, grep all task aggregations and date comparisons to catch family instances early.
- **Bug family exhaustion** - when fixing a bug, identify its family and fix every remaining instance in one pass.
- **Contradiction detection** - compare what each output says about the same task and flag disagreements (e.g., Gantt shows overshoot but capacity shows zero).
- **Consumer audits after modification** - when a fix modifies a data structure, grep every downstream consumer to make sure nothing else broke.

## Test Suite

157 tests across 6 tiers. The first four tiers are standard (unit, function, integration, end-to-end). Tiers 5 and 6 came out of specific bugs - fixing one consumer kept breaking another, so those tiers trace a single task's state through every output channel at once.

| Tier | Focus | What it covers |
|------|-------|----------------|
| 1 | Pure unit tests | `norm_date()`, `clean_str()`, `parse_date()`, `get_week_start()`, `priority_sort_key()` - no I/O, no fixtures |
| 2 | Function tests | `is_working_day()`, `count_working_days()`, `get_end_date()`, `normalize_columns()`, `validate_data()` - constructed data, no Excel |
| 3 | Integration tests | `load_team()`, `load_workstreams()`, `load_tasks()`, `load_public_holidays()`, `load_leave()` - temp Excel round-trips |
| 4 | End-to-end tests | Full pipeline: Excel input to schedule calculation to chart rendering to PNG output |
| 5 | Cross-consumer regression | A single task state (e.g., early-finish Complete) traced through capacity, variance, confidence, deadlines, priority totals, and date filtering at the same time |
| 6 | Production readiness | Worst-case scenarios: all-same-status inputs, multi-person overlapping tasks, late-finish capacity extension, and interaction effects between features |

Tier 5 was added after a fix that trimmed a data structure for capacity correctness accidentally zeroed out the variance label - both consumers passed their own tests, but nobody tested them together. Tier 6 was added after the final bug showed that early-finish and late-finish logic needed to be symmetric, and that extended capacity data had to survive concurrent task filtering, confidence exclusions, and timeline bounds.

```bash
pytest test_capacity_planner.py -v
```

## Metrics

| Metric | Value |
|--------|-------|
| Automated tests | 157 |
| Test tiers | 6 |
| Review rounds | 20 |
| Bugs found & fixed | 42 |
| False positives rejected | 7 |
| Bug families identified | 5 |
| Review patterns developed | 21 |
| Clean rounds before ship | 3 |
