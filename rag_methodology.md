# RAG Status Methodology
### Project Health Reporting Agent — Zycus Professional Services

---

## What is RAG?

RAG stands for **Red / Amber / Green** — a traffic-light system used to communicate project health at a glance:

| Status | Meaning |
|--------|---------|
| 🟢 **Green** | Project is on track. No significant risks. Proceeding as planned. |
| 🟡 **Amber** | Project is at risk. Issues exist but are being managed. Monitoring required. |
| 🔴 **Red** | Project is in trouble. Immediate escalation or corrective action required. |

---

## The 5-Dimension Scoring Framework

We evaluate every project across **5 weighted dimensions**. Each dimension scores 0, 1, or 2 points. The final weighted score determines the RAG status.

---

### Dimension 1 — Schedule Health (Weight: 30%)

Measures whether the project timeline is being followed.

| Score | Criteria |
|-------|----------|
| **2 — Green** | Project's `Schedule Health` field = Green. Timeline is on track or ahead of baseline. |
| **1 — Amber** | `Schedule Health` = Yellow. Minor delays (<10% of total phase duration). |
| **0 — Red** | `Schedule Health` = Red. Significant slippage (>10%) or baseline dates missed. |

> **Data Source:** `Schedule Health` column in the Project Plan sheet + `Summary` sheet.

---

### Dimension 2 — Task Completion Rate (Weight: 25%)

Compares actual completion to what was expected at this point in the project timeline.

**Formula:**
```
Expected % Complete = (Today − Project Start Date) / (Project End Date − Project Start Date)
Gap = Expected % Complete − Actual % Complete
```

| Score | Criteria |
|-------|----------|
| **2 — Green** | Gap ≤ 0% (project is at or ahead of expected progress) |
| **1 — Amber** | Gap is between 1%–15% (slightly behind, recoverable) |
| **0 — Red** | Gap > 15% (significantly behind expected trajectory) |

> **Data Source:** `% Complete` and `Project Start/End Date` from the `Summary` sheet.

---

### Dimension 3 — Blocker & Risk Exposure (Weight: 20%)

Measures the number and severity of active blockers and on-hold tasks.

| Score | Criteria |
|-------|----------|
| **2 — Green** | No tasks `On Hold`. `At Risk` field = Low or None. |
| **1 — Amber** | Tasks `On Hold` exist but are non-critical. `At Risk` = Medium. |
| **0 — Red** | Critical-path tasks are `On Hold`, OR `At Risk` = High and unmitigated. |

> **Data Source:** `On Hold?`, `At Risk?`, and `Critical?` columns in the Project Plan sheet.

---

### Dimension 4 — Overdue Critical Tasks (Weight: 15%)

Flags tasks that are on the critical path and have passed their due date without completion.

A task is considered **overdue** if:
- `End Date` < Today's Date, AND
- `Status` ≠ `Completed`, AND
- `Critical?` = Yes

| Score | Criteria |
|-------|----------|
| **2 — Green** | Zero overdue critical tasks. |
| **1 — Amber** | 1–2 overdue non-critical tasks (critical path unaffected). |
| **0 — Red** | 1 or more overdue critical tasks. |

> **Data Source:** `End Date`, `Status`, `Critical?` columns.

---

### Dimension 5 — Qualitative Sentiment (Weight: 10%)

Interprets free-text comments and status notes to assess stakeholder confidence and hidden risks that do not show up in structured data.

The AI agent reads all available `Status Comment`, `Comments`, and the `Comments` sheet entries and evaluates language signals:

| Score | Criteria |
|-------|----------|
| **2 — Green** | Comments signal smooth progress, completed workshops, positive client collaboration. |
| **1 — Amber** | Mixed signals: pending dependencies, client data delays, minor disagreements. |
| **0 — Red** | Escalation language, stalled client actions, unresolved blockers, or resource constraints. |

> **Data Source:** `Status Comment` column + `Comments` sheet (parsed by Gemini LLM).

---

## Final RAG Calculation

```
Final Score = (D1 × 0.30) + (D2 × 0.25) + (D3 × 0.20) + (D4 × 0.15) + (D5 × 0.10)
Max possible score = (2 × 0.30) + (2 × 0.25) + (2 × 0.20) + (2 × 0.15) + (2 × 0.10) = 2.0
```

| Final Score | RAG Status |
|-------------|-----------|
| **1.50 – 2.00** | 🟢 **GREEN** — On Track |
| **0.80 – 1.49** | 🟡 **AMBER** — At Risk |
| **0.00 – 0.79** | 🔴 **RED** — Critical |

---

## Messy Data Handling

Real-world project plans are often incomplete or contain formula errors. The agent handles these gracefully:

| Problem | Handling |
|---------|----------|
| `#UNPARSEABLE` date values | Fall back to `Baseline Start/Finish` columns. If still missing, estimate from `Duration`. |
| Missing `% Complete` | Treated as 0% for scoring (conservative assumption). |
| Blank `Owner` / `Assigned To` | Flagged as a minor risk; adds a note in the report. |
| `Not Applicable` tasks | Excluded from all scoring calculations. |
| Empty `Comments` sheet | Agent uses quantitative signals only; notes "No qualitative data available." |
| `On Hold?` / `Not Applicable?` mixed with text/boolean | Agent normalizes to True/False using string matching (`"yes"`, `"true"`, `"1"`). |

---

## Key Assumptions

1. **Today's Date** is sourced from the `Summary` sheet (`Today's Date` field) to allow historical report replay.
2. **Critical tasks** are identified via the `Critical?` column (value = `Yes`).
3. **Expected progress** is calculated linearly based on elapsed time relative to total project duration.
4. **LLM Override:** The Gemini AI agent may upgrade or downgrade the computed RAG status by one level if strong qualitative evidence from comments justifies it. The reasoning is always explained in plain English.
5. **At Risk = High** in the `Summary` sheet is treated as a strong Amber→Red signal regardless of other scores.

---

## Example Scoring: UniSan Project (Project Plan B.xlsx)

| Dimension | Value | Score |
|-----------|-------|-------|
| Schedule Health | Red (Summary field) | 0 pts |
| Task Completion | Actual 44% vs Expected ~60% → Gap 16% | 0 pts |
| Blocker Exposure | At Risk = High | 0 pts |
| Overdue Critical Tasks | To be computed by agent | TBD |
| Qualitative Sentiment | No comments available | 1 pt (neutral) |
| **Final Score** | **(0×0.30)+(0×0.25)+(0×0.20)+(0×0.15)+(1×0.10)** | **0.10 → 🔴 RED** |

---

*Document Version: 1.0 | Prepared for Zycus AI Engineer Intern Assignment*
