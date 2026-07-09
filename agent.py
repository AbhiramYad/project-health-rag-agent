"""
Project Health Reporting Agent
================================
Reads project Excel files, computes a weighted RAG score across 5 dimensions,
calls Gemini LLM for plain-English reasoning, and writes weekly health reports.

Usage:
    python agent.py

Requirements:
    pip install pandas openpyxl google-genai python-dotenv
    Add your key to .env file: GEMINI_API_KEY=your_key_here
"""

import os
import sys
import json
import re
from datetime import date, datetime
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types

# Load environment variables from .env file
load_dotenv()

# Fix Windows console encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
GEMINI_MODEL = "gemini-2.5-flash"
OUTPUT_DIR = "outputs"
TODAY_OVERRIDE = None  # Set to "YYYY-MM-DD" to replay a past date

PROJECT_FILES = {
    "UniSan_S2P": "Project Plan B.xlsx",
    "Outokumpu_S2P": "S2P Project.xlsx",
}

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def safe_float(val, default=0.0):
    """Convert value to float safely."""
    try:
        f = float(val)
        return f if not pd.isna(f) else default
    except Exception:
        return default


def safe_date(val):
    """Parse a date value from various formats, return None if unparseable."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    s = str(val).strip()
    if "#" in s or s == "" or s.lower() == "nan":
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def normalize_bool(val):
    """Normalize yes/no/true/false/1/0 to bool."""
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    return s in ("yes", "true", "1", "y")


# ─────────────────────────────────────────────
# LOAD & EXTRACT
# ─────────────────────────────────────────────

def load_project(filepath: str) -> dict:
    """
    Load and extract all relevant data from a project Excel file.
    Returns a structured dict with project metadata, task stats, and comments.
    """
    xls = pd.ExcelFile(filepath)
    data = {}

    # ── Summary Sheet ──────────────────────────────────────
    if "Summary" in xls.sheet_names:
        summary_df = pd.read_excel(xls, sheet_name="Summary")
        summary_df.columns = ["key", "value"]
        summary = dict(zip(summary_df["key"].astype(str), summary_df["value"]))
    else:
        summary = {}

    data["project_manager"] = str(summary.get("Project Manager", "Unknown"))
    data["project_start"]   = safe_date(summary.get("Project Start Date"))
    data["project_end"]     = safe_date(summary.get("Project End Date"))
    data["pct_complete"]    = safe_float(summary.get("% Complete", 0)) * 100  # stored as 0.xx
    data["at_risk"]         = str(summary.get("At Risk", "Unknown"))
    data["project_stage"]   = str(summary.get("Project Stage", "Unknown"))
    data["project_status"]  = str(summary.get("Project Status", "Unknown"))
    data["schedule_health"] = str(summary.get("Schedule Health", "Unknown"))
    data["today_date"]      = safe_date(summary.get("Today's Date")) or date.today()

    not_started = safe_float(summary.get("Not Started", 0))
    in_progress  = safe_float(summary.get("In Progress", 0))
    completed    = safe_float(summary.get("Completed", 0))
    on_hold      = safe_float(summary.get("On Hold", 0))
    data["task_counts"] = {
        "not_started": int(not_started),
        "in_progress":  int(in_progress),
        "completed":    int(completed),
        "on_hold":      int(on_hold),
        "total":        int(not_started + in_progress + completed + on_hold),
    }

    # ── Plan Sheet ────────────────────────────────────────
    plan_sheet = [s for s in xls.sheet_names if s not in ("Summary", "Comments")]
    if plan_sheet:
        plan_df = pd.read_excel(xls, sheet_name=plan_sheet[0])
    else:
        plan_df = pd.DataFrame()

    data["plan_columns"] = list(plan_df.columns) if not plan_df.empty else []
    data["project_name"] = "Unknown"

    # Extract project name from first non-null value in Project Name column
    if "Project Name" in plan_df.columns:
        names = plan_df["Project Name"].dropna()
        if not names.empty:
            data["project_name"] = str(names.iloc[0])

    # Overdue critical tasks
    overdue_critical = []
    overdue_noncritical = []
    today = data["today_date"]

    if not plan_df.empty:
        for _, row in plan_df.iterrows():
            status   = str(row.get("Status", "")).strip()
            critical = normalize_bool(row.get("Critical ?", False))
            end_raw  = row.get("End Date") or row.get("Baseline Finish")
            end_date = safe_date(end_raw)
            task_name = str(row.get("Task Name", "")).strip()

            if end_date and end_date < today and status not in ("Completed", "Not Applicable", "On Hold"):
                if critical:
                    overdue_critical.append({"task": task_name, "due": str(end_date)})
                else:
                    overdue_noncritical.append({"task": task_name, "due": str(end_date)})

    data["overdue_critical"]    = overdue_critical[:10]   # cap at 10
    data["overdue_noncritical"] = overdue_noncritical[:5]

    # On-hold tasks
    on_hold_tasks = []
    if not plan_df.empty and "On Hold?" in plan_df.columns:
        on_hold_df = plan_df[plan_df["On Hold?"].apply(normalize_bool)]
        on_hold_tasks = on_hold_df["Task Name"].dropna().astype(str).tolist()[:5]
    data["on_hold_tasks"] = on_hold_tasks

    # Sample status comments
    comments_list = []
    if not plan_df.empty and "Status Comment" in plan_df.columns:
        raw = plan_df["Status Comment"].dropna().astype(str)
        comments_list += [c for c in raw.tolist() if c.strip() and c.lower() != "nan"][:10]
    if not plan_df.empty and "Comments" in plan_df.columns:
        raw2 = plan_df["Comments"].dropna().astype(str)
        comments_list += [c for c in raw2.tolist() if c.strip() and c.lower() != "nan"][:5]

    # Comments sheet
    if "Comments" in xls.sheet_names:
        c_df = pd.read_excel(xls, sheet_name="Comments")
        if not c_df.empty:
            # Second column tends to hold the actual comment text
            for col in c_df.columns[1:2]:
                raw3 = c_df[col].dropna().astype(str)
                comments_list += [c for c in raw3.tolist() if c.strip() and c.lower() != "nan"][:10]

    data["comments"] = list(set(comments_list))[:15]

    return data


# ─────────────────────────────────────────────
# RAG SCORING ENGINE
# ─────────────────────────────────────────────

def compute_rag_score(p: dict) -> dict:
    """
    Compute the weighted 5-dimension RAG score.
    Returns scores dict and preliminary RAG status.
    """
    scores = {}
    notes  = {}

    today         = p["today_date"]
    proj_start    = p["project_start"]
    proj_end      = p["project_end"]
    pct_complete  = p["pct_complete"]
    sched_health  = p["schedule_health"].strip().lower()
    at_risk       = p["at_risk"].strip().lower()
    overdue_crit  = p["overdue_critical"]
    overdue_non   = p["overdue_noncritical"]
    on_hold_count = len(p["on_hold_tasks"])

    # ── D1: Schedule Health (30%) ─────────────────────
    if sched_health == "green":
        scores["schedule_health"] = 2
        notes["schedule_health"]  = "Schedule is Green — project is on track."
    elif sched_health in ("yellow", "amber"):
        scores["schedule_health"] = 1
        notes["schedule_health"]  = "Schedule is Yellow — minor delays detected."
    else:
        scores["schedule_health"] = 0
        notes["schedule_health"]  = f"Schedule is {sched_health.title()} — significant slippage."

    # ── D2: Task Completion Rate (25%) ───────────────
    if proj_start and proj_end and proj_end != proj_start:
        elapsed   = (today - proj_start).days
        total     = (proj_end - proj_start).days
        expected  = round((elapsed / total) * 100, 1)
    else:
        expected  = pct_complete  # fallback: no penalty

    gap = round(expected - pct_complete, 1)
    notes["completion"] = f"Actual: {pct_complete:.1f}% | Expected: {expected:.1f}% | Gap: {gap:.1f}%"

    if gap <= 0:
        scores["completion"] = 2
    elif gap <= 15:
        scores["completion"] = 1
    else:
        scores["completion"] = 0

    # ── D3: Blocker & Risk Exposure (20%) ────────────
    if at_risk in ("high",):
        scores["blockers"] = 0
        notes["blockers"]  = f"At Risk = High with {on_hold_count} tasks On Hold."
    elif on_hold_count > 0 or at_risk in ("medium", "moderate"):
        scores["blockers"] = 1
        notes["blockers"]  = f"{on_hold_count} tasks On Hold. At Risk = {at_risk.title()}."
    else:
        scores["blockers"] = 2
        notes["blockers"]  = "No significant blockers or on-hold tasks."

    # ── D4: Overdue Critical Tasks (15%) ──────────────
    if len(overdue_crit) > 0:
        scores["overdue"] = 0
        notes["overdue"]  = f"{len(overdue_crit)} overdue critical task(s): {[t['task'][:40] for t in overdue_crit[:3]]}"
    elif len(overdue_non) > 0:
        scores["overdue"] = 1
        notes["overdue"]  = f"No overdue critical tasks. {len(overdue_non)} non-critical task(s) overdue."
    else:
        scores["overdue"] = 2
        notes["overdue"]  = "No overdue tasks detected."

    # ── D5: Qualitative Sentiment (10%) ───────────────
    # Rule-based keyword scan; LLM will refine later
    red_keywords   = ["escalat", "stall", "block", "halt", "critical delay", "unresolved", "reject"]
    amber_keywords = ["pending", "delay", "impact", "depend", "waiting", "hold", "partial"]
    green_keywords = ["complet", "success", "on track", "smooth", "deliver", "done", "achiev"]

    comments_text = " ".join(p["comments"]).lower()
    red_hits   = sum(1 for k in red_keywords   if k in comments_text)
    amber_hits = sum(1 for k in amber_keywords if k in comments_text)
    green_hits = sum(1 for k in green_keywords if k in comments_text)

    if not comments_text.strip():
        scores["sentiment"] = 1
        notes["sentiment"]  = "No qualitative comments available — neutral assumption."
    elif red_hits >= 2:
        scores["sentiment"] = 0
        notes["sentiment"]  = f"Comments contain {red_hits} escalation/blocker signals."
    elif amber_hits > green_hits:
        scores["sentiment"] = 1
        notes["sentiment"]  = f"Mixed signals: {amber_hits} delay/pending vs {green_hits} positive indicators."
    else:
        scores["sentiment"] = 2
        notes["sentiment"]  = f"Comments are mostly positive ({green_hits} green signals detected)."

    # ── Final Score ────────────────────────────────────
    weights = {
        "schedule_health": 0.30,
        "completion":      0.25,
        "blockers":        0.20,
        "overdue":         0.15,
        "sentiment":       0.10,
    }
    final = sum(scores[k] * weights[k] for k in weights)
    final = round(final, 3)

    if final >= 1.50:
        preliminary = "Green"
    elif final >= 0.80:
        preliminary = "Amber"
    else:
        preliminary = "Red"

    return {
        "scores":      scores,
        "notes":       notes,
        "final_score": final,
        "preliminary": preliminary,
        "expected_pct": expected if proj_start and proj_end else None,
        "gap":         gap,
    }


# ─────────────────────────────────────────────
# LLM PROMPT & CALL
# ─────────────────────────────────────────────

def build_prompt(project_name: str, p: dict, scored: dict) -> str:
    """Build a structured prompt for the Gemini LLM."""
    comments_block = "\n".join(f"  - {c}" for c in p["comments"]) or "  (No comments available)"
    on_hold_block  = "\n".join(f"  - {t}" for t in p["on_hold_tasks"]) or "  None"
    overdue_block  = "\n".join(
        f"  - {t['task']} (due {t['due']})" for t in p["overdue_critical"]
    ) or "  None"

    return f"""You are a senior project delivery analyst at Zycus Professional Services.
Your job is to validate a pre-computed project health RAG status and provide clear, executive-level reasoning.

PROJECT CONTEXT
================
Project Name     : {project_name}
Project Manager  : {p['project_manager']}
Current Phase    : {p['project_stage']}
Project Status   : {p['project_status']}
Timeline         : {p['project_start']} to {p['project_end']}
Today's Date     : {p['today_date']}
At Risk Level    : {p['at_risk']}

PROGRESS
  Actual % Complete    : {p['pct_complete']:.1f}%
  Expected % Complete  : {scored.get('expected_pct', 'N/A')}%
  Gap (Expected-Actual): {scored.get('gap', 'N/A')}%

TASK BREAKDOWN
  Completed   : {p['task_counts']['completed']}
  In Progress : {p['task_counts']['in_progress']}
  Not Started : {p['task_counts']['not_started']}
  On Hold     : {p['task_counts']['on_hold']}
  Total Tasks : {p['task_counts']['total']}

SCHEDULE HEALTH (from system): {p['schedule_health']}

OVERDUE CRITICAL TASKS:
{overdue_block}

ON-HOLD TASKS:
{on_hold_block}

QUALITATIVE COMMENTS / STATUS NOTES:
{comments_block}

PRE-COMPUTED RAG SCORE
  Schedule Health  : {scored['scores']['schedule_health']}/2 - {scored['notes']['schedule_health']}
  Completion Rate  : {scored['scores']['completion']}/2 - {scored['notes']['completion']}
  Blocker Exposure : {scored['scores']['blockers']}/2 - {scored['notes']['blockers']}
  Overdue Critical : {scored['scores']['overdue']}/2 - {scored['notes']['overdue']}
  Sentiment        : {scored['scores']['sentiment']}/2 - {scored['notes']['sentiment']}
  FINAL SCORE      : {scored['final_score']:.3f} / 2.000
  PRELIMINARY RAG  : {scored['preliminary']}

INSTRUCTIONS:
1. Review the project context and pre-computed RAG score above.
2. You MAY override the preliminary RAG by one level (e.g. Amber to Red) ONLY if strong qualitative evidence clearly justifies it.
3. Output your response in EXACTLY this format (no extra text before or after, no markdown formatting):

RAG Status: [Red/Amber/Green]
Summary: [2-3 sentences of plain-English executive summary suitable for a VP. Mention the project name.]
Key Risks:
- [Risk 1]
- [Risk 2]
- [Risk 3]
Recommended Actions:
- [Action 1]
- [Action 2]
Data Quality Notes: [Any missing/messy data fields that affected the analysis, or write None.]
"""


def call_gemini(prompt: str, api_key: str) -> str:
    """Send prompt to Gemini and return the text response."""
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.3,
            max_output_tokens=2048,  # increased to prevent truncation
        ),
    )
    return response.text.strip()


# ─────────────────────────────────────────────
# PARSE LLM RESPONSE
# ─────────────────────────────────────────────

def parse_llm_response(text: str) -> dict:
    """Extract structured fields from LLM response text."""
    result = {
        "rag_status": "Unknown",
        "summary": "",
        "key_risks": [],
        "recommended_actions": [],
        "data_quality_notes": "",
        "raw": text,
    }

    # RAG Status
    m = re.search(r"RAG Status:\s*(Red|Amber|Green|Yellow)", text, re.IGNORECASE)
    if m:
        result["rag_status"] = m.group(1).capitalize()
        if result["rag_status"] == "Yellow":
            result["rag_status"] = "Amber"

    # Summary - capture until next section header
    m = re.search(r"Summary:\s*(.+?)(?=\n(?:Key Risks|Recommended Actions|Data Quality)|$)", text, re.DOTALL | re.IGNORECASE)
    if m:
        result["summary"] = m.group(1).strip()

    # Key Risks - capture bullet lines between Key Risks and Recommended Actions
    m = re.search(r"Key Risks:(.+?)(?=\nRecommended Actions:|\nData Quality|$)", text, re.DOTALL | re.IGNORECASE)
    if m:
        result["key_risks"] = [
            line.lstrip("- *•123456789.").strip()
            for line in m.group(1).strip().splitlines()
            if line.strip() and not line.strip().startswith("Key")
        ]
        result["key_risks"] = [r for r in result["key_risks"] if len(r) > 3]

    # Recommended Actions
    m = re.search(r"Recommended Actions:(.+?)(?=\nData Quality|$)", text, re.DOTALL | re.IGNORECASE)
    if m:
        result["recommended_actions"] = [
            line.lstrip("- *•123456789.").strip()
            for line in m.group(1).strip().splitlines()
            if line.strip() and not line.strip().startswith("Recommended")
        ]
        result["recommended_actions"] = [r for r in result["recommended_actions"] if len(r) > 3]

    # Data Quality Notes
    m = re.search(r"Data Quality Notes:\s*(.+)", text, re.DOTALL | re.IGNORECASE)
    if m:
        result["data_quality_notes"] = m.group(1).strip().split("\n")[0].strip()

    return result


# ─────────────────────────────────────────────
# SAVE REPORT
# ─────────────────────────────────────────────

RAG_EMOJI = {"Green": "🟢", "Amber": "🟡", "Red": "🔴", "Unknown": "⚪"}

def save_report(project_key: str, p: dict, scored: dict, parsed: dict, report_date: str) -> str:
    """Write the weekly health report to a Markdown file."""
    out_dir = os.path.join(OUTPUT_DIR, report_date)
    os.makedirs(out_dir, exist_ok=True)
    filepath = os.path.join(out_dir, f"{project_key}_health_report.md")

    rag    = parsed["rag_status"]
    emoji  = RAG_EMOJI.get(rag, "⚪")
    pre    = scored["preliminary"]
    overridden = " *(LLM Override)*" if rag != pre else ""

    lines = [
        f"# Project Health Report — {project_key.replace('_', ' ')}",
        f"> **Report Date:** {report_date}  |  **Project Manager:** {p['project_manager']}",
        "",
        f"## {emoji} RAG Status: **{rag}**{overridden}",
        "",
        "---",
        "",
        "## Executive Summary",
        "",
        parsed["summary"],
        "",
        "---",
        "",
        "## Scoring Breakdown",
        "",
        f"| Dimension | Score | Weight | Note |",
        f"|-----------|-------|--------|------|",
        f"| Schedule Health  | {scored['scores']['schedule_health']}/2 | 30% | {scored['notes']['schedule_health']} |",
        f"| Task Completion  | {scored['scores']['completion']}/2 | 25% | {scored['notes']['completion']} |",
        f"| Blocker Exposure | {scored['scores']['blockers']}/2 | 20% | {scored['notes']['blockers']} |",
        f"| Overdue Critical | {scored['scores']['overdue']}/2 | 15% | {scored['notes']['overdue']} |",
        f"| Sentiment        | {scored['scores']['sentiment']}/2 | 10% | {scored['notes']['sentiment']} |",
        f"| **TOTAL**        | **{scored['final_score']:.3f}/2.000** | 100% | Preliminary: **{pre}** → Final: **{rag}** |",
        "",
        "---",
        "",
        "## Project Snapshot",
        "",
        f"| Field | Value |",
        f"|-------|-------|",
        f"| Phase | {p['project_stage']} |",
        f"| Timeline | {p['project_start']} → {p['project_end']} |",
        f"| Actual % Complete | {p['pct_complete']:.1f}% |",
        f"| Expected % Complete | {scored.get('expected_pct', 'N/A')}% |",
        f"| Progress Gap | {scored.get('gap', 'N/A')}% behind |",
        f"| Tasks Completed | {p['task_counts']['completed']} / {p['task_counts']['total']} |",
        f"| Tasks On Hold | {p['task_counts']['on_hold']} |",
        f"| At Risk Level | {p['at_risk']} |",
        "",
        "---",
        "",
        "## Key Risks",
        "",
    ]
    for risk in parsed["key_risks"]:
        lines.append(f"- ⚠️ {risk}")
    lines += [
        "",
        "---",
        "",
        "## Recommended Actions",
        "",
    ]
    for action in parsed["recommended_actions"]:
        lines.append(f"- ✅ {action}")
    lines += [
        "",
        "---",
        "",
        "## Overdue Critical Tasks",
        "",
    ]
    if p["overdue_critical"]:
        for t in p["overdue_critical"]:
            lines.append(f"- 🚨 **{t['task']}** — Due: {t['due']}")
    else:
        lines.append("- ✅ No overdue critical tasks.")
    lines += [
        "",
        "---",
        "",
        "## Data Quality Notes",
        "",
        parsed.get("data_quality_notes", "None"),
        "",
        "---",
        "",
        "*Generated by Project Health Reporting Agent | Zycus Professional Services*",
    ]

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    # Also save raw JSON for synthesis pipeline
    json_path = os.path.join(out_dir, f"{project_key}_health_report.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({
            "project_key":   project_key,
            "project_name":  p.get("project_name", project_key),
            "report_date":   report_date,
            "rag_status":    rag,
            "final_score":   scored["final_score"],
            "preliminary":   pre,
            "pct_complete":  p["pct_complete"],
            "expected_pct":  scored.get("expected_pct"),
            "gap":           scored.get("gap"),
            "at_risk":       p["at_risk"],
            "project_stage": p["project_stage"],
            "task_counts":   p["task_counts"],
            "overdue_critical": p["overdue_critical"],
            "on_hold_tasks": p["on_hold_tasks"],
            "summary":       parsed["summary"],
            "key_risks":     parsed["key_risks"],
            "recommended_actions": parsed["recommended_actions"],
            "scores":        scored["scores"],
            "notes":         scored["notes"],
        }, f, indent=2)

    return filepath


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("❌ ERROR: GEMINI_API_KEY environment variable not set.")
        print("   Run: set GEMINI_API_KEY=your_key_here")
        return

    report_date = TODAY_OVERRIDE or date.today().strftime("%Y-%m-%d")
    sep = "=" * 60
    print(f"\n{sep}")
    print(f"  Project Health Reporting Agent")
    print(f"  Report Date: {report_date}")
    print(f"{sep}\n")

    for project_key, filepath in PROJECT_FILES.items():
        if not os.path.exists(filepath):
            print(f"[!] File not found: {filepath} - skipping.\n")
            continue

        print(f"[+] Processing: {filepath}")

        # 1. Load data
        print("   -> Loading and extracting project data...")
        p = load_project(filepath)
        print(f"   -> Project: {p['project_name']} | PM: {p['project_manager']}")
        print(f"   -> Progress: {p['pct_complete']:.1f}% complete | Stage: {p['project_stage']}")

        # 2. Compute score
        print("   -> Computing RAG score...")
        scored = compute_rag_score(p)
        print(f"   -> Pre-score: {scored['final_score']:.3f}/2.0 -> Preliminary: {scored['preliminary']}")

        # 3. Call LLM
        print("   -> Calling Gemini LLM for reasoning...")
        prompt   = build_prompt(project_key, p, scored)
        llm_text = call_gemini(prompt, api_key)
        parsed   = parse_llm_response(llm_text)
        print(f"   -> Final RAG: {parsed['rag_status']}")

        # 4. Save report
        out_path = save_report(project_key, p, scored, parsed, report_date)
        print(f"   -> Report saved: {out_path}\n")

    print(f"{sep}")
    print(f"  [OK] All reports generated in: {OUTPUT_DIR}/{report_date}/")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
