# Project Health Reporting Agent

An AI-powered system that automatically reads project plan Excel files, computes a weighted **RAG (Red/Amber/Green)** health status, generates plain-English reports using Gemini LLM, and synthesizes a monthly executive PowerPoint presentation for VP-level stakeholders.

Built as part of the **Zycus AI Engineer Internship Assignment**.

---

## What It Does

| Phase | What Happens |
|-------|-------------|
| **Weekly** | Agent reads Excel project plans → computes RAG score → calls Gemini LLM → saves a Markdown health report |
| **Monthly** | Synthesis script reads all weekly reports → finds cross-project trends → generates a 7-slide PowerPoint deck |
| **Scheduled** | Scheduler runs the agent automatically every Monday at 9:00 AM |

---

## Project Structure

```
zycus/
├── Project Plan B.xlsx          ← Input: UniSan S2P project data
├── S2P Project.xlsx             ← Input: Outokumpu S2P project data
├── agent.py                     ← Main AI agent (weekly reports)
├── synthesis.py                 ← Monthly synthesis + PPTX generator
├── scheduler.py                 ← Weekly auto-runner (Bonus)
├── rag_methodology.md           ← Phase 1: RAG framework document
├── .env                         ← Your API key (never committed)
├── .env.example                 ← Template for .env
├── README.md                    ← This file
└── outputs/
    ├── 2026-07-08/
    │   ├── UniSan_S2P_health_report.md
    │   ├── UniSan_S2P_health_report.json
    │   ├── Outokumpu_S2P_health_report.md
    │   └── Outokumpu_S2P_health_report.json
    └── monthly_presentation.pptx
```

---

## Setup

### 1. Clone the repository
```bash
git clone https://github.com/AbhiramYad/project-health-rag-agent.git
cd project-health-rag-agent
```

### 2. Install dependencies
```bash
pip install pandas openpyxl google-genai python-dotenv schedule python-pptx
```

### 3. Set your API key
Get a free Gemini API key from [aistudio.google.com](https://aistudio.google.com).

Create a `.env` file in the project root:
```
GEMINI_API_KEY=your_gemini_api_key_here
```

---

## Usage

### Run the Weekly Health Agent
```bash
python agent.py
```
Generates `outputs/YYYY-MM-DD/<project>_health_report.md` and `.json` for each project.

### Generate the Monthly Executive Presentation
```bash
python synthesis.py
```
Generates `outputs/monthly_presentation.pptx` — a 7-slide branded PowerPoint.

### Run the Weekly Scheduler (Bonus)
```bash
# Runs every Monday at 9:00 AM automatically
python scheduler.py

# Test immediately without waiting
python scheduler.py --now
```

---

## RAG Methodology

The agent scores each project across **5 weighted dimensions**:

| Dimension | Weight | What it measures |
|-----------|--------|-----------------|
| Schedule Health | 30% | Red/Yellow/Green from project system |
| Task Completion Rate | 25% | Actual % vs expected % based on elapsed time |
| Blocker & Risk Exposure | 20% | On-hold tasks and At Risk level |
| Overdue Critical Tasks | 15% | Tasks past due date on the critical path |
| Qualitative Sentiment | 10% | Keyword analysis of status comments |

**Score → RAG mapping:**
- `1.50 – 2.00` → 🟢 **Green** (On Track)
- `0.80 – 1.49` → 🟡 **Amber** (At Risk)
- `0.00 – 0.79` → 🔴 **Red** (Critical)

The Gemini LLM then validates this score against qualitative context and may override by one level if strong evidence justifies it.

Full methodology: [`rag_methodology.md`](rag_methodology.md)

---

## Sample Outputs (July 2026)

### UniSan S2P (Project Plan B.xlsx)
- **RAG: 🔴 Red**
- PM: Rajat Bothra | Stage: Training Phase I
- 44% complete vs 58.8% expected — 14.8% behind schedule
- At Risk: High | 5 non-critical tasks overdue

### Outokumpu S2P (S2P Project.xlsx)
- **RAG: 🔴 Red** *(LLM override from Amber — qualitative risks)*
- PM: Aftab Hashambhai | Stage: Configuration and Build Phase
- 71% complete | Parallel phase delays, pending client data (JDE mapping)

---

## Executive Presentation

The monthly presentation (`monthly_presentation.pptx`) contains 7 slides:

| # | Slide | Contents |
|---|-------|---------|
| 1 | Cover | Title, month, portfolio RAG summary counts |
| 2 | Executive Summary | RAG dashboard card for each project |
| 3 | Trends | Cross-project delivery patterns |
| 4 | Emerging Risks | Systemic risks requiring leadership attention |
| 5 | Project Deep-Dives | Per-project snapshot with risks |
| 6 | Recommendations | 5 strategic actions for leadership |
| 7 | Appendix | Full weekly RAG log table |

---

## Design Decisions

### Why no Vector Database?
The term "RAG" in this project means **Red/Amber/Green** (a project management status system) — not Retrieval-Augmented Generation. The project data fits entirely within Gemini's context window, so no vector store is needed.

### Why Gemini instead of GPT-4?
Gemini offers a **free tier** via Google AI Studio with generous rate limits, requires no payment setup, and the SDK was already available in the Python environment.

### Why pure Python scripts instead of a web framework?
The assignment asks for a "working agent runnable on a weekly schedule" — not a web application. Pure scripts are the cleanest, most portable answer and avoid unnecessary complexity.

### Why `python-pptx` for slides?
Generates native, **fully editable** `.pptx` files that a VP can open directly in PowerPoint or Google Slides and modify with minimal effort — exactly what the assignment asks for.

### Messy Data Handling
Real project plans contain formula errors (`#UNPARSEABLE`), missing fields, and inconsistent formats. The agent handles all of these gracefully with safe parsers and fallback logic detailed in [`rag_methodology.md`](rag_methodology.md).

---

## Tech Stack

| Component | Tool | License |
|-----------|------|---------|
| Language | Python 3.x | PSF (Free) |
| Excel Parsing | `pandas` + `openpyxl` | BSD/MIT (Free) |
| LLM | Google Gemini (`google-genai`) | Free tier |
| Slide Generation | `python-pptx` | MIT (Free) |
| Scheduling | `schedule` | MIT (Free) |
| Env Management | `python-dotenv` | BSD (Free) |

---

*Built by Abhiram Yadav M — Zycus AI Engineer Intern Assignment*
