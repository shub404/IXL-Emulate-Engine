# IXL Emulate Engine

> **IIT Madras Summer Research Internship Project**
>
> A resilient browser automation engine that systematically extracts Grade 5 mathematics curriculum content from IXL.com — capturing question text, answer options, visual diagrams, and step-by-step explanations into a structured dataset for downstream AI and TTS model training.

---

## Overview

This project was developed during a summer research internship at **IIT Madras**. The goal: build a structured, AI-consumable dataset of math questions that can be transformed into natural-language teacher narrations to train **Text-to-Speech (TTS) models** — enabling a system that dictates math problems exactly as a teacher would to students in a classroom.

The scraper navigates every skill under the IXL Grade 5 Math curriculum, extracts 3 questions per skill, and produces:

- Structured question data (text + options + correct answers) in an **Excel workbook**
- Visual diagrams (number lines, fraction bars, pie charts, coordinate planes, etc.) as **PNG screenshots**
- Step-by-step **explanation text** files
- All assets synced to **Google Drive**

---

## Downstream Pipeline

```
IXL Platform
     │
     ▼
IXL Emulate Engine  ──────────────────────────────────────────────────
     │                                                                │
     ▼                                                                ▼
Excel Database                                            Google Drive
(question text,                                     (diagram PNGs + explanation
 options, answers)                                   text files per question)
     │
     ▼
Prompt Engineering Layer
(convert structured Q&A into natural teacher-speech prompts)
     │
     ▼
TTS Model Training
(train model to dictate math problems as a teacher would to students)
```

---

## Key Features

- **Anti-Bot Human Emulation** — runs a visible Chromium browser, navigates via UI clicks and breadcrumbs, uses real user-agent strings; never touches the API
- **JavaScript DOM Walker** — custom JS text extractor handles fractions (`.vFrac`, `.old-fraction-in-text`), vertical arithmetic, fill-in-the-blank inputs, and Unicode math symbols
- **Multi-Type Diagram Capture** — detects and screenshots 50+ diagram signal types: number lines, fraction bars, pie charts, ten frames, area models, cube trains, Venn diagrams, coordinate planes, calendars, and more
- **Four-Fallback Extraction** — each of question text, options, and correct answer has 3–4 independent extraction strategies; the pipeline degrades gracefully instead of crashing
- **Async Drive Uploads** — diagram screenshots are uploaded in background threads (5-worker pool) without blocking question progression
- **Resume Mode** — supports resuming a partial scrape from any skill URL without re-scraping already-captured data
- **Explanation Extraction** — captures REMEMBER and SOLVE tabs from post-answer explanations, with inline `[image: url]` references for embedded diagrams
- **Excel with Hyperlinks** — output workbook has clickable Drive folder links for diagram columns

---

## Architecture

```
python scraper.py
  │
  ├─ _choose_mode()              → Mode 1 (full) or Mode 2 (resume from START_URL)
  ├─ _init_drive_service()       → Google OAuth2 → global Drive client
  ├─ setup_dir()                 → create ixl_diagrams/ and ixl_explanations/
  ├─ init_excel()                → create 11-column Excel workbook (idempotent)
  ├─ _read_max_question_id()     → resume: find highest existing Ques ID
  │
  └─ Playwright Chromium session (headless=False)
       ├─ Login to IXL
       ├─ Navigate to Grade 5 Math index
       │
       └─ for each category → for each skill:
            └─ extract_and_advance()          ← core loop, 3 questions/skill
                 ├─ Poll for DOM stabilization (new question loaded)
                 ├─ extract_question_text()   → JS walker + normalize Unicode
                 ├─ extract_options()         → 4-type option extraction
                 ├─ extract_diagrams_screenshots()
                 │    └─ _extract_from_scope()   → screenshot + Drive upload
                 ├─ Submit blank answer to reveal correct answer
                 ├─ extract_correct_answer()  → 4-strategy extraction
                 ├─ _screenshot_answer_diagrams()
                 ├─ scrape_explanation()      → .txt file + inline image refs
                 ├─ append_to_excel()         → row with hyperlinked Drive URLs
                 └─ Click "Got it" → advance to next question
            │
            └─ Breadcrumb click → back to Grade 5 index
```

### Core Components

| Component | File Location | Role |
|---|---|---|
| `run_scraper()` | line 1937 | Entry point — full orchestration |
| `extract_and_advance()` | line 1827 | Per-skill question loop |
| `_extract_from_scope()` | line 1383 | Core screenshot capture engine |
| `extract_question_text()` | line 746 | Text extraction with 3 fallbacks |
| `_MATH_WALKER_JS` | line 279 | JavaScript DOM tree walker for math text |
| `_EXPL_WALKER_JS` | line 463 | DOM walker for explanation text (with `@@IMG:N@@` markers) |
| `DIAGRAM_SIGNALS` | line 46 | 50+ CSS selectors that indicate a visual diagram |

---

## Output Schema

### Excel Workbook (`ixl_grade5_questions[test].xlsx`)

| Column | Header | Description |
|---|---|---|
| A | Ques ID | Globally unique, ever-increasing integer |
| B | # | Skill serial number |
| C | Category | Curriculum category name |
| D | Skill Name | IXL skill name |
| E | Question No | 1, 2, or 3 within the skill |
| F | Question Text | Extracted and normalized question string |
| G | Question Diagram | Hyperlink → Google Drive folder of question PNGs |
| H | Question Options | Newline-separated answer option strings |
| I | Option Diagrams | Hyperlink → Google Drive folder of option PNGs |
| J | Correct Answer | Extracted correct answer string |
| K | Answer Diagram | Hyperlink → Google Drive folder of answer PNGs |

### Google Drive Folder Layout

```
DRIVE_FOLDER_ID/
├── {skill}_q{n}_qdiag_{ts}/          ← question diagrams
├── {skill}_q{n}_optdiag_{ts}/        ← option diagrams
└── {skill}_q{n}_ansdiag_{ts}/        ← answer diagrams

EXPL_TXT_DRIVE_FOLDER_ID/
└── {skill}_explanation_q{n}.txt      ← explanation text (REMEMBER + SOLVE tabs)

EXPL_IMG_DRIVE_FOLDER_ID/
└── {skill}_q{n}_explimg_{ts}/
    ├── {skill}_q{n}_remember_img{n}.png
    └── {skill}_q{n}_solve_img{n}.png
```

---

## Technology Stack

| Technology | Version | Purpose |
|---|---|---|
| Python | 3.x | Orchestration language |
| [Playwright](https://playwright.dev/python/) | latest | Browser automation and DOM interaction |
| [openpyxl](https://openpyxl.readthedocs.io/) | latest | Excel file creation and row appending |
| Google Drive API v3 | — | Cloud storage for screenshots and text files |
| google-auth-oauthlib | latest | OAuth2 desktop flow for Drive access |
| concurrent.futures | stdlib | Thread pool for async Drive uploads |

---

## Prerequisites

- Python 3.8+
- A valid IXL account (email + password)
- A Google Cloud project with **Drive API** enabled and a downloaded `client_secret.json` (OAuth2 Desktop credentials)
- Three Google Drive folders created in advance (for diagrams, explanation text, and explanation images) — note their folder IDs

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/shub404/IXL-Emulate-Engine.git
cd IXL-Emulate-Engine

# 2. Install Python dependencies
pip install playwright openpyxl google-auth-oauthlib google-api-python-client

# 3. Install Playwright's Chromium browser
playwright install chromium

# 4. Place your Google OAuth credentials file
cp /path/to/your/client_secret.json .
```

---

## Configuration

Open `scraper.py` and update the constants at the top of the file:

```python
# IXL credentials
EMAIL    = "your_email@example.com"
PASSWORD = "your_password"

# Google Drive folder IDs (from your Drive)
DRIVE_FOLDER_ID          = "your_diagrams_root_folder_id"
EXPL_TXT_DRIVE_FOLDER_ID = "your_explanation_text_folder_id"
EXPL_IMG_DRIVE_FOLDER_ID = "your_explanation_images_folder_id"

# Scraping parameters
QUESTIONS_PER_SKILL = 3          # questions to extract per skill
TARGET_URL = "https://www.ixl.com/math/grade-5"  # change grade here

# For Mode 2 resume: set this to the skill URL you want to resume from
START_URL = "https://www.ixl.com/math/grade-5/some-skill-name"
```

---

## Usage

```bash
python scraper.py
```

On first run, a browser window will open to complete Google OAuth2 authorization. A `token.json` file is saved and auto-refreshed on subsequent runs.

**Select a mode when prompted:**

```
[1] Full scrape — start from the beginning of Grade 5 Math
[2] Resume     — skip to START_URL and continue from there
```

The scraper runs visibly in a Chromium window. Do not close the browser while it is running. Progress is printed to the terminal as each skill completes.

---

## Project Structure

```
IXL-Emulate-Engine/
├── scraper.py                      # Entire codebase (2078 lines, single file)
├── client_secret.json              # Google OAuth2 credentials (not committed)
├── token.json                      # Cached OAuth token (auto-generated)
├── ixl_grade5_questions[test].xlsx # Output Excel database (auto-generated)
├── ixl_diagrams/                   # Local PNG screenshot cache
│   └── {skill}_q{n}_{type}_{ts}.png
├── ixl_explanations/               # Local explanation .txt cache
│   └── {skill}_explanation_q{n}.txt
└── docs/
    ├── ARCHITECTURE_SUMMARY.md
    ├── PROJECT_KNOWLEDGE_MAP.md
    └── FUNCTION_REGISTRY.json
```

---

## Extending the Scraper

| Goal | What to Change |
|---|---|
| **Different grade** | `TARGET_URL`, `EXCEL_FILENAME`, `ws.title` in `init_excel()` |
| **More questions per skill** | `QUESTIONS_PER_SKILL` constant (line 25) |
| **New diagram type** | Add CSS selector to `DIAGRAM_SIGNALS` (line 46) and `L1_INTEGRATED`/`L1_MULTI` in `_extract_from_scope()` (line 1527/1570); add class to `_DIAGRAM_CLASSES` in both JS walkers |
| **New option type** | Add fallback block in `extract_options()` (line 827) |
| **Different Drive folders** | Update the three `*_DRIVE_FOLDER_ID` constants |
| **Resume from specific skill** | Set `START_URL` (line 37) and run Mode 2 |

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---|---|---|
| Empty question text | New diagram class absorbing text node | Check `_DIAGRAM_CLASSES` in `_MATH_WALKER_JS` (line 281) |
| Missing diagram screenshots | New diagram type not in signal list | Add CSS selector to `DIAGRAM_SIGNALS` and `L1_INTEGRATED` |
| Duplicate screenshots | Overlap threshold too low | Adjust threshold in `_boxes_overlap()` (line 1095) |
| `PermissionError` on Excel | File open in Excel | Close the Excel file; the script will retry automatically |
| Drive auth failure | `token.json` expired or missing | Delete `token.json` and re-run to re-authorize |
| Session killed mid-scrape | IXL anti-bot detection triggered | Increase `page.wait_for_timeout` delays in `extract_and_advance()` |
| Wrong question extracted | Active question detection failing | Check `_get_active_question()` Y-coordinate sort (line 724) |

---

## Research Context

This project was completed as part of a **Summer Research Internship at IIT Madras**. The dataset produced by this scraper serves as the raw input for a downstream pipeline that:

1. Processes structured Q&A data from the Excel output
2. Converts each question into natural-language teacher narrations (prompt engineering)
3. Feeds those narrations into a **TTS model training pipeline** — the goal being a system that speaks math problems and explanations exactly as a teacher would dictate them to students

The scraped content covers the full Grade 5 IXL Math curriculum across all categories and skills, with 3 question samples per skill.

---

## Acknowledgements

- Developed during Summer Research Internship at **IIT Madras**
- Built on top of [Microsoft Playwright](https://playwright.dev/python/) for browser automation
- Data sourced from [IXL Learning](https://www.ixl.com) (Grade 1-5 Math curriculum)
