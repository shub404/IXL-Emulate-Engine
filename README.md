# IXL Emulate Engine

> **IIT Madras Summer Research Internship Project**
>
> A Playwright-based scraper that extracts math questions, answer choices, diagrams, correct answers, and explanations from IXL. Used to collect curriculum data across Grades 1–5 and stores the results in Excel and Google Drive for educational research.

---

## Overview

This project was built during a summer research internship at **IIT Madras**. The objective was to collect a structured dataset of math questions from IXL's Grades 1–5 curriculum - including question text, answer options, correct answers, visual diagrams, and step-by-step explanations - to support later research on teacher-style spoken narration of mathematics.

IXL is a dynamically rendered SPA that actively resists automated access. The main technical challenge was extracting content that varies significantly in format - some questions are plain text, some involve fractions or vertical arithmetic, and many include visual diagrams (number lines, fraction bars, pie charts, coordinate planes, etc.) that cannot be captured through text extraction alone.

The scraper handles all of this in a single file and saves results to:

- An **Excel workbook** with one row per question (text, options, correct answer, Drive links)
- **PNG screenshots** of diagrams, uploaded to **Google Drive**
- **Explanation text files** (.txt) for each question's REMEMBER and SOLVE walkthrough tabs

The collected dataset was intended to support later work on converting math questions into teacher-style narrations for Text-to-Speech model training.

---

## Output

```
IXL.com
    │
    ▼
scraper.py
    │
    ├──────────────────────┬─────────────────────────┐
    ▼                      ▼                         ▼
Excel Workbook       Diagram PNGs             Explanation Text
(questions, options, (number lines,           (.txt files with
 correct answers)     pie charts, etc.)        REMEMBER + SOLVE
                      → Google Drive)          → Google Drive)
    │
    ▼
Structured dataset for educational research
```

---

## Challenges Solved

- **Dynamically rendered content** - IXL loads questions via JavaScript without page refreshes; a DOM stabilization check polls until the new question is confirmed loaded before extraction starts
- **Math-specific text formats** - fractions, vertical arithmetic, fill-in-the-blank inputs, and Unicode symbols (×, ÷, –) all needed special handling in the text extractor
- **Visual diagrams alongside text** - 50+ distinct diagram types (number lines, pie charts, area models, Venn diagrams, cube trains, calendars, etc.) needed to be detected and screenshotted separately from the question text
- **Pre-rendered phantom elements** - IXL renders upcoming questions below the viewport; the scraper had to identify and extract only the topmost (currently active) question
- **Multiple question formats** - each extraction function (question text, options, correct answer) has 3–4 independent methods tried in order, so an unexpected format falls through to the next method rather than failing silently
- **Session integrity** - navigating via breadcrumbs instead of direct URLs preserves the session state and avoids anti-bot redirects
- **Concurrent uploads without blocking** - diagram screenshots are uploaded to Google Drive in background threads so scraping continues without waiting for each upload
- **Resumable scraping** - if a run is interrupted, Mode 2 lets you resume from any specific skill URL without re-scraping already-captured data
- **Post-answer explanation extraction** - after submitting a blank answer to reveal the correct answer, the scraper switches to the explanation tab and captures both text and any inline diagrams

---

## Dataset Collected


| Metric | Count |
|---|---|
| Grades covered | 1–5 |
| Categories scraped | 200+ |
| Skills scraped | 2000+ |
| Questions collected | 6000+ |
| Diagram screenshots generated | 10000+ |
| Explanation files generated | 6000+ |

---

## Features

- **Playwright-Based Website Interaction** - runs a visible Chromium browser, navigates by clicking UI elements and breadcrumbs, uses a real user-agent string
- **Custom Text Extraction Logic** - a custom JavaScript DOM walker handles fractions (`.vFrac`, `.old-fraction-in-text`), vertical arithmetic, fill-in-the-blank inputs, and Unicode math symbols
- **Diagram Detection and Screenshot Capture** - detects and screenshots 50+ diagram types: number lines, fraction bars, pie charts, ten frames, area models, cube trains, Venn diagrams, coordinate planes, calendars, and more
- **Multiple Extraction Methods** - each of question text, options, and correct answer has 3–4 independent extraction strategies tried in sequence
- **Background Uploads to Google Drive** - screenshots are uploaded in background threads (5-worker pool) without pausing question extraction
- **Resume Mode** - supports resuming a partial scrape from any skill URL without re-scraping already-captured data
- **Explanation Extraction** - captures REMEMBER and SOLVE tabs from post-answer explanations, with inline `[image: url]` references where diagrams appear in the explanation
- **Excel with Hyperlinks** - output workbook has clickable Google Drive folder links in the diagram columns

---

## How It Works

```
python scraper.py
  │
  ├─ Choose mode: full scrape or resume from a specific skill URL
  ├─ Google OAuth2 → connect to Drive
  ├─ Create output directories and Excel file
  │
  └─ Chromium browser (visible window)
       ├─ Log in to IXL
       ├─ Navigate to the target grade's math index
       │
       └─ For each category → for each skill:
            ├─ Click the skill link
            └─ For each of 3 questions:
                 ├─ Wait until the new question has loaded
                 ├─ Extract question text
                 ├─ Extract answer options
                 ├─ Screenshot any diagrams in the question
                 ├─ Submit a blank answer to reveal the correct answer
                 ├─ Extract the correct answer
                 ├─ Screenshot any diagrams in the answer
                 ├─ Extract the explanation (REMEMBER + SOLVE tabs)
                 └─ Append everything as a row in the Excel file
            │
            └─ Click the breadcrumb to return to the grade index
```

### Key Functions

| Function | What it does |
|---|---|
| `run_scraper()` | Starts the scraper, handles login and the category/skill loops |
| `extract_and_advance()` | Runs the extraction sequence for each question within a skill |
| `extract_question_text()` | Extracts the question string; tries 3 methods in order |
| `extract_options()` | Extracts answer choices; handles 4 different option formats |
| `extract_correct_answer()` | Extracts the correct answer after submission; tries 4 methods |
| `extract_diagrams_screenshots()` | Finds and screenshots diagrams in the question and option tiles |
| `_extract_from_scope()` | Low-level screenshot function that deduplicates overlapping captures |
| `scrape_explanation()` | Extracts the REMEMBER and SOLVE explanation tabs |
| `append_to_excel()` | Writes one question's data as a row in the Excel workbook |
| `_MATH_WALKER_JS` | JavaScript string that walks the DOM to extract math-formatted text |

---

## Output Schema

### Excel Workbook

| Column | Header | Description |
|---|---|---|
| A | Ques ID | Unique question ID, increments across runs |
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

| Technology | Purpose |
|---|---|
| Python 3.x | Main language |
| [Playwright](https://playwright.dev/python/) | Browser automation and DOM interaction |
| [openpyxl](https://openpyxl.readthedocs.io/) | Excel file creation and row appending |
| Google Drive API v3 | Cloud storage for screenshots and explanation files |
| google-auth-oauthlib | OAuth2 desktop flow for Drive access |
| concurrent.futures (stdlib) | Background thread pool for Drive uploads |

---

## Prerequisites

- Python 3.8+
- A valid IXL account (email + password)
- A Google Cloud project with **Drive API** enabled and a downloaded `client_secret.json` (OAuth2 Desktop credentials)
- Three Google Drive folders created in advance (for diagrams, explanation text, and explanation images) - copy their folder IDs

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
QUESTIONS_PER_SKILL = 3
TARGET_URL = "https://www.ixl.com/math/grade-5"  # change to target grade

# For Mode 2 resume: set this to the skill URL to resume from
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
[1] Full scrape - start from the beginning of the grade
[2] Resume     - skip to START_URL and continue from there
```

The scraper runs in a visible Chromium window. Do not close the browser while it is running. Progress is printed to the terminal as each skill completes.

---

## Project Structure

```
IXL-Emulate-Engine/
├── scraper.py                      # Main scraping logic
├── client_secret.json              # Google OAuth2 credentials (not committed)
├── token.json                      # Cached OAuth token (auto-generated)
├── ixl_grade5_questions[test].xlsx # Output Excel file (auto-generated)
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
| **Different grade** | `TARGET_URL`, `EXCEL_FILENAME`, sheet title in `init_excel()` |
| **More questions per skill** | `QUESTIONS_PER_SKILL` constant |
| **New diagram type** | Add CSS selector to `DIAGRAM_SIGNALS` and to `_extract_from_scope()`; add class name to `_DIAGRAM_CLASSES` in both JS walker strings |
| **New option format** | Add a new fallback block in `extract_options()` |
| **Different Drive folders** | Update the three `*_DRIVE_FOLDER_ID` constants |
| **Resume from specific skill** | Set `START_URL` and run Mode 2 |

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---|---|---|
| Empty question text | A new diagram class is absorbing the text node | Check `_DIAGRAM_CLASSES` inside `_MATH_WALKER_JS` |
| Missing diagram screenshots | New diagram type not in the signal list | Add its CSS selector to `DIAGRAM_SIGNALS` and `_extract_from_scope()` |
| Duplicate screenshots | Overlap threshold is too low | Adjust the threshold in `_boxes_overlap()` |
| `PermissionError` on Excel | File is open in Excel | Close the Excel file; the script will retry automatically |
| Drive auth failure | `token.json` expired or missing | Delete `token.json` and re-run to re-authorize |
| Session killed mid-scrape | IXL detected automated access | Increase the `page.wait_for_timeout` delays in `extract_and_advance()` |
| Wrong question extracted | Active question detection failing | Check the Y-coordinate sort logic in `_get_active_question()` |

---

## Research Context

This project was completed as part of a **Summer Research Internship at IIT Madras**. The dataset collected by this scraper was intended to support later experiments involving teacher-style spoken narration of mathematics questions - where each question would be converted into natural speech and used to train or evaluate a Text-to-Speech model.

The scraper covers the full IXL Grades 1–5 Math curriculum across all categories and skills, collecting 3 question samples per skill.

---

## Acknowledgements

- Developed during Summer Research Internship at **IIT Madras**
- Built with [Microsoft Playwright](https://playwright.dev/python/) for browser automation
- Data sourced from [IXL Learning](https://www.ixl.com) (Grades 1–5 Math curriculum)
