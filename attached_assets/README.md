# Curriculum Automation Script

Automates curriculum generation and QA for a list of skills by driving ChatGPT and Google Sheets through a real Chrome window via Playwright.

---

## What It Does

1. Reads skill names from a Google Sheet (skips rows already marked **Completed**)
2. Opens a new ChatGPT conversation in the **Plan** project and generates a curriculum table
3. Opens a new ChatGPT conversation in the **QA** project, runs a QA report, and scores it
4. If the score is below 98, asks the Plan project to fix the curriculum and re-runs QA (up to 2 retries)
5. Writes the approved curriculum rows to a **Curriculum** tab in the Google Sheet
6. Saves the Plan URL + QA URL to column D and sets column F to **Completed**

---

## Requirements

### Python version

**Python 3.10 or higher** (the script uses `int | None` union type syntax).

### pip packages

Install all three with one command:

```
pip install playwright pandas pyperclip
```

| Package | What it's used for |
|---|---|
| `playwright` | Controls Chrome via the Chrome DevTools Protocol (CDP) |
| `pandas` | Parses the TSV clipboard dump from Google Sheets |
| `pyperclip` | Reads and writes the system clipboard |

> `asyncio`, `re`, `io`, `time`, `socket`, `subprocess`, `sys`, `os` are all part of the Python standard library — no install needed.

### Playwright browser driver

After installing the `playwright` package, run this **once** to download the Chromium driver:

```
playwright install chromium
```

### Google Chrome (real browser)

The script attaches to your **real Chrome installation** (not Playwright's bundled Chromium) so it can use your saved login sessions for ChatGPT and Google Sheets.

- Chrome must be installed at `C:\Program Files\Google\Chrome\Application\chrome.exe`
- If yours is in a different location, update `CHROME_EXE` at the top of the script

---

## First-Time Setup

### 1 — Create a separate Chrome profile for the script

The script uses a dedicated Chrome profile so it never touches your main browser data.

```
CHROME_PROFILE_PATH = r"C:\Users\YourName\AppData\Local\Google\ChromeDebugSession"
```

Change `YourName` to your actual Windows username. Chrome will create this folder automatically on first run.

> **Do not point this at your real Default profile.** Chrome 120+ blocks remote debugging on the main profile.

### 2 — Edit the configuration block

Open the script and fill in the top section:

```python
SHEET_URL          = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/..."
PLAN_PROJECT_URL   = "https://chatgpt.com/g/g-p-XXXX/project"
QA_PROJECT_URL     = "https://chatgpt.com/g/g-p-YYYY/project"
CHROME_PROFILE_PATH = r"C:\Users\YourName\AppData\Local\Google\ChromeDebugSession"
CHROME_EXE         = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
```

### 3 — Log in on first run

On the very first run, the script detects that the profile folder does not exist and automatically opens two tabs:

- **Tab 1** — ChatGPT (log in with your account)
- **Tab 2** — Google Sheet (log in / grant access)

Once you are logged into both, press **Enter** in the terminal. Your sessions are saved in the profile folder from that point on — you will not need to log in again.

---

## Running the Script

```
python curriculum_automation.py
```

Press **Enter** when prompted to launch Chrome. The script runs fully automatically after that.

---

## Google Sheet Layout Expected

| Column | Content |
|---|---|
| A | (anything / row number) |
| B | **Skill name** — what the script reads |
| C | (anything) |
| D | ChatGPT links — written by the script |
| E | (anything) |
| F | **Status** — written by the script (`Completed`, `QA Failed`, etc.) |

Rows where column F already contains **Completed** or **Done** are skipped on re-runs.

The script creates a **Curriculum** tab automatically if it does not exist and appends rows to it on every run.

---

## Resuming After a Crash

Simply re-run the script. It re-reads the sheet, skips every row that already has **Completed** in column F, and picks up from where it left off. The Curriculum tab counter is also re-initialised by reading the existing row count so no data is overwritten.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `chrome.exe not found` | Update `CHROME_EXE` to match your actual Chrome path |
| Debug port never opens | Make sure all existing Chrome windows are closed before running |
| `Clipboard empty after 3 attempts` | Click on the Google Sheet in Chrome and try again — the sheet must have focus |
| `Could not find a visible ChatGPT input box` | ChatGPT may have changed its UI; check that you are logged in inside the debug profile |
| Script crashes on skill 2+ | Usually a timeout from ChatGPT being slow; the script marks the row as `Error:...` and continues automatically |
