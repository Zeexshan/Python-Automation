# =============================================================================
# CURRICULUM AUTOMATION SCRIPT v2 (FIXED)
#
# WORKFLOW PER SKILL:
#   1. Read skill names from Google Sheet (Status column = blank)
#   2. Generate curriculum via ChatGPT Plan project
#   3. Run QA via ChatGPT QA project (analyze → report)
#   4. If score < 98: fix curriculum → Re-QA (same conversations, max 2 retries)
#   5. If score >= 98: write curriculum rows to "Curriculum" tab
#   6. Save Plan URL + QA URL to Column D of Tracker
#   7. Update Status column
#
# SETUP (run once):
#   pip install playwright pandas pyperclip
#   playwright install chromium
# =============================================================================

import asyncio
import io
import re
import time
import socket
import subprocess
import sys
import os
import shutil
import pyperclip
import pandas as pd
from playwright.async_api import async_playwright

# =============================================================================
# CONFIGURATION — Edit before running
# =============================================================================

SHEET_URL = "https://docs.google.com/spreadsheets/d/18w9QnS6U-7scYa4RwYWpFBgr9PXh0_V8mL-XGezsZkU/edit?gid=0#gid=0"

PLAN_PROJECT_URL = (
    "https://chatgpt.com/g/g-p-6a33ebfaaab881919d7c821a7091f00e-structure-plan-tanveer/project"
    # "https://chatgpt.com/g/g-p-6a35212b0f888191ac467cce1d44a860/project"
)
QA_PROJECT_URL = (
    "https://chatgpt.com/g/g-p-6a33ebd7254881918d82fcaf9208e631-structure-qa-tanveer/project"
    # "https://chatgpt.com/g/g-p-6a3521435c74819184a0fdd3ee12a134/project"
)

# Your real Chrome profile (used as the SOURCE for the one-time copy below).
CHROME_USER_DATA_DIR = r"C:\Users\acer\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE_DIR   = "Profile 5"

# The script's dedicated Chrome profile directory.
# MUST differ from CHROME_USER_DATA_DIR — Chrome 120+ blocks remote debugging
# on the default User Data path.
# On first run: Profile 5 is COPIED here once, carrying all your logins.
# On every subsequent run: this folder is reused as-is (no re-copy).
# Result: Google's "verify it's you" check appears only on the very first run.
CHROME_DEBUG_DATA_DIR = r"C:\Users\acer\AppData\Local\Google\ChromeDebugSession"

# Path to chrome.exe — this is the standard location
CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Debugging port Playwright will connect to
CHROME_DEBUG_PORT = 9222

QA_PASS_THRESHOLD = 98   # Minimum score to accept curriculum
MAX_REQA_RETRIES  = 2    # How many fix+reQA attempts before giving up
DELAY_BETWEEN_SKILLS = 5  # Seconds between skills

# =============================================================================
# PROMPTS
# =============================================================================

PLAN_PROMPT = """\
Generate a finalized curriculum structure for the following course Skill :

{skill_name}

Follow these rules strictly:

1. Create the curriculum in this table format only:

| Chapter | Lesson | Concepts to be Covered |
| ------- | ------ | ---------------------- |

2. Use only one input: the course name/topic provided above. Do not ask for multiple inputs.
3. Keep chapter names short and clear. Do not write "Chapter 1," "Chapter 2," or numbering in the chapter column.
4. Keep lesson names short and clear. Do not write "Lesson 1," "Lesson 2," or numbering in the lesson column.
5. There should be exactly 4 chapters and one "End Course Assessment" chapter at the end. Each main chapter except end course must have a minimum of 3 lessons and a maximum of 5 lessons.
6. After the lessons of every main chapter, add one Skill Mastery Challenge row in this exact format:
   Skill Mastery Challenge: {{Chapter Name}}
7. For every Skill Mastery Challenge row, the "Concepts to be Covered" column should mention:
   Interview questions and answers related to that chapter.
8. Add one final chapter with the chapter name exactly as:
   End Course Assessment
9. Under End Course Assessment, add exactly these two lessons:
   End Course Quiz Assessment: {skill_name}
   End Course Summative Assessment: {skill_name}
10. For the End Course Quiz Assessment concepts, mention a full-course quiz covering all major chapters, tools, workflows, concepts, safety, troubleshooting, and mastery-level technical understanding.
11. For the End Course Summative Assessment concepts, mention a final project or practical assessment where learners apply the full course workflow end-to-end.
12. Add a "Concepts to be Covered" column for every lesson. Concepts should be specific, practical, technically deep, and aligned with the lesson name.
13. Generate the curriculum so that the skill is covered at mastery level. The curriculum should not stay at surface-level awareness. It should include deep technical foundations, practical implementation workflows, advanced usage patterns, debugging, optimization, production readiness, best practices, safety, troubleshooting, and real-world application.
14. Use web searches if required to make the curriculum current, accurate, and aligned with industry practices, official documentation, latest tools, and real-world workflows.
15. Do not add an activity column.
16. Do not add long descriptions outside the table.
17. Do not use optional wording like "or," "optional," or "recommended."
18. Make the curriculum progression logical from beginner to advanced.
19. Ensure the curriculum feels certification-ready, industry-relevant, and technically rigorous.
20. Output style: short chapter names, short lesson names, specific concepts, mastery-level depth, clean table only, no document canvas, no extra explanation after the table.

Now generate the curriculum."""

QA_ANALYZE_PROMPT = """\
I have uploaded the source document "Curriculum QA System - v1.0" in this source.

Before doing any curriculum QA work, analyze this source document carefully and treat it as the governing Curriculum QA standard for this chat.

Do not generate a QA report yet.

Your task is to:
1. Understand the Curriculum QA framework, scoring system, rubric, report format, review scope, re-QA rules, and issue-writing standards from the source document.
2. Confirm how you will apply the source document for future curriculum skeleton reviews in this chat.
3. Identify what is in scope for Curriculum QA.
4. Identify what is out of scope and should be left for Content QA.
5. Confirm how you will review chapters, lessons, concept coverage, skill progression, and mastery readiness.
6. Confirm how you will avoid penalizing items that are intentionally deferred for later review.
7. Confirm how you will handle re-QA after fixes.

Keep your response concise and operational. Do not rewrite or reproduce the whole source document."""

# QA_REPORT_PROMPT is split into header + footer so the curriculum text
# can be pasted as a SEPARATE clipboard operation (avoids large single-paste failures).
QA_REPORT_PROMPT_HEADER = """\
Using the attached source document "Curriculum QA System - v1.0" as the governing Curriculum QA standard, generate a Curriculum QA report for the following course skeleton:
Short Course Name : {skill_name}
Curriculum :
\"\"\""""

QA_REPORT_PROMPT_FOOTER = """\
\"\"\"
Review the provided curriculum skeleton, including chapters, lessons, and concepts to be covered.
Follow the source document strictly.
Before writing the report:
1. Identify the curriculum artifact type.
2. Confirm the review scope.
3. Apply the Curriculum QA rubric from the source document.
4. Review chapter progression, lesson sequencing, concept coverage, beginner-to-advanced flow, skill mastery readiness, and real-world relevance.
5. Identify only essential curriculum structure issues.
6. Do not perform Content QA.
7. Do not review grammar, quiz answer correctness, lesson content accuracy, code correctness, visual quality, or detailed assessment criteria unless explicitly requested.
8. Do not penalize missing items that are intentionally planned for a later stage.
Generate the Curriculum QA report in the approved format:
1. Final QA Summary
2. Final Verdict
3. Score Breakdown
4. Essential Changes Required
5. Chapter and Lesson Structure Review
6. Skill Progression Review
7. Content Strengths
8. What Not to Change Unnecessarily
9. Required Fixes Before Re-QA
Important instructions:
- Focus only on curriculum structure and skill mastery progression.
- Do not merge this with Content QA.
- Do not invent issues if no essential curriculum issue exists.
- If no essential chapter or lesson structure issue is found, state that clearly.
- Use section names, chapter names, lesson names, and concept blocks as issue locations.
- Do not use page numbers as the main locator.
- In the Essential Changes Required table, include: Issue ID, Priority, Section / Content Block, Current Issue, Required Change, and Acceptance Criteria.
- Keep issues specific and actionable.
- If a Skill Mastery Challenge is defined as interview questions and answers, accept that format unless the source document or prompt says otherwise.
- If course-level title, target learner, prerequisites, certification outcome, quiz distribution, or final assessment evaluation criteria are marked as deferred, do not penalize them."""

PLAN_FIX_PROMPT = """\
The curriculum you generated for "{skill_name}" received a QA score of {score}/100 and did not meet the required threshold.

Here is the QA report with required fixes:

{qa_report}

Please regenerate the full curriculum table for "{skill_name}" with all the required fixes applied.

Follow the same original rules strictly:
- Exactly 4 main chapters + End Course Assessment
- Each main chapter: 3–5 lessons + 1 Skill Mastery Challenge row
- End Course Assessment: exactly 2 lessons (Quiz + Summative)
- Clean markdown table only — no extra text before or after
- Mastery-level depth, beginner-to-advanced progression

Output the corrected curriculum table now."""

# REQA_PROMPT split into header + footer — curriculum pasted separately.
REQA_PROMPT_HEADER = """\
Using the attached source document "Curriculum QA System - v1.0" and the previous Curriculum QA report in this chat, perform re-QA for the updated curriculum skeleton:
Short Course Name : {skill_name}
Curriculum :
\"\"\""""

REQA_PROMPT_FOOTER = """\
\"\"\"
Your task is to verify whether the previously listed curriculum issues have been fixed.
Do not generate a completely new QA report from scratch unless the curriculum skeleton has changed substantially.
Follow these rules:
1. Check the previously listed curriculum issues.
2. Verify whether each issue is Fixed, Still Open, Partially Fixed, Not Applicable, or Newly Introduced.
3. Do not add new issues from unchanged curriculum sections.
4. Add new issues only if:
   - New chapters or lessons were added,
   - Existing chapters or lessons were changed,
   - A fix introduced a new sequencing or coverage issue,
   - The review scope changed,
   - Previously unavailable curriculum information became available.
5. Do not perform Content QA.
6. Do not review quiz answer correctness, grammar, code, visuals, or detailed lesson content unless explicitly requested.
7. Update the score and final verdict if needed.
8. Keep the response concise and action-oriented.
Do not include a long tracker table unless explicitly requested."""


# =============================================================================
# CHATGPT HELPERS
# =============================================================================

_STOP_SELECTOR = (
    "[data-testid='stop-button'], "
    "button[aria-label*='Stop generating'], "
    "button[aria-label*='Stop']"
)


async def wait_for_response(page, timeout=360):
    """
    Wait until ChatGPT finishes generating.

    Strategy:
      1. Wait up to 20s for the STOP button to APPEAR — this confirms the
         model actually started generating.  Without this step the function
         can return immediately on copy/regenerate buttons that were already
         visible from the previous assistant message.
      2. Then poll until the stop button DISAPPEARS (generation complete).

    If the stop button never appears (ultra-fast or already done) we fall
    back to a short extra sleep and return.
    """
    await asyncio.sleep(2)   # brief pause for UI to react after send

    # Phase 1 — wait for generation to START
    try:
        await page.wait_for_selector(_STOP_SELECTOR, timeout=20000)
    except Exception:
        # Stop button never appeared — response may have been instantaneous.
        await asyncio.sleep(2)
        return

    # Phase 2 — wait for generation to FINISH
    deadline = time.time() + timeout
    while time.time() < deadline:
        stop = page.locator(_STOP_SELECTOR)
        if await stop.count() == 0:
            await asyncio.sleep(1)   # one extra tick to let DOM settle
            return
        await asyncio.sleep(1.5)
    raise TimeoutError(f"ChatGPT did not finish within {timeout}s")


_CLIPBOARD_SENTINEL = "||CURRICULUM_BOT_SENTINEL_CLEAR||"


async def get_last_response(page):
    """
    Return the RAW MARKDOWN of the last assistant message by clicking
    ChatGPT's Copy button.

    KEY: We write a sentinel to the clipboard FIRST, then click Copy,
    then POLL until the clipboard differs from the sentinel.  Without
    this, pyperclip.paste() returns whatever WE last copied (e.g. the
    PLAN_PROMPT) because the button click is asynchronous.

    Falls back to inner_text() if no copy button is found or clipboard
    never updates — inner_text() strips markdown pipes but is better than
    returning the wrong text entirely.
    """
    await asyncio.sleep(1)  # let copy button render

    copy_btn_selectors = [
        "button[data-testid='copy-turn-action-button']",
        "button[aria-label='Copy']",
        "button[aria-label*='copy' i]",
    ]

    for sel in copy_btn_selectors:
        btns = page.locator(sel)
        n = await btns.count()
        if n > 0:
            try:
                # 1. Write a sentinel so we can detect clipboard change
                pyperclip.copy(_CLIPBOARD_SENTINEL)
                # 2. Click the LAST copy button (= last assistant message)
                await btns.nth(n - 1).click()
                # 3. Poll until clipboard content changes from the sentinel
                for _ in range(40):          # up to 4 seconds
                    await asyncio.sleep(0.1)
                    text = pyperclip.paste()
                    if text != _CLIPBOARD_SENTINEL and text.strip():
                        return text.strip()
                # If we reach here the click didn't update the clipboard
            except Exception:
                pass

    # Fallback: inner_text (loses markdown table pipes)
    msgs = page.locator("[data-message-author-role='assistant']")
    count = await msgs.count()
    if count == 0:
        raise ValueError("No assistant messages found.")
    return await msgs.nth(count - 1).inner_text()


async def find_visible_input(page, timeout=20000):
    """
    Return the first VISIBLE ChatGPT input box, or raise if none found.
    Checks visibility explicitly — count()>0 alone matches hidden elements
    which cause click() to hang for 30s before timing out.
    """
    # Order matters: most specific first so we don't land on a hidden clone.
    input_selectors = [
        "#prompt-textarea",
        "[data-testid='prompt-textarea']",
        # Project landing page input (e.g. "New chat in structure plan tanveer")
        "div[contenteditable='true'][placeholder]",
        "textarea[placeholder]",
        # Broader fallback — only if visible
        "div[contenteditable='true']",
    ]

    deadline = time.time() + timeout / 1000
    while time.time() < deadline:
        for sel in input_selectors:
            locs = page.locator(sel)
            n = await locs.count()
            for i in range(n):
                loc = locs.nth(i)
                try:
                    if await loc.is_visible():
                        return loc
                except Exception:
                    continue
        await asyncio.sleep(0.5)

    raise ValueError(
        "Could not find a visible ChatGPT input box after "
        f"{timeout/1000:.0f}s. Check that ChatGPT is loaded and logged in."
    )


async def _click_send_or_enter(page):
    """
    Wait for the send button to become enabled, then click it.
    Falls back to Enter key if button stays disabled or isn't found.
    Shared by send_message and send_message_in_parts.
    """
    send_sel = "button[data-testid='send-button']"
    for _ in range(30):          # up to 15 s
        btn = page.locator(send_sel)
        if await btn.count() > 0:
            disabled = await btn.first.get_attribute("disabled")
            if disabled is None:   # no disabled attr → button is active
                await btn.first.click()
                await asyncio.sleep(1)
                return
        await asyncio.sleep(0.5)
    # Fallback
    await page.keyboard.press("Enter")
    await asyncio.sleep(1)


async def send_message(page, text):
    """Paste a single text block into ChatGPT input and send."""
    box = await find_visible_input(page)
    await box.click()
    await asyncio.sleep(0.5)

    pyperclip.copy(text)
    await page.keyboard.press("Control+a")
    await page.keyboard.press("Control+v")
    await asyncio.sleep(1.2)

    await _click_send_or_enter(page)


async def send_message_in_parts(page, *parts):
    """
    Build one ChatGPT message from multiple text parts and send it.

    Each part is injected directly into the focused contenteditable input
    using document.execCommand('insertText').  This avoids two problems
    that clipboard+keyboard navigation caused:
      1. Control+End moving focus OUT of the input → subsequent pastes
         hit the wrong element and the page scrolls endlessly.
      2. pyperclip.paste() returning our OWN last clipboard write instead
         of ChatGPT's response (fixed separately in get_last_response).

    execCommand('insertText') appends text at the current cursor position
    and fires the React synthetic input/change events so ChatGPT's React
    state updates correctly — the Send button becomes clickable.
    """
    box = await find_visible_input(page)
    await box.click()
    await asyncio.sleep(0.5)

    # Clear any existing content
    await page.keyboard.press("Control+a")
    await page.keyboard.press("Delete")
    await asyncio.sleep(0.3)

    for part in parts:
        if not part:
            continue
        # Inject text directly into the focused element — no clipboard involved,
        # no keyboard navigation, no focus loss.
        await page.evaluate(
            "text => document.execCommand('insertText', false, text)",
            part
        )
        await asyncio.sleep(0.4)  # let React process each chunk

    await _click_send_or_enter(page)


async def open_new_chat(page, project_url):
    """
    Navigate to the ChatGPT project page and open a fresh conversation.
    On project landing pages the input IS the new-chat entry point — no
    separate 'New chat' button click is needed.
    Returns the conversation URL (captured after send, so call page.url later).
    """
    await page.goto(project_url, wait_until="domcontentloaded")
    await asyncio.sleep(3)

    # Only click a 'New chat' button if we're already INSIDE a conversation
    # (i.e. the URL has /c/ in it), not on the project landing page.
    if "/c/" in page.url:
        new_chat_selectors = [
            "[data-testid='create-new-chat-button']",
            "button[aria-label='New chat']",
            "a[aria-label='New chat']",
        ]
        for sel in new_chat_selectors:
            btn = page.locator(sel).first
            try:
                if await btn.is_visible():
                    await btn.click()
                    await asyncio.sleep(2)
                    break
            except Exception:
                continue

    # Wait for a visible input — raises if ChatGPT isn't ready
    await find_visible_input(page, timeout=20000)
    return page.url


def extract_score(text):
    """
    Parse QA score from response. Returns int or None.

    Uses the LAST match rather than the first — re-QA responses often quote
    the old score early in the text (e.g. 'Previous score: 34/100') before
    stating the updated score at the end ('Updated score: 100/100').
    Taking the last match reliably returns the final/updated score.
    """
    patterns = [
        r"(?:final\s+)?(?:total\s+)?(?:overall\s+)?score[:\s]+(\d{1,3})\s*/\s*100",
        r"(\d{1,3})\s*/\s*100",
        r"(?:score|total)[:\s]+(\d{1,3})\b",
    ]
    for pat in patterns:
        matches = re.findall(pat, text.lower())
        if matches:
            # Take the last match — it's the updated/final score in re-QA replies
            v = int(matches[-1])
            if 0 <= v <= 100:
                return v
    return None


def parse_table(text, skill_name):
    """
    Extract rows from markdown table OR TSV → [[skill, chapter, lesson, concepts], ...]

    ChatGPT's copy button sometimes returns the table as tab-separated text
    (TSV) rather than raw markdown — especially when the response is fetched
    after navigating back to plan_url.  We detect which format was received
    and parse accordingly.
    """
    rows = []

    # ── Detect format ─────────────────────────────────────────────────────────
    sample_lines = [l for l in text.strip().split("\n") if l.strip()][:10]
    has_pipes = any("|" in ln for ln in sample_lines)
    has_tabs  = any("\t" in ln for ln in sample_lines)

    # ── TSV path ──────────────────────────────────────────────────────────────
    if has_tabs and not has_pipes:
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            parts = line.split("\t")
            if len(parts) < 3:
                continue
            chapter  = parts[0].strip()
            lesson   = parts[1].strip()
            concepts = parts[2].strip() if len(parts) > 2 else ""
            # Skip header row
            if chapter.lower() in ("chapter", "") and lesson.lower() in ("lesson", ""):
                continue
            if chapter or lesson:
                rows.append([skill_name, chapter, lesson, concepts])
        return rows

    # ── Markdown pipe table path ───────────────────────────────────────────────
    for line in text.split("\n"):
        line = line.strip()
        if not line.startswith("|"):
            continue
        if re.match(r"^\|\s*[-:]+\s*\|", line):
            continue
        parts = [p.strip() for p in line.strip("|").split("|")]
        if len(parts) < 3:
            continue
        chapter  = parts[0]
        lesson   = parts[1]
        concepts = parts[2] if len(parts) > 2 else ""
        if chapter.lower() in ("chapter", "") and lesson.lower() in ("lesson", ""):
            continue
        if chapter or lesson:
            rows.append([skill_name, chapter, lesson, concepts])
    return rows


# =============================================================================
# GOOGLE SHEETS HELPERS
# =============================================================================

async def focus_sheet(page):
    """
    Wait for Google Sheets grid to be ready and give it keyboard focus.
    Uses the specific grid container ID to avoid matching toolbar elements.
    """
    # #waffle-grid-container is the actual spreadsheet grid div.
    # [class*='waffle'] is too broad — it also matches invisible toolbar buttons.
    try:
        await page.wait_for_selector(
            "#waffle-grid-container, .grid-container",
            timeout=30000
        )
    except Exception:
        pass  # page may still be usable; proceed

    await asyncio.sleep(2)

    # Try the specific grid container first
    grid = page.locator("#waffle-grid-container").first
    if await grid.count() > 0 and await grid.is_visible():
        await grid.click(position={"x": 200, "y": 100})
    else:
        # Fallback: click at a fixed position that is always inside the cell area
        # (below the toolbar ~140px, left column ~200px)
        await page.mouse.click(200, 200)

    await asyncio.sleep(0.5)


async def read_skills_from_sheet(page):
    """
    Returns list of (row_number, skill_name) for rows where
    Col B (skill) is non-empty and Col F (status) is blank.
    Row numbers are 1-based including header (so data row 1 = sheet row 2).
    """
    # Use domcontentloaded — Google Sheets never reaches networkidle because
    # it continuously makes background requests; networkidle hangs forever.
    await page.goto(SHEET_URL, wait_until="domcontentloaded")
    await focus_sheet(page)

    tsv = ""
    for attempt in range(3):
        await page.keyboard.press("Escape")
        await asyncio.sleep(0.5)
        await page.keyboard.press("Control+Home")
        await asyncio.sleep(1)
        await page.keyboard.press("Control+a")
        await asyncio.sleep(1)
        await page.keyboard.press("Control+c")
        await asyncio.sleep(2)

        tsv = pyperclip.paste()
        if tsv.strip():
            break
        print(f"    Clipboard empty on attempt {attempt + 1}, retrying...")
        # Re-focus the grid and try again
        await focus_sheet(page)

    if not tsv.strip():
        raise ValueError("Clipboard empty after 3 attempts — could not read sheet.")

    df = pd.read_csv(io.StringIO(tsv), sep="\t", header=0, dtype=str).fillna("")

    # Statuses that mean "done — skip this row".
    # Everything else (blank, QA Failed, Score Parse Error, Unexpected Error…)
    # is treated as pending so it gets retried on the next run.
    DONE_STATUSES = {"completed", "done"}

    skills = []
    for i, row in df.iterrows():
        skill = row.iloc[1].strip() if len(row) > 1 else ""
        status = row.iloc[5].strip() if len(row) > 5 else ""
        if skill and status.lower() not in DONE_STATUSES:
            skills.append((i + 2, skill))  # +2: 1-based + header row

    return skills


def _parse_cell_address(cell_address):
    """
    Parse 'D30' → (col_index=4, row_index=30).
    col_index is 1-based (A=1, B=2, ..., Z=26, AA=27, ...).
    """
    m = re.match(r'^([A-Za-z]+)(\d+)$', cell_address.strip())
    if not m:
        raise ValueError(f"Invalid cell address: {cell_address!r}")
    col_str, row_str = m.group(1).upper(), m.group(2)
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return col, int(row_str)


async def navigate_to_cell(page, cell_address):
    """
    Navigate to a cell using Ctrl+Home then arrow keys.

    WHY THIS APPROACH:
      - Ctrl+G → Chrome intercepts it as "Find Next", types end up in the
        browser find bar instead of a Sheets dialog.
      - Name Box selectors (.cell-input, .docs-name-box, etc.) all map to
        the FORMULA BAR, not the Name Box. Clicking the formula bar puts the
        active cell into edit mode, so the address is typed as cell content
        ("A1" literally appears in the cell) rather than used for navigation.
      - Ctrl+Home is a Sheets shortcut (not intercepted by Chrome) that always
        lands on A1. From there, plain arrow key presses are deterministic.
    """
    col, row = _parse_cell_address(cell_address)

    # Exit any edit mode / close any open dialog or Chrome find bar.
    await page.keyboard.press("Escape")
    await asyncio.sleep(0.2)
    await page.keyboard.press("Escape")
    await asyncio.sleep(0.2)

    # Always start from A1 — gives us a known, repeatable origin.
    await page.keyboard.press("Control+Home")
    await asyncio.sleep(0.4)

    # Move right to the target column (A is already col 1, so we need col-1 presses).
    for _ in range(col - 1):
        await page.keyboard.press("ArrowRight")
        await asyncio.sleep(0.02)

    # Move down to the target row (row 1 is already active, so row-1 presses).
    for _ in range(row - 1):
        await page.keyboard.press("ArrowDown")
        await asyncio.sleep(0.02)

    await asyncio.sleep(0.3)


async def write_cell(page, cell_address, value):
    """
    Navigate to a cell and write a value.

    Uses keyboard.type() with F2 to enter edit mode explicitly.
    For multiline values (e.g. plan_url + newline + qa_url), each newline
    is sent as Alt+Enter — the Google Sheets shortcut for an in-cell line
    break. Plain Enter would confirm the cell and jump to the next row.
    """
    await navigate_to_cell(page, cell_address)

    # F2 enters edit mode without changing the cell's content.
    await page.keyboard.press("F2")
    await asyncio.sleep(0.2)

    value_str = str(value)
    if "\n" in value_str:
        lines = value_str.split("\n")
        for i, line in enumerate(lines):
            await page.keyboard.type(line)
            if i < len(lines) - 1:
                await page.keyboard.press("Alt+Enter")  # in-cell line break
                await asyncio.sleep(0.1)
    else:
        await page.keyboard.type(value_str)

    # Tab confirms the cell and moves right — does NOT jump to the next row.
    await page.keyboard.press("Tab")
    await asyncio.sleep(0.4)


async def save_sheet(page):
    await page.keyboard.press("Control+s")
    await asyncio.sleep(1)


async def click_sheet_tab(page, tab_name):
    """Click a sheet tab by name."""
    # FIX 7: Use role-based tab detection with text fallback
    tab = page.get_by_role("tab", name=tab_name)
    if await tab.count() > 0:
        await tab.click()
        await asyncio.sleep(1.5)
        return True

    # Fallback: find by text in the bottom tab bar
    tab2 = page.locator(f"text='{tab_name}'").last
    if await tab2.count() > 0:
        await tab2.click()
        await asyncio.sleep(1.5)
        return True

    return False


async def create_curriculum_tab(page):
    """
    Create a new 'Curriculum' sheet tab and add a header row.
    Returns True on success, False if creation failed.
    """
    print("    📋 'Curriculum' tab not found — creating it...")

    # Click the "+" (Add sheet) button at the bottom of Google Sheets
    add_btn_selectors = [
        "[aria-label='Add sheet']",
        ".docs-sheet-add-button",
        "[data-tooltip='Add sheet']",
        "div.docs-sheet-active-tab + div",   # button right after last tab
    ]
    clicked = False
    for sel in add_btn_selectors:
        btn = page.locator(sel)
        if await btn.count() > 0:
            try:
                await btn.first.click()
                await asyncio.sleep(1.5)
                clicked = True
                break
            except Exception:
                continue

    if not clicked:
        print("    ❌ Could not find 'Add sheet' button.")
        return False

    # Rename the new tab to "Curriculum":
    # Double-click the currently selected (active) tab to enter rename mode
    active_tab_selectors = [
        ".docs-sheet-active-tab .docs-sheet-tab-name",
        ".docs-sheet-active-tab",
        "[aria-selected='true']",
    ]
    renamed = False
    for sel in active_tab_selectors:
        tab = page.locator(sel).last
        if await tab.count() > 0:
            try:
                await tab.dblclick()
                await asyncio.sleep(0.8)
                # Select all existing text and replace
                await page.keyboard.press("Control+a")
                await asyncio.sleep(0.2)
                await page.keyboard.type("Curriculum")
                await page.keyboard.press("Enter")
                await asyncio.sleep(1)
                renamed = True
                break
            except Exception:
                continue

    if not renamed:
        print("    ⚠️  Tab created but could not rename — leaving as-is.")

    # After a tab rename, keyboard focus is on the tab bar, NOT the cell grid.
    # focus_sheet() clicks the grid container so that Ctrl+Home (inside
    # navigate_to_cell) is intercepted by Google Sheets, not by Chrome.
    await focus_sheet(page)

    # Add header row: Short Course Name | Chapter | Lesson | Concepts to be Covered
    await navigate_to_cell(page, "A1")
    header_tsv = "Short Course Name\tChapter\tLesson\tConcepts to be Covered\tLinks"
    pyperclip.copy(header_tsv)
    await page.keyboard.press("Control+v")
    await asyncio.sleep(1.5)
    await save_sheet(page)
    print("    ✅ 'Curriculum' tab created with header row.")
    return True


# ---------------------------------------------------------------------------
# Row counter for the Curriculum tab.
# We track the next empty row explicitly rather than using Ctrl+End, because
# Ctrl+End on a newly-created Google Sheet (which pre-allocates 1000 rows)
# jumps to row 1000 even when the sheet is mostly empty.
# ---------------------------------------------------------------------------
_curriculum_next_row: int | None = None


async def _read_curriculum_row_count(page) -> int:
    """
    Count how many DATA rows the Curriculum tab already has (header excluded).
    Used on script restart when the tab already exists from a previous run.
    Returns 0 if the tab is empty or has only a header.
    """
    await focus_sheet(page)
    await page.keyboard.press("Escape")
    await asyncio.sleep(0.3)
    await page.keyboard.press("Control+Home")
    await asyncio.sleep(1)
    await page.keyboard.press("Control+a")
    await asyncio.sleep(1)
    await page.keyboard.press("Control+c")
    await asyncio.sleep(2)

    tsv = pyperclip.paste()
    if not tsv.strip():
        return 0

    lines = [ln for ln in tsv.split("\n") if ln.strip()]
    return max(0, len(lines) - 1)   # subtract 1 for the header row


async def append_rows_to_curriculum_tab(page, rows, plan_url="", qa_url=""):
    """
    Paste new curriculum rows at the bottom of the Curriculum tab.
    Creates the tab (with headers) automatically if it does not exist.
    Rows format: [skill_name, chapter, lesson, concepts]
    A 5th "Links" column is appended to every row using plan_url + qa_url.

    Uses a module-level counter (_curriculum_next_row) to track the exact
    target row so we never rely on Ctrl+End (which overshoots on empty sheets).
    On the first call of a fresh session the counter is initialised to 2
    (row after the header). On a restart the counter is initialised by
    reading the tab to count its existing rows.
    """
    global _curriculum_next_row

    found = await click_sheet_tab(page, "Curriculum")
    if not found:
        created = await create_curriculum_tab(page)
        if not created:
            print("    ❌ Could not create 'Curriculum' tab — skipping write.")
            return
        _curriculum_next_row = 2   # header is row 1, first data row is row 2

    elif _curriculum_next_row is None:
        # Tab already exists (script was restarted) — figure out where data ends.
        existing = await _read_curriculum_row_count(page)
        # existing=0 → header only → next row is 2
        # existing=22 → rows 1-23 used → next row is 24
        _curriculum_next_row = existing + 2

    # Build links value — same format as Tracker column D
    links_value = f"1. {plan_url}\n2. {qa_url}" if plan_url or qa_url else ""

    # Navigate directly to the known next empty row and paste the TSV block.
    await navigate_to_cell(page, f"A{_curriculum_next_row}")

    tsv = "\n".join(
        "\t".join(str(c) for c in row) + "\t" + links_value
        for row in rows
    )
    pyperclip.copy(tsv)
    await page.keyboard.press("Control+v")
    await asyncio.sleep(2)

    _curriculum_next_row += len(rows)

    await save_sheet(page)
    print(f"    ✅ Wrote {len(rows)} rows to 'Curriculum' tab.")


async def update_tracker_row(page, row_number, status_text, plan_url, qa_url):
    """
    Update Tracker tab:
      - Col D (ChatGPT Page Link) = "1. {plan_url}\n2. {qa_url}"
      - Col F (Status) = status_text
    """
    # Go back to the first / main sheet tab
    await page.goto(SHEET_URL, wait_until="domcontentloaded")
    await focus_sheet(page)

    # Write ChatGPT links to col D
    links_value = f"1. {plan_url}\n2. {qa_url}"
    await write_cell(page, f"D{row_number}", links_value)

    # Write status to col F
    await write_cell(page, f"F{row_number}", status_text)

    await save_sheet(page)


# =============================================================================
# MAIN SKILL PIPELINE
# =============================================================================

async def process_skill(page, row_number, skill_name, idx, total):
    print(f"\n{'='*62}")
    print(f"  [{idx}/{total}] {skill_name}  (sheet row {row_number})")
    print(f"{'='*62}")

    plan_url = ""
    qa_url   = ""

    # ── 1. Generate curriculum ────────────────────────────────────
    print("  Step 1 │ Generating curriculum (Plan project)...")
    plan_url = await open_new_chat(page, PLAN_PROJECT_URL)
    await send_message(page, PLAN_PROMPT.format(skill_name=skill_name))
    await wait_for_response(page)
    curriculum = await get_last_response(page)
    plan_url = page.url          # capture after conversation starts
    print(f"          Plan URL: {plan_url}")
    print(f"          Curriculum length: {len(curriculum)} chars")

    # ── 2. QA — analyze then report ──────────────────────────────
    print("  Step 2 │ Running QA (QA project)...")
    qa_url = await open_new_chat(page, QA_PROJECT_URL)
    await send_message(page, QA_ANALYZE_PROMPT)
    await wait_for_response(page)
    print("          QA analyzer acknowledged.")

    await send_message_in_parts(
        page,
        QA_REPORT_PROMPT_HEADER.format(skill_name=skill_name),
        "\n" + curriculum + "\n",
        QA_REPORT_PROMPT_FOOTER,
    )
    await wait_for_response(page)
    qa_report = await get_last_response(page)
    qa_url = page.url
    print(f"          QA URL: {qa_url}")

    score = extract_score(qa_report)
    print(f"          QA Score: {score}/100")

    # ── 3. Re-QA loop if needed ───────────────────────────────────
    retry = 0
    while (score is None or score < QA_PASS_THRESHOLD) and retry < MAX_REQA_RETRIES:
        retry += 1
        print(f"  Step 3 │ Re-QA attempt {retry}/{MAX_REQA_RETRIES}...")

        # Ask Plan project (same conversation) to fix curriculum
        await page.goto(plan_url, wait_until="domcontentloaded", timeout=60000)
        await asyncio.sleep(2)
        fix_prompt = PLAN_FIX_PROMPT.format(
            skill_name=skill_name,
            score=score if score is not None else "unknown",
            qa_report=qa_report
        )
        await send_message(page, fix_prompt)
        await wait_for_response(page)
        curriculum = await get_last_response(page)
        print(f"          Fixed curriculum: {len(curriculum)} chars")

        # Ask QA project (same conversation) to re-QA
        await page.goto(qa_url, wait_until="domcontentloaded", timeout=60000)
        await asyncio.sleep(2)
        await send_message_in_parts(
            page,
            REQA_PROMPT_HEADER.format(skill_name=skill_name),
            "\n" + curriculum + "\n",
            REQA_PROMPT_FOOTER,
        )
        await wait_for_response(page)
        qa_report = await get_last_response(page)

        score = extract_score(qa_report)
        print(f"          Re-QA Score: {score}/100")

    # ── 4. Write results ──────────────────────────────────────────
    if score is None:
        status = "Score Parse Error"
        result_label = "⚠️  Score Parse Error"
    elif score >= QA_PASS_THRESHOLD:
        # After re-QA retries the page may have navigated away from plan_url,
        # and the fallback inner_text() path strips markdown table pipes so
        # parse_table returns 0 rows.  Navigate back to plan_url and pull the
        # final approved curriculum fresh so the copy-button path gets a clean
        # attempt at the raw markdown.
        if retry > 0:
            print("          Fetching final curriculum from Plan project...")
            await page.goto(plan_url, wait_until="domcontentloaded", timeout=60000)
            await asyncio.sleep(2)
            curriculum = await get_last_response(page)
            print(f"          Final curriculum: {len(curriculum)} chars")

        rows = parse_table(curriculum, skill_name)
        if len(rows) == 0:
            # Diagnostic: show a snippet so we know what parse_table received
            print(f"  ⚠️  parse_table returned 0 rows — first 300 chars received:")
            print(f"      {repr(curriculum[:300])}")

        await page.goto(SHEET_URL, wait_until="domcontentloaded")
        await focus_sheet(page)
        await append_rows_to_curriculum_tab(page, rows, plan_url=plan_url, qa_url=qa_url)
        status = "Completed"
        result_label = f"✅ Passed ({score}/100, {len(rows)} rows written)"
    else:
        status = f"QA Failed ({score})"
        result_label = f"❌ Failed ({score}/100 after {retry} retries)"

    # Save links + status to Tracker
    await update_tracker_row(page, row_number, status, plan_url, qa_url)

    print(f"  Result │ {result_label}")
    return skill_name, score, result_label


# =============================================================================
# ENTRY POINT
# =============================================================================

LOG_FILE = "chrome_launch.log"


def log(msg):
    """Print to console and append to chrome_launch.log for diagnostics."""
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")


def prepare_debug_profile():
    """
    Copy Profile 5 → ChromeDebugSession/Default on the very first run only.

    WHY copy-once instead of copy-every-run:
      - Copying every run overwrites the "verified device" token Google stores
        in the profile after the "verify it's you" check, so Google would ask
        for the OTP on every single run.
      - Copying just once means the verification is saved permanently in
        ChromeDebugSession and never asked again.

    WHY copy at all (not use Profile 5 directly):
      - Chrome 120+ rejects --remote-debugging-port when --user-data-dir points
        to the real Chrome User Data directory (the default location).
      - A separate directory bypasses this restriction.
    """
    dst_profile = os.path.join(CHROME_DEBUG_DATA_DIR, "Default")

    # Already set up — skip copy entirely
    if os.path.exists(dst_profile):
        return

    src_profile = os.path.join(CHROME_USER_DATA_DIR, CHROME_PROFILE_DIR)
    src_state   = os.path.join(CHROME_USER_DATA_DIR, "Local State")
    dst_state   = os.path.join(CHROME_DEBUG_DATA_DIR, "Local State")

    if not os.path.exists(src_profile):
        print(f"\n❌ Source profile not found: {src_profile}")
        print("   Check CHROME_USER_DATA_DIR and CHROME_PROFILE_DIR in the config.")
        sys.exit(1)

    print(f"\n  Copying '{CHROME_PROFILE_DIR}' → {CHROME_DEBUG_DATA_DIR}")
    print("  (This runs only once — takes 15–60 s depending on profile size...)")

    os.makedirs(CHROME_DEBUG_DATA_DIR, exist_ok=True)
    shutil.copytree(src_profile, dst_profile)

    if os.path.exists(src_state):
        shutil.copy2(src_state, dst_state)

    print("  ✅ Profile copied. Chrome will ask 'verify it's you' this one time.")
    print("     Tick 'Don't ask again on this device' before confirming.")
    print()


def delete_singleton_lock():
    """
    Remove Chrome's SingletonLock file from the debug data directory.
    If Chrome was killed hard, this file can linger and prevent the next
    Chrome launch from opening the profile — causing the debug port to
    silently never open.
    """
    lock_path = os.path.join(CHROME_DEBUG_DATA_DIR, "Default", "SingletonLock")
    lock_parent = os.path.join(CHROME_DEBUG_DATA_DIR, "SingletonLock")
    for p in (lock_path, lock_parent):
        if os.path.exists(p):
            try:
                os.remove(p)
                log(f"  Removed stale lock: {p}")
            except OSError as e:
                log(f"  Could not remove lock {p}: {e}")


def launch_chrome_with_debug_port():
    """
    Start Chrome as a normal subprocess with a remote debugging port.
    Playwright will connect to it via CDP instead of launching it directly.
    """
    # Open log fresh each run
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"=== Chrome launch log — {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")

    # Kill ALL existing Chrome processes. If ANY Chrome window is open without
    # --remote-debugging-port, the new subprocess merges into it and the debug
    # port never opens.
    log("  Killing any existing Chrome processes...")
    result = subprocess.run(
        ["taskkill", "/F", "/IM", "chrome.exe"],
        capture_output=True, text=True
    )
    log(f"  taskkill stdout: {result.stdout.strip()}")
    log(f"  taskkill stderr: {result.stderr.strip()}")

    # Wait long enough for processes AND file handles to fully release.
    # 1 second is often not enough — profile lock files can linger.
    log("  Waiting 4s for Chrome processes to fully terminate...")
    time.sleep(4)

    # Remove any stale SingletonLock left by the killed Chrome.
    # This file prevents a new Chrome from using the same profile.
    delete_singleton_lock()

    # Locate chrome.exe
    exe_candidates = [
        CHROME_EXE,
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    chrome_exe = next((p for p in exe_candidates if os.path.exists(p)), None)
    if not chrome_exe:
        log("❌ chrome.exe not found. Set CHROME_EXE at the top of the script.")
        sys.exit(1)
    log(f"  Using Chrome: {chrome_exe}")

    args = [
        chrome_exe,
        f"--remote-debugging-port={CHROME_DEBUG_PORT}",
        # Required on Chrome 120+ — without this the debug port accepts no connections
        "--remote-allow-origins=*",
        f"--user-data-dir={CHROME_DEBUG_DATA_DIR}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
    ]
    log(f"  Launch args: {' '.join(args)}")

    # Redirect Chrome stdout/stderr to the log file so we can see crash reasons
    log_fh = open(LOG_FILE, "a", encoding="utf-8")
    proc = subprocess.Popen(args, stdout=log_fh, stderr=log_fh)
    log(f"  Chrome launched (pid {proc.pid}) on port {CHROME_DEBUG_PORT}.")

    # Poll port in a loop. Also check the process is still alive each second.
    log("  Waiting for Chrome debug port to open (up to 45s)...")
    for i in range(45):
        time.sleep(1)

        # If Chrome process already exited, it failed to start
        if proc.poll() is not None:
            log(f"❌ Chrome process exited early (code {proc.returncode}). "
                f"Check {LOG_FILE} for details.")
            log_fh.close()
            sys.exit(1)

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            if s.connect_ex(("127.0.0.1", CHROME_DEBUG_PORT)) == 0:
                log(f"  ✅ Debug port {CHROME_DEBUG_PORT} open after {i + 1}s.")
                log_fh.close()
                return proc

        if i % 5 == 4:
            log(f"  ... still waiting ({i + 1}s elapsed) ...")

    log(f"❌ Chrome debug port never opened after 45s. "
        f"Check {LOG_FILE} for Chrome error output.")
    log_fh.close()
    sys.exit(1)


async def main():
    print("=" * 62)
    print("  CURRICULUM AUTOMATION  v2  (FIXED)")
    print("=" * 62)

    # Copy Profile 5 → ChromeDebugSession on first run only.
    # On subsequent runs this is a no-op (folder already exists).
    first_time = not os.path.exists(
        os.path.join(CHROME_DEBUG_DATA_DIR, "Default")
    )
    prepare_debug_profile()

    if first_time:
        print()
        print("  *** FIRST-TIME SETUP — read before pressing Enter ***")
        print()
        print("  Your Profile 5 was just copied to the debug folder.")
        print("  When Chrome opens, Google may show 'verify it's you'.")
        print()
        print("  → Enter the OTP from your boss to confirm.")
        print("  → Tick 'Don't ask again on this device' before confirming.")
        print()
        print("  That's it — you won't be asked again on future runs.")
        print()
    input("  → Press Enter to launch Chrome... ")

    results = []

    # Launch Chrome with debug port so Playwright can attach
    chrome_proc = launch_chrome_with_debug_port()

    # Give the boss time to verify / log in without the script zooming off.
    print()
    print("  Chrome is open.")
    print("  Complete any verification steps in Chrome (OTP, 'verify it's you', etc.)")
    print("  Once Chrome looks normal and you're ready to go:")
    print()
    input("  🍳  Should we start cooking? Press Enter to begin...  ")
    print()

    async with async_playwright() as p:
        print("Connecting Playwright to Chrome...")
        browser = await p.chromium.connect_over_cdp(
            f"http://127.0.0.1:{CHROME_DEBUG_PORT}"
        )
        # Use the existing context (has your logged-in sessions)
        if browser.contexts:
            context = browser.contexts[0]
        else:
            context = await browser.new_context()

        page = await context.new_page()

        # Read pending skills
        print("Reading skills from Google Sheet...")
        try:
            skills = await read_skills_from_sheet(page)
        except Exception as e:
            print(f"❌ Could not read sheet: {e}")
            await browser.close()
            return

        if not skills:
            print("No pending skills found (all rows already have a Status).")
            await browser.close()
            return

        skill_list = [s[1] for s in skills]
        print(f"Found {len(skills)} skills to process:")
        for i, s in enumerate(skill_list, 1):
            print(f"  {i}. {s}")
        print()

        # Process each skill
        for i, (row_number, skill_name) in enumerate(skills, 1):
            try:
                result = await process_skill(page, row_number, skill_name, i, len(skills))
                results.append(result)
            except Exception as e:
                err_msg = str(e)[:60]
                print(f"  ❌ Unexpected error: {err_msg}")
                results.append((skill_name, "N/A", f"❌ Error: {err_msg}"))
                try:
                    await update_tracker_row(
                        page, row_number, f"Error: {err_msg[:40]}", "", ""
                    )
                except Exception:
                    pass

            if i < len(skills):
                print(f"\n  Waiting {DELAY_BETWEEN_SKILLS}s before next skill...")
                await asyncio.sleep(DELAY_BETWEEN_SKILLS)

        await browser.close()

    # Kill Chrome subprocess if we launched it
    if chrome_proc is not None:
        chrome_proc.terminate()

    # Summary
    print("\n" + "=" * 62)
    print("  FINAL SUMMARY")
    print("=" * 62)
    print(f"{'Skill':<44} {'Score':<7} Result")
    print("-" * 62)
    for skill, score, label in results:
        print(f"{skill:<44} {str(score):<7} {label}")
    print("=" * 62)


if __name__ == "__main__":
    asyncio.run(main())
