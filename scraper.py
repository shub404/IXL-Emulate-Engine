import os
import time
import base64
import urllib.parse
from urllib.parse import urljoin
import requests
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

BASE_URL = "https://in.ixl.com"
TARGET_URL = "https://in.ixl.com/maths/class-iii"
LOGIN_URL = "https://in.ixl.com/signin"
EMAIL = "parkerhouston411@kacad"
PASSWORD = "81party"
QUESTIONS_PER_SKILL = 3
EXCEL_FILENAME = "ixl_grade3_questions(1).xlsx"
IMAGE_DIR = "ixl_diagrams"

# ── Mode 2: set this to the skill URL you want to resume from ─────────────────
START_URL = "https://www.ixl.com/math/grade-3/multiplication-facts-for-2-3-4-5-and-10-sorting"

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

JUNK_LABELS = {"scratchpad", "eraser", "highlighter - blue",
               "pencil - black", "pencil", "highlighter"}

ICON_MAX = 32

# ── gate signals: ONLY class-based, specific to IXL figure containers ─────────
# svg[role=...] removed — too broad, catches UI icons and nav elements
DIAGRAM_SIGNALS = [
    ".multiplication-model-container",
    ".guide-counting-qm",
    ".open-number-line",
    ".dc-fraction-strip-model",
    ".pie-chart",
    ".vector-image-wrapper",
    ".shape",
    "canvas",
    '[role="figure"]',
    "table.old-table",
    "table.qTabularGrid",
    "svg:has(g.grid-region)",
    "div.table:has([data-testid='area-model-cell'])",
]

Q_SCOPE_PARTS = [
    ".question-and-submission-view .secContent",
]

# ── image dimensions in Excel ─────────────────────────────────────────────────
EXCEL_IMG_MAX_W = 200   # px — max width when embedding in Excel
EXCEL_IMG_MAX_H = 150   # px — max height when embedding in Excel
EXCEL_ROW_HEIGHT_PER_IMG = 115  # Excel row height units per embedded image


def setup_dir():
    os.makedirs(IMAGE_DIR, exist_ok=True)


def init_excel():
    if os.path.exists(EXCEL_FILENAME):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Grade 3 Maths"
    headers = ["#", "Category", "Skill Name", "Question No",
               "Question Text", "Answer Options", "Correct Answer",
               "Question Diagrams", "Option Diagrams"]
    header_font = Font(name="Calibri", bold=True, size=11)

    for col, label in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    widths = {"A": 6, "B": 30, "C": 45, "D": 12,
              "E": 70, "F": 35, "G": 30, "H": 30, "I": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:I1"
    wb.save(EXCEL_FILENAME)


def _scale_image_for_excel(img_path):
    """
    Load a PNG, scale it to fit within EXCEL_IMG_MAX_W x EXCEL_IMG_MAX_H
    preserving aspect ratio, return an openpyxl Image object.
    """
    from PIL import Image as PILImage
    with PILImage.open(img_path) as pil_img:
        orig_w, orig_h = pil_img.size
        scale = min(EXCEL_IMG_MAX_W / orig_w, EXCEL_IMG_MAX_H / orig_h, 1.0)
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)
        resized = pil_img.resize((new_w, new_h), PILImage.LANCZOS)
        buf = BytesIO()
        resized.save(buf, format="PNG")
        buf.seek(0)
    xl_img = XLImage(buf)
    xl_img.width  = new_w
    xl_img.height = new_h
    return xl_img


def append_to_excel(row_data, q_diagram_paths, opt_diagram_paths, ans_diagram_paths=None):
    """
    Writes a data row and embeds diagram images directly into columns H and I+.
    row_data must have 7 values (columns A–G); H and I are handled via images.
    If ans_diagram_paths is provided, col G gets images (vertically) instead of text.
    """
    try:
        wb = load_workbook(EXCEL_FILENAME)
        ws = wb.active
        current_row = ws.max_row + 1
        cell_font = Font(name="Calibri", size=11)

        # ── write columns A–G (skip G if answer images are provided) ─────────
        for col_idx, value in enumerate(row_data, start=1):
            if col_idx == 7 and ans_diagram_paths:
                # Col G will hold images instead — write border/alignment only
                cell = ws.cell(row=current_row, column=col_idx)
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top")
                continue
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            cell.font = cell_font
            cell.border = BORDER
            if col_idx in [1, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        # ── columns H and I get borders even if no images ─────────────────────
        for col_idx in [8, 9]:
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top")

        # ── embed question diagrams into column H ─────────────────────────────
        h_col_letter = get_column_letter(8)
        q_img_count = 0
        for img_path in q_diagram_paths:
            if not os.path.exists(img_path):
                continue
            try:
                xl_img = _scale_image_for_excel(img_path)
                anchor_row = current_row + q_img_count
                cell_ref = f"{h_col_letter}{anchor_row}"
                ws.add_image(xl_img, cell_ref)

                ws.cell(row=anchor_row, column=8).border = BORDER
                ws.cell(row=anchor_row, column=8).alignment = Alignment(vertical="top")
                ws.row_dimensions[anchor_row].height = max(
                    ws.row_dimensions[anchor_row].height or 0,
                    xl_img.height * 0.75 + 6
                )
                q_img_count += 1
            except Exception as e:
                print(f"     [!] Could not embed image {img_path}: {e}")

        # ── embed answer diagrams vertically in column G ─────────────────────
        if ans_diagram_paths:
            g_col_letter = get_column_letter(7)
            ans_img_count = 0
            for img_path in ans_diagram_paths:
                if not os.path.exists(img_path):
                    continue
                try:
                    xl_img = _scale_image_for_excel(img_path)
                    anchor_row = current_row + ans_img_count
                    ws.add_image(xl_img, f"{g_col_letter}{anchor_row}")
                    cell = ws.cell(row=anchor_row, column=7)
                    cell.border = BORDER
                    cell.alignment = Alignment(vertical="top")
                    ws.row_dimensions[anchor_row].height = max(
                        ws.row_dimensions[anchor_row].height or 0,
                        xl_img.height * 0.75 + 6
                    )
                    ans_img_count += 1
                except Exception as e:
                    print(f"     [!] Could not embed answer image {img_path}: {e}")

        # ── embed option diagrams horizontally: I, J, K… (one image per column) ──
        opt_col = 9  # start at column I
        for img_path in opt_diagram_paths:
            if not os.path.exists(img_path):
                continue
            try:
                xl_img = _scale_image_for_excel(img_path)
                col_letter = get_column_letter(opt_col)
                ws.add_image(xl_img, f"{col_letter}{current_row}")
                cell = ws.cell(row=current_row, column=opt_col)
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top")
                if ws.column_dimensions[col_letter].width < 30:
                    ws.column_dimensions[col_letter].width = 30
                ws.row_dimensions[current_row].height = max(
                    ws.row_dimensions[current_row].height or 0,
                    xl_img.height * 0.75 + 6
                )
                opt_col += 1
            except Exception as e:
                print(f"     [!] Could not embed image {img_path}: {e}")

        # ── ensure the base row is visible even if there are no images ────────
        if not q_diagram_paths and not opt_diagram_paths:
            ws.row_dimensions[current_row].height = 40
        elif not ws.row_dimensions[current_row].height:
            ws.row_dimensions[current_row].height = 40

        wb.save(EXCEL_FILENAME)

    except PermissionError:
        print(f"\n[!] Close {EXCEL_FILENAME} in Excel immediately.")
        input("Press ENTER here once closed to resume saving data...")
        wb = load_workbook(EXCEL_FILENAME)
        ws = wb.active
        current_row = ws.max_row + 1
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=current_row, column=col_idx, value=value)
        wb.save(EXCEL_FILENAME)

_MATH_WALKER_JS = """
    const _DIAGRAM_CLASSES = [
        'dc-fraction-strip-model', 'open-number-line', 'graphingBaseContainer',
        'pie-chart', 'multiplication-model-container', 'guide-counting-qm',
        'vector-image-wrapper', 'parking-lot', 'old-table', 'binsContainer', 'qTabularGrid', 'table'
    ];
    let out = '';
    const walk = (node) => {
        for (const child of node.childNodes) {
            if (child.nodeType === Node.TEXT_NODE) {
                out += child.textContent;
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                if (child.getAttribute('aria-hidden') === 'true') continue;
                if (child.classList && _DIAGRAM_CLASSES.some(c => child.classList.contains(c))) continue;
                const tag = child.tagName.toLowerCase();
                if (tag === 'input' && child.classList.contains('fillIn')) {
                    out += '__';
                } else if (tag === 'div' && child.classList && child.classList.contains('drop-slot')) {
                    out += '___';
                } else if (tag === 'table' && child.hasAttribute('audioalt')) {
                    // Old-style fraction table on option tiles: <table audioalt="2/3">
                    out += child.getAttribute('audioalt');
                } else if (child.classList && child.classList.contains('old-fraction-in-text')) {
                    // Old-style fraction in question text: span.old-fraction-in-text > table
                    const tbl = child.querySelector('table');
                    if (tbl && tbl.hasAttribute('audioalt')) {
                        out += tbl.getAttribute('audioalt');
                    } else if (tbl) {
                        const rows = tbl.querySelectorAll('tr');
                        const num = rows.length > 0 ? rows[0].textContent.trim() : '?';
                        const den = rows.length > 1 ? rows[1].textContent.trim() : '?';
                        out += num + '/' + den;
                    }
                } else if (child.classList && child.classList.contains('vFrac')) {
                    const numEl = child.querySelector('.numerator');
                    const denEl = child.querySelector('.denominator');
                    const numText = (numEl && numEl.querySelector('input.fillIn')) ? '__'
                                  : (numEl ? numEl.textContent.trim() : '?');
                    const denText = (denEl && denEl.querySelector('input.fillIn')) ? '__'
                                  : (denEl ? denEl.textContent.trim() : '?');
                    out += numText + '/' + denText;
                } else if (child.classList && child.classList.contains('vertArith')) {
                    const rows = [...child.querySelectorAll('.vertArithRow')];
                    const operands = [];
                    let operator = '';
                    let answerBlanks = 0;
                    for (const row of rows) {
                        if (row.getAttribute('role') === 'group') {
                            answerBlanks = row.querySelectorAll('input.fillIn').length;
                            continue;
                        }
                        const opCell = row.querySelector('.vertArithCell.operator');
                        if (opCell) {
                            let sym = opCell.textContent.trim();
                            sym = sym === '–' ? '-' : sym === '÷' ? '/' : sym;
                            if (!operator) operator = sym;
                        }
                        let numStr = '';
                        for (const cell of row.querySelectorAll('.vertArithCell')) {
                            if (cell.classList.contains('operator')) continue;
                            const txt = cell.querySelector('.txt');
                            if (txt) { numStr += txt.textContent.trim(); }
                            else if (cell.classList.contains('rtlCell')) {
                                const d = cell.textContent.trim();
                                if (d) numStr += d;
                            }
                        }
                        if (numStr) operands.push(numStr);
                    }
                    if (operands.length > 0) {
                        const op = operator || '+';
                        
                          += operands.join(' ' + op + ' ');
                        if (answerBlanks > 0)
                            out += ' = ' + Array(answerBlanks).fill('__').join(' ');
                    }
                } else if (child.classList && child.classList.contains('old-vertiArith')) {
                    const rows = [...child.querySelectorAll('table tr')];
                    const entries = [];
                    let operator = '';
                    for (const row of rows) {
                        const cells = [...row.querySelectorAll('td')];
                        let numStr = '';
                        for (let i = 0; i < cells.length; i++) {
                            if (i === 0) {
                                const t = cells[i].textContent.trim();
                                if (t && t !== ' ') {
                                    operator = t === '–' ? '-' : t === '÷' ? '/' : t;
                                }
                                continue;
                            }
                            const fi = cells[i].querySelector('input.fillIn');
                            if (fi) { numStr += '__'; }
                            else {
                                const inner = cells[i].querySelector('div') || cells[i];
                                const t = inner.textContent.trim();
                                if (t && t !== ' ') numStr += t;
                            }
                        }
                        if (numStr) entries.push(numStr);
                    }
                    if (entries.length >= 2) {
                        const op = operator || '+';
                        out += entries[0] + ' ' + op + ' ' + entries[1];
                        if (entries.length >= 3) out += ' = ' + entries[2];
                    }
                } else {
                    walk(child);
                }
            }
        }
    };
"""

"""Only activates for fill-in questions; formats fractions as num/den."""
def _reconstruct_with_blanks(page, root_selector):

    js = f"""
    (sel) => {{
        const root = document.querySelector(sel);
        if (!root) return null;
        if (!root.querySelector('input.fillIn')) return null;
        {_MATH_WALKER_JS}
        walk(root);
        return out;
    }}
    """
    try:
        return page.evaluate(js, root_selector)
    except Exception:
        return None

    """Extract text from selector with fractions as num/den; works for all question types."""
def _extract_math_text(page, root_selector):

    js = f"""
    (sel) => {{
        const root = document.querySelector(sel);
        if (!root) return null;
        {_MATH_WALKER_JS}
        walk(root);
        return out.trim() || null;
    }}
    """
    try:
        return page.evaluate(js, root_selector)
    except Exception:
        return None


def extract_question_text(page):
    for sel in (".question-and-submission-view .math.section",
                ".question-and-submission-view .ixl-practice-crate",
                ".ixl-practice-crate"):
        rebuilt = _reconstruct_with_blanks(page, sel)
        # To clean the string and remove additional spaces
        if rebuilt is not None:
            text = " ".join(rebuilt.replace("\n", " ").split())
            return text

    question_text = ""
    hdr     = page.locator(".secHdr").first
    content = page.locator(".secContent").first
    crate   = page.locator(".ixl-practice-crate").first

    try:
        parts = []
        if hdr.count() > 0 and hdr.is_visible():
            hdr_text = _extract_math_text(page, ".secHdr")
            parts.append(hdr_text if hdr_text else hdr.inner_text())
        if content.count() > 0 and content.is_visible():
            content_text = _extract_math_text(page, ".secContent")
            if content_text and content_text.strip():
                parts.append(content_text)
        if parts:
            question_text = "\n".join(
                " ".join(p.replace("\n", " ").split()) for p in parts if p.strip()
            )
        elif crate.count() > 0 and crate.is_visible():
            crate_text = _extract_math_text(page, ".ixl-practice-crate")
            # Accept the entire dirty string if no question appears
            # Clean the text by removing 'Submit' and extra spaces
            question_text = " ".join(
                (crate_text or crate.inner_text()).split("Submit")[0].replace("\n", " ").split()
            )
    except Exception as e:
        print(f"     [!] text read failed: {e}")

    return question_text


def extract_options(page):
    options = []
    tiles = page.locator(
        ".question-and-submission-view .SelectableTile, "
        ".ixl-practice-crate .SelectableTile"
    ).all()
    _walker_js = f"el => {{ {_MATH_WALKER_JS} walk(el); return out.trim(); }}"
    for tile in tiles:
        try:
            label = tile.get_attribute("aria-label") or ""
            if not label.strip():
                # No aria-label: old-style fraction tiles use audioalt on the inner table.
                # Extract via JS walker on .GeneticallyModified content.
                gm = tile.locator(".GeneticallyModified").first
                if gm.count() > 0:
                    try:
                        label = gm.evaluate(_walker_js) or ""
                    except Exception:
                        label = gm.inner_text()
            if label and label.strip():
                options.append(label.strip())
        except Exception:
            continue
    # Drag-and-drop questions: options are draggable tiles in .parking-lot
    if not options:
        drag_tiles = page.locator(
            ".question-and-submission-view .parking-lot .draggable-tile, "
            ".ixl-practice-crate .parking-lot .draggable-tile"
        ).all()
        for tile in drag_tiles:
            try:
                try:
                    label = tile.evaluate(_walker_js) or ""
                except Exception:
                    label = tile.inner_text()
                label = " ".join(label.replace("\n", " ").split())
                if label and label.strip():
                    options.append(label.strip())
            except Exception:
                continue

    # Sorting drag-and-drop: tiles in .ddItemBankDropSlot
    if not options:
        bank_slots = page.locator(
            ".question-and-submission-view .ddItemBankDropSlot, "
            ".ixl-practice-crate .ddItemBankDropSlot"
        ).all()
        for slot in bank_slots:
            try:
                content = slot.locator(".itemContent").first
                target = content if content.count() > 0 else slot
                label = target.evaluate(_walker_js) or ""
                label = " ".join(label.replace("\n", " ").split())
                if label:
                    options.append(label.strip())
            except Exception:
                continue

    seen, unique = set(), []
    for o in options:
        if o not in seen:
            seen.add(o)
            unique.append(o)
    return "\n".join(unique)


def extract_correct_answer(page):
    answer = ""
    box = page.locator(".answer-box .correct-answer").first
    if box.count() == 0:
        box = page.locator(".answer-box").first

    try:
        tiles = box.locator(".SelectableTile").all() if box.count() > 0 else []
        for tile in tiles:
            sr      = tile.locator(".sr-only").first
            sr_text = sr.inner_text().strip() if sr.count() > 0 else ""
            cls     = tile.get_attribute("class") or ""
            if sr_text.lower().startswith("correct answer") or "selected" in cls.split():
                gm = tile.locator(".GeneticallyModified").first
                if gm.count() > 0:
                    try:
                        walker_js = f"""
                        el => {{
                            {_MATH_WALKER_JS}
                            walk(el);
                            return out.trim();
                        }}
                        """
                        val = gm.evaluate(walker_js)
                    except Exception:
                        val = " ".join(gm.inner_text().replace("\n", " ").split())
                    val = " ".join((val or "").replace("\n", " ").split())
                    if val:
                        answer = val
                        break
    except Exception:
        pass

    if not answer and box.count() > 0:
        try:
            inputs = box.locator("input.fillIn").all()
            vals = []
            for inp in inputs:
                v = inp.input_value() or inp.get_attribute("value") or ""
                if v.strip():
                    vals.append(v.strip())
            if vals:
                answer = ", ".join(vals)
        except Exception:
            pass

    if not answer and box.count() > 0:
        try:
            txt    = box.inner_text().replace("\n", " ").strip()
            answer = " ".join(txt.split())
        except Exception:
            pass

    return answer


# ─────────────────────────────────────────────────────────────────────────────
#  DIAGRAM EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def _safe_bbox(element):
    try:
        return element.bounding_box()
    except Exception:
        return None


def _is_junk_label(label: str) -> bool:
    return (label or "").strip().lower() in JUNK_LABELS


def _is_too_small(bb) -> bool:
    if bb is None:
        return True
    w, h = bb["width"], bb["height"]
    if w < 1 or h < 1:
        return True
    if w <= ICON_MAX and h <= ICON_MAX:
        return True
    return False


def _boxes_overlap(a, b, threshold=0.70):
    if a is None or b is None:
        return False
    ax1, ay1 = a["x"], a["y"]
    ax2, ay2 = ax1 + a["width"],  ay1 + a["height"]
    bx1, by1 = b["x"], b["y"]
    bx2, by2 = bx1 + b["width"],  by1 + b["height"]
    inter_x    = max(0, min(ax2, bx2) - max(ax1, bx1))
    inter_y    = max(0, min(ay2, by2) - max(ay1, by1))
    inter_area = inter_x * inter_y
    b_area     = b["width"] * b["height"]
    if b_area <= 0:
        return False
    return (inter_area / b_area) >= threshold


def _wait_for_element_painted(element, retries=6, delay_ms=300):
    for _ in range(retries):
        bb = _safe_bbox(element)
        if bb and bb["width"] > 1 and bb["height"] > 1:
            return bb
        time.sleep(delay_ms / 1000)
    return None


def _screenshot_element(element, path):
    try:
        element.screenshot(path=path)
        return True
    except Exception as e:
        print(f"     [!] screenshot failed ({path}): {e}")
        return False


def _collect_units_from_container(container, scope_label):
    for sel in (
        ".guide-counting-clickable-image-container",
        ".vector-image-wrapper[role='img']",
        ".shape",
        ".pie-chart",
    ):
        units = container.locator(sel).all()
        if units:
            print(f"       [{scope_label}] repeating container → {len(units)} units via '{sel}'")
            return units
    return [container]


def _question_has_diagram(page):
    """
    Gate: checks ONLY inside .secHdr and .secContent.
    Each scope part queried independently — no comma-joining with signals.
    """
    for scope_part in Q_SCOPE_PARTS:
        for signal in DIAGRAM_SIGNALS:
            try:
                if page.locator(f"{scope_part} {signal}").count() > 0:
                    return True
            except Exception:
                pass
    return False


def _tile_has_diagram(tile):
    """Gate: checks signals directly on a single option tile locator."""
    for signal in DIAGRAM_SIGNALS:
        try:
            if tile.locator(signal).count() > 0:
                return True
        except Exception:
            pass
    return False



def extract_diagrams_screenshots(page, question_index, skill_name):
    slug = skill_name.replace(" ", "_")[:40]
    ts = int(time.time())

    q_paths = []
    opt_paths = []
    if _question_has_diagram(page):
        q_paths = _extract_from_scope(
            page=page,
            scope_parts=Q_SCOPE_PARTS,
            root_locator=None,
            scope_label="Q",
            prefix=f"{slug}_q{question_index + 1}",
            ts=ts,
        )

    # _keep_largest intentionally removed: L1_INTEGRATED phantom de-dupe handles single-type
    # diagrams; L1_MULTI diagrams (number lines, pie charts) may have multiple real instances.

    # Find active question container to avoid duplicate phantom options
    _q_candidates = []
    _seen_q_coords = set()
    for _scope in (".question-and-submission-view", ".ixl-practice-crate"):
        try:
            for _el in page.locator(_scope).all():
                _bb = _safe_bbox(_el)
                if _bb is None or _bb["width"] < 1 or _bb["height"] < 1 or _bb["y"] < 0:
                    continue
                _coord = (round(_bb["x"]), round(_bb["y"]), round(_bb["width"]), round(_bb["height"]))
                if _coord not in _seen_q_coords:
                    _seen_q_coords.add(_coord)
                    _q_candidates.append((_bb["y"], _el))
        except Exception:
            pass

    active_tiles = []
    if _q_candidates:
        _q_candidates.sort(key=lambda t: t[0])
        _active_q = _q_candidates[0][1]
        active_tiles = _active_q.locator(".SelectableTile").all()

    for t_idx, tile in enumerate(active_tiles):
        tile_class = tile.get_attribute("class") or ""
        if "TEXT" in tile_class.split():
            continue
        if not _tile_has_diagram(tile):
            continue

        bb = _wait_for_element_painted(tile)
        if bb and bb["width"] > 2 and bb["height"] > 2:
            idx = len(opt_paths) + 1
            path = os.path.join(IMAGE_DIR, f"{slug}_q{question_index + 1}_opt{idx}_{ts}.png")
            if _screenshot_element(tile, path):
                opt_paths.append(path)
                print(f"       [Opt{idx}] SAVED tile screenshot: {path}")

    # Sorting drag-and-drop bins: find the live binsContainer (topmost y = active
    # question), then screenshot only its direct .bin children. IXL pre-renders
    # upcoming questions below the viewport, so a flat .binsContainer .bin query
    # returns N×3 results; we must scope to the single live container first.
    _bin_candidates = []
    _seen_bin_coords = set()
    for _scope in (".question-and-submission-view", ".ixl-practice-crate"):
        try:
            for _el in page.locator(f"{_scope} .binsContainer").all():
                _bb = _safe_bbox(_el)
                if _bb is None or _bb["width"] < 1 or _bb["height"] < 1 or _bb["y"] < 0:
                    continue
                _coord = (round(_bb["x"]), round(_bb["y"]),
                          round(_bb["width"]), round(_bb["height"]))
                if _coord not in _seen_bin_coords:
                    _seen_bin_coords.add(_coord)
                    _bin_candidates.append((_bb["y"], _el))
        except Exception:
            pass

    if _bin_candidates:
        _bin_candidates.sort(key=lambda t: t[0])
        _active_container = _bin_candidates[0][1]
        for b_idx, bin_el in enumerate(_active_container.locator(".bin").all()):
            try:
                bb = _wait_for_element_painted(bin_el)
                if bb is None or bb["width"] < 2 or bb["height"] < 2:
                    continue
                path = os.path.join(IMAGE_DIR,
                                    f"{slug}_q{question_index + 1}_bin{b_idx + 1}_{ts}.png")
                if _screenshot_element(bin_el, path):
                    opt_paths.append(path)
                    print(f"       [Bin{b_idx + 1}] SAVED bin screenshot: {path}")
            except Exception as e:
                print(f"     [!] bin screenshot failed: {e}")

    return q_paths, opt_paths
def _extract_from_scope(page, scope_parts, root_locator, scope_label, prefix, ts):
    paths      = []
    seen_boxes = []

    def already_captured(bb):
        for seen in seen_boxes:
            if _boxes_overlap(seen, bb):
                return True
        return False

    def do_screenshot(element, tag, layer="?", signal="?"):
        bb = _wait_for_element_painted(element)
        if _is_too_small(bb):
            print(f"       [{scope_label}] DROP-small  layer={layer} sig={signal} bb={bb}")
            return
        if already_captured(bb):
            print(f"       [{scope_label}] DROP-overlap layer={layer} sig={signal} "
                  f"bb=({round(bb['x'])},{round(bb['y'])},{round(bb['width'])},{round(bb['height'])})")
            return
        try:
            meta = element.evaluate(
                "el => ({tag: el.tagName, cls: el.getAttribute('class'), "
                "role: el.getAttribute('role'), al: el.getAttribute('aria-label')})"
            )
        except Exception as e:
            meta = {"tag": "?", "cls": f"<eval failed: {e}>", "role": "?", "al": "?"}
        idx  = len(paths) + 1
        path = os.path.join(IMAGE_DIR, f"{prefix}_{tag}_{ts}_{idx}.png")
        if _screenshot_element(element, path):
            paths.append(path)
            seen_boxes.append(bb)
            print(f"       [{scope_label}] SAVED#{idx} layer={layer} sig={signal} "
                  f"tag={meta.get('tag')} cls={meta.get('cls')} role={meta.get('role')} "
                  f"al={meta.get('al')} "
                  f"bb=({round(bb['x'])},{round(bb['y'])},{round(bb['width'])},{round(bb['height'])})")

    def _is_really_visible(el):
        try:
            if not el.is_visible():
                return False
        except Exception:
            return False
        bb = _safe_bbox(el)
        if bb is None:
            return False
        if bb["y"] < 0 or bb["x"] < 0:
            return False
        if bb["width"] < 1 or bb["height"] < 1:
            return False
        return True

    def get_elements(signal):
        if root_locator is not None:
            try:
                return [el for el in root_locator.locator(signal).all()
                        if _is_really_visible(el)]
            except Exception:
                return []
        else:
            results     = []
            seen_coords = set()
            for part in scope_parts:
                try:
                    els = page.locator(f"{part} {signal}").all()
                    for el in els:
                        if not _is_really_visible(el):
                            continue
                        bb = _safe_bbox(el)
                        coord_key = (round(bb["x"]), round(bb["y"]),
                                     round(bb["width"]), round(bb["height"]))
                        if coord_key not in seen_coords:
                            seen_coords.add(coord_key)
                            results.append(el)
                except Exception:
                    pass
            return results

    L1_INTEGRATED = {
        ".open-number-line", ".dc-fraction-strip-model",
        "table.old-table", "table.qTabularGrid",
        "svg:has(g.grid-region)",
        "div.table:has([data-testid='area-model-cell'])",
    }
    # These can have multiple real instances per question (e.g. two number lines for
    # equivalence questions, or a standalone pie chart) — no phantom de-dupe applied.
    L1_MULTI      = {".graphingBaseContainer", ".pie-chart"}
    L1_REPEATING  = {".multiplication-model-container", ".guide-counting-qm"}

    for container_sel in list(L1_INTEGRATED) + list(L1_MULTI) + list(L1_REPEATING):
        containers = get_elements(container_sel)

        # For integrated single-figure diagrams in the QUESTION scope, IXL
        # pre-renders upcoming-question copies stacked below the live one.
        # Keep only the topmost (smallest y) — that's the active diagram.
        # L1_MULTI types are exempt: a question may show several legitimately.
        if container_sel in L1_INTEGRATED and root_locator is None and len(containers) > 1:
            containers = sorted(
                containers,
                key=lambda c: (_safe_bbox(c) or {"y": 1e9})["y"]
            )[:1]
            print(f"       [{scope_label}] {container_sel}: kept topmost of "
                  f"{container_sel} (phantom de-dupe)")

        for container in containers:
            bb_container = _safe_bbox(container)
            if already_captured(bb_container):
                continue
            if container_sel in L1_INTEGRATED or container_sel in L1_MULTI:
                do_screenshot(container, "fig", layer="L1-int", signal=container_sel)
            else:
                units = _collect_units_from_container(container, scope_label)
                for unit in units:
                    do_screenshot(unit, "fig", layer="L1-rep", signal=container_sel)

    return paths

# ─────────────────────────────────────────────────────────────────────────────

def _screenshot_answer_bins(page, question_index, skill_name, ts):
    """After submission, screenshot each bin (with placed tiles) as the correct answer.
    Distinguishes the answer-state container from the question-state container by
    checking which binsContainer has tiles (draggableElement) placed inside its bins."""
    slug = skill_name.replace(" ", "_")[:40]
    paths = []
    try:
        all_candidates = []
        seen_coords = set()
        for scope in (".answer-box", ".question-and-submission-view", ".ixl-practice-crate"):
            try:
                for el in page.locator(f"{scope} .binsContainer").all():
                    bb = _safe_bbox(el)
                    if bb is None or bb["width"] < 1 or bb["height"] < 1 or bb["y"] < 0:
                        continue
                    coord = (round(bb["x"]), round(bb["y"]),
                             round(bb["width"]), round(bb["height"]))
                    if coord not in seen_coords:
                        seen_coords.add(coord)
                        all_candidates.append((bb["y"], el))
            except Exception:
                pass

        if not all_candidates:
            return paths

        # Prefer a container that has tiles placed inside bins (answer state).
        # The question-state container has empty binContentDropSlots; the answer-state
        # one has draggableElement tiles inside .bin elements.
        with_tiles = []
        for y, el in all_candidates:
            try:
                if el.locator(".bin .draggableElement").count() > 0:
                    with_tiles.append((y, el))
            except Exception:
                pass

        candidates = with_tiles if with_tiles else all_candidates
        candidates.sort(key=lambda t: t[0])
        container = candidates[0][1]

        for b_idx, bin_el in enumerate(container.locator(".bin").all()):
            bb = _wait_for_element_painted(bin_el)
            if bb is None or bb["width"] < 2 or bb["height"] < 2:
                continue
            path = os.path.join(IMAGE_DIR,
                                f"{slug}_q{question_index + 1}_ans_bin{b_idx + 1}_{ts}.png")
            if _screenshot_element(bin_el, path):
                paths.append(path)
                print(f"       [AnsBin{b_idx + 1}] SAVED answer bin: {path}")
    except Exception as e:
        print(f"     [!] answer bin screenshot failed: {e}")
    return paths


def extract_and_advance(page, category_name, skill_name, serial_tracker):
    previous_question_text = ""

    for i in range(QUESTIONS_PER_SKILL):
        print(f"  -> Processing Question {i+1}...")

        try:
            page.wait_for_selector(
                ".ixl-practice-crate, .math.section, .question-component",
                state="visible", timeout=15000
            )
            page.wait_for_timeout(800)

            if previous_question_text:
                for _ in range(20):
                    candidate = extract_question_text(page)
                    if candidate and candidate != previous_question_text:
                        break
                    page.wait_for_timeout(500)
        except Exception as e:
            print(f"     [!] Failed to stabilize DOM for question {i+1}: {e}")
            break

        question_text = extract_question_text(page)
        options_text  = extract_options(page)
        q_diagrams, opt_diagrams = extract_diagrams_screenshots(page, i, skill_name)
        previous_question_text = question_text

        if not question_text:
            print(f"     [!] WARNING: empty question text on Q{i+1}")

        correct_answer = ""
        try:
            submit_btn = page.locator(
                'button[data-cy="question-submit-button"], div.question button.submit'
            ).first
            submit_btn.wait_for(state="visible", timeout=10000)
            submit_btn.click()

            popup_btn = page.locator(
                'button[data-cy="incomplete-answer-popover-submit-button"]'
            ).first
            popup_btn.wait_for(state="visible", timeout=5000)
            popup_btn.click()

            page.wait_for_selector(".answer-box", state="visible", timeout=10000)
            page.wait_for_timeout(600)
            correct_answer = extract_correct_answer(page)
        except Exception as e:
            print(f"     [!] Could not read correct answer on Q{i+1}: {e}")

        ts_ans = int(time.time())
        ans_diagrams = _screenshot_answer_bins(page, i, skill_name, ts_ans)

        # row_data = columns A–G only (7 values)
        # H and I+ are handled by append_to_excel via image embedding
        row_vals = [
            serial_tracker[0] if i == 0 else "",
            category_name     if i == 0 else "",
            skill_name        if i == 0 else "",
            i + 1,
            question_text,
            options_text,
            correct_answer,
        ]
        append_to_excel(row_vals, q_diagrams, opt_diagrams,
                        ans_diagram_paths=ans_diagrams if ans_diagrams else None)

        if i < QUESTIONS_PER_SKILL - 1:
            try:
                got_it_btn = page.locator(
                    'button:has-text("Got it"), .next-problem button'
                ).first
                got_it_btn.wait_for(state="visible", timeout=15000)
                got_it_btn.click()
                got_it_btn.wait_for(state="hidden", timeout=10000)
            except Exception as e:
                print(f"     [!] Transition error on Q{i+1}: {e}")
                break
        else:
            try:
                got_it_btn = page.locator(
                    'button:has-text("Got it"), .next-problem button'
                ).first
                if got_it_btn.count() > 0 and got_it_btn.is_visible():
                    got_it_btn.click()
                    page.wait_for_timeout(500)
            except Exception:
                pass

    serial_tracker[0] += 1

def _choose_mode():
    print("\n  IXL Scraper\n")
    print("  [1]  Scrape entire website from scratch")
    print(f"  [2]  Start from a specific skill URL - {' '.join(START_URL.split('/')[-1].split('.')[0].split('-')).title()}")
    while True:
        choice = input("\n  Enter mode (1 or 2): ").strip()
        if choice in ("1", "2"):
            return int(choice)
        print("  Invalid choice. Please enter 1 or 2.")


def run_scraper():
    mode = _choose_mode()

    if mode == 2:
        # Normalise START_URL to just its path for comparison
        start_path = urllib.parse.urlparse(START_URL).path.rstrip("/")
        print(f"\n[Mode 2] Will skip skills until: {START_URL}")
        reached_start = False
    else:
        start_path    = None
        reached_start = True  # Mode 1: start immediately

    setup_dir()
    init_excel()

    with sync_playwright() as p:
        print("\nInitializing browser context...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Authenticating...")
        page.goto(LOGIN_URL)
        page.fill("input[type='email'], input[name='username']", EMAIL)
        page.fill("input[type='password'], input[name='password']", PASSWORD)
        page.keyboard.press("Enter")
        page.wait_for_timeout(6000)

        print(f"Navigating to target directory: {TARGET_URL}")
        page.goto(TARGET_URL)
        page.wait_for_selector("div.skill-tree-category", state="attached", timeout=15000)

        category_elements = page.locator("div.skill-tree-category").all()
        print(f"Found {len(category_elements)} category blocks.")

        serial_tracker = [1]

        for cat_index in range(len(category_elements)):
            cat_block = page.locator("div.skill-tree-category").nth(cat_index)

            header_loc = cat_block.locator(".skill-tree-skills-header").first
            if not header_loc.is_visible():
                continue

            code_text = (header_loc.locator(".category-code").inner_text().strip()
                         if header_loc.locator(".category-code").count() > 0 else "")
            name_text = (header_loc.locator(".category-name").inner_text().strip()
                         if header_loc.locator(".category-name").count() > 0 else "")
            full_category_name = f"{code_text} {name_text}".strip()

            skill_links = cat_block.locator("a.skill-tree-skill-link")
            skill_count = skill_links.count()

            for skill_index in range(skill_count):
                current_skill = (page.locator("div.skill-tree-category").nth(cat_index)
                                 .locator("a.skill-tree-skill-link").nth(skill_index))

                title_span = current_skill.locator("span.skill-tree-skill-name").first
                skill_name = (title_span.inner_text().strip()
                              if title_span.count() > 0
                              else current_skill.inner_text().strip())

                # Mode 2: skip skills until we hit the START_URL
                if not reached_start:
                    skill_href = (current_skill.get_attribute("href") or "").rstrip("/")
                    if skill_href == start_path:
                        reached_start = True
                        print(f"\n[Mode 2] Resuming from: {skill_name}")
                    else:
                        print(f"  [SKIP] {full_category_name} / {skill_name}")
                        continue

                print(f"\n[{full_category_name}] Entering skill: {skill_name}")

                current_skill.scroll_into_view_if_needed()
                current_skill.click()

                extract_and_advance(page, full_category_name, skill_name, serial_tracker)

                print("  -> Retreating to main directory...")
                try:
                    breadcrumb = page.locator(
                        'a.breadcrumb-link:has-text("Third grade"), '
                        'a.breadcrumb-link[href="/math/grade-3"], '
                        'a.breadcrumb-link[href="/maths/class-iii"]'
                    ).first
                    breadcrumb.wait_for(state="visible", timeout=10000)
                    breadcrumb.click()
                    page.wait_for_selector("div.skill-tree-category",
                                           state="attached", timeout=15000)
                except Exception as e:
                    print(f"  [!] Failed to use breadcrumb. Forcing URL reload. {e}")
                    page.goto(TARGET_URL)
                    page.wait_for_selector("div.skill-tree-category",
                                           state="attached", timeout=15000)

        print("\nAll topics processed. Closing driver.")
        browser.close()

if __name__ == "__main__":
    run_scraper()