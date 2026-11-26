import os
import re
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_KNOWLEDGE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "data", "trained_formats", "knowledge.json")
)

# --- STYLING CONSTANTS (Matched to Huliot templates) ---
FONT_FAMILY = "Calibri"
FONT_TITLE = Font(name=FONT_FAMILY, size=14, bold=True)
FONT_HEADER_MAIN = Font(name=FONT_FAMILY, size=11, bold=True)
FONT_HEADER_SUB = Font(name=FONT_FAMILY, size=11, bold=True)
FONT_REGULAR = Font(name=FONT_FAMILY, size=11, bold=False)
FONT_BOLD = Font(name=FONT_FAMILY, size=11, bold=True)

FILL_HEADER = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Soft blue-grey
FILL_HIGHLIGHT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Soft yellow

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

THIN_BORDER_SIDE = Side(style="thin", color="A6A6A6")
BORDER_ALL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)

NUM_FMT_CURRENCY = '_ "₹"* #,##0_ ;_ "₹"* -#,##0_ ;_ "₹"* "-"??_ ;_ @_ '
NUM_FMT_PERCENT = "0%"
NUM_FMT_INT = "#,##0"


def eval_qty(val):
    """Safely evaluates numeric quantities or formulas like =20*34."""
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str) and val.startswith("="):
        expr = val.lstrip("=")
        try:
            if re.match(r"^[\d\s\*\+\-\/\(\)]+$", expr):
                return int(eval(expr))
        except Exception:
            pass
    return None


def find_matching_sheet(wb, patterns):
    """Finds worksheet matching any of regex patterns."""
    for pattern in patterns:
        for sname in wb.sheetnames:
            if sname == "Summary Sheet":
                continue
            if re.search(pattern, sname, re.IGNORECASE):
                return wb[sname]
    return None


def discover_options_in_sheet(ws):
    """
    Dynamically scans rows 6–25 of any BOQ sheet to discover product options and their Total column.
    Works regardless of exact header row position (fixes Kingdom 57, Satva, Aster 108, etc.)
    Returns: dict of {option_name: col_letter} e.g. {'Ultra Silent': 'BS', 'HT Pro': 'BZ', 'Hybrid': 'CG'}
    """
    options = {}
    if ws is None:
        return options

    max_col = min(ws.max_column, 500)  # safety limit

    # Step 1: Find the option header row (where 'Ultra Silent', 'HT Pro', 'Hybrid' labels appear)
    # Require the keyword to appear in column >= 4 OR appear in 2+ columns on the same row
    # This avoids false matches like 'Material : HT Pro' in metadata rows (always col 1)
    opt_header_row = None
    for r in range(6, 30):
        keyword_cols = []
        for c in range(1, max_col + 1):
            txt = str(ws.cell(r, c).value or "").lower()
            if ("ultra silent" in txt or "ht pro" in txt or "htpro" in txt or
                    "ht-pro" in txt or "hybrid" in txt or "drain pro" in txt or "decilo" in txt):
                keyword_cols.append(c)
        # Must find keyword in col>=4 OR appear 2+ times (ruling out metadata rows in col 1)
        if keyword_cols and (max(keyword_cols) >= 4 or len(keyword_cols) >= 2):
            opt_header_row = r
            break

    if not opt_header_row:
        return options

    # Step 2: Find the 'Total' subheader row (within 4 rows below opt_header_row)
    total_row = None
    for r in range(opt_header_row, min(opt_header_row + 5, ws.max_row + 1)):
        for c in range(1, max_col + 1):
            if str(ws.cell(r, c).value or "").strip().lower() == "total":
                total_row = r
                break
        if total_row:
            break

    if not total_row:
        return options

    # Step 3: Map each option label to its Total column
    for c in range(1, max_col + 1):
        txt = " ".join([str(ws.cell(r, c).value or "") for r in range(opt_header_row, total_row)]).strip()
        opt_key = None
        if "internal - ultra silent" in txt.lower() or "hybrid" in txt.lower():
            opt_key = "Hybrid"
        elif "ultra silent" in txt.lower() or "u.s" in txt.lower():
            opt_key = "Ultra Silent"
        elif "ht pro" in txt.lower() or "htpro" in txt.lower() or "ht-pro" in txt.lower():
            opt_key = "HT Pro"
        elif "drain pro" in txt.lower():
            opt_key = "Drain Pro"
        elif "decilo" in txt.lower():
            opt_key = "Decilo"

        if opt_key and opt_key not in options:
            # Find the Total col in the window [c, c+12]
            for tc in range(c, min(c + 12, max_col + 1)):
                if str(ws.cell(total_row, tc).value or "").strip().lower() == "total":
                    options[opt_key] = get_column_letter(tc)
                    break

    return options


def is_summary_sheet(ws):
    """
    Detects whether a worksheet is a Summary Sheet by tab name OR internal content
    (e.g. 'Summary', 'Summary Sheet', 'Project Summary', or contains 'Discount Structure' / 'Grand Total Cost').
    Prevents Summary tabs from ever being misclassified as Tower/BOQ sheets.
    """
    if ws is None:
        return False
    sl = ws.title.strip().lower()
    if "summary" in sl:
        return True
    for r in range(1, min(26, ws.max_row + 1)):
        for c in range(1, min(16, ws.max_column + 1)):
            val = str(ws.cell(r, c).value or "").strip().lower()
            if "discount structure" in val or "grand total cost" in val or "huliot summary sheet" in val:
                return True
    return False


def classify_workbook_sheets(wb):
    """
    Logic-first sheet classifier that inspects internal BOQ structure rather than fragile tab name regexes.
    Returns dict:
      - 'toilet_split': ws or None (explicit split Toilet sheet like '(T)')
      - 'kitchen_split': ws or None (explicit split Kitchen sheet like '(K)')
      - 'diversion_sheets': list of ws (Diversion BOQ sheets)
      - 'rainwater_sheets': list of ws (Rainwater BOQ sheets)
      - 'typical_sheets': list of ws (Combined Typical Floor BOQ sheets - 1 per project or 1 per tower)
    """
    toilet_split = None
    kitchen_split = None
    diversion_sheets = []
    rainwater_sheets = []
    typical_sheets = []

    for sname in wb.sheetnames:
        sl = sname.strip().lower()
        if "design sheet" in sl or "flat matrix" in sl or sl == "sheet":
            continue
        ws = wb[sname]
        if is_summary_sheet(ws):
            continue
        opts = discover_options_in_sheet(ws)
        if not opts:
            continue  # Not a priced BOQ sheet

        if "diversion" in sl:
            diversion_sheets.append(ws)
        elif "rain" in sl or "rwp" in sl:
            rainwater_sheets.append(ws)
        elif re.search(r"\(\s*t\s*\)|\(\s*toilet\s*\)|typical.*toilet|g to \d+", sl):
            toilet_split = ws
        elif re.search(r"\(\s*k\s*\)|\(\s*kitchen\s*\)|typical.*kitchen", sl):
            kitchen_split = ws
        else:
            typical_sheets.append(ws)

    return {
        "toilet_split": toilet_split,
        "kitchen_split": kitchen_split,
        "diversion_sheets": diversion_sheets,
        "rainwater_sheets": rainwater_sheets,
        "typical_sheets": typical_sheets,
    }


def clean_tower_name(sname):
    """Normalizes any tower sheet name (e.g. 'Typical Floor BOQ Tower 1', 'T2', 'Typical Floor BOQTower(Central)') into a clean display name."""
    s = sname.strip()
    s_clean = re.sub(r"^typical\s*floor\s*boq\s*", "", s, flags=re.I).strip(" _-:")
    if not s_clean:
        s_clean = s
    s_clean = re.sub(r"\(([^)]+)\)", r" \1", s_clean).strip()
    s_clean = re.sub(r"\s+", " ", s_clean)
    tm = re.match(r"^t\s*[-_]?\s*(\d+)$", s_clean, re.I)
    if tm:
        s_clean = f"Tower {tm.group(1)}"
    return s_clean


def resolve_boq_cell(ws, target_option):
    """
    Discovers the exact cell coordinate for an option's Total in a single-shaft BOQ sheet (e.g. Toilet or Kitchen).
    """
    if ws is None:
        return None

    # Step 1: Discover Total Row
    total_row = None
    for r in range(15, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "").strip().lower()
        if "no of floor" in val or "no. of floor" in val:
            total_row = r
            break

    if not total_row:
        for r in range(ws.max_row, 14, -1):
            cell_val = str(ws.cell(r, 20).value or "")
            if cell_val.startswith("="):
                total_row = r
                break

    if not total_row:
        total_row = 63

    # Step 2: Discover Column
    opts = discover_options_in_sheet(ws)
    opt_col = opts.get(target_option)

    if not opt_col:
        # Fallback alias search
        for k, v in opts.items():
            if target_option.lower() in k.lower() or k.lower() in target_option.lower():
                opt_col = v
                break

    if not opt_col:
        return None

    sheet_ref = f"'{ws.title}'"
    return f"{sheet_ref}!${opt_col}${total_row}"


def extract_project_towers(wb, ref_ws=None, default_floors=20, default_toilets=500, default_kitchens=250, sheet_cls=None):
    """
    Discovers multi-tower configuration across sheets using logic-first inspection:
    0. Check Multi-Sheet Towers: multiple Typical Floor BOQ sheets regardless of sheet tab naming
    1. Check FLAT MATRIX sheet (like Astrana)
    2. Check 'Sheet' sheet (like Kalpataru)
    3. Check Typical Floor BOQ row 12 headers (like Avinea)
    Returns list of dicts: [{'name': 'Tower 1', 'floors': 35, 'toilets': 700, 'kitchens': 280}, ...]
    """
    towers = []
    if sheet_cls is None:
        sheet_cls = classify_workbook_sheets(wb)

    # 0. Check Multi-Sheet Towers: multiple Typical Floor BOQ sheets (regardless of naming)
    is_split_tk = (sheet_cls["toilet_split"] is not None and sheet_cls["kitchen_split"] is not None)
    typical_sheets = sheet_cls["typical_sheets"]
    if not is_split_tk and len(typical_sheets) > 1:
        def tower_sort_key(ws_obj):
            m = re.search(r"\d+", ws_obj.title)
            return (0, int(m.group(0))) if m else (1, ws_obj.title)
        typical_sheets_sorted = sorted(typical_sheets, key=tower_sort_key)

        for ws in typical_sheets_sorted:
            ts = ws.title
            t_row = None
            k_row = None
            flrs = default_floors
            for r in range(15, ws.max_row + 1):
                v = str(ws.cell(r, 1).value or "").lower()
                if "no of floor" in v or "no. of floor" in v:
                    t_row = r
                    m = re.search(r"(\d+)", v)
                    if m:
                        flrs = int(m.group(1))
                    break

            # Count shafts in row 13
            t_shafts = 0
            k_shafts = 0
            for c in range(4, ws.max_column + 1):
                sh = str(ws.cell(13, c).value or "").upper()
                if re.search(r"SH|T\d|S\d|TOILET", sh):
                    t_shafts += 1
                elif re.search(r"K\d|K-|KITCHEN", sh):
                    k_shafts += 1

            # Check immediate rows below t_row for Kitchen cost (avoiding Tax 18% / GST rows)
            if t_row:
                for r in range(t_row + 1, min(t_row + 8, ws.max_row + 1)):
                    lbl = str(ws.cell(r, 1).value or "").strip().lower()
                    row_txt = " ".join([str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)]).lower()
                    if ("kitchen" in lbl or ("*18%" in row_txt and "tax" not in row_txt and "gst" not in row_txt)) and not k_row:
                        k_row = r
                        break

            opts = discover_options_in_sheet(ws)
            clean_name = clean_tower_name(ts)
            towers.append({
                "name": clean_name,
                "sheet": ts,
                "floors": flrs,
                "toilets": t_shafts * flrs if t_shafts > 0 else (default_toilets // len(typical_sheets_sorted)),
                "kitchens": k_shafts * flrs if k_shafts > 0 else 0,
                "t_row": t_row,
                "k_row": k_row,
                "opts": opts,
                "is_multi_sheet": True
            })
        if towers:
            return towers

    # 1. FLAT MATRIX
    for s in wb.sheetnames:
        if 'FLAT' in s.upper() and 'MATRIX' in s.upper():
            ws = wb[s]
            for r in range(1, 20):
                d_val = str(ws.cell(r, 4).value or '').strip()
                if 'TOWER' in d_val.upper() and not '+' in d_val:
                    flr = ws.cell(r, 7).value
                    t_cnt = ws.cell(r, 8).value
                    k_cnt = ws.cell(r, 9).value
                    towers.append({
                        'name': d_val,
                        'floors': flr if isinstance(flr, (int, float)) else default_floors,
                        'toilets': t_cnt if isinstance(t_cnt, (int, float)) else (default_toilets // 4),
                        'kitchens': k_cnt if isinstance(k_cnt, (int, float)) else (default_kitchens // 4),
                        'flr_ref': f"'{s}'!G{r}",
                        't_ref': f"'{s}'!H{r}",
                        'k_ref': f"'{s}'!I{r}"
                    })
            if towers:
                break

    # 2. 'Sheet' (like Kalpataru)
    if not towers:
        for s in wb.sheetnames:
            if s.strip().lower() == 'sheet':
                ws = wb[s]
                for r in range(1, 25):
                    d_val = str(ws.cell(r, 4).value or '').strip()
                    e_val = str(ws.cell(r, 5).value or '').strip()
                    if 'TOWER' in d_val.upper() and not 'TOTAL' in e_val.upper() and d_val not in [x['name'] for x in towers]:
                        t_tot = None
                        k_tot = None
                        flrs = 0
                        t_cell_row = None
                        for tr in range(r, r+8):
                            if str(ws.cell(tr, 4).value or '').strip() == d_val or 'TOTAL FOR ' + d_val in str(ws.cell(tr, 5).value or '').upper():
                                if isinstance(ws.cell(tr, 8).value, (int, float)):
                                    flrs += int(ws.cell(tr, 8).value)
                            if 'TOTAL FOR ' + d_val in str(ws.cell(tr, 5).value or '').upper():
                                t_tot = ws.cell(tr, 9).value
                                k_tot = ws.cell(tr, 10).value
                                t_cell_row = tr
                                break
                        if t_tot is not None:
                            towers.append({
                                'name': d_val,
                                'floors': flrs if flrs > 0 else default_floors,
                                'toilets': t_tot if isinstance(t_tot, (int, float)) else 300,
                                'kitchens': k_tot if isinstance(k_tot, (int, float)) else 100,
                                'flr_ref': None,
                                't_ref': f"'{s}'!I{t_cell_row}" if t_cell_row else None,
                                'k_ref': f"'{s}'!J{t_cell_row}" if t_cell_row else None
                            })
                if towers:
                    break

    # 3. BOQ Sheet row 12 headers
    if not towers:
        rws = ref_ws
        if not rws:
            for s in wb.sheetnames:
                if 'typical' in s.lower() and 'boq' in s.lower():
                    rws = wb[s]
                    break
        if rws:
            cols_found = []
            for c in range(1, rws.max_column+1):
                v = str(rws.cell(12, c).value or '').strip()
                if 'TOWER' in v.upper():
                    cols_found.append((c, v))
            if len(cols_found) > 1:
                for c, name in cols_found:
                    clean_name = re.split(r'\(', name)[0].strip()
                    if clean_name not in [x['name'] for x in towers]:
                        towers.append({
                            'name': clean_name,
                            'floors': default_floors,
                            'toilets': None,
                            'kitchens': None
                        })

    # Fill fallbacks if partial
    if towers:
        n = len(towers)
        for t in towers:
            if not t.get('floors'):
                t['floors'] = default_floors
            if not t.get('toilets'):
                t['toilets'] = default_toilets // n
            if not t.get('kitchens'):
                t['kitchens'] = default_kitchens // n
    return towers


def analyze_workbook(wb_path, knowledge_db_path=None, original_filename=None):
    """
    Parses workbook to extract project metadata, archetype, detected options, and counts.
    If a matching entry is found in knowledge.json, learned discounts/quantities are used as defaults.
    """
    wb = openpyxl.load_workbook(wb_path, data_only=False)

    # Resolve knowledge_db_path relative to project data/trained_formats/
    if knowledge_db_path is None or knowledge_db_path == "knowledge.json":
        knowledge_db_path = DEFAULT_KNOWLEDGE_PATH
    has_summary = "Summary Sheet" in wb.sheetnames

    sheet_cls = classify_workbook_sheets(wb)
    toilet_ws = sheet_cls["toilet_split"]
    kitchen_ws = sheet_cls["kitchen_split"]
    diversion_sheets = sheet_cls["diversion_sheets"]
    diversion_ws = diversion_sheets[0] if diversion_sheets else None
    typical_sheets = sheet_cls["typical_sheets"]
    combined_ws = typical_sheets[0] if typical_sheets else None

    is_split_tk = (toilet_ws is not None and kitchen_ws is not None)
    is_hospital = ("hospital" in wb_path.lower()) or (toilet_ws is not None and kitchen_ws is None and len(typical_sheets) == 0)
    is_combined_single_sheet = (not is_split_tk and not is_hospital and len(typical_sheets) == 1)
    has_diversion = len(diversion_sheets) > 0

    ref_ws = toilet_ws or combined_ws or diversion_ws or (wb[wb.sheetnames[0]] if wb.sheetnames else None)

    # 1. Project Metadata
    project_name = "Huliot Project"
    consultant = ""
    client = ""
    system_design = "SINGLE STACK"
    num_floors_text = ""

    if ref_ws:
        for r in range(1, 15):
            for c in range(1, 10):
                txt = str(ref_ws.cell(r, c).value or "").strip()
                if "project name" in txt.lower():
                    m = re.split(r":|--", txt, 1)
                    if len(m) > 1 and m[1].strip():
                        project_name = m[1].strip()
                elif "consultant" in txt.lower():
                    m = re.split(r":|--", txt, 1)
                    if len(m) > 1:
                        consultant = m[1].strip()
                elif "client" in txt.lower():
                    m = re.split(r":|--", txt, 1)
                    if len(m) > 1:
                        client = m[1].strip()
                elif "system design" in txt.lower():
                    m = re.split(r":|--", txt, 1)
                    if len(m) > 1:
                        system_design = m[1].strip()

    # Fallback project name from filename
    if project_name == "Huliot Project":
        base = os.path.basename(wb_path).replace(".xlsx", "")
        base = re.sub(r"huliot\s*boq\s*(summary)?", "", base, flags=re.I).strip(" _-")
        base = re.sub(r"_\d{2}\.\d{2}\.\d{4}.*", "", base).strip(" _-")
        if base:
            project_name = base

    # 2. Extract Shafts and Floors
    t_shafts = 0
    k_shafts = 0
    floors_count = 20

    if ref_ws:
        for c in range(4, ref_ws.max_column + 1):
            sub_hdr = str(ref_ws.cell(14, c).value or "").strip().lower()
            if sub_hdr in ("total", "rate", "amount", "discount", "qty", "quantity"):
                break  # Stop at pricing columns so repeated option blocks are not double-counted
            val = str(ref_ws.cell(13, c).value or "").strip()
            if re.search(r"SH|T[-_]?\d|S[-_]?\d|TOILET", val, re.IGNORECASE):
                t_shafts += 1
            elif re.search(r"K[-_]?\d|K-|KITCHEN", val, re.IGNORECASE):
                k_shafts += 1

        for r in range(15, ref_ws.max_row + 1):
            val = str(ref_ws.cell(r, 1).value or "").strip()
            if "no of floor" in val.lower() or "no. of floor" in val.lower():
                num_floors_text = val
                m = re.search(r"(\d+)\s*(?:typical|floor|flr)", val, re.IGNORECASE)
                if m:
                    floors_count = int(m.group(1))
                elif "+" in val:
                    rhs = val.split(":", 1)[-1] if ":" in val else val
                    parts = [p.strip() for p in rhs.split("+") if p.strip()]
                    if parts:
                        floors_count = len(parts)
                break

    # 3. Extract Explicit Quantities from bottom notes (e.g. Nos of Toilet =20*34 or inline 'No of Toilets : 33  No of Kitchens : 28')
    est_toilets = None
    est_kitchens = None
    explicit_no_kitchen = False
    is_commercial = ("commercial" in (original_filename or wb_path).lower())
    if ref_ws:
        for r in range(40, ref_ws.max_row + 1):
            lbl = str(ref_ws.cell(r, 1).value or "").strip().lower()
            if "building composition" in lbl and "commercial" in lbl:
                is_commercial = True
            if ("nos of toilet" in lbl or "no of toilet" in lbl or "no. of toilet" in lbl) and not est_toilets:
                est_toilets = eval_qty(ref_ws.cell(r, 2).value)
                if not est_toilets:
                    mt = re.search(r"nos?\.?\s*of\s*toilets?\s*[:=\-]\s*(\d+)", lbl, re.IGNORECASE)
                    if mt:
                        est_toilets = int(mt.group(1))
            if ("nos of kitchen" in lbl or "no of kitchen" in lbl or "no. of kitchen" in lbl) and est_kitchens is None:
                if re.search(r"nos?\.?\s*of\s*kitchens?\s*[:=\-]\s*(n\.?a\.?|nil|none|0)\b", lbl, re.IGNORECASE):
                    explicit_no_kitchen = True
                    est_kitchens = 0
                else:
                    est_kitchens = eval_qty(ref_ws.cell(r, 2).value)
                    if est_kitchens is None:
                        mk = re.search(r"nos?\.?\s*of\s*kitchens?\s*[:=\-]\s*(\d+)", lbl, re.IGNORECASE)
                        if mk:
                            est_kitchens = int(mk.group(1))

    if not est_toilets:
        est_toilets = (t_shafts * floors_count) if t_shafts > 0 else 500
    if est_kitchens is None:
        if explicit_no_kitchen or (is_commercial and k_shafts == 0):
            est_kitchens = 0
        else:
            est_kitchens = (k_shafts * floors_count) if k_shafts > 0 else (est_toilets // 2)

    # 4. Discover Options
    discovered_opts = discover_options_in_sheet(ref_ws)
    detected_options = list(discovered_opts.keys())
    if not detected_options:
        detected_options = ["HT Pro", "Hybrid", "Ultra Silent"]
    else:
        # Standardize order: HT Pro, Hybrid, Ultra Silent, others
        ordered = []
        for o in ["HT Pro", "Hybrid", "Ultra Silent", "Drain Pro", "Decilo"]:
            if o in detected_options:
                ordered.append(o)
        for o in detected_options:
            if o not in ordered:
                ordered.append(o)
        detected_options = ordered

    # 5. Extract Baseline Discount from line item row 15
    baseline_discount = 0.55
    if ref_ws:
        for c in range(1, ref_ws.max_column + 1):
            if str(ref_ws.cell(14, c).value or "").strip().lower() == "discount":
                v = ref_ws.cell(15, c).value
                if isinstance(v, (int, float)) and 0 < v < 1:
                    baseline_discount = float(v)
                    break

    # 6. Extract Multi-Tower Configuration
    towers = extract_project_towers(wb, ref_ws, floors_count, est_toilets, est_kitchens, sheet_cls=sheet_cls)
    is_multi_sheet_towers = any(t.get("is_multi_sheet") for t in towers)

    if is_multi_sheet_towers and towers:
        # Clean Project Name: remove specific 'Tower 1' or 'Tower 2' and make 'VTP Verve (Tower 2 to 6)' or 'VTP Sierra (7 Towers)'
        clean_pname = re.sub(r"(?:tower|t)\s*[-_]?\s*[\d,\s-]+", "", project_name, flags=re.I).strip(" _-:")
        if not clean_pname:
            clean_pname = project_name
        first_m = re.search(r"\d+", towers[0]["name"])
        last_m = re.search(r"\d+", towers[-1]["name"])
        if first_m and last_m:
            project_name = f"{clean_pname} (Tower {first_m.group(0)} to {last_m.group(0)})"
        else:
            project_name = f"{clean_pname} ({len(towers)} Towers)"

        # Sum up toilets and kitchens across all tower sheets
        est_toilets = sum(t["toilets"] for t in towers)
        est_kitchens = sum(t["kitchens"] for t in towers)
        floors_count = max(t["floors"] for t in towers)
        num_floors_text = f"{floors_count} Typical Floors"

        # Detected options union across all tower sheets
        all_opts = set()
        for t in towers:
            all_opts.update(t.get("opts", {}).keys())
        if all_opts:
            ordered = []
            for o in ["HT Pro", "Hybrid", "Ultra Silent", "Drain Pro", "Decilo"]:
                if o in all_opts:
                    ordered.append(o)
            for o in all_opts:
                if o not in ordered:
                    ordered.append(o)
            detected_options = ordered

    # 7. Active Learning: Check knowledge.json for learned discounts / quantities
    knowledge_match = False
    knowledge_source = None
    knowledge_match_type = None
    learned_discounts = None
    k_entry, match_type = find_knowledge_match(wb_path, wb, knowledge_db_path=knowledge_db_path, original_filename=original_filename)
    if k_entry:
        knowledge_match = True
        knowledge_source = k_entry.get("project_title") or k_entry.get("filename")
        knowledge_match_type = match_type
        # Only auto-apply discounts & quantities from verified full_summary or generated_summary entries with strong match
        if k_entry.get("training_type", "full_summary") in ("full_summary", "generated_summary") and match_type in ("exact_filename", "sheet_fingerprint"):
            k_discs = k_entry.get("discounts", {})
            if k_discs:
                learned_discounts = {}
                for tag in ["P", "F", "M", "L", "C"]:
                    if tag in k_discs and isinstance(k_discs[tag].get("discount"), (int, float)):
                        learned_discounts[tag] = float(k_discs[tag]["discount"])
                if "P" in learned_discounts:
                    baseline_discount = learned_discounts["P"]
            k_qty = k_entry.get("quantities", {})
            if isinstance(k_qty.get("toilets"), (int, float)) and k_qty["toilets"] > 0:
                est_toilets = int(k_qty["toilets"])
            if isinstance(k_qty.get("kitchens"), (int, float)) and k_qty["kitchens"] > 0:
                est_kitchens = int(k_qty["kitchens"])

    # 8. Health Check & Warnings for UI Live Preview
    warnings = []
    rainwater_sheets = sheet_cls.get("rainwater_sheets", [])
    if len(typical_sheets) == 0 and toilet_ws is None and len(diversion_sheets) == 0:
        warnings.append("कोई भी Typical या Diversion BOQ शीट नहीं मिली। कृपया शीट फ़ॉर्मेट जांचें।")
    if not est_toilets or est_toilets <= 0:
        warnings.append("Toilets की गिनती 0 डिटेक्ट हुई है — कृपया मैन्युअल रूप से जांचें।")
    if est_toilets and est_kitchens and est_toilets > 0 and est_kitchens > 0:
        if est_kitchens > est_toilets * 1.5 or est_toilets > est_kitchens * 6:
            warnings.append("Toilets per Floor और Kitchens per Floor match नहीं हो रहे, कृपया Source BOQ चेक करें")
    if not detected_options:
        warnings.append("कोई भी Product Option (HT Pro / Hybrid / Ultra Silent) डिटेक्ट नहीं हुआ।")

    archetype = "Hospital / Commercial" if is_hospital else (
        "Split (T)+(K)" if is_split_tk else (
            "Multi-Sheet Towers" if is_multi_sheet_towers else (
                "Single Combined Sheet" if is_combined_single_sheet else "Standard"
            )
        )
    )

    # 9. Format Confidence Score (Compare against saved formats in knowledge.json)
    arch_key = "hospital" if is_hospital else (
        "split_tk" if is_split_tk else (
            "multi_sheet_towers" if is_multi_sheet_towers else "combined"
        )
    )
    format_labels = {
        "split_tk": "Format A (Split T+K)",
        "combined": "Format B (Combined)",
        "multi_sheet_towers": "Format C (Multi-Tower)",
        "hospital": "Format D (Hospital/Commercial)",
    }
    conf_res = compute_format_confidence(
        wb_path=wb_path,
        wb=wb,
        arch_key=arch_key,
        detected_options=detected_options,
        has_diversion=has_diversion,
        knowledge_db_path=knowledge_db_path,
        original_filename=original_filename,
        format_labels=format_labels,
    )

    return {
        "project_name": project_name,
        "consultant": consultant,
        "client": client,
        "system_design": system_design,
        "archetype": archetype,
        "format_confidence": conf_res["confidence"],
        "detected_format_label": conf_res["format_label"],
        "matched_format_project": conf_res["matched_project"],
        "is_new_format": conf_res["is_new_format"],
        "has_summary_sheet": has_summary,
        "is_hospital": is_hospital,
        "is_combined_single_sheet": is_combined_single_sheet,
        "is_multi_sheet_towers": is_multi_sheet_towers,
        "has_diversion": has_diversion,
        "has_rainwater": len(rainwater_sheets) > 0,
        "baseline_discount": baseline_discount,
        "learned_discounts": learned_discounts,
        "knowledge_match": knowledge_match,
        "knowledge_source": knowledge_source,
        "knowledge_match_type": knowledge_match_type,
        "sheets_found": {
            "toilet_sheet": toilet_ws.title if toilet_ws else None,
            "kitchen_sheet": kitchen_ws.title if kitchen_ws else None,
            "diversion_sheet": diversion_ws.title if diversion_ws else None,
            "combined_sheet": combined_ws.title if combined_ws else None
        },
        "sheet_classification": {
            "typical_count": len(typical_sheets) + (1 if toilet_ws else 0),
            "diversion_count": len(diversion_sheets),
            "rainwater_count": len(rainwater_sheets),
            "is_split_tk": is_split_tk
        },
        "shafts": {"toilet_shafts": t_shafts, "kitchen_shafts": k_shafts},
        "floors_info": num_floors_text,
        "floors_count": floors_count,
        "estimated_toilets": est_toilets,
        "estimated_kitchens": est_kitchens,
        "detected_options": detected_options,
        "towers": towers,
        "num_towers": len(towers) if len(towers) > 1 else 1,
        "warnings": warnings
    }

def generate_summary_sheet(wb_path, output_path, project_name=None, num_toilets=None, num_kitchens=None, num_floors=None, num_towers=None, discounts=None, options=None, original_filename=None, include_rainwater=True):
    """
    Constructs the executive Summary Sheet with 100% dynamic formulas and zero errors.
    Handles Multi-sheet projects, Single combined-sheet projects, Hospital projects, and Multi-Tower projects.
    Includes unit economics: Typical Floor Cost, Per Toilet (Bathroom) Cost, Per Flat Cost, and Tower-wise breakdowns.
    """
    wb = openpyxl.load_workbook(wb_path)

    if "Summary Sheet" in wb.sheetnames:
        del wb["Summary Sheet"]

    ws = wb.create_sheet(title="Summary Sheet", index=0)
    ws.views.sheetView[0].showGridLines = True

    # Sheet discovery (Logic-first)
    sheet_cls = classify_workbook_sheets(wb)
    toilet_ws = sheet_cls["toilet_split"]
    kitchen_ws = sheet_cls["kitchen_split"]
    diversion_sheets = sheet_cls["diversion_sheets"]
    diversion_ws = diversion_sheets[0] if diversion_sheets else None
    rainwater_sheets = sheet_cls.get("rainwater_sheets", [])
    typical_sheets = sheet_cls["typical_sheets"]
    combined_ws = typical_sheets[0] if typical_sheets else None

    is_split_tk = (toilet_ws is not None and kitchen_ws is not None)
    is_hospital = ("hospital" in wb_path.lower()) or (toilet_ws is not None and kitchen_ws is None and len(typical_sheets) == 0)
    is_combined = (not is_split_tk and not is_hospital and len(typical_sheets) == 1)
    has_diversion = len(diversion_sheets) > 0
    use_rainwater = bool(include_rainwater and len(rainwater_sheets) > 0)

    meta = analyze_workbook(wb_path, original_filename=original_filename)
    if not project_name:
        project_name = meta["project_name"]
    if num_toilets is None:
        num_toilets = meta["estimated_toilets"]
    if num_kitchens is None:
        num_kitchens = meta["estimated_kitchens"]
    if not num_floors:
        num_floors = meta.get("floors_count", 20)
    if not options:
        options = meta["detected_options"]

    # Treat Commercial / No-Kitchen projects identically to Hospital (no Kitchen Cost columns)
    if int(num_kitchens or 0) == 0:
        is_hospital = True

    towers = meta.get("towers", [])
    if not num_towers:
        num_towers = meta.get("num_towers", 1)

    base_disc = meta.get("baseline_discount", 0.55)
    default_discounts = {
        "P": base_disc,
        "F": base_disc,
        "M": base_disc,
        "L": min(base_disc, 0.45) if base_disc >= 0.45 else base_disc,
        "C": min(base_disc, 0.45) if base_disc >= 0.45 else base_disc,
    }
    if meta.get("learned_discounts"):
        default_discounts.update(meta["learned_discounts"])
    if discounts:
        default_discounts.update(discounts)

    # 1. TITLE BLOCK (Row 2)
    start_col = "C"
    if is_hospital:
        end_col = "L" if use_rainwater else "K"
    elif has_diversion:
        end_col = "O" if use_rainwater else "N"
    else:
        end_col = "M" if use_rainwater else "L"
    ws.merge_cells(f"{start_col}2:{end_col}2")
    title_cell = ws[f"{start_col}2"]
    title_cell.value = f"Huliot Summary Sheet - {project_name}"
    title_cell.font = FONT_TITLE
    title_cell.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 31.2

    # 2. MAIN SUMMARY TABLE HEADERS (Rows 3 & 4)
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 24

    scope_parts = ["Typical Floor"]
    if has_diversion:
        scope_parts.append("Diversion")
    if use_rainwater:
        scope_parts.append("Rainwater")
    scope_name = " + ".join(scope_parts)

    if is_hospital:
        ws.merge_cells("C3:C4"); ws["C3"] = "Option"
        ws.merge_cells("D3:D4"); ws["D3"] = "Tower"
        ws.merge_cells("E3:E4"); ws["E3"] = "Tower Cost"
        ws.merge_cells("F3:F4"); ws["F3"] = "Per Toilet Cost"
        ws.merge_cells("G3:G4"); ws["G3"] = "Diversion Cost"
        ws.merge_cells("H3:H4"); ws["H3"] = "Per Toilet Cost"
        if use_rainwater:
            ws.merge_cells("I3:I4"); ws["I3"] = "Rainwater Cost"
            ws.merge_cells("J3:J4"); ws["J3"] = "Grand Total Cost"
            ws.merge_cells("K3:K4"); ws["K3"] = "Typical Floor Cost"
            ws.merge_cells("L3:L4"); ws["L3"] = "Per Toilet Cost"
        else:
            ws.merge_cells("I3:I4"); ws["I3"] = "Grand Total Cost"
            ws.merge_cells("J3:J4"); ws["J3"] = "Typical Floor Cost"
            ws.merge_cells("K3:K4"); ws["K3"] = "Per Toilet Cost"
    elif has_diversion:
        ws.merge_cells("C3:C4"); ws["C3"] = "Option"
        ws.merge_cells("D3:D4"); ws["D3"] = "Tower"
        ws.merge_cells("E3:F3"); ws["E3"] = "Toilet Cost"; ws["E4"] = "Tower Cost"; ws["F4"] = "Per Toilet Cost"
        ws.merge_cells("G3:H3"); ws["G3"] = "Kitchen Cost"; ws["G4"] = "Tower Cost"; ws["H4"] = "Per Kitchen Cost"
        ws.merge_cells("I3:J3"); ws["I3"] = "Diversion Cost"; ws["I4"] = "Tower Cost"; ws["J4"] = "Per Toilet Cost"
        if use_rainwater:
            ws.merge_cells("K3:K4"); ws["K3"] = "Rainwater Cost"
            ws.merge_cells("L3:L4"); ws["L3"] = "Grand Total Cost"
            ws.merge_cells("M3:M4"); ws["M3"] = "Typical Floor Cost"
            ws.merge_cells("N3:N4"); ws["N3"] = "Per Toilet Cost"
            ws.merge_cells("O3:O4"); ws["O3"] = "Per Flat Cost"
        else:
            ws.merge_cells("K3:K4"); ws["K3"] = "Grand Total Cost"
            ws.merge_cells("L3:L4"); ws["L3"] = "Typical Floor Cost"
            ws.merge_cells("M3:M4"); ws["M3"] = "Per Toilet Cost"
            ws.merge_cells("N3:N4"); ws["N3"] = "Per Flat Cost"
    else:
        ws.merge_cells("C3:C4"); ws["C3"] = "Option"
        ws.merge_cells("D3:D4"); ws["D3"] = "Tower"
        ws.merge_cells("E3:F3"); ws["E3"] = "Toilet Cost"; ws["E4"] = "Tower Cost"; ws["F4"] = "Per Toilet Cost"
        ws.merge_cells("G3:H3"); ws["G3"] = "Kitchen Cost"; ws["G4"] = "Tower Cost"; ws["H4"] = "Per Kitchen Cost"
        if use_rainwater:
            ws.merge_cells("I3:I4"); ws["I3"] = "Rainwater Cost"
            ws.merge_cells("J3:J4"); ws["J3"] = "Grand Total Cost"
            ws.merge_cells("K3:K4"); ws["K3"] = "Typical Floor Cost"
            ws.merge_cells("L3:L4"); ws["L3"] = "Per Toilet Cost"
            ws.merge_cells("M3:M4"); ws["M3"] = "Per Flat Cost"
        else:
            ws.merge_cells("I3:I4"); ws["I3"] = "Grand Total Cost"
            ws.merge_cells("J3:J4"); ws["J3"] = "Typical Floor Cost"
            ws.merge_cells("K3:K4"); ws["K3"] = "Per Toilet Cost"
            ws.merge_cells("L3:L4"); ws["L3"] = "Per Flat Cost"

    for r in [3, 4]:
        for c in range(3, openpyxl.utils.column_index_from_string(end_col) + 1):
            cell = ws.cell(r, c)
            cell.font = FONT_HEADER_MAIN
            cell.alignment = ALIGN_CENTER
            cell.fill = FILL_HEADER
            cell.border = BORDER_ALL

    # Parameter Rows (Fixed, no collision)
    row_toilets = 18
    row_kitchens = 19
    row_floors = 20
    row_towers = 21

    # 3. RESOLVE SOURCE REFERENCES
    # If single combined sheet:
    comb_opts = {}
    comb_t_row = None
    comb_k_row = None
    comb_tot_row = None
    if is_combined and combined_ws:
        comb_opts = discover_options_in_sheet(combined_ws)
        # Step 1: Find toilet_row (marked by 'no of floor')
        for r in range(15, combined_ws.max_row + 1):
            val = str(combined_ws.cell(r, 1).value or "").strip().lower()
            if "no of floor" in val or "no. of floor" in val:
                comb_t_row = r
                break

        # Step 2: Dynamically inspect immediate rows below toilet_row for Kitchen & Total (skipping Tax/GST rows)
        if comb_t_row and not is_hospital:
            for r in range(comb_t_row + 1, min(comb_t_row + 8, combined_ws.max_row + 1)):
                lbl = str(combined_ws.cell(r, 1).value or "").strip().lower()
                row_txt = " ".join([str(combined_ws.cell(r, c).value or "") for c in range(1, combined_ws.max_column + 1)]).lower()
                is_tax_row = ("tax" in row_txt or "gst" in row_txt)
                row_has_18 = any("*18%" in str(combined_ws.cell(r, c).value or "") for c in range(1, combined_ws.max_column + 1))
                if not is_tax_row and ("kitchen" in lbl or row_has_18) and not comb_k_row:
                    comb_k_row = r
                elif ("building composition" in lbl or "grand total" in lbl) and not comb_tot_row:
                    comb_tot_row = r

    # 4. FILL DATA ROWS
    cur_row = 5
    for opt in options:
        ws.row_dimensions[cur_row].height = 20
        ws[f"C{cur_row}"] = opt
        ws[f"D{cur_row}"] = scope_name
        ws[f"C{cur_row}"].font = FONT_BOLD
        ws[f"D{cur_row}"].font = FONT_REGULAR
        ws[f"C{cur_row}"].border = BORDER_ALL
        ws[f"D{cur_row}"].border = BORDER_ALL

        # Discover references based on archetype
        is_multi_sheet_towers = any(t.get("is_multi_sheet") for t in towers)

        if is_multi_sheet_towers:
            t_cells = [f"'{tw['sheet']}'!${tw['opts'][opt]}${tw['t_row']}" for tw in towers if tw.get('opts', {}).get(opt) and tw.get('t_row')]
            k_cells = [f"'{tw['sheet']}'!${tw['opts'][opt]}${tw['k_row']}" for tw in towers if tw.get('opts', {}).get(opt) and tw.get('k_row')]
            t_ref = f"SUM({','.join(t_cells)})" if t_cells else None
            k_ref = f"SUM({','.join(k_cells)})" if k_cells else None
        elif is_combined and combined_ws:
            opt_col = comb_opts.get(opt)
            if not opt_col:
                for k, v in comb_opts.items():
                    if opt.lower() in k.lower() or k.lower() in opt.lower():
                        opt_col = v
                        break
            sheet_title = f"'{combined_ws.title}'"
            t_ref = f"{sheet_title}!${opt_col}${comb_t_row}" if (opt_col and comb_t_row) else None
            k_ref = f"{sheet_title}!${opt_col}${comb_k_row}" if (opt_col and comb_k_row) else None
        else:
            t_ref = resolve_boq_cell(toilet_ws, opt)
            k_ref = resolve_boq_cell(kitchen_ws, opt)

        # Resolve Diversion if present (independent of typical floor sheet layout)
        div_ref = "0"
        if has_diversion and diversion_sheets:
            d_cells = []
            for d_ws in diversion_sheets:
                c_ref = resolve_boq_cell(d_ws, opt)
                if not c_ref:
                    c_ref = resolve_boq_cell(d_ws, "Ultra Silent")
                if c_ref:
                    d_cells.append(c_ref)
            if len(d_cells) > 1:
                div_ref = f"SUM({','.join(d_cells)})"
            elif len(d_cells) == 1:
                div_ref = d_cells[0]

        # Resolve Rainwater if enabled and present
        rwp_ref = "0"
        if use_rainwater and rainwater_sheets:
            r_cells = []
            for r_ws in rainwater_sheets:
                c_ref = resolve_boq_cell(r_ws, opt)
                if not c_ref:
                    c_ref = resolve_boq_cell(r_ws, "Ultra Silent")
                if not c_ref:
                    r_opts = discover_options_in_sheet(r_ws)
                    if r_opts:
                        c_ref = resolve_boq_cell(r_ws, list(r_opts.keys())[0])
                if c_ref:
                    r_cells.append(c_ref)
            if len(r_cells) > 1:
                rwp_ref = f"SUM({','.join(r_cells)})"
            elif len(r_cells) == 1:
                rwp_ref = r_cells[0]

        # SELF-VALIDATION & FALLBACK SEARCH: Never silently write '=0'
        if not t_ref:
            # Search all sheets for this option
            for sname in wb.sheetnames:
                if sname != "Summary Sheet":
                    ws_try = wb[sname]
                    opts_try = discover_options_in_sheet(ws_try)
                    if opt in opts_try:
                        for tr in range(15, ws_try.max_row + 1):
                            lbl = str(ws_try.cell(tr, 1).value or "").lower()
                            if "no of floor" in lbl or "total" in lbl:
                                t_ref = f"'{ws_try.title}'!${opts_try[opt]}${tr}"
                                break
                    if t_ref:
                        break

        if not k_ref and not is_hospital:
            # Search for kitchen row in the same sheet as toilet
            if t_ref and "!" in t_ref:
                target_sheet = t_ref.split("!")[0].strip("'")
                if target_sheet in wb.sheetnames:
                    ws_target = wb[target_sheet]
                    col_match = re.search(r"\$([A-Z]+)\$(\d+)", t_ref)
                    if col_match:
                        c_let = col_match.group(1)
                        t_row_idx = int(col_match.group(2))
                        # Check the 6 rows immediately below toilet row
                        for kr in range(t_row_idx + 1, min(t_row_idx + 7, ws_target.max_row + 1)):
                            lbl = str(ws_target.cell(kr, 1).value or "").lower()
                            row_has_18 = any("*18%" in str(ws_target.cell(kr, c).value or "") for c in range(1, ws_target.max_column + 1))
                            if ("kitchen" in lbl or row_has_18) and "nos of" not in lbl:
                                k_ref = f"'{target_sheet}'!${c_let}${kr}"
                                break

        t_ref = t_ref or "0"
        k_ref = k_ref or "0"

        if not is_hospital and not has_diversion:
            ws[f"E{cur_row}"] = f"={t_ref}"
            ws[f"F{cur_row}"] = f"=E{cur_row}/$D${row_toilets}"
            ws[f"G{cur_row}"] = f"={k_ref}"
            ws[f"H{cur_row}"] = f"=G{cur_row}/$D${row_kitchens}"
            if use_rainwater:
                ws[f"I{cur_row}"] = f"={rwp_ref}"
                ws[f"J{cur_row}"] = f"=E{cur_row}+G{cur_row}+I{cur_row}"
                ws[f"K{cur_row}"] = f"=J{cur_row}/$D${row_floors}"
                ws[f"L{cur_row}"] = f"=J{cur_row}/$D${row_toilets}"
                ws[f"M{cur_row}"] = f"=J{cur_row}/$D${row_kitchens}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K", "L", "M"]
            else:
                ws[f"I{cur_row}"] = f"=E{cur_row}+G{cur_row}"
                ws[f"J{cur_row}"] = f"=I{cur_row}/$D${row_floors}"
                ws[f"K{cur_row}"] = f"=I{cur_row}/$D${row_toilets}"
                ws[f"L{cur_row}"] = f"=I{cur_row}/$D${row_kitchens}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K", "L"]

            for col in cols_to_fmt:
                cell = ws[f"{col}{cur_row}"]
                cell.font = FONT_REGULAR
                cell.alignment = ALIGN_RIGHT
                cell.border = BORDER_ALL
                cell.number_format = NUM_FMT_CURRENCY

        elif has_diversion and not is_hospital:
            ws[f"E{cur_row}"] = f"={t_ref}"
            ws[f"F{cur_row}"] = f"=E{cur_row}/$D${row_toilets}"
            ws[f"G{cur_row}"] = f"={k_ref}"
            ws[f"H{cur_row}"] = f"=G{cur_row}/$D${row_kitchens}"
            ws[f"I{cur_row}"] = f"={div_ref}"
            ws[f"J{cur_row}"] = f"=I{cur_row}/$D${row_toilets}"
            if use_rainwater:
                ws[f"K{cur_row}"] = f"={rwp_ref}"
                ws[f"L{cur_row}"] = f"=E{cur_row}+G{cur_row}+I{cur_row}+K{cur_row}"
                ws[f"M{cur_row}"] = f"=L{cur_row}/$D${row_floors}"
                ws[f"N{cur_row}"] = f"=L{cur_row}/$D${row_toilets}"
                ws[f"O{cur_row}"] = f"=L{cur_row}/$D${row_kitchens}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
            else:
                ws[f"K{cur_row}"] = f"=E{cur_row}+G{cur_row}+I{cur_row}"
                ws[f"L{cur_row}"] = f"=K{cur_row}/$D${row_floors}"
                ws[f"M{cur_row}"] = f"=K{cur_row}/$D${row_toilets}"
                ws[f"N{cur_row}"] = f"=K{cur_row}/$D${row_kitchens}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]

            for col in cols_to_fmt:
                cell = ws[f"{col}{cur_row}"]
                cell.font = FONT_REGULAR
                cell.alignment = ALIGN_RIGHT
                cell.border = BORDER_ALL
                cell.number_format = NUM_FMT_CURRENCY

        elif is_hospital:
            ws[f"E{cur_row}"] = f"={t_ref}"
            ws[f"F{cur_row}"] = f"=E{cur_row}/$D${row_toilets}"
            ws[f"G{cur_row}"] = f"={div_ref}"
            ws[f"H{cur_row}"] = f"=G{cur_row}/$D${row_toilets}"
            if use_rainwater:
                ws[f"I{cur_row}"] = f"={rwp_ref}"
                ws[f"J{cur_row}"] = f"=E{cur_row}+G{cur_row}+I{cur_row}"
                ws[f"K{cur_row}"] = f"=J{cur_row}/$D${row_floors}"
                ws[f"L{cur_row}"] = f"=J{cur_row}/$D${row_toilets}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K", "L"]
            else:
                ws[f"I{cur_row}"] = f"=E{cur_row}+G{cur_row}"
                ws[f"J{cur_row}"] = f"=I{cur_row}/$D${row_floors}"
                ws[f"K{cur_row}"] = f"=I{cur_row}/$D${row_toilets}"
                cols_to_fmt = ["E", "F", "G", "H", "I", "J", "K"]

            for col in cols_to_fmt:
                cell = ws[f"{col}{cur_row}"]
                cell.font = FONT_REGULAR
                cell.alignment = ALIGN_RIGHT
                cell.border = BORDER_ALL
                cell.number_format = NUM_FMT_CURRENCY

        cur_row += 1

    # 5. DISCOUNT STRUCTURE (Rows 11 to 16)
    ws.merge_cells("C11:D11")
    ws["C11"] = "Discount Structure"
    ws["C11"].font = FONT_HEADER_SUB
    ws["C11"].alignment = ALIGN_CENTER
    ws["C11"].fill = FILL_HEADER
    ws["C11"].border = BORDER_ALL
    ws["D11"].border = BORDER_ALL

    categories = [
        ("P", "Pipe", default_discounts.get("P", 0.55)),
        ("F", "Fittings", default_discounts.get("F", 0.55)),
        ("M", "MFT", default_discounts.get("M", 0.55)),
        ("L", "Lubricant", default_discounts.get("L", 0.45)),
        ("C", "Clamps", default_discounts.get("C", 0.45))
    ]

    r_idx = 12
    for tag, name, val in categories:
        ws[f"B{r_idx}"] = tag
        ws[f"C{r_idx}"] = name
        ws[f"D{r_idx}"] = val

        ws[f"C{r_idx}"].font = FONT_REGULAR
        ws[f"D{r_idx}"].font = FONT_BOLD
        ws[f"C{r_idx}"].border = BORDER_ALL
        ws[f"D{r_idx}"].border = BORDER_ALL
        ws[f"D{r_idx}"].number_format = NUM_FMT_PERCENT
        ws[f"D{r_idx}"].alignment = ALIGN_RIGHT
        r_idx += 1

    ws.column_dimensions["B"].width = 0.0
    ws.column_dimensions["B"].hidden = True

    # 6. QUANTITY INPUTS BLOCK (Rows 18 to 21)
    ws[f"C{row_toilets}"] = "No of Toilets"
    ws[f"D{row_toilets}"] = int(num_toilets)
    ws[f"C{row_toilets}"].font = FONT_BOLD
    ws[f"D{row_toilets}"].font = FONT_BOLD
    ws[f"C{row_toilets}"].border = BORDER_ALL
    ws[f"D{row_toilets}"].border = BORDER_ALL
    ws[f"D{row_toilets}"].number_format = NUM_FMT_INT
    ws[f"D{row_toilets}"].alignment = ALIGN_RIGHT

    ws[f"C{row_kitchens}"] = "No of Kitchens"
    ws[f"D{row_kitchens}"] = int(num_kitchens)
    ws[f"C{row_kitchens}"].font = FONT_BOLD
    ws[f"D{row_kitchens}"].font = FONT_BOLD
    ws[f"C{row_kitchens}"].border = BORDER_ALL
    ws[f"D{row_kitchens}"].border = BORDER_ALL
    ws[f"D{row_kitchens}"].number_format = NUM_FMT_INT
    ws[f"D{row_kitchens}"].alignment = ALIGN_RIGHT

    ws[f"C{row_floors}"] = "No of Typical Floors"
    ws[f"D{row_floors}"] = int(num_floors)
    ws[f"C{row_floors}"].font = FONT_BOLD
    ws[f"D{row_floors}"].font = FONT_BOLD
    ws[f"C{row_floors}"].border = BORDER_ALL
    ws[f"D{row_floors}"].border = BORDER_ALL
    ws[f"D{row_floors}"].number_format = NUM_FMT_INT
    ws[f"D{row_floors}"].alignment = ALIGN_RIGHT

    ws[f"C{row_towers}"] = "No of Towers"
    ws[f"D{row_towers}"] = int(num_towers)
    ws[f"C{row_towers}"].font = FONT_BOLD
    ws[f"D{row_towers}"].font = FONT_BOLD
    ws[f"C{row_towers}"].border = BORDER_ALL
    ws[f"D{row_towers}"].border = BORDER_ALL
    ws[f"D{row_towers}"].number_format = NUM_FMT_INT
    ws[f"D{row_towers}"].alignment = ALIGN_RIGHT

    # 7. TOWER-WISE BREAKDOWN TABLE (Works for BOTH Single Tower and Multi-Tower!)
    towers_to_render = towers if len(towers) > 1 else [{
        "name": towers[0]["name"] if towers else "Tower 1 (Typical Floor)",
        "floors": num_floors,
        "toilets": num_toilets,
        "kitchens": num_kitchens,
        "flr_ref": f"$D${row_floors}",
        "t_ref": f"$D${row_toilets}",
        "k_ref": f"$D${row_kitchens}",
        "is_multi_sheet": False
    }]

    # Interactive Option Dropdown Selector (Row 23)
    opt_sel_row = 23
    pref_opt = "Ultra Silent" if "Ultra Silent" in options else options[-1]
    ws.row_dimensions[opt_sel_row].height = 22
    ws.merge_cells(f"C{opt_sel_row}:D{opt_sel_row}")
    ws[f"C{opt_sel_row}"] = "Selected Option (Dropdown):"
    ws[f"C{opt_sel_row}"].font = FONT_HEADER_SUB
    ws[f"C{opt_sel_row}"].alignment = ALIGN_CENTER
    ws[f"C{opt_sel_row}"].fill = FILL_HEADER
    ws[f"C{opt_sel_row}"].border = BORDER_ALL
    ws[f"D{opt_sel_row}"].border = BORDER_ALL

    ws[f"E{opt_sel_row}"] = pref_opt
    ws[f"E{opt_sel_row}"].font = FONT_BOLD
    ws[f"E{opt_sel_row}"].alignment = ALIGN_CENTER
    ws[f"E{opt_sel_row}"].fill = FILL_HIGHLIGHT
    ws[f"E{opt_sel_row}"].border = BORDER_ALL

    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
    dv.promptTitle = "Select Product Option"
    dv.prompt = "Choose an option to dynamically recalculate the breakdown below."
    ws.add_data_validation(dv)
    dv.add(ws[f"E{opt_sel_row}"])

    ws.merge_cells(f"F{opt_sel_row}:K{opt_sel_row}")
    ws[f"F{opt_sel_row}"] = "← Dropdown se Option change karein; niche sabhi calculations auto-update honge"
    ws[f"F{opt_sel_row}"].font = Font(name=FONT_FAMILY, size=10, italic=True, color="475569")
    ws[f"F{opt_sel_row}"].alignment = ALIGN_LEFT

    tb_start = 24
    ws.row_dimensions[tb_start].height = 24
    ws.merge_cells(f"C{tb_start}:K{tb_start}")
    tb_title = ws[f"C{tb_start}"]
    tb_title.value = "Tower-Wise & Typical Floor-Wise Cost Breakdown (Dynamic by Selected Option)"
    tb_title.font = FONT_HEADER_SUB
    tb_title.alignment = ALIGN_CENTER
    tb_title.fill = FILL_HEADER
    tb_title.border = BORDER_ALL

    ws.row_dimensions[tb_start + 1].height = 22
    t_hdrs = [
        ("C", "Tower Name"), ("D", "Floors"), ("E", "Toilets"), ("F", "Kitchens"),
        ("G", "Toilet Cost"), ("H", "Kitchen Cost"), ("I", "Total Tower Cost"),
        ("J", "Typical Floor Cost"), ("K", "Per Flat Cost")
    ]
    for col, lbl in t_hdrs:
        cell = ws[f"{col}{tb_start + 1}"]
        cell.value = lbl
        cell.font = FONT_HEADER_MAIN
        cell.alignment = ALIGN_CENTER
        cell.fill = FILL_HEADER
        cell.border = BORDER_ALL

    last_opt_row = 4 + len(options)

    def _build_option_switch_formula(tw_dict, row_key, selector_ref="$E$23"):
        pairs = []
        for o in options:
            o_col = tw_dict.get("opts", {}).get(o)
            r_num = tw_dict.get(row_key)
            if o_col and r_num:
                pairs.append((o, f"'{tw_dict['sheet']}'!${o_col}${r_num}"))
        if not pairs:
            return "=0"
        if len(pairs) == 1:
            return f"={pairs[0][1]}"
        expr = pairs[-1][1]
        for o_name, cell_ref in reversed(pairs[:-1]):
            expr = f'IF({selector_ref}="{o_name}",{cell_ref},{expr})'
        return f"={expr}"

    for i, tw in enumerate(towers_to_render):
        r = tb_start + 2 + i
        ws.row_dimensions[r].height = 20
        ws[f"C{r}"] = tw["name"]

        # Dynamic formula links to source sheets if available, else numeric values
        if tw.get("flr_ref"):
            ws[f"D{r}"] = f"={tw['flr_ref']}"
        else:
            ws[f"D{r}"] = int(tw["floors"]) if isinstance(tw.get("floors"), (int, float)) else 20

        if tw.get("t_ref"):
            ws[f"E{r}"] = f"={tw['t_ref']}"
        else:
            ws[f"E{r}"] = int(tw["toilets"]) if isinstance(tw.get("toilets"), (int, float)) else 250

        if tw.get("k_ref"):
            ws[f"F{r}"] = f"={tw['k_ref']}"
        else:
            ws[f"F{r}"] = int(tw["kitchens"]) if isinstance(tw.get("kitchens"), (int, float)) else 100

        if tw.get("is_multi_sheet") and tw.get("t_row"):
            ws[f"G{r}"] = _build_option_switch_formula(tw, "t_row")
            ws[f"H{r}"] = _build_option_switch_formula(tw, "k_row") if tw.get("k_row") else "=0"
        elif is_hospital:
            ws[f"G{r}"] = f"=INDEX($F$5:$F${last_opt_row},MATCH($E${opt_sel_row},$C$5:$C${last_opt_row},0))*E{r}"
            ws[f"H{r}"] = "=0"
        else:
            ws[f"G{r}"] = f"=INDEX($F$5:$F${last_opt_row},MATCH($E${opt_sel_row},$C$5:$C${last_opt_row},0))*E{r}"
            ws[f"H{r}"] = f"=INDEX($H$5:$H${last_opt_row},MATCH($E${opt_sel_row},$C$5:$C${last_opt_row},0))*F{r}"
        ws[f"I{r}"] = f"=G{r}+H{r}"
        ws[f"J{r}"] = f"=I{r}/D{r}"
        ws[f"K{r}"] = f"=IF(F{r}>0, I{r}/F{r}, I{r}/E{r})"

        ws[f"C{r}"].font = FONT_BOLD
        ws[f"C{r}"].border = BORDER_ALL
        for c in ["D", "E", "F"]:
            ws[f"{c}{r}"].font = FONT_REGULAR
            ws[f"{c}{r}"].border = BORDER_ALL
            ws[f"{c}{r}"].number_format = NUM_FMT_INT
            ws[f"{c}{r}"].alignment = ALIGN_RIGHT
        for c in ["G", "H", "I", "J", "K"]:
            ws[f"{c}{r}"].font = FONT_REGULAR
            ws[f"{c}{r}"].border = BORDER_ALL
            ws[f"{c}{r}"].number_format = NUM_FMT_CURRENCY
            ws[f"{c}{r}"].alignment = ALIGN_RIGHT

    last_tb_row = tb_start + 1 + len(towers_to_render)

    # Total Row (only when multiple towers exist)
    if len(towers_to_render) > 1:
        tot_r = tb_start + 2 + len(towers_to_render)
        last_tb_row = tot_r
        ws.row_dimensions[tot_r].height = 22
        first_r = tb_start + 2
        last_r = tot_r - 1
        ws[f"C{tot_r}"] = "Project Total (All Towers)"
        ws[f"D{tot_r}"] = f"=MAX(D{first_r}:D{last_r})"
        ws[f"E{tot_r}"] = f"=SUM(E{first_r}:E{last_r})"
        ws[f"F{tot_r}"] = f"=SUM(F{first_r}:F{last_r})"
        ws[f"G{tot_r}"] = f"=SUM(G{first_r}:G{last_r})"
        ws[f"H{tot_r}"] = f"=SUM(H{first_r}:H{last_r})"
        ws[f"I{tot_r}"] = f"=SUM(I{first_r}:I{last_r})"
        ws[f"J{tot_r}"] = f"=SUM(J{first_r}:J{last_r})"
        ws[f"K{tot_r}"] = f"=IF(F{tot_r}>0, I{tot_r}/F{tot_r}, I{tot_r}/E{tot_r})"

        ws[f"C{tot_r}"].font = FONT_BOLD
        ws[f"C{tot_r}"].border = BORDER_ALL
        for c in ["D", "E", "F"]:
            ws[f"{c}{tot_r}"].font = FONT_BOLD
            ws[f"{c}{tot_r}"].border = BORDER_ALL
            ws[f"{c}{tot_r}"].number_format = NUM_FMT_INT
            ws[f"{c}{tot_r}"].alignment = ALIGN_RIGHT
        for c in ["G", "H", "I", "J", "K"]:
            ws[f"{c}{tot_r}"].font = FONT_BOLD
            ws[f"{c}{tot_r}"].border = BORDER_ALL
            ws[f"{c}{tot_r}"].number_format = NUM_FMT_CURRENCY
            ws[f"{c}{tot_r}"].alignment = ALIGN_RIGHT

    # 7B. OPTION COMPARISON TABLE (₹ & % Difference vs Base Option)
    if len(options) >= 2:
        if is_hospital:
            gt_col = "J" if use_rainwater else "I"
            pt_col = "L" if use_rainwater else "K"
            pf_col = pt_col
        elif has_diversion:
            gt_col = "L" if use_rainwater else "K"
            pt_col = "N" if use_rainwater else "M"
            pf_col = "O" if use_rainwater else "N"
        else:
            gt_col = "J" if use_rainwater else "I"
            pt_col = "L" if use_rainwater else "K"
            pf_col = "M" if use_rainwater else "L"

        comp_start = last_tb_row + 2
        ws.row_dimensions[comp_start].height = 24
        ws.merge_cells(f"C{comp_start}:I{comp_start}")
        ws[f"C{comp_start}"] = f"Option Comparison vs Base ({options[0]})"
        ws[f"C{comp_start}"].font = FONT_HEADER_SUB
        ws[f"C{comp_start}"].alignment = ALIGN_CENTER
        ws[f"C{comp_start}"].fill = FILL_HEADER
        ws[f"C{comp_start}"].border = BORDER_ALL

        comp_hdrs = [
            ("C", "Option"), ("D", "Grand Total Cost"), ("E", "Diff vs Base (₹)"),
            ("F", "% Diff vs Base"), ("G", "Per Toilet Cost"), ("H", "Per Toilet Diff (₹)"),
            ("I", "Per Flat Diff (₹)")
        ]
        ws.row_dimensions[comp_start + 1].height = 22
        for col, lbl in comp_hdrs:
            cell = ws[f"{col}{comp_start + 1}"]
            cell.value = lbl
            cell.font = FONT_HEADER_MAIN
            cell.alignment = ALIGN_CENTER
            cell.fill = FILL_HEADER
            cell.border = BORDER_ALL

        base_comp_row = comp_start + 2
        for idx, opt in enumerate(options):
            src_r = 5 + idx
            cr = comp_start + 2 + idx
            ws.row_dimensions[cr].height = 20
            ws[f"C{cr}"] = opt
            ws[f"C{cr}"].font = FONT_BOLD
            ws[f"C{cr}"].border = BORDER_ALL

            ws[f"D{cr}"] = f"={gt_col}{src_r}"
            ws[f"E{cr}"] = f"=D{cr}-$D${base_comp_row}"
            ws[f"F{cr}"] = f"=IF($D${base_comp_row}>0,(D{cr}-$D${base_comp_row})/$D${base_comp_row},0)"
            ws[f"G{cr}"] = f"={pt_col}{src_r}"
            ws[f"H{cr}"] = f"=G{cr}-$G${base_comp_row}"
            ws[f"I{cr}"] = f"={pf_col}{src_r}-{pf_col}5"

            for c in ["D", "E", "G", "H", "I"]:
                ws[f"{c}{cr}"].font = FONT_REGULAR
                ws[f"{c}{cr}"].border = BORDER_ALL
                ws[f"{c}{cr}"].number_format = NUM_FMT_CURRENCY
                ws[f"{c}{cr}"].alignment = ALIGN_RIGHT
            ws[f"F{cr}"].font = FONT_BOLD
            ws[f"F{cr}"].border = BORDER_ALL
            ws[f"F{cr}"].number_format = NUM_FMT_PERCENT
            ws[f"F{cr}"].alignment = ALIGN_RIGHT

    # 8. COLUMN WIDTHS (Auto-Fit Tower Name & Scope columns so text never truncates)
    col_widths = {
        "C": 22, "D": 22, "E": 20, "F": 18, "G": 20, "H": 18,
        "I": 22, "J": 20, "K": 20, "L": 24, "M": 24, "N": 24, "O": 24
    }
    merged_coords = set()
    for rng in ws.merged_cells.ranges:
        for mr in range(rng.min_row, rng.max_row + 1):
            for mc in range(rng.min_col, rng.max_col + 1):
                merged_coords.add((mr, mc))

    for col in ["C", "D"]:
        c_idx = openpyxl.utils.column_index_from_string(col)
        max_len = col_widths[col]
        for r in range(3, ws.max_row + 1):
            if (r, c_idx) in merged_coords:
                continue
            val = str(ws.cell(r, c_idx).value or "")
            if not val.startswith("="):
                max_len = max(max_len, len(val) + 4)
        col_widths[col] = min(max_len, 55)

    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # 9. TWO-WAY DYNAMIC LINK: WIRE BOQ DISCOUNT CELLS TO SUMMARY SHEET
    link_boq_discounts_to_summary(wb)

    wb.save(output_path)
    return output_path


def link_boq_discounts_to_summary(wb):
    """
    Wires every line item discount cell across all BOQ sheets
    to dynamic formulas referencing the Summary Sheet Discount Structure ($D$12 to $D$16).
    
    Category Mapping:
    - Pipe: ='Summary Sheet'!$D$12
    - Fittings: ='Summary Sheet'!$D$13
    - MFT: ='Summary Sheet'!$D$14
    - Lubricant: ='Summary Sheet'!$D$15
    - Clamps: ='Summary Sheet'!$D$16
    """
    for sname in wb.sheetnames:
        if sname == "Summary Sheet":
            continue
        sheet = wb[sname]

        # Find discount columns in row 14
        disc_cols = [c for c in range(1, sheet.max_column + 1) if str(sheet.cell(14, c).value or "").strip().lower() == "discount"]
        if not disc_cols:
            continue

        # Find line items range (rows 15 up to row before 'No of Floor')
        end_row = None
        for r in range(15, sheet.max_row + 1):
            lbl = str(sheet.cell(r, 1).value or "").strip().lower()
            if "no of floor" in lbl or "no. of floor" in lbl:
                end_row = r - 1
                break
        if not end_row:
            end_row = min(sheet.max_row, 120)

        for r in range(15, end_row + 1):
            desc = str(sheet.cell(r, 2).value or "").lower().strip()
            if not desc:
                desc = str(sheet.cell(r, 1).value or "").lower().strip()

            if "clamp" in desc:
                ref = "='Summary Sheet'!$D$16"
            elif "lubricant" in desc:
                ref = "='Summary Sheet'!$D$15"
            elif "floor drain" in desc or "trap" in desc:
                ref = "='Summary Sheet'!$D$14"
            elif "inspection pipe" in desc or "inspection" in desc or "access pipe" in desc or "cleanout" in desc:
                ref = "='Summary Sheet'!$D$13"  # Domain Rule: Inspection Pipe is a FITTING (F), NOT a pipe!
            elif "pipe" in desc:
                ref = "='Summary Sheet'!$D$12"
            else:
                ref = "='Summary Sheet'!$D$13"

            for dc in disc_cols:
                # Wire up formula so changing discount on Summary Sheet updates BOQ instantly
                sheet.cell(r, dc).value = ref



def load_knowledge_db(knowledge_db_path=None):
    """Loads the persistent knowledge database. Returns {} if not found or corrupt."""
    if knowledge_db_path is None or knowledge_db_path == "knowledge.json":
        knowledge_db_path = DEFAULT_KNOWLEDGE_PATH
    if not os.path.exists(knowledge_db_path):
        return {}
    try:
        with open(knowledge_db_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def find_knowledge_match(wb_path, wb, knowledge_db_path="knowledge.json", original_filename=None):
    """
    Finds a matching entry in knowledge.json for the given workbook.
    Matches by:
      1. Exact filename (using original_filename if provided)
      2. Sorted sheetnames fingerprint (same sheets = same project variant)
      3. Fuzzy project name match (excluding generic BOQ/builder words)
    Returns (entry, match_type) or (None, None).
    """
    db = load_knowledge_db(knowledge_db_path)
    projects = db.get("learned_projects", [])
    if not projects:
        return None, None

    fname = os.path.basename(original_filename or wb_path)
    current_sheets = sorted([s for s in wb.sheetnames if "summary" not in s.lower()])

    # Match 1: Exact filename
    for p in projects:
        if p.get("filename", "").strip() == fname:
            return p, "exact_filename"

    # Match 2: Sheetnames fingerprint (structural match - same project, different date)
    for p in projects:
        learned_sheets = sorted([s for s in p.get("sheetnames", []) if "summary" not in s.lower()])
        if learned_sheets and learned_sheets == current_sheets:
            return p, "sheet_fingerprint"

    # Match 3: Strict fuzzy project name match (ignore generic BOQ/builder terms)
    stop_words = {
        "huliot", "boq", "summary", "for", "the", "by", "and", "vtp", "tower",
        "project", "typical", "floor", "commercial", "area", "toilet", "toilets",
        "kitchen", "std", "length", "with", "diversion", "sheet", "punawale", "road"
    }
    fname_words = set(re.findall(r"[a-z]+", fname.lower())) - stop_words
    for p in projects:
        pname_words = set(re.findall(r"[a-z]+", p.get("project_title", "").lower())) - stop_words
        pfname_words = set(re.findall(r"[a-z]+", p.get("filename", "").lower())) - stop_words
        overlap = fname_words & (pname_words | pfname_words)
        if len(overlap) >= 2:
            return p, "fuzzy_name"

    return None, None


def compute_format_confidence(wb_path, wb, arch_key, detected_options, has_diversion, knowledge_db_path, original_filename, format_labels):
    """
    Compares the uploaded workbook against all saved formats in knowledge.json
    and returns a realistic Confidence Score (0-100) plus the matched Format label.
    """
    db = load_knowledge_db(knowledge_db_path)
    projects = db.get("learned_projects", [])
    default_label = format_labels.get(arch_key, "Standard Format")
    if not projects:
        return {
            "confidence": 0,
            "format_label": default_label,
            "matched_project": None,
            "is_new_format": True,
        }

    fname = os.path.basename(original_filename or wb_path)
    cur_sheets = {s.strip().lower() for s in wb.sheetnames if "summary" not in s.lower()}
    cur_opts = {o.strip().lower() for o in (detected_options or [])}

    best_score = 0
    best_entry = None

    for p in projects:
        if p.get("filename", "").strip() == fname:
            best_score = 100
            best_entry = p
            break

        p_sheets = {s.strip().lower() for s in p.get("sheetnames", []) if "summary" not in s.lower()}
        if cur_sheets or p_sheets:
            union_s = cur_sheets | p_sheets
            sheet_sim = (len(cur_sheets & p_sheets) / len(union_s)) if union_s else 0.0
        else:
            sheet_sim = 0.0

        p_arch = p.get("archetype")
        if p_arch:
            arch_match = 1.0 if p_arch == arch_key else 0.0
        else:
            # Infer archetype for older knowledge entries without explicit archetype field
            inferred_arch = "split_tk" if any("(t)" in s for s in p_sheets) else (
                "multi_sheet_towers" if len(p_sheets) > 2 else "combined"
            )
            arch_match = 1.0 if inferred_arch == arch_key else 0.0

        p_opts = {str(o.get("option", "")).strip().lower() for o in p.get("options", []) if o.get("option")}
        if cur_opts or p_opts:
            union_o = cur_opts | p_opts
            opt_sim = (len(cur_opts & p_opts) / len(union_o)) if union_o else 0.0
        else:
            opt_sim = 0.0

        p_has_div = p.get("has_diversion")
        if p_has_div is None:
            p_has_div = any("diversion" in s for s in p_sheets)
        div_match = 1.0 if bool(p_has_div) == bool(has_diversion) else 0.0

        score = round(100 * (0.50 * sheet_sim + 0.25 * arch_match + 0.15 * opt_sim + 0.10 * div_match))
        if cur_sheets and p_sheets and cur_sheets == p_sheets:
            score = max(score, 96)

        if score > best_score:
            best_score = score
            best_entry = p

    matched_arch = (best_entry or {}).get("archetype") or arch_key
    label = format_labels.get(matched_arch, default_label)
    return {
        "confidence": int(best_score),
        "format_label": label,
        "matched_project": (best_entry or {}).get("project_title") or (best_entry or {}).get("filename"),
        "is_new_format": best_score < 70,
    }


def train_from_workbook(wb_path, knowledge_db_path=None, original_filename=None, training_type_override=None):
    """
    Ingests either:
      1. A completed workbook with any Summary Sheet ('Summary', 'Summary Sheet', 'Project Summary', etc.) -> 'full_summary'
      2. A raw BOQ workbook without Summary Sheet -> 'raw_layout' (learns sheet structure, options, towers, & runs Health Check)
    Safety rules:
      - Health Check Gate: rejects empty/non-BOQ files (0 BOQ sheets or 0 options) to prevent knowledge pollution.
      - Protection Rule: raw_layout never overwrites an existing verified full_summary entry.
    """
    if knowledge_db_path is None or knowledge_db_path == "knowledge.json":
        knowledge_db_path = DEFAULT_KNOWLEDGE_PATH
    wb = openpyxl.load_workbook(wb_path, data_only=False)
    fname = os.path.basename(original_filename or wb_path)

    # Flexible Summary Sheet detection
    summary_ws = None
    for sname in wb.sheetnames:
        if is_summary_sheet(wb[sname]):
            summary_ws = wb[sname]
            break

    sheet_cls = classify_workbook_sheets(wb)
    is_split_tk = (sheet_cls["toilet_split"] is not None and sheet_cls["kitchen_split"] is not None)
    is_hospital = ("hospital" in fname.lower()) or (sheet_cls["toilet_split"] is not None and sheet_cls["kitchen_split"] is None and len(sheet_cls["typical_sheets"]) == 0)
    archetype = "hospital" if is_hospital else ("split_tk" if is_split_tk else ("multi_sheet_towers" if len(sheet_cls["typical_sheets"]) > 1 else "combined"))
    towers = extract_project_towers(wb, sheet_cls=sheet_cls)

    if summary_ws is not None:
        # --- MODE 1: FULL SUMMARY TRAINING ---
        ws = summary_ws
        project_title = str(ws["C2"].value or ws["B2"].value or "Unknown Project")

        options_learned = []
        for r in range(4, 12):
            opt_name = str(ws.cell(r, 3).value or ws.cell(r, 2).value or "").strip()
            if opt_name and opt_name.lower() not in ["option", "discount structure", "none", "tower"]:
                formulas = {}
                for c in range(1, 16):
                    val = str(ws.cell(r, c).value or "")
                    if val.startswith("="):
                        formulas[get_column_letter(c)] = val
                if formulas:
                    options_learned.append({"option": opt_name, "formulas": formulas})

        discounts_learned = {}
        for r in range(10, 22):
            tag = str(ws.cell(r, 2).value or "").strip()
            name = str(ws.cell(r, 3).value or "").strip()
            val = ws.cell(r, 4).value
            if tag and isinstance(val, (int, float)):
                discounts_learned[tag] = {"name": name, "discount": val}

        toilets_count = None
        kitchens_count = None
        for r in range(15, 26):
            lbl = str(ws.cell(r, 3).value or ws.cell(r, 2).value or "").lower()
            val = ws.cell(r, 4).value or ws.cell(r, 3).value
            if "toilets" in lbl:
                toilets_count = val
            elif "kitchens" in lbl or "flats" in lbl:
                kitchens_count = val

        entry = {
            "training_type": training_type_override or "full_summary",
            "summary_sheet_name": summary_ws.title,
            "project_title": project_title,
            "filename": fname,
            "archetype": archetype,
            "tower_count": len(towers) if towers else 1,
            "tower_names": [t["name"] for t in towers],
            "has_diversion": len(sheet_cls["diversion_sheets"]) > 0,
            "diversion_sheet_count": len(sheet_cls["diversion_sheets"]),
            "has_rainwater": len(sheet_cls.get("rainwater_sheets", [])) > 0,
            "rainwater_sheet_count": len(sheet_cls.get("rainwater_sheets", [])),
            "rainwater_sheet_names": [rws.title for rws in sheet_cls.get("rainwater_sheets", [])],
            "options": options_learned,
            "discounts": discounts_learned,
            "quantities": {"toilets": toilets_count, "kitchens": kitchens_count},
            "sheetnames": wb.sheetnames
        }
        warnings = []
    else:
        # --- MODE 2: RAW BOQ LAYOUT TRAINING ---
        meta = analyze_workbook(wb_path, knowledge_db_path=knowledge_db_path, original_filename=original_filename)
        total_boq_sheets = (
            len(sheet_cls["typical_sheets"])
            + len(sheet_cls["diversion_sheets"])
            + (1 if sheet_cls["toilet_split"] else 0)
            + (1 if sheet_cls["kitchen_split"] else 0)
        )
        # Health Check Gate: prevent polluting knowledge.json with invalid/empty Excel files
        if total_boq_sheets == 0 or not meta.get("detected_options"):
            return {
                "status": "error",
                "message": "यह फाइल वैध Huliot BOQ फॉर्मेट नहीं लगती (0 priced BOQ शीट्स / 0 Options मिले)। Knowledge DB में सेव नहीं किया गया।"
            }

        entry = {
            "training_type": "raw_layout",
            "project_title": meta["project_name"],
            "filename": fname,
            "archetype": archetype,
            "tower_count": meta["num_towers"],
            "tower_names": [t["name"] for t in towers] if towers else ["Typical Floor"],
            "has_diversion": meta["has_diversion"],
            "diversion_sheet_count": len(sheet_cls["diversion_sheets"]),
            "has_rainwater": meta["has_rainwater"],
            "rainwater_sheet_count": len(sheet_cls.get("rainwater_sheets", [])),
            "rainwater_sheet_names": [rws.title for rws in sheet_cls.get("rainwater_sheets", [])],
            "options": [{"option": o, "formulas": {}} for o in meta["detected_options"]],
            "quantities": {
                "toilets": meta["estimated_toilets"],
                "kitchens": meta["estimated_kitchens"]
            },
            "floors_count": meta["floors_count"],
            "baseline_discount": meta["baseline_discount"],
            "sheetnames": wb.sheetnames
        }
        warnings = meta.get("warnings", [])

    entry["trained_at"] = datetime.now().isoformat(timespec="seconds")

    db = {"learned_projects": [], "summary_count": 0}
    if os.path.exists(knowledge_db_path):
        try:
            with open(knowledge_db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception:
            pass

    existing_idx = next((i for i, p in enumerate(db["learned_projects"]) if p.get("filename") == entry["filename"]), None)
    if existing_idx is not None:
        existing_type = db["learned_projects"][existing_idx].get("training_type", "full_summary")
        # Protection Rule: raw_layout must NEVER overwrite a verified full_summary
        if entry["training_type"] == "raw_layout" and existing_type == "full_summary":
            return {
                "status": "success",
                "training_type": "raw_layout",
                "preserved_existing_summary": True,
                "warnings": warnings,
                "learned": entry
            }
        db["learned_projects"][existing_idx] = entry
    else:
        db["learned_projects"].append(entry)

    db["summary_count"] = len(db["learned_projects"])

    with open(knowledge_db_path, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)

    # Lesson 7: Auto-Backup on every new format train
    backup_dir = os.path.dirname(os.path.abspath(knowledge_db_path))
    os.makedirs(backup_dir, exist_ok=True)
    backup_filename = f"knowledge_backup_auto_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.json"
    backup_path = os.path.join(backup_dir, backup_filename)
    try:
        with open(backup_path, "w", encoding="utf-8") as bf:
            json.dump(db, bf, indent=2, ensure_ascii=False)
    except Exception:
        backup_path = None

    return {
        "status": "success",
        "training_type": entry["training_type"],
        "auto_backup_file": backup_path,
        "warnings": warnings,
        "learned": entry
    }


def verify_generated_summary(output_path):
    """
    Step 4 (Zero-Error Gate): Inspects the generated Summary Sheet before serving download.
    Verifies:
      - No formula cell is '=0' or contains Excel error tokens (#REF!, #DIV/0!, #VALUE!, #NAME?)
      - Every referenced sheet/cell coordinate exists in the workbook
      - Parameter cells (D18 Toilets, D20 Floors, D21 Towers) are > 0
    Returns dict: {'ok': bool, 'warnings': list[str], 'checked_cells': int}
    """
    wb = openpyxl.load_workbook(output_path, data_only=False)
    if "Summary Sheet" not in wb.sheetnames:
        return {"ok": False, "warnings": ["Summary Sheet missing in generated file"], "checked_cells": 0}

    ws = wb["Summary Sheet"]
    warnings = []
    checked = 0
    error_tokens = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!")

    # Check parameter counts (D18..D21)
    for r, lbl in [(18, "No of Toilets"), (20, "No of Typical Floors"), (21, "No of Towers")]:
        v = ws[f"D{r}"].value
        if not isinstance(v, (int, float)) or v <= 0:
            warnings.append(f"{lbl} (D{r}) invalid: {v}")

    # Check option rows (rows 5..9)
    for r in range(5, 10):
        opt = ws.cell(r, 3).value
        if not opt:
            continue
        for c in range(5, 16):
            val = ws.cell(r, c).value
            if val is None:
                continue
            checked += 1
            sval = str(val).strip()
            if sval == "=0":
                warnings.append(f"{get_column_letter(c)}{r} ({opt}): Formula is '=0'")
            if any(err in sval for err in error_tokens):
                warnings.append(f"{get_column_letter(c)}{r} ({opt}): Error token in {sval}")
            # Verify referenced sheet names exist
            for m in re.finditer(r"'([^']+)'!\$[A-Z]+\$\d+", sval):
                ref_sheet = m.group(1)
                if ref_sheet not in wb.sheetnames:
                    warnings.append(f"{get_column_letter(c)}{r} ({opt}): Missing sheet '{ref_sheet}'")

    return {"ok": len(warnings) == 0, "warnings": warnings, "checked_cells": checked}


def learn_from_generation(generated_path, knowledge_db_path=None, original_filename=None):
    """
    Auto-Learn on Generate: Persists user-confirmed discounts, quantities, and generated formulas
    from a newly generated Summary Sheet into knowledge.json, keeping raw input files tagged as 'raw_layout'.
    """
    return train_from_workbook(
        generated_path,
        knowledge_db_path=knowledge_db_path,
        original_filename=original_filename,
        training_type_override="raw_layout",
    )
