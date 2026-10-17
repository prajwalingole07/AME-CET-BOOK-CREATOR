import os
import re
import io
import copy
import shutil
from difflib import SequenceMatcher
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.oxml import OxmlElement

SKIP_CATS = ["DEFOPENS", "DEFOBCS", "DEFROBCS", "DEFLOPEN", "DEFLOBC"]
TABLE_COLS = ["Branch (Intake)", "Gender", "OPEN", "SC", "ST", "VJ-DT",
              "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC", "EWS", "TFWS", "AI (JEE)"]
CAT_ORDER = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC", "EWS", "TFWS", "AI (JEE)"]
REGULAR_CATS = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC"]
AI_COL = "AI (JEE)"
# Columns that hold ONE value per branch (not per gender) -> merged + centred
SHARED_COLS = ["EWS", "TFWS", "AI (JEE)"]
AI_EXCEL_DEFAULT = r"E:\OFFICE PERSONAL DATA\ENGG 2027\MH_AI_OPEN_CLOSE_2026_SINGLE_ROUND.xlsx"

def find_code_col(df):
    for c in df.columns:
        su = str(c).upper()
        if 'CODE' in su and ('INST' in su or 'COLLEGE' in su):
            return c
    for c in df.columns:
        su = str(c).upper()
        if 'CODE' in su and 'BRANCH' not in su and 'COURSE' not in su:
            return c
    return None

def find_inst_name_col(df):
    for c in df.columns:
        su = str(c).upper()
        if ('INSTITUTE' in su or 'COLLEGE' in su) and 'CODE' not in su:
            return c
    for c in df.columns:
        su = str(c).upper()
        if 'NAME' in su and 'BRANCH' not in su and 'COURSE' not in su and 'CITY' not in su and 'CODE' not in su:
            return c
    return None

def find_col(df, keywords):
    for col in df.columns:
        col_str = str(col).upper()
        if all(kw.upper() in col_str for kw in keywords):
            return col
    return None

def parse_int(val):
    if val is None or pd.isna(val): return None
    s = str(val).strip()
    if not s: return None
    match = re.search(r'-?\d+', s)
    if match:
        try: return int(match.group())
        except: return None
    return None

def get_cutoff_cell_from_row(row):
    """
    Picks the best round (highest percentile) across R1..R4 and returns
    "rank" or "rank\n(percentile%)" (same two-line layout as the AI (JEE) column).
    """
    pairs = []
    for idx in range(1, 5):
        rk_val = parse_int(row.get(f'R{idx} RANK'))
        pc_val = row.get(f'R{idx} PERCENTILE')
        if rk_val is not None:
            try:
                pc_num = float(pc_val) if pd.notna(pc_val) else -1.0
            except (TypeError, ValueError):
                pc_num = -1.0
            pairs.append((rk_val, pc_num))

    if not pairs:
        return "-----"

    pairs.sort(key=lambda x: (x[1], -x[0]), reverse=True)
    rank, pct = pairs[0]
    if pct and pct > 0:
        return f"{rank}\n({pct:.2f}%)"
    return str(rank)

# ============================================================
# AI (JEE) EXCEL SUPPORT  -  single round rank / percentile
# ============================================================
_AI_FILE_CACHE = {"key": None, "rows": {}, "colleges": {}}

def norm_ai_code(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    s = str(v).strip()
    if s.endswith('.0'): s = s[:-2]
    m = re.match(r'^0*(\d+)', s)
    return m.group(1) if m else s

def branch_key(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return ""
    s = str(s).lower().replace('&', ' and ')
    s = s.replace('-', '').replace(',', '')
    s = re.sub(r'[\(\)\[\]]', ' ', s)
    s = re.sub(r'[^a-z0-9 ]+', ' ', s)
    s = re.sub(r'\bengg\b', 'engineering', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def compact_key(s):
    return branch_key(s).replace(' ', '')

def names_match(cet_name, ai_name):
    ck = compact_key(cet_name)
    ak = compact_key(ai_name)
    if not ck or not ak: return False
    if ck == ak: return True
    raw = str(cet_name or '').strip()
    if raw.endswith('...'):
        prefix = compact_key(raw.rstrip('.'))
        if prefix and ak.startswith(prefix): return True
    raw2 = str(ai_name or '').strip()
    if raw2.endswith('...'):
        prefix2 = compact_key(raw2.rstrip('.'))
        if prefix2 and ck.startswith(prefix2): return True
    return False

def names_similar(cet_name, ai_name):
    """Same as names_match but also accepts a clear prefix form, e.g.
    'Electronics and Communication' vs 'Electronics and Communication Engineering'."""
    if names_match(cet_name, ai_name):
        return True
    ck = compact_key(cet_name)
    ak = compact_key(ai_name)
    if len(ck) < 14 or len(ak) < 14:
        return False
    return ck.startswith(ak) or ak.startswith(ck)

def read_ai_excel(ai_path):
    """Reads the AI (JEE) single-round excel. Cached by path + mtime."""
    if not ai_path or not os.path.exists(ai_path):
        return {}, {}
    try:
        mtime = str(os.path.getmtime(ai_path))
    except OSError:
        return {}, {}
    if _AI_FILE_CACHE["key"] == (ai_path, mtime):
        return _AI_FILE_CACHE["rows"], _AI_FILE_CACHE["colleges"]

    df = pd.read_excel(ai_path)
    code_col = find_col(df, ["INST", "CODE"]) or find_code_col(df)
    name_col = find_col(df, ["INSTITUTE", "NAME"]) or find_inst_name_col(df)
    city_col = find_col(df, ["CITY"])
    course_col = find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    choice_col = find_col(df, ["CHOICE", "CODE"])
    rank_col = find_col(df, ["AI", "MERIT"]) or find_col(df, ["RANK"])
    pct_col = find_col(df, ["PERCENTILE"])
    status_col = find_col(df, ["STATUS"])
    uni_col = find_col(df, ["UNIVERSITY"])

    rows_by_college = {}
    colleges = {}
    for _, row in df.iterrows():
        code = norm_ai_code(row.get(code_col)) if code_col else ""
        if not code: continue
        course = str(row.get(course_col, '')).strip() if course_col else ""
        if not course or course.lower() == 'nan': continue

        rank = parse_int(row.get(rank_col)) if rank_col else None
        if rank is None: continue
        pct = None
        if pct_col and pd.notna(row.get(pct_col)):
            try: pct = float(row.get(pct_col))
            except (TypeError, ValueError): pct = None

        rows_by_college.setdefault(code, []).append({
            "course": course,
            "code": norm_ai_code(row.get(choice_col)) if choice_col else "",
            "rank": rank,
            "pct": pct,
        })

        if code not in colleges:
            def clean(col):
                if not col: return ""
                v = row.get(col)
                if pd.isna(v): return ""
                return str(v).strip()
            colleges[code] = {
                "status": clean(status_col),
                "university": clean(uni_col),
                "city": clean(city_col),
                "name": clean(name_col),
            }

    _AI_FILE_CACHE["key"] = (ai_path, mtime)
    _AI_FILE_CACHE["rows"] = rows_by_college
    _AI_FILE_CACHE["colleges"] = colleges
    return rows_by_college, colleges

# ------------------------------------------------------------
# Institute directory (University + Autonomous status)
# ------------------------------------------------------------
DEFAULT_COLLEGE_DIRECTORY = r"C:\Users\prajw\Downloads\Maharashtra_Engineering_Colleges_Universities_with_University.xlsx"
_COLLEGE_DIR_CACHE = {"key": None, "data": {}}

def load_college_directory(path=None):
    """Reads the institute directory excel (Institute Code, Institute Name,
    Status, University).  Returns {code_or_name: {university, status, name}}.
    Cached by path + mtime."""
    path = path or DEFAULT_COLLEGE_DIRECTORY
    if not path or not os.path.exists(path):
        return {}
    try:
        mtime = str(os.path.getmtime(path))
    except OSError:
        return {}
    if _COLLEGE_DIR_CACHE["key"] == (path, mtime):
        return _COLLEGE_DIR_CACHE["data"]

    out = {}
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print(f"college directory read failed: {e}")
        return {}
    code_col = find_col(df, ["INSTITUTE", "CODE"]) or find_col(df, ["CODE"])
    name_col = find_col(df, ["INSTITUTE", "NAME"]) or find_col(df, ["NAME"])
    status_col = find_col(df, ["STATUS"])
    uni_col = find_col(df, ["UNIVERSITY"])

    for _, row in df.iterrows():
        def clean(col):
            if not col: return ""
            v = row.get(col)
            if pd.isna(v): return ""
            return str(v).strip()
        rec = {
            "university": clean(uni_col),
            "status": clean(status_col),
            "name": clean(name_col),
        }
        code = clean(code_col).replace('.0', '')
        if code:
            out[code] = rec
        if rec["name"]:
            out.setdefault(rec["name"].lower(), rec)

    _COLLEGE_DIR_CACHE["key"] = (path, mtime)
    _COLLEGE_DIR_CACHE["data"] = out
    return out

def enrich_from_directory(data, directory):
    """Overlays University + autonomous Status from the institute directory onto a
    parsed college (matched by Institute Code, then by name).  Also refreshes the
    chip values and the Word info-box summary."""
    if not data or not directory:
        return data
    rec = directory.get(str(data.get("college_code") or "").strip())
    if not rec:
        rec = directory.get(str(data.get("college_name") or "").strip().lower())
    if not rec:
        return data

    uni = rec.get("university") or ""
    st = rec.get("status") or ""
    if uni:
        data["university"] = uni
    if st:
        data["status"] = st

    wm = data.setdefault("web_meta", {})
    if uni:
        wm["university"] = uni
    if st:
        wm["status"] = st
    if data.get("university"):
        wm["summary"] = f"{data['university']} | {data.get('status', '')}".strip(" |")
    return data

def build_ai_index(cet_df, ai_path):
    """
    Matches every AI (JEE) row to the exact CET branch name of the same college.
    Pass 1: branch name match (exact or truncated).  Pass 2: choice code match,
    accepted only when the labels agree (exact or a clear prefix form).
    Returns: ({college_code: {cet_branch_name: "rank\n(percentile%)"}}, {college_code: meta})
    """
    rows_by_college, colleges_meta = read_ai_excel(ai_path)
    if not rows_by_college:
        return {}, {}

    code_col = find_code_col(cet_df)
    branch_col = find_col(cet_df, ["BRANCH", "NAME"]) or find_col(cet_df, ["COURSE", "NAME"]) or find_col(cet_df, ["BRANCH"])
    branch_code_col = find_col(cet_df, ["BRANCH", "CODE"])
    if not code_col or not branch_col:
        return {}, {}

    cet_branches = {}
    keep = [c for c in [code_col, branch_col, branch_code_col] if c]
    for _, r in cet_df[keep].drop_duplicates().iterrows():
        c_code = norm_ai_code(r.get(code_col))
        b_name = str(r.get(branch_col, '')).strip() if pd.notna(r.get(branch_col)) else ""
        if not c_code or not b_name or b_name.lower() == 'nan': continue
        entry = cet_branches.setdefault(c_code, {"names": [], "by_key": {}, "by_code": {}, "trunc": []})
        if b_name not in entry["names"]:
            entry["names"].append(b_name)
            entry["by_key"].setdefault(compact_key(b_name), b_name)
            if b_name.endswith('...'):
                entry["trunc"].append(b_name)
        if branch_code_col is not None and pd.notna(r.get(branch_code_col)):
            bc = norm_ai_code(r.get(branch_code_col))
            if bc:
                entry["by_code"].setdefault(bc, b_name)

    matches = {}
    for c_code, rows in rows_by_college.items():
        info = cet_branches.get(c_code)
        if not info: continue
        picked = {}
        claimed = set()
        pending = []
        for row in rows:
            target = info["by_key"].get(compact_key(row["course"]))
            if not target:
                ak = compact_key(row["course"])
                for t in info["trunc"]:
                    if ak and ak.startswith(compact_key(t.rstrip('.'))):
                        target = t
                        break
            if target:
                prev = picked.get(target)
                pct = row["pct"] if row["pct"] is not None else -1.0
                if prev is None or pct > prev["pct"]:
                    picked[target] = {"rank": row["rank"], "pct": pct}
                claimed.add(target)
            else:
                pending.append(row)
        for row in pending:
            cands = []
            if row["code"] and row["code"] in info["by_code"]:
                cands.append(info["by_code"][row["code"]])
            if not cands and row["code"]:
                suf = row["code"][-5:]
                for bc, bn in info["by_code"].items():
                    if bc.endswith(suf) and bn not in cands:
                        cands.append(bn)
            target = None
            exact = info["by_code"].get(row["code"]) if row["code"] else None
            for c in cands:
                if not names_similar(c, row["course"]):
                    continue
                if c in claimed and c != exact:
                    continue
                target = c
                break
            if not target:
                continue
            prev = picked.get(target)
            pct = row["pct"] if row["pct"] is not None else -1.0
            if prev is None or pct > prev["pct"]:
                picked[target] = {"rank": row["rank"], "pct": pct}
            claimed.add(target)
        if picked:
            matches[c_code] = {
                b: f"{v['rank']}\n({v['pct']:.2f}%)" if v["pct"] is not None else str(v["rank"])
                for b, v in picked.items()
            }
    return matches, colleges_meta

def apply_ai_to_branches(branches, ai_matches):
    """Writes the AI (JEE) rank+percentile into every quota table (both genders)."""
    if not ai_matches: return 0
    added = 0
    for b_name, b_data in branches.items():
        val = ai_matches.get(b_name)
        if not val: continue
        added += 1
        for key in ("HU_M", "HU_F", "OHU_M", "OHU_F", "STATE_M", "STATE_F"):
            b_data.setdefault(key, {})[AI_COL] = val
    return added

def check_quota_has_regular_data(branches, mode):
    regular_cols = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC"]
    m_key = f"{mode}_M"
    f_key = f"{mode}_F"
    for b_name, b_data in branches.items():
        m_dict = b_data.get(m_key, {})
        f_dict = b_data.get(f_key, {})
        for col in regular_cols:
            v_m = m_dict.get(col)
            if v_m and str(v_m).strip() not in ("-----", "", "None", "nan"):
                return True
            v_f = f_dict.get(col)
            if v_f and str(v_f).strip() not in ("-----", "", "None", "nan"):
                return True
    return False

CATEGORY_MAPPING = {
    "OPEN": "OPEN",
    "SC": "SC",
    "ST": "ST",
    "VJ": "VJ-DT",
    "VJDT": "VJ-DT",
    "VJ-DT": "VJ-DT",
    "NT1": "NT-1 (B)",
    "NT-1": "NT-1 (B)",
    "NT2": "NT-2 (C)",
    "NT-2": "NT-2 (C)",
    "NT3": "NT-3 (D)",
    "NT-3": "NT-3 (D)",
    "OBC": "OBC",
    "SEBC": "SEBC",
}

def classify_and_route(cat):
    c = str(cat).strip().upper()
    if not c: return None
    
    # 1. STRICT FILTER: Discard ANY Defence, PWD, Orphan categories completely
    if c.startswith(("DEF", "PWD", "ORPHAN")): return None
    if "DEF" in c or "PWD" in c or "ORPHAN" in c: return None
    
    # 2. Special state-level quotas
    if "EWS" in c: return ("Male", "EWS", "SPECIAL")
    if "TFWS" in c: return ("Male", "TFWS", "SPECIAL")
    if c == "MI" or "MINORITY" in c: return ("Male", "MI", "SPECIAL")
    if "AI" in c or "JEE" in c: return ("Male", "AI (JEE)", "SPECIAL")
    
    # 3. Regular MHT-CET categories must start with G (General/Male) or L (Ladies/Female)
    if not (c.startswith("G") or c.startswith("L")): return None
    
    # 4. Must end with quota suffix H (Home University), O (Other than HU), or S (State Level)
    if not (c.endswith("H") or c.endswith("O") or c.endswith("S")): return None
    
    gender = "Female" if c.startswith("L") else "Male"
    dest = "HU" if c.endswith("H") else ("OHU" if c.endswith("O") else "STATE")
    
    mid = c[1:-1]
    base = CATEGORY_MAPPING.get(mid)
    if not base: return None
    
    return (gender, base, dest)

def fetch_college_web_info(college_name):
    """Fallback web scraper to enrich missing college information."""
    info = {
        "nba_grading": "Available in MHT-CET Record",
        "total_seats": "Available",
        "website": "",
        "summary": "Data extracted from MHT-CET Cutoff Database."
    }
    return info

def scan_colleges_from_excel(excel_path, ai_path=None):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File not found at path: {excel_path}")
        
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])
    
    colleges_dict = {}
    
    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip() if code_col and pd.notna(row.get(code_col)) else ""
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ""
        c_city = str(row.get(city_col, '')).strip() if city_col and pd.notna(row.get(city_col)) else ""
        b_name = str(row.get(branch_col, '')).strip() if branch_col and pd.notna(row.get(branch_col)) else ""
        cat_val = str(row.get(cat_col, '')).strip().upper() if cat_col and pd.notna(row.get(cat_col)) else ""
        
        if c_code.endswith('.0'): c_code = c_code[:-2]
        if not c_code and not c_name: continue
        
        key = c_code if c_code else c_name
        if key not in colleges_dict:
            colleges_dict[key] = {
                "code": c_code,
                "name": c_name if c_name else f"Institute {c_code}",
                "city": c_city,
                "branches": set(),
                "categories": set(),
                "has_hu": False,
                "has_ohu": False,
                "has_state": False,
                "has_ews": False,
                "has_tfws": False,
                "has_ai": norm_ai_code(c_code) in ai_matches or key in ai_matches,
                "display": f"{c_code} - {c_name}" if c_code and c_name and c_code != c_name else (c_name or c_code)
            }
            
        c_obj = colleges_dict[key]
        if b_name and b_name.lower() != "nan":
            c_obj["branches"].add(b_name)
        if cat_val:
            c_obj["categories"].add(cat_val)
            route = classify_and_route(cat_val)
            if route:
                _, _, dest = route
                if dest == "HU": c_obj["has_hu"] = True
                elif dest == "OHU": c_obj["has_ohu"] = True
                elif dest == "STATE": c_obj["has_state"] = True
                elif dest == "SPECIAL":
                    if "EWS" in cat_val: c_obj["has_ews"] = True
                    if "TFWS" in cat_val: c_obj["has_tfws"] = True
            
    colleges_list = []
    for k, v in colleges_dict.items():
        colleges_list.append({
            "code": v["code"],
            "name": v["name"],
            "city": v["city"],
            "branches_count": len(v["branches"]),
            "display": v["display"],
            "has_hu": v["has_hu"],
            "has_ohu": v["has_ohu"],
            "has_state": v["has_state"],
            "has_ews": v["has_ews"],
            "has_tfws": v["has_tfws"],
            "has_ai": v["has_ai"]
        })
        
    colleges_list.sort(key=lambda x: (x["code"] or "99999", x["name"]))
    return colleges_list, len(df)

def fast_precache_all(excel_path, ai_path=None, directory_path=None):
    """
    Vectorized single-pass ultra-fast pre-cache engine for all colleges.
    Accurately routes EWS/TFWS only to available quotas and avoids phantom empty tables.
    """
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    directory = load_college_directory(directory_path)
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])

    # First pass: detect regular quota presence per college
    college_quota_meta = {}
    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip().replace('.0','') if code_col and pd.notna(row.get(code_col)) else ''
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ''
        cat = str(row.get(cat_col, '')).strip().upper() if cat_col and pd.notna(row.get(cat_col)) else ''
        if not c_code and not c_name: continue
        key = c_code if c_code else c_name
        if key not in college_quota_meta:
            college_quota_meta[key] = {"has_hu": False, "has_ohu": False, "has_state": False, "has_ews": False, "has_tfws": False}
        m = college_quota_meta[key]
        route = classify_and_route(cat)
        if route:
            _, _, dest = route
            if dest == "HU": m["has_hu"] = True
            elif dest == "OHU": m["has_ohu"] = True
            elif dest == "STATE": m["has_state"] = True
            elif dest == "SPECIAL":
                if "EWS" in cat: m["has_ews"] = True
                if "TFWS" in cat: m["has_tfws"] = True

    colleges_cache = {}

    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip().replace('.0','') if code_col and pd.notna(row.get(code_col)) else ''
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ''
        b_name = str(row.get(branch_col, '')).strip() if branch_col and pd.notna(row.get(branch_col)) else ''
        cat = str(row.get(cat_col, '')).strip() if cat_col and pd.notna(row.get(cat_col)) else ''
        
        if not c_code and not c_name: continue
        key = c_code if c_code else c_name
        
        if key not in colleges_cache:
            full_name = f"{c_code} - {c_name}" if c_code and c_name and c_code != c_name else (c_name or c_code)
            q_info = college_quota_meta.get(key, {})
            ext = ai_meta.get(norm_ai_code(c_code)) or ai_meta.get(key) or {}
            row_city = ""
            if city_col and pd.notna(row.get(city_col)):
                row_city = str(row.get(city_col)).strip()
                if row_city.lower() == 'nan': row_city = ""
            colleges_cache[key] = {
                "college_code": c_code,
                "college_name": c_name,
                "full_name": full_name,
                "city": row_city or ext.get("city", ""),
                "status": ext.get("status", ""),
                "university": ext.get("university", ""),
                "branches": {},
                "has_hu": q_info.get("has_hu", False),
                "has_ohu": q_info.get("has_ohu", False),
                "has_state": q_info.get("has_state", False),
                "has_ews": q_info.get("has_ews", False),
                "has_tfws": q_info.get("has_tfws", False),
                "has_ai": False,
                "ai_branches": 0,
                "blank_warnings": [],
                "web_meta": {
                    "nba_grading": "Available in MHT-CET Record",
                    "total_seats": "Available",
                    "website": "",
                    "university": ext.get("university", ""),
                    "status": ext.get("status", ""),
                    "city": ext.get("city", ""),
                    "summary": (f"{ext.get('university')} | {ext.get('status', '')}".strip(" |")
                                if ext.get("university") else f"MHT-CET Cutoff record for {full_name}.")
                }
            }
            
        c_entry = colleges_cache[key]
        if not b_name or b_name.lower() == "nan": continue
        
        res = classify_and_route(cat)
        if not res: continue
        
        gender, base, dest = res
        cell = get_cutoff_cell_from_row(row)
        if cell == "-----": continue
        
        b_dict = c_entry["branches"]
        if b_name not in b_dict:
            b_dict[b_name] = {"HU_M": {}, "HU_F": {}, "OHU_M": {}, "OHU_F": {}, "STATE_M": {}, "STATE_F": {}}
            
        b = b_dict[b_name]
        k = "HU_M" if gender == "Male" else "HU_F"
        k_ohu = "OHU_M" if gender == "Male" else "OHU_F"
        k_state = "STATE_M" if gender == "Male" else "STATE_F"
        
        q_meta = college_quota_meta.get(key, {})
        
        if dest == "SPECIAL":
            # Only put EWS/TFWS in tables that actually have regular quota!
            if q_meta.get("has_hu"):
                if base not in b[k]: b[k][base] = cell
            if q_meta.get("has_ohu"):
                if base not in b[k_ohu]: b[k_ohu][base] = cell
            if q_meta.get("has_state"):
                if base not in b[k_state]: b[k_state][base] = cell
            # If college has no regular HU/OHU/State (rare edge case), fallback to State
            if not q_meta.get("has_hu") and not q_meta.get("has_ohu") and not q_meta.get("has_state"):
                if base not in b[k_state]: b[k_state][base] = cell
        elif dest == "HU":
            if base not in b[k]: b[k][base] = cell
        elif dest == "OHU":
            if base not in b[k_ohu]: b[k_ohu][base] = cell
        elif dest == "STATE":
            if base not in b[k_state]: b[k_state][base] = cell

    for key, c_entry in colleges_cache.items():
        branches = c_entry["branches"]
        ai_map = ai_matches.get(norm_ai_code(key)) or ai_matches.get(key)
        added = apply_ai_to_branches(branches, ai_map)
        c_entry["has_ai"] = added > 0
        c_entry["ai_branches"] = added
        c_entry["has_hu"] = check_quota_has_regular_data(branches, "HU")
        c_entry["has_ohu"] = check_quota_has_regular_data(branches, "OHU")
        c_entry["has_state"] = check_quota_has_regular_data(branches, "STATE")
        
        warnings = []
        for b_name, b_data in branches.items():
            ews_val = b_data["HU_M"].get("EWS") or b_data["HU_F"].get("EWS") or b_data["STATE_M"].get("EWS") or b_data["STATE_F"].get("EWS")
            tfws_val = b_data["HU_M"].get("TFWS") or b_data["HU_F"].get("TFWS") or b_data["STATE_M"].get("TFWS") or b_data["STATE_F"].get("TFWS")
            if not ews_val or ews_val == "-----":
                warnings.append(f"EWS rank is blank for branch: {b_name}")
            if not tfws_val or tfws_val == "-----":
                warnings.append(f"TFWS rank is blank for branch: {b_name}")
        c_entry["blank_warnings"] = warnings
        enrich_from_directory(c_entry, directory)

    return colleges_cache

def parse_college_data(excel_path, college_query, mode="ALL", ai_path=None, directory_path=None):
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])
    
    college_df = df
    if college_query:
        q_str = str(college_query).strip().lower()
        matches = []
        if code_col:
            matches.append(df[code_col].astype(str).str.lower().str.replace('.0','',regex=False) == q_str)
        if inst_col:
            matches.append(df[inst_col].astype(str).str.lower().str.contains(q_str))
        if matches:
            combined_filter = matches[0]
            for m in matches[1:]:
                combined_filter = combined_filter | m
            filtered = df[combined_filter]
            if not filtered.empty:
                college_df = filtered

    code_val = ""
    inst_val = ""
    if code_col and not college_df[code_col].dropna().empty:
        code_val = str(college_df[code_col].dropna().iloc[0]).strip().replace('.0','')
    if inst_col and not college_df[inst_col].dropna().empty:
        inst_val = str(college_df[inst_col].dropna().iloc[0]).strip()
        
    full_name = f"{code_val} - {inst_val}" if code_val and inst_val and code_val != inst_val else (inst_val or code_val)
    if not full_name: full_name = "Selected College"

    # First pass: detect regular quota presence per college using classify_and_route
    college_df_cats = college_df[cat_col].dropna().astype(str).str.upper().tolist() if cat_col else []
    has_hu = False
    has_ohu = False
    has_state = False
    has_ews = any("EWS" in c for c in college_df_cats)
    has_tfws = any("TFWS" in c for c in college_df_cats)
    for c in college_df_cats:
        r = classify_and_route(c)
        if r:
            _, _, d = r
            if d == "HU": has_hu = True
            elif d == "OHU": has_ohu = True
            elif d == "STATE": has_state = True

    branches = {}
    cur_branch = ""
    
    for _, row in college_df.iterrows():
        if branch_col and pd.notna(row.get(branch_col)):
            val = str(row[branch_col]).strip()
            if val and val.lower() != "nan":
                cur_branch = val
                
        if not cur_branch: continue
        
        cat = row.get(cat_col)
        if pd.isna(cat): continue
        
        res = classify_and_route(cat)
        if not res: continue
        
        gender, base, dest = res
        cell = get_cutoff_cell_from_row(row)
        
        if cell == "-----": continue
        
        if cur_branch not in branches:
            branches[cur_branch] = {"HU_M": {}, "HU_F": {}, "OHU_M": {}, "OHU_F": {}, "STATE_M": {}, "STATE_F": {}}
        
        b = branches[cur_branch]
        k = "HU_M" if gender == "Male" else "HU_F"
        k_ohu = "OHU_M" if gender == "Male" else "OHU_F"
        k_state = "STATE_M" if gender == "Male" else "STATE_F"
        
        if dest == "SPECIAL":
            if has_hu:
                if base not in b[k]: b[k][base] = cell
            if has_ohu:
                if base not in b[k_ohu]: b[k_ohu][base] = cell
            if has_state:
                if base not in b[k_state]: b[k_state][base] = cell
            if not has_hu and not has_ohu and not has_state:
                if base not in b[k_state]: b[k_state][base] = cell
        elif dest == "HU":
            if base not in b[k]: b[k][base] = cell
        elif dest == "OHU":
            if base not in b[k_ohu]: b[k_ohu][base] = cell
        elif dest == "STATE":
            if base not in b[k_state]: b[k_state][base] = cell

    # AI (JEE) single-round data is merged into every quota table
    ai_key = norm_ai_code(code_val)
    ai_map = ai_matches.get(ai_key) or ai_matches.get(str(college_query).strip())
    ai_added = apply_ai_to_branches(branches, ai_map)
    ext = ai_meta.get(ai_key) or {}

    city_val = ""
    if city_col and not college_df[city_col].dropna().empty:
        city_val = str(college_df[city_col].dropna().iloc[0]).strip()
        if city_val.lower() == 'nan': city_val = ""
    if not city_val: city_val = ext.get("city", "")

    blank_warnings = []
    for b_name, b_data in branches.items():
        ews_val = b_data["HU_M"].get("EWS") or b_data["HU_F"].get("EWS") or b_data["STATE_M"].get("EWS") or b_data["STATE_F"].get("EWS")
        tfws_val = b_data["HU_M"].get("TFWS") or b_data["HU_F"].get("TFWS") or b_data["STATE_M"].get("TFWS") or b_data["STATE_F"].get("TFWS")
        
        if not ews_val or ews_val == "-----":
            blank_warnings.append(f"EWS rank is blank for branch: {b_name}")
        if not tfws_val or tfws_val == "-----":
            blank_warnings.append(f"TFWS rank is blank for branch: {b_name}")

    web_meta = fetch_college_web_info(full_name)
    web_meta["university"] = ext.get("university", "")
    web_meta["status"] = ext.get("status", "")
    web_meta["city"] = city_val
    if ext.get("university"):
        web_meta["summary"] = f"{ext.get('university')} | {ext.get('status', '')}".strip(" |")
    
    # Final verification: verify that regular data actually exists in table
    has_hu = check_quota_has_regular_data(branches, "HU")
    has_ohu = check_quota_has_regular_data(branches, "OHU")
    has_state = check_quota_has_regular_data(branches, "STATE")
    
    result = {
        "college_code": code_val,
        "college_name": inst_val,
        "full_name": full_name,
        "city": city_val,
        "status": ext.get("status", ""),
        "university": ext.get("university", ""),
        "has_hu": has_hu,
        "has_ohu": has_ohu,
        "has_state": has_state,
        "has_ews": has_ews,
        "has_tfws": has_tfws,
        "has_ai": ai_added > 0,
        "ai_branches": ai_added,
        "branches": branches,
        "blank_warnings": blank_warnings,
        "web_meta": web_meta
    }
    enrich_from_directory(result, load_college_directory(directory_path))
    return result

# ============================================================
# WORD DOCX GENERATOR
# ============================================================
def set_cell_shading(cell, color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))

def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    b = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                  f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'</w:tcBorders>')
    tcPr.append(b)

def drop_border(cell, side):
    """Removes one side of a cell's border (used where two cells are merged)."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        return
    el = borders.find(qn(f'w:{side}'))
    if el is not None:
        el.set(qn('w:val'), 'nil')

def make_run(p, text, bold=False, size=8):
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    run.font._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    return run

# JoSAA-document layout constants: the per-college Word file must match the
# JoSAA downloads page for page - 0.4" margins, every table exactly filling
# the usable width (11.69 - 2*0.4 = 10.89") with a hair to spare, centred.
DOC_MARGIN_IN = 0.4
DOC_TABLE_WIDTH_IN = 10.85          # JoSAA uses 10.85" tables on 10.89" usable
DOC_TITLE = "MHT-CET 2026-27  |  OPENING & CLOSING RANKS  (Source: Excel)"


def _fit_widths(widths):
    """Scale a column-width list so it sums to DOC_TABLE_WIDTH_IN."""
    total = float(sum(widths))
    if total <= 0:
        return list(widths)
    f = DOC_TABLE_WIDTH_IN / total
    return [w * f for w in widths]


def create_docx_header_logo(doc, logo_path):
    if logo_path and os.path.exists(logo_path):
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run()
            run.add_picture(logo_path, width=Inches(0.95))
        except Exception as e:
            print(f"Error adding logo to docx: {e}")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    run = make_run(p, DOC_TITLE, bold=True, size=10.5)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

def create_info_box(doc, name, web_meta=None):
    tbl = doc.add_table(rows=3, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    nba = web_meta.get("nba_grading", "Available in MHT-CET Record") if web_meta else "Available in MHT-CET Record"
    seats = web_meta.get("total_seats", "Available") if web_meta else "Available"
    university = (web_meta or {}).get("university", "")
    status = (web_meta or {}).get("status", "")
    city = (web_meta or {}).get("city", "")
    
    texts = [
        [("College Code & Name: ", False), (name, False)],
        [("NBA Grading: ", True), (nba, False)],
        [("Total Seats: ", True), (seats, False)],
        [("Available Cut-off: ", True), ("CET R1-R4 + AI (JEE) (Rank + %ile)", False)]
    ]
    widths = [Inches(x) for x in _fit_widths([4.0, 2.0, 1.5, 2.5])]
    
    for i, parts in enumerate(texts):
        cell = tbl.cell(0, i)
        cell.width = widths[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        for txt, b in parts: make_run(p, txt, bold=b)
        set_cell_shading(cell, "F2F2F2"); set_cell_border(cell)

    merged = tbl.cell(1, 0).merge(tbl.cell(2, 3))
    merged.text = ""
    p0 = merged.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0); p0.paragraph_format.space_after = Pt(0)
    parts = []
    if university: parts.append(("University: ", True)); parts.append((university, False))
    if status: parts.append(("    Status: ", True)); parts.append((status, False))
    if city: parts.append(("    City: ", True)); parts.append((city, False))
    for txt, b in parts:
        make_run(p0, txt, bold=b, size=7)
    for p in merged.paragraphs[1:]:
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    set_cell_shading(merged, "F2F2F2"); set_cell_border(merged)

def create_rank_table(doc, title, branches, mode):
    m_key, f_key = ("HU_M", "HU_F") if mode == "HU" else ("OHU_M", "OHU_F") if mode == "OHU" else ("STATE_M", "STATE_F")
    
    # Check if there is ANY regular data for this table
    has_any_regular = False
    for b_name, b_data in branches.items():
        m_data, f_data = b_data[m_key], b_data[f_key]
        if any(m_data.get(c) for c in REGULAR_CATS) or any(f_data.get(c) for c in REGULAR_CATS):
            has_any_regular = True
            break
            
    if not has_any_regular:
        return # Skip generating table if all regular cells are empty/dashed!

    def fill_cell(p, val, bold=False, left=False):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        if "\n" in str(val):
            top, bottom = str(val).split("\n", 1)
            make_run(p, top, bold=bold)
            br = p.add_run()
            br.font.size = Pt(7)
            br.add_break()
            make_run(p, bottom, bold=bold, size=7)
        else:
            make_run(p, val, bold=bold)

    shared_cols = [2 + CAT_ORDER.index(c) for c in SHARED_COLS]
    data_rows = []
    shared_pairs = []          # (male_data_row, female_data_row, [shared values])
    for b_name, b_data in branches.items():
        m_data, f_data = b_data[m_key], b_data[f_key]
        has_m = any(m_data.get(c) for c in CAT_ORDER)
        has_f = any(f_data.get(c) for c in CAT_ORDER)
        
        if has_m:
            row = [b_name, "Male"] + [m_data.get(c, "-----") for c in CAT_ORDER]
            data_rows.append(row)
        if has_f:
            branch_label = "" if has_m else b_name
            row = [branch_label, "Female"] + [f_data.get(c, "-----") for c in CAT_ORDER]
            if has_m:
                # EWS / TFWS / AI (JEE) are one value per branch -> shown once
                # in the vertically merged cell created below.
                shared_vals = []
                for c in SHARED_COLS:
                    v = m_data.get(c) or f_data.get(c) or "-----"
                    shared_vals.append(v)
                    row[2 + CAT_ORDER.index(c)] = ""
                shared_pairs.append((len(data_rows) - 1, len(data_rows), shared_vals))
            data_rows.append(row)
            
    if not data_rows: return

    tbl = doc.add_table(rows=len(data_rows)+1, cols=14)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    # JoSAA proportions, scaled so the 14 columns fill exactly
    # DOC_TABLE_WIDTH_IN - before this they summed to 11.86" and the table
    # ran past the page margins.
    w = _fit_widths([2.7, 0.7, 0.72, 0.72, 0.65, 0.65, 0.8, 0.8, 0.8, 0.65,
                     0.72, 0.65, 0.65, 0.65])
    for r in tbl.rows:
        for i, wi in enumerate(w): r.cells[i].width = Inches(wi)

    tbl._tbl.addprevious(
        parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="left"/><w:spacing w:before="120" w:after="60"/></w:pPr>'
                  f'<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/></w:rPr>'
                  f'<w:t>{title}</w:t></w:r></w:p>'))

    for i, col in enumerate(TABLE_COLS):
        c = tbl.rows[0].cells[i]; c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        make_run(p, col, bold=True)
        set_cell_shading(c, "4472C4"); set_cell_border(c)
        for run in p.runs:
            run._element.get_or_add_rPr().append(parse_xml(f'<w:color {nsdecls("w")} w:val="FFFFFF"/>'))

    for ri, rd in enumerate(data_rows):
        bg = "FFFFFF" if ri % 2 == 0 else "F2F7FC"
        for ci in range(14):
            c = tbl.rows[ri+1].cells[ci]; c.text = ""
            p = c.paragraphs[0]
            val = rd[ci] if ci < len(rd) else ""
            fill_cell(p, val, bold=(ci == 0), left=(ci == 0))
            set_cell_shading(c, bg); set_cell_border(c)

    # EWS / TFWS / AI (JEE): one value, centred across the Male+Female rows
    for m_idx, f_idx, vals in shared_pairs:
        for k, ci in enumerate(shared_cols):
            top = tbl.rows[m_idx + 1].cells[ci]
            bot = tbl.rows[f_idx + 1].cells[ci]
            merged = top.merge(bot)
            merged.text = ""
            fill_cell(merged.paragraphs[0], vals[k])
            merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            drop_border(top, 'bottom')     # no line through the merged cell
            drop_border(bot, 'top')

def add_page_break(doc):
    p = doc.add_paragraph()
    p._element.clear()
    p._element.append(parse_xml(f'<w:r {nsdecls("w")}><w:br w:type="page"/></w:r>'))

def _init_docx_document():
    doc = Document()
    for s in doc.sections:
        s.page_width, s.page_height = Cm(29.7), Cm(21.0)
        s.top_margin = s.bottom_margin = Inches(DOC_MARGIN_IN)
        s.left_margin = s.right_margin = Inches(DOC_MARGIN_IN)

    st = doc.styles['Normal']
    st.font.name = 'Arial'; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    st.paragraph_format.space_before = Pt(0); st.paragraph_format.space_after = Pt(0)
    return doc


def _add_college_to_doc(doc, college_data, mode="ALL", logo_path=None):
    """Append ONE college (logo header + info box + quota tables) to doc."""
    name = college_data["full_name"]
    branches = college_data["branches"]
    web_meta = college_data.get("web_meta", {})

    has_hu = college_data.get("has_hu", True)
    has_ohu = college_data.get("has_ohu", True)
    has_state = college_data.get("has_state", True)

    create_docx_header_logo(doc, logo_path)
    create_info_box(doc, name, web_meta)

    if mode == "HU":
        if has_hu:
            create_rank_table(doc, "Home University Cut-Off", branches, "HU")
    elif mode == "OHU":
        if has_ohu:
            create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
    elif mode == "STATE":
        # A college without State quota gets header + info box only -
        # exactly what the UI shows (an empty-state notice), not ALL tables.
        if has_state:
            create_rank_table(doc, "State Level Cut-Off", branches, "STATE")
    elif mode in ("HU_OHU", "4"):
        if has_hu:
            create_rank_table(doc, "Home University Cut-Off", branches, "HU")
        if has_ohu:
            if has_hu: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
    else:  # ALL
        rendered_count = 0
        if has_hu:
            create_rank_table(doc, "Home University Cut-Off", branches, "HU")
            rendered_count += 1
        if has_ohu:
            if rendered_count > 0: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
            rendered_count += 1
        if has_state:
            if rendered_count > 0: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "State Level Cut-Off", branches, "STATE")


def _safe_docx_name(name, fallback="College_Cutoff"):
    safe_name = re.sub(r'[<>:"/\\|?*]', '', str(name or '')).strip().strip('.').strip()
    return safe_name or fallback


def build_college_docx(college_data, mode="ALL", logo_path=None):
    """Build the college cut-off Word file fully in memory.

    Returns (bytes, filename).  No disk writes - works on read-only
    filesystems (Vercel) and never leaves stale files behind.
    Same content as the on-screen tables: logo header, info box, and one
    rank table per quota (HU / OHU / STATE) with Branch | Gender rows,
    CET categories and the merged EWS / TFWS / AI (JEE) cells.
    """
    doc = _init_docx_document()
    _add_college_to_doc(doc, college_data, mode=mode, logo_path=logo_path)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), f"{_safe_docx_name(college_data.get('full_name'))}.docx"


def build_book_docx(college_data_list, mode="ALL", logo_path=None, book_title="CET_Cutoff_Book"):
    """Build ONE combined Word book for many colleges fully in memory.

    Every college is rendered with the SAME current-data layout as the
    single-college download (logo header + info box + quota tables) and is
    separated from the next by a page break.  Returns (bytes, filename).
    """
    doc = _init_docx_document()
    first = True
    for college_data in (college_data_list or []):
        if not first:
            add_page_break(doc)
        first = False
        _add_college_to_doc(doc, college_data, mode=mode, logo_path=logo_path)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), f"{_safe_docx_name(book_title, 'CET_Cutoff_Book')}.docx"


def _is_cet_college_table(tbl):
    """True when a table carries a CET info block ('College Code & Name:')."""
    try:
        for row in tbl.rows:
            for c in row.cells:
                if 'College Code' in c.text:
                    return True
    except Exception:
        pass
    return False


# ------------------------------------------------------------------
# Template book updater: replace DATA inside the template's CET tables
# (in-place, exactly like the JoSAA book_updater - the table format,
#  info box, branch order and header rows are all preserved).
# ------------------------------------------------------------------
_TPL_ABBR = {
    "eng": "engineering", "comm": "communication", "sci": "science",
    "tech": "technology", "info": "information", "comp": "computer",
    "electron": "electronics", "elec": "electrical", "mech": "mechanical",
    "chem": "chemical", "tc": "telecommunication", "telecom": "telecommunication",
    "instru": "instrumentation", "inst": "instrumentation", "prod": "production",
    "it": "information technology", "cse": "computer science",
    "auto": "automobile", "biotech": "biotechnology", "civil": "civil",
    "aids": "artificial intelligence and data science",
    "aiml": "artificial intelligence and machine learning",
    "ai": "artificial intelligence", "ml": "machine learning",
    "ds": "data science", "iot": "internet of things",
    "vlsi": "vlsi", "mba": "business administration",
}
_TPL_NOISE = set("and of the in with for to a an on by at from as is are "
                 "branch course engineering technology".split())


def _tpl_norm_text(s):
    s = str(s or "").replace("\xa0", " ")
    s = s.lower().replace("&", " and ")
    s = re.sub(r"\(\s*\d+\s*\)", " ", s)          # drop intake "(60)"
    s = re.sub(r"\[\s*\d+\s*\]", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _tpl_tokens(s):
    toks = []
    for w in _tpl_norm_text(s).split():
        if w in _TPL_NOISE or len(w) < 2:
            continue
        toks.extend(_TPL_ABBR.get(w, w).split())
    return toks


def _branch_sim(a, b):
    """Similarity in [0,1] between a template branch label and a data branch name."""
    A, B = _tpl_tokens(a), _tpl_tokens(b)
    if not A or not B:
        return 0.0
    if A == B:
        return 1.0
    used, hit = set(), 0
    for x in A:
        for j, y in enumerate(B):
            if j in used:
                continue
            if x == y or (min(len(x), len(y)) >= 5 and
                          SequenceMatcher(None, x, y).ratio() >= 0.85):
                used.add(j); hit += 1; break
    if hit:
        cont = hit / min(len(A), len(B))
        jac = hit / (len(A) + len(B) - hit)
        if cont >= 0.999:
            return max(jac, 0.6)
        if cont >= 0.5:                       # one side's words fold into the other
            return max(jac, 0.5)
        if jac >= 0.4 and hit >= 2:
            return jac
        return 0.0
    # concatenated form (e.g. 'petro chemical' == 'petrochemical')
    ca, cb = "".join(A), "".join(B)
    if SequenceMatcher(None, ca, cb).ratio() >= 0.9:
        return 0.5
    return 0.0


def _tc_text(tc):
    return "".join(t.text or "" for t in tc.iter(qn('w:t'))).strip()


def _set_tc_text(tc, text):
    """Replace a cell's text keeping the first run's formatting (JoSAA style)."""
    text = str(text).replace("\n", " ")
    ps = tc.findall(qn('w:p'))
    if not ps:
        p = parse_xml(f'<w:p {nsdecls("w")}/>')
        tc.append(p)
        ps = [p]
    for extra in ps[1:]:
        tc.remove(extra)
    p = ps[0]
    rpr = None
    runs = p.findall(qn('w:r'))
    if runs:
        rp = runs[0].find(qn('w:rPr'))
        if rp is not None:
            rpr = copy.deepcopy(rp)
    for child in list(p):
        if child.tag != qn('w:pPr'):
            p.remove(child)
    r = parse_xml(f'<w:r {nsdecls("w")}/>')
    if rpr is not None:
        r.append(rpr)
    t = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve"/>')
    t.text = text
    r.append(t)
    p.append(r)


def _quota_from_template(info_text):
    """Read the quota out of the template's 'Available Cut-off : ...' info row."""
    l = info_text.lower()
    has_hu = re.search(r'\bhu\b', l) is not None or 'hu+' in l or 'hu +' in l
    has_ohu = re.search(r'\bohu\b', l) is not None
    if has_hu and has_ohu:
        return ['HU', 'OHU']
    if has_ohu:
        return ['OHU']
    if has_hu:
        return ['HU']
    return ['STATE']


def _extract_college(trs, hidx):
    """Pull (college name, code) from the template info rows."""
    for tr in trs[:hidx]:
        cells = tr.findall(qn('w:tc'))
        texts = [_tc_text(tc) for tc in cells]
        for i, tx in enumerate(texts):
            if 'College Code' in tx:
                for tx2 in texts[i + 1:]:
                    m = re.search(r'(\d{3,6})\s*-\s*(.+)', tx2)
                    if m:
                        return m.group(2).strip(), m.group(1)
    # fallback: any cell holding 'code-name'
    for tr in trs[:hidx]:
        for tc in tr.findall(qn('w:tc')):
            m = re.search(r'(\d{3,6})\s*-\s*([A-Za-z].+)', _tc_text(tc))
            if m:
                return m.group(2).strip(), m.group(1)
    return "", ""


def _find_cet_header(trs):
    for i, tr in enumerate(trs):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < 13:
            continue
        c0 = _tc_text(tcs[0]).lower()
        c1 = _tc_text(tcs[1]).lower() if len(tcs) > 1 else ""
        if c0.startswith('branch') and 'gender' in c1:
            return i
    return None


def _update_cet_table(tbl, college_data):
    """Replace the rank cells of ONE template CET table with current data."""
    trs = tbl._tbl.findall(qn('w:tr'))
    hidx = _find_cet_header(trs)
    if hidx is None:
        return 0

    info = " ".join(_tc_text(tc) for tr in trs[:hidx] for tc in tr.findall(qn('w:tc')))
    quotas = _quota_from_template(info)
    branches = college_data.get('branches', {}) or {}

    lookups = []
    for q in quotas:
        rows = []
        for bname, bd in branches.items():
            rows.append({
                'name': bname,
                'M': bd.get(q + '_M', {}) or {},
                'F': bd.get(q + '_F', {}) or {},
            })
        lookups.append(rows)

    def best_row(label):
        best, bs = None, 0.0
        for rows in lookups:
            for r in rows:
                sc = _branch_sim(label, r['name'])
                if sc > bs:
                    bs, best = sc, r
        return best if bs >= 0.4 else None

    cur = None
    updated = 0
    for tr in trs[hidx + 1:]:
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < 13:
            continue
        lbl = _tc_text(tcs[0])
        gender = _tc_text(tcs[1]).lower()
        if lbl:
            cur = best_row(lbl)
        g = 'F' if 'female' in gender else 'M'
        mdata = cur['M'] if cur else {}
        fdata = cur['F'] if cur else {}
        for i, cat in enumerate(CAT_ORDER):
            ci = 2 + i
            if ci >= len(tcs):
                break
            if cat in SHARED_COLS:
                if g == 'F':
                    continue            # merged cell -> written on the male row
                val = mdata.get(cat) or fdata.get(cat)
            else:
                val = (mdata if g == 'M' else fdata).get(cat)
            _set_tc_text(tcs[ci], val if val else "-----")
            updated += 1
    return updated


def _update_template_meta(tbl, college_data, hidx):
    """Refresh University / Status text in the template info box."""
    wm = college_data.get('web_meta', {}) or {}
    uni = college_data.get('university') or wm.get('university') or ""
    st = college_data.get('status') or wm.get('status') or ""
    if not (uni or st):
        return
    for tr in tbl._tbl.findall(qn('w:tr'))[:hidx]:
        for tc in tr.findall(qn('w:tc')):
            t = _tc_text(tc)
            if uni and t.startswith('University'):
                _set_tc_text(tc, f"University:  {uni}")
            elif st and t.startswith('Status'):
                _set_tc_text(tc, f"Status: {st}")


def _set_college_identity(tbl_el, hidx, college_data):
    """Write 'CODE-Name' into the info box 'College Code & Name' value cell."""
    code = str(college_data.get('college_code') or '').strip()
    name = college_data.get('full_name') or college_data.get('college_name') or ''
    name = re.sub(r'^\s*\d{3,6}\s*-\s*', '', name).strip()
    if not name:
        return
    val = f"{code}-{name}" if code else name
    trs = tbl_el.findall(qn('w:tr'))
    lim = hidx if hidx is not None else len(trs)
    for tr in trs[:lim]:
        cells = tr.findall(qn('w:tc'))
        texts = [_tc_text(tc) for tc in cells]
        for i, tx in enumerate(texts):
            if 'College Code' in tx:
                for tc2, tx2 in zip(cells[i + 1:], texts[i + 1:]):
                    if re.search(r'\d{3,6}\s*-\s*\S', tx2):
                        _set_tc_text(tc2, val)
                        return
                if i + 1 < len(cells):
                    _set_tc_text(cells[i + 1], val)
                    return
    for tr in trs[:lim]:
        for tc in tr.findall(qn('w:tc')):
            if re.search(r'\d{3,6}\s*-\s*[A-Za-z]', _tc_text(tc)):
                _set_tc_text(tc, val)
                return


def _write_cet_row(tcs, label, gender, mdata, fdata):
    """Fill one cloned data row (label / gender / category ranks)."""
    if len(tcs) < 3:
        return 0
    _set_tc_text(tcs[0], label)
    _set_tc_text(tcs[1], 'Male' if gender == 'M' else 'Female')
    n = 0
    for i, cat in enumerate(CAT_ORDER):
        ci = 2 + i
        if ci >= len(tcs):
            break
        if cat in SHARED_COLS:
            if gender == 'F':
                continue            # merged cell -> written on the male row
            val = mdata.get(cat) or fdata.get(cat)
        else:
            val = (mdata if gender == 'M' else fdata).get(cat)
        _set_tc_text(tcs[ci], val if val else "-----")
        n += 1
    return n


_BLANK_INFO_PREFIXES = ('ame grading', 'total seats', 'fee', 'address',
                        'contact', 'web', 'email', 'hostel', 'established')


def _blank_unknown_info(tbl_el, hidx):
    """Blank info-box fields we have no data for (address, web, fee, ...)."""
    trs = tbl_el.findall(qn('w:tr'))
    lim = hidx if hidx is not None else len(trs)
    for tr in trs[:lim]:
        for tc in tr.findall(qn('w:tc')):
            t = _tc_text(tc).strip()
            low = t.lower()
            if any(low.startswith(p) for p in _BLANK_INFO_PREFIXES):
                label = t.split(':', 1)[0]
                label = re.sub(r'\s+\d.*$', '', label).strip()
                _set_tc_text(tc, f"{label}:" if ':' in t else label)


def _build_generated_cet(tbl_el, college_data):
    """Rebuild a cloned blueprint table with the college's own branches."""
    trs = tbl_el.findall(qn('w:tr'))
    hidx = _find_cet_header(trs)
    if hidx is None:
        return 0

    info = " ".join(_tc_text(tc) for tr in trs[:hidx] for tc in tr.findall(qn('w:tc')))
    template_quotas = _quota_from_template(info)
    # fall back through every available quota so generated tables still fill
    quotas = list(template_quotas)
    for q in ('STATE', 'HU', 'OHU'):
        if q not in quotas:
            quotas.append(q)

    branches = college_data.get('branches', {}) or {}

    def merge(g):
        out = {}
        for q in quotas:
            for cat, val in (bd.get(f"{q}_{g}") or {}).items():
                if val and cat not in out:
                    out[cat] = val
        return out

    prepared = []
    for bname, bd in branches.items():
        prepared.append((bname, merge('M'), merge('F')))
    if not prepared:
        return 0

    data_trs = trs[hidx + 1:]
    proto_m = copy.deepcopy(data_trs[0]) if data_trs else None
    proto_f = None
    for tr in data_trs[1:]:
        if len(tr.findall(qn('w:tc'))) >= 13:
            proto_f = copy.deepcopy(tr)
            break
    if proto_m is None or proto_f is None:
        return 0

    anchor = trs[hidx]
    for tr in data_trs:
        tbl_el.remove(tr)

    updated = 0
    for bname, mdata, fdata in prepared:
        mr = copy.deepcopy(proto_m)
        anchor.addnext(mr)
        anchor = mr
        updated += _write_cet_row(mr.findall(qn('w:tc')), bname, 'M', mdata, fdata)

        fr = copy.deepcopy(proto_f)
        anchor.addnext(fr)
        anchor = fr
        updated += _write_cet_row(fr.findall(qn('w:tc')), '', 'F', mdata, fdata)
    return updated


def _renumber_cet(tbl_el, n):
    """Set the sequential 'Sr.' number cell of a template CET table."""
    trs = tbl_el.findall(qn('w:tr'))
    hidx = _find_cet_header(trs)
    limit = hidx if hidx is not None else len(trs)
    for tr in trs[:limit]:
        for tc in tr.findall(qn('w:tc')):
            if re.fullmatch(r'\d+\s*\.', _tc_text(tc).strip()):
                _set_tc_text(tc, f"{n}.")
                return


CANON_BASE = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)",
              "NT-3 (D)", "OBC", "SEBC"]
MISSING = "-----"
MI_MARKERS = ("thakur", "sanghvi", "conceicao rodrigues", "vivekanand",
              "k.j. somaiya", "kj somaiya", "thadomal", "mit academy",
              "dattajirao kadam")
MI_CODES = {"3176", "3199", "3184", "3185", "3209", "3182", "6146", "6222"}


def _canon_header(mi):
    return (["Branch (Intake)", "Gender"] + CANON_BASE +
            [("MI" if mi else "EWS"), "TFWS", "AI (JEE)"])


def _is_mi_college(cd):
    code = re.sub(r'\D', '', str(cd.get('college_code') or cd.get('code') or ''))
    if code and int(code) in {int(c) for c in MI_CODES}:
        return True
    name = (cd.get('full_name') or cd.get('college_name') or '').lower()
    return any(m in name for m in MI_MARKERS)


def _strip_intake(label):
    return re.sub(r'\s*\(\s*\d+\s*\)\s*$', '', str(label or '')).strip()


def _rank_val(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ('', MISSING, 'None', 'nan', 'NaN', '-', '--', '----'):
        return None
    m = re.search(r'\d[\d,]*', s)
    return m.group(0).replace(',', '') if m else None


def _tc_pr(tc):
    return tc.find(qn('w:tcPr'))


def _is_vmerge_cont(tc):
    pr = _tc_pr(tc)
    if pr is None:
        return False
    vm = pr.find(qn('w:vMerge'))
    if vm is None:
        return False
    val = vm.get(qn('w:val'))
    return val is None or val == 'continue'


def _is_hmerge_cont(tc):
    pr = _tc_pr(tc)
    if pr is None:
        return False
    hm = pr.find(qn('w:hMerge'))
    if hm is None:
        return False
    val = hm.get(qn('w:val'))
    return val is None or val == 'continue'


def _row_texts(tr):
    return [_tc_text(tc) for tc in tr.findall(qn('w:tc'))]


def _block_info_row(tr):
    return any('College Code' in t for t in _row_texts(tr))


def _row_is_empty(tr):
    return all(not t.strip() for t in _row_texts(tr))


def _extract_college_from_block(trs):
    for tr in trs[:6]:
        cells = tr.findall(qn('w:tc'))
        texts = [_tc_text(tc) for tc in cells]
        for i, tx in enumerate(texts):
            if 'College Code' in tx:
                for tx2 in texts[i + 1:]:
                    m = re.search(r'(\d{3,6})\s*-\s*(.+)', tx2)
                    if m:
                        return m.group(1), m.group(2).strip()
    for tr in trs[:6]:
        for tc in tr.findall(qn('w:tc')):
            m = re.search(r'(\d{3,6})\s*-\s*([A-Za-z].+)', _tc_text(tc))
            if m:
                return m.group(1), m.group(2).strip()
    return "", ""


def _best_branch(label, branches):
    best, bs = None, 0.0
    for bname in branches:
        sc = _branch_sim(label, bname)
        if sc > bs:
            bs, best = sc, bname
    return best if bs >= 0.4 else None


def _shared_rank(bd, quota, gender, cat):
    for q in (quota, 'STATE', 'HU', 'OHU'):
        v = _rank_val((bd.get(f"{q}_{gender}") or {}).get(cat))
        if v:
            return v
    return None


def _update_block(block, cd, mi):
    """Replace the header + rank cells of one college block (in place)."""
    branches = cd.get('branches', {}) or {}
    canon = _canon_header(mi)
    shared_cat = 'MI' if mi else 'EWS'
    info = " ".join(_tc_text(tc) for tr in block for tc in tr.findall(qn('w:tc')))
    cur_q = _quota_from_template(info)
    cur_q = cur_q[0] if isinstance(cur_q, list) else cur_q

    cur_branch = None
    seen_header = False
    header_tr = None
    hdr_off = 0
    cells = rows = 0
    matched_branches = set()
    for tr in block:
        tcs = tr.findall(qn('w:tc'))
        if not tcs:
            continue
        texts = [_tc_text(tc) for tc in tcs]

        hi = None
        for i in range(len(tcs) - 1):
            if (texts[i].strip().lower().startswith('branch') and
                    'gender' in texts[i + 1].strip().lower()):
                hi = i
                break
        if hi is not None:
            hdr_off = hi
            for k, tc in enumerate(tcs):
                ci = k - hi
                if 0 <= ci < len(canon) and not _is_hmerge_cont(tc):
                    _set_tc_text(tc, canon[ci])
            seen_header = True
            header_tr = tr
            continue
        if not seen_header:
            continue                       # skip the info-box rows

        off = hdr_off
        low = " ".join(t.strip().lower() for t in texts)
        is_div = len(tcs) <= off + 3
        if is_div and ('home university' in low or 'state' in low):
            if 'other' in low:
                cur_q = 'OHU'
            elif 'home university' in low:
                cur_q = 'HU'
            else:
                cur_q = 'STATE'
            nxt = tr.getnext()
            if header_tr is not None and nxt is not None:
                nt = _row_texts(nxt)
                if (nt and not nt[0].strip().lower().startswith('branch')
                        and len(nxt.findall(qn('w:tc'))) >= off + 13):
                    tr.addnext(copy.deepcopy(header_tr))   # repeat full header
            continue

        if len(tcs) < off + 3:
            continue

        g1 = texts[off + 1].strip().lower()
        if 'female' in g1:
            g = 'F'
        elif 'male' in g1:
            g = 'M'
        else:
            g = 'M' if texts[off].strip() else 'F'
        if texts[off].strip():
            cur_branch = _strip_intake(texts[off])
        bname = _best_branch(cur_branch, branches) if cur_branch else None
        if bname:
            matched_branches.add(bname)
        bd = branches.get(bname) if bname else None

        if not (_is_vmerge_cont(tcs[off + 1]) or _is_hmerge_cont(tcs[off + 1])):
            _set_tc_text(tcs[off + 1], 'Male' if g == 'M' else 'Female')

        for i, cat in enumerate(CANON_BASE):
            ci = off + 2 + i
            if ci >= len(tcs):
                break
            if _is_vmerge_cont(tcs[ci]) or _is_hmerge_cont(tcs[ci]):
                continue
            v = _rank_val((bd.get(f"{cur_q}_{g}") or {}).get(cat)) if bd else None
            _set_tc_text(tcs[ci], v or MISSING)
            cells += 1
        for n_off, cat in ((off + 11, shared_cat), (off + 12, 'TFWS'),
                           (off + 13, 'AI (JEE)')):
            if n_off >= len(tcs):
                continue
            if _is_vmerge_cont(tcs[n_off]) or _is_hmerge_cont(tcs[n_off]):
                continue
            if bd and n_off == off + 11:
                # minority seat may be stored as MI or (in some Excel years) EWS
                v = (_shared_rank(bd, cur_q, g, shared_cat) or
                     _shared_rank(bd, cur_q, g, 'EWS' if mi else 'MI'))
            else:
                v = _shared_rank(bd, cur_q, g, cat) if bd else None
            _set_tc_text(tcs[n_off], v or MISSING)
            cells += 1
        rows += 1

    missing = [b for b in branches if b not in matched_branches]
    return cells, rows, missing


def _update_block_meta(block, cd):
    wm = cd.get('web_meta', {}) or {}
    uni = cd.get('university') or wm.get('university') or ""
    st = cd.get('status') or wm.get('status') or ""
    if not (uni or st):
        return
    for tr in block:
        for tc in tr.findall(qn('w:tc')):
            t = _tc_text(tc)
            if uni and t.startswith('University'):
                _set_tc_text(tc, f"University:  {uni}")
            elif st and t.startswith('Status'):
                _set_tc_text(tc, f"Status: {st}")


def _renumber_block(block, n):
    for tr in block:
        for tc in tr.findall(qn('w:tc')):
            if re.fullmatch(r'\d+\s*\.', _tc_text(tc).strip()):
                _set_tc_text(tc, f"{n}.")
                return


def _append_after(el, anchor, node):
    if anchor is not None:
        anchor.addnext(node)
    else:
        el.append(node)
    return node


def _page_break_before(el):
    """Force the element to start on a fresh page (via the separator paragraph)."""
    prev = el.getprevious()
    if prev is not None and prev.tag == qn('w:p'):
        p = prev
    else:
        p = OxmlElement('w:p')
        el.addprevious(p)
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        p.insert(0, ppr)
    if ppr.find(qn('w:pageBreakBefore')) is None:
        pbb = OxmlElement('w:pageBreakBefore')
        ppr.insert(0, pbb)


def _mark_repeat_header(tbl_el):
    """Mark the table's header row so Word repeats it when the table overflows
    onto a following page."""
    for tr in tbl_el.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < 13:
            continue
        texts = [_tc_text(tc).strip().lower() for tc in tcs]
        is_hdr = any(texts[i].startswith('branch') and i + 1 < len(texts)
                     and 'gender' in texts[i + 1] for i in range(len(texts) - 1))
        if not is_hdr:
            continue
        trpr = tr.find(qn('w:trPr'))
        if trpr is None:
            trpr = OxmlElement('w:trPr')
            tr.insert(0, trpr)
        if trpr.find(qn('w:tblHeader')) is None:
            trpr.append(OxmlElement('w:tblHeader'))
        break


def _is_womens_college(cd):
    """True only when the college has female data but no male data at all."""
    branches = cd.get('branches', {}) or {}
    has_f = False
    for bd in branches.values():
        for g in ('M', 'F'):
            for q in ('HU', 'OHU', 'STATE'):
                d = bd.get(f"{q}_{g}") or {}
                if any(_rank_val(v) for v in d.values()):
                    if g == 'M':
                        return False
                    has_f = True
    return has_f


def _gen_row(proto, label, gender):
    r = copy.deepcopy(proto)
    tcs = r.findall(qn('w:tc'))
    if tcs:
        _set_tc_text(tcs[0], label)
    if len(tcs) > 1:
        _set_tc_text(tcs[1], gender)
    return r


def _gen_sections(cd):
    secs = []
    if cd.get('has_hu'):
        secs.append(('HU', 'Home University'))
    if cd.get('has_ohu'):
        secs.append(('OHU', 'Other than Home University'))
    if cd.get('has_state') or not secs:
        secs.append(('STATE', 'State Quota'))
    return secs


def build_generated_table(proto_el, hu_div, ohu_div, cd, mi):
    """Clone a clean template CET table and rebuild it for a college that has no
    block in the template, using that college's own branches (JoSAA technique:
    clone existing rows only, never restyle)."""
    if proto_el is None:
        return None
    el = copy.deepcopy(proto_el)
    trs = el.findall(qn('w:tr'))
    hidx = _find_cet_header(trs)
    if hidx is None:
        return None
    header_tr = trs[hidx]
    data_trs = trs[hidx + 1:]
    proto_m = proto_f = None
    for r in data_trs:
        tcs = r.findall(qn('w:tc'))
        if not tcs:
            continue
        if _is_vmerge_cont(tcs[0]):
            if proto_f is None:
                proto_f = r
        elif proto_m is None:
            proto_m = r
    if proto_m is None:
        proto_m = data_trs[0] if data_trs else None
    if proto_f is None:
        proto_f = proto_m
    if proto_m is None:
        return None
    for r in data_trs:
        el.remove(r)

    branches = cd.get('branches', {}) or {}
    if not branches:
        return None
    womens = _is_womens_college(cd)
    secs = _gen_sections(cd)
    show_div = not (len(secs) == 1 and secs[0][0] == 'STATE')

    el.remove(header_tr)
    anchor = trs[hidx - 1] if hidx > 0 else None
    for q, label in secs:
        if show_div and hu_div is not None and ohu_div is not None:
            div = copy.deepcopy(hu_div if q == 'HU' else ohu_div)
            dtcs = div.findall(qn('w:tc'))
            if dtcs:
                _set_tc_text(dtcs[0], label)
            anchor = _append_after(el, anchor, div)
        anchor = _append_after(el, anchor, copy.deepcopy(header_tr))
        for bname in branches:
            anchor = _append_after(
                el, anchor, _gen_row(proto_m, bname, 'Female' if womens else 'Male'))
            if not womens:
                anchor = _append_after(el, anchor, _gen_row(proto_f, '', 'Female'))

    new_trs = el.findall(qn('w:tr'))
    new_hidx = _find_cet_header(new_trs)
    _set_college_identity(el, new_hidx if new_hidx is not None else len(new_trs), cd)
    _blank_unknown_info(el, new_hidx)
    lbl = ('HU+OHU' if (cd.get('has_hu') and cd.get('has_ohu'))
           else 'HU' if cd.get('has_hu')
           else 'OHU' if cd.get('has_ohu') else 'State')
    for tr in new_trs[:(new_hidx or 0)]:
        for tc in tr.findall(qn('w:tc')):
            if _tc_text(tc).strip().lower().startswith('available'):
                _set_tc_text(tc, f"Available Cut-off : {lbl} Quota")
    _update_block(new_trs, cd, mi)
    return el


def build_template_book(template_path, college_data_list, mode="ALL",
                        logo_path=None, book_title="CET_Cutoff_Book",
                        progress_cb=None, report_out=None):
    """Update the template book by replacing the DATA inside its CET tables only.

    The template's cover / intro / index / info boxes / table format / header
    styling / branch order / seat counts are preserved exactly (the JoSAA
    book_updater technique - only cell text changes; no restyling).  Every
    college *block* is updated independently, so multi-college tables are
    handled correctly.  Blocks whose college is not selected are dropped and
    JoSAA IIT/NIT/IIIT tables are dropped.  Returns (bytes, filename).
    """
    if progress_cb:
        progress_cb(30, "Loading Word template ...")
    doc = Document(template_path)

    by_code, by_name = {}, {}
    for cd in (college_data_list or []):
        for field in ('college_code', 'code'):
            d = re.sub(r'\D', '', str(cd.get(field) or ''))
            if d:
                by_code[int(d)] = cd
        nm = cd.get('full_name') or cd.get('college_name') or ''
        nm = re.sub(r'^\s*\d{3,6}\s*-\s*', '', nm).strip()
        if nm:
            by_name.setdefault(_tpl_norm_text(nm), cd)
    name_items = list(by_name.items())

    def _match_cd(code, name):
        # The template carries a numeric code for every college; the CET Excel
        # uses the same number (leading zeros only).  Match on that number and
        # never fuzzy-match a known-but-unselected code to a different college.
        d = re.sub(r'\D', '', code or '')
        if d:
            return by_code.get(int(d))
        k = _tpl_norm_text(name)
        if k and k in by_name:
            return by_name[k]
        return None

    report = {"tables": 0, "blocks": 0, "removed_tables": 0, "cells": 0,
              "rows": 0, "mi_tables": 0, "unmatched_blocks": [],
              "missing_branches": [], "skipped_selected": []}

    cet_tables, josaa_els = [], []
    for tbl in doc.tables:
        if _is_cet_college_table(tbl):
            cet_tables.append(tbl)
        elif len(tbl.columns) >= 20:
            josaa_els.append(tbl._tbl)

    # Blueprint for colleges that have no block in the template: a clean
    # single-college CET table (info box + header + a male/female row pair) plus
    # the divider rows.  Everything is cloned from the template - never rebuilt.
    proto_el = hu_div = ohu_div = None
    for tbl in cet_tables:
        trs = tbl._tbl.findall(qn('w:tr'))
        if (proto_el is None and trs and _block_info_row(trs[0])
                and _find_cet_header(trs) is not None):
            proto_el = tbl._tbl
        for tr in trs:
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) != 1:
                continue
            t = _tc_text(tcs[0]).strip().lower()
            if t == 'home university' and hu_div is None:
                hu_div = tr
            elif 'other than home university' in t and ohu_div is None:
                ohu_div = tr

    # element just before the CET section: the empty-paragraph cleanup below
    # only runs from here on, so the cover / index pages are never touched.
    cet_anchor = cet_tables[0]._tbl.getprevious() if cet_tables else None

    matched_ids = set()
    kept_els = []
    sr_no = 0
    total = len(cet_tables)
    for ti, tbl in enumerate(cet_tables):
        if progress_cb:
            progress_cb(35 + int(45 * (ti + 1) / max(1, total)),
                        f"Updating table {ti + 1}/{total} ...")
        el = tbl._tbl
        trs = el.findall(qn('w:tr'))
        starts = [i for i, tr in enumerate(trs) if _block_info_row(tr)]
        if not starts:
            continue
        bounds = [(starts[j], starts[j + 1] if j + 1 < len(starts) else len(trs))
                  for j in range(len(starts))]

        remove_rows = []
        any_kept = False
        for (s, e) in bounds:
            block = trs[s:e]
            code, name = _extract_college_from_block(block)
            cd = _match_cd(code, name)
            if cd is None or id(cd) in matched_ids:
                remove_rows.extend(block)
                if cd is None:
                    report["unmatched_blocks"].append(name or code)
                continue
            matched_ids.add(id(cd))
            any_kept = True
            if os.environ.get("CET_DEBUG"):
                report.setdefault("pairs", []).append(
                    f"{(name or code)[:40]} -> {cd.get('college_code')}")
            for tr in block:
                if _row_is_empty(tr):
                    remove_rows.append(tr)
            mi = _is_mi_college(cd)
            if mi:
                report["mi_tables"] += 1
            sr_no += 1
            c, r, missing = _update_block(block, cd, mi)
            report["cells"] += c
            report["rows"] += r
            report["blocks"] += 1
            report["missing_branches"].extend(
                f"{name}: {b}" for b in missing)
            _update_block_meta(block, cd)
            _renumber_block(block, sr_no)

        if not any_kept:
            # no selected college in this table (it only held intro links or
            # unselected colleges) -> drop the whole table
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
            report["removed_tables"] += 1
            continue

        for tr in remove_rows:
            el.remove(tr)
        report["tables"] += 1
        kept_els.append(el)

    # ---- generate tables for selected colleges that have no template block ----
    gen_anchor = None
    for tbl in reversed(cet_tables):
        if tbl._tbl.getparent() is not None:
            gen_anchor = tbl._tbl
            break
    if gen_anchor is None:
        gen_anchor = cet_anchor
    generated = 0
    for cd in (college_data_list or []):
        if id(cd) in matched_ids:
            continue
        try:
            gen = build_generated_table(proto_el, hu_div, ohu_div, cd,
                                        _is_mi_college(cd))
        except Exception:
            gen = None
        if gen is None:
            continue
        sr_no += 1
        _renumber_block(gen.findall(qn('w:tr')), sr_no)
        if gen_anchor is not None and gen_anchor.getparent() is not None:
            gen_anchor.addnext(gen)
        elif cet_anchor is not None:
            cet_anchor.addnext(gen)
        else:
            doc.element.body.append(gen)
        gen_anchor = gen
        if _is_mi_college(cd):
            report["generated_mi"] = report.get("generated_mi", 0) + 1
        for tr in gen.findall(qn('w:tr')):
            t = _row_texts(tr)
            if len(t) >= 2 and t[1].strip().lower() in ('male', 'female'):
                report["rows"] += 1
        matched_ids.add(id(cd))
        generated += 1
    report["generated_tables"] = generated
    report["tables"] += generated

    report["skipped_selected"] = [
        cd.get('full_name') or cd.get('college_name')
        for cd in (college_data_list or []) if id(cd) not in matched_ids]

    for el in josaa_els:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # Removed tables (CET + the JoSAA tables) leave their spacer paragraphs
    # behind, piling up into blank pages.  Collapse any run of >=3 empty
    # paragraphs down to one (no text / no image / no section mark), keeping the
    # one separator so adjacent tables never merge in Word.
    def _is_blank_p(el):
        if el.tag != qn('w:p'):
            return False
        if ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip():
            return False
        if el.findall('.//' + qn('w:drawing')) or el.findall('.//' + qn('w:pict')):
            return False
        if el.findall('.//' + qn('w:sectPr')):
            return False
        return True

    collapse = []
    run = []
    node = list(doc.element.body)[0] if len(doc.element.body) else None
    while node is not None:
        nxt = node.getnext()
        if _is_blank_p(node):
            run.append(node)
        else:
            if len(run) >= 3:
                collapse.extend(run[1:])
            run = []
        node = nxt
    if len(run) >= 3:
        collapse.extend(run[1:])
    for n in collapse:
        n.getparent().remove(n)
    report["collapsed_paras"] = len(collapse)

    # ---- keep only the college tables ----
    cet_now = [t for t in doc.tables if _is_cet_college_table(t)]
    # drop any intro rows the template merged into a college's table
    for t in cet_now:
        trs = t._tbl.findall(qn('w:tr'))
        first_info = next((i for i, tr in enumerate(trs) if _block_info_row(tr)),
                          None)
        if first_info:
            for tr in trs[:first_info]:
                t._tbl.remove(tr)
    # delete everything above the first college table (cover / index / intro)
    if cet_now:
        body_el = doc.element.body
        first_el = cet_now[0]._tbl
        for node in list(body_el):
            if node is first_el:
                break
            body_el.remove(node)
    # remove the decorative page border that the template drew round every page
    for sect in doc.sections:
        sp = sect._sectPr
        pb = sp.find(qn('w:pgBorders'))
        if pb is not None:
            sp.remove(pb)
    # one table per page: a table that overflows simply continues on the next
    # page with its repeated header row
    for t in cet_now[1:]:
        _page_break_before(t._tbl)
    for t in cet_now:
        _mark_repeat_header(t._tbl)
    report["cet_tables_final"] = len(cet_now)

    report["selected"] = len(college_data_list or [])
    if report_out is not None:
        try:
            report_out.clear()
            report_out.update(report)
        except Exception:
            pass

    if progress_cb:
        progress_cb(90, "Saving the combined book ...")
    if os.environ.get("CET_DEBUG"):
        import json as _json
        for p in report.get("pairs", []):
            print("PAIR", p)
        print("MISSING", len(report["missing_branches"]))
        for m in report["missing_branches"][:80]:
            print("  m:", m)
        print("UNMATCHED", len(report["unmatched_blocks"]))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), f"{_safe_docx_name(book_title, 'CET_Cutoff_Book')}.docx"


def generate_word_document(college_data, mode="ALL", logo_path=None, output_folder="College_Data"):
    """Back-compat wrapper: build in memory, then persist to output_folder."""
    blob, file_name = build_college_docx(college_data, mode=mode, logo_path=logo_path)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    file_path = os.path.join(output_folder, file_name)
    with open(file_path, "wb") as f:
        f.write(blob)
    return file_path, os.path.basename(file_path)
