"""
Confluence IT Purchase 2026 — Scraper & SAP Upload Generator
=============================================================
Flow:
  1. Scrape Confluence page → raw table
  2. Read SAP template .xls → extract header row as column definitions
  3. Match Confluence columns to template headers (case-insensitive)
  4. Save as .xls with template headers

KRAFTON Brand Colors: #000000 / #ffffff / #f9423a
No log file output — GUI log only.
"""

import sys
import threading
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

# ─────────────────────────────────────────────
# VERSION
# ─────────────────────────────────────────────
VERSION = "0.1.1"

# ─────────────────────────────────────────────
# HARDCODED CONNECTION (not shown in UI)
# ─────────────────────────────────────────────
DEFAULT_BASE_URL       = "https://krafton.atlassian.net"
DEFAULT_PAGE_ID        = "864695451"
DEFAULT_CONFLUENCE_URL = "https://krafton.atlassian.net/wiki/spaces/PublishingGroup/pages/864695451/IT+Purchase+-+2026"
SESSION_CACHE_FILE     = "scraper_session.json"
DEFAULT_EMAIL     = "jmlee@krafton.com"
DEFAULT_API_TOKEN = "ATATT3xFfGF0ZpVz5H3LflT0ndS2Yktm2RKwg-J4K8o4uj9HWTeopDaHl9KbLGGx6ldIyiOA1_Q69ZeDbGdyFpTfnKyc9aqw1oDgQWMdwoRbB2XSFrn4_KElpbqSYoRbFPKfVFmlkgqxtnW78rj8UgA43VPd5Kn_vOThXlsRGsD9Rqi37LpsEUQ=1EA82E42"

# ─────────────────────────────────────────────
# DEFAULTS
# ─────────────────────────────────────────────
DEFAULT_OUTPUT   = r"C:\Users\Ajak\Desktop\Confluence Scrapper"
DEFAULT_TEMPLATE     = r"C:\Users\Ajak\Desktop\Confluence Scrapper\Asset registration master_1400_template.xls"
DEFAULT_VENDOR_MASTER  = r"C:\Users\Ajak\Desktop\Confluence Scrapper\SAP Vendor Master_20260313.xlsx"
DEFAULT_TEAM_MEMBERS   = r"C:\Users\Ajak\Desktop\Confluence Scrapper\IT_Team_Members.xlsx"
FILE_PREFIX_RAW  = "IT_Purchase_RAW"
FILE_PREFIX_SAP  = "IT_Purchase_SAP_Upload"

# ─────────────────────────────────────────────
# CLAUDE API
# ─────────────────────────────────────────────
CLAUDE_API_KEY   = "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"
CLAUDE_API_URL   = "https://api.anthropic.com/v1/messages"
CLAUDE_MODEL     = "claude-opus-4-5"



# ─────────────────────────────────────────────
# SNIPE-IT API
# ─────────────────────────────────────────────
SNIPEIT_BASE_URL = "http://100.68.20.67"
SNIPEIT_API_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxMyIsImp0aSI6ImVkOWFhMTY1YjFlY2Q5ZTM1NDFkNjI2OGIzODc4NjY2YWE5ODhlOGFkNTZjOTQwN2I0ODhlYmNmZDlhMzEyNzZmZDM5MjViNWUwY2Y1ZDU5IiwiaWF0IjoxNzczNzU4MTAwLjk0NTA3MSwibmJmIjoxNzczNzU4MTAwLjk0NTA3MiwiZXhwIjoyMjQ3MTQzNzAwLjkzOTgzMiwic3ViIjoiNjIzNyIsInNjb3BlcyI6W119.maV0TcqPK6t-a1S0uA2vdtoJY0wSuQabl6C7NAQ3bppzOSHYdsUyTRB54RYwp4t-v0Wr_VQIe1prqk9hU1qb9l3thmdRf0Wn5me0lD_4mFPakGT9X6vKT--TiLHayhF4XRSp2Dtkp11bcyFgP0XtgLebNbiLH-N47d7CzuoEQ0zW4W3-YE4GIYNKBXibJGyNdncRd8cLh2iEHSQTZ_-hh76vP2cOUaFIqFMflW3W1-KZIQbk7m7dWUZ1bxYSUYhME4mzDwylEu7O--563gWjkLprARCA9tmskwtT_leALKNXh55c_rQb427LPWIbu9AvXYEyed8WGHV1tmn4Mt3Ss-YGSMbdRWqIsLVZxMLvJHzvWdKcDqpZob7YIG4cD6x8OwC3RTGOHZlTfAVKKC3gZoA7G-cgGHf8YXi2X0sCJlQbJVmECKsKDQBibu7EbgKnPcgFq-aaa2OtiL98WneLMc1g6xUSyxrqJWOrBjPBfWdwE6umHIvTuPyJbnGuBVIyB5cHQpBWaJYsT_2UM4yI8uulvxOZVjeY_F5yp9YcLNSNzC7S5Hg8aRd-RqgJ6Plwv6peXN57GpEUNvak02OpjlZb4xxl33qG596-oHtDfpioooVuvZ71dhkQdyEDvXIZzm5d4sPWd5chWkwk0xjZRBiKZZYdeb55Uus9r83LHz8"

# ─────────────────────────────────────────────
# ASSET CLASS DEFINITIONS (hardcoded)
# ─────────────────────────────────────────────
ASSET_CLASSES = {
    "1010": "Office Equipment_IT",
    "1011": "Office Equipment_PC/Monitor",
    "1012": "Office Equipment_Server&Network",
    "1020": "Office Equipment_General",
    "2101": "Software_Others",
    "2900": "Assets under construction",
}

ASSET_CLASS_SYSTEM_PROMPT = """당신은 기업의 IT 자산 관리 및 회계 처리 전문가입니다. 제공된 Confluence page를 참조하여 IT 구매 내역을 분석하고, 정의된 자산 분류 체계(Asset Class)에 따라 가장 적합한 코드를 매칭하는 업무를 수행합니다.

[Asset Class Definitions] 각 항목을 분류할 때 아래의 기준을 엄격히 따르세요:
* 1011 (Office Equipment_PC/Monitor): 노트북, 데스크탑(워크스테이션), 1000유로 이상의 컴퓨터 주변 기기 등(예: MacBook, Dell Laptop, GeForce RTX 5070)
* 1012 (Office Equipment_Server&Network): 서버 장비, 스위치, 라우터, 방화벽, 스토리지, Rack 등 네트워크 인프라 장비.
* 1010 (Office Equipment_IT): PC/서버/네트워크를 제외한 IT 관련 기기.
* 1020 (Office Equipment_General): 사무용 가구, 가전, 비품. (예: 의자, 책상, 공기청정기, 커피머신)
* 2101 (Software_Others): 소프트웨어 라이선스, Perpetual License, 설치형 프로그램 구매 비용. (예: Perforce Perpetual License)
* 2900 (Assets under construction): 선급금 처리된 프로젝트, 아직 구축 중인 미완성 자산, 장기 프로젝트 착수금. Snipe IT Asset Tag가 빈칸이면 주로 이 Class로 설정해줘.

[Analysis Logic]
1. Description 우선: 품목 명칭에서 'PC', 'Laptop' 모델명이 발견되면 1011로 분류합니다. (예. Lenovo, Apple Mac, Wacom, 5080, 5070 등)
2. Server 관련 장비등의 모델명등이 있으면 1012로 구분해줘
3. Vendor 참고: Vendor가 'Cisco', 'Arista' 등 네트워크 전문 업체인 경우 1012일 확률이 높습니다.
4. URL 확인: Kissflow URL이나 비고란에 '구독', 'License'가 언급되면 2101로 분류합니다.
5. 금액 및 목적: 단가가 낮고 소모성인 주변기기는 자산이 아닌 비용으로 구분 합니다.

반드시 아래 6개 코드 중 하나만 숫자로 응답하세요. 설명 없이 코드 번호만:
1010, 1011, 1012, 1020, 2101, 2900"""

# ─────────────────────────────────────────────
# KRAFTON BRAND THEME
# ─────────────────────────────────────────────
BG       = "#000000"
PANEL    = "#111111"
PANEL2   = "#1a1a1a"
BORDER   = "#2e2e2e"
ACCENT   = "#f9423a"
ACCENT_H = "#ff5e57"
TEXT     = "#ffffff"
TEXT_DIM = "#777777"
LOG_BG   = "#0a0a0a"
SUCCESS  = "#4cdb72"
ERROR    = "#f9423a"
WARNING  = "#ffb347"

FONT_MONO  = ("Consolas", 9)
FONT_UI    = ("Segoe UI", 9)
FONT_LABEL = ("Segoe UI", 9)


# ═════════════════════════════════════════════
# SCRAPER
# ═════════════════════════════════════════════

def fetch_page_html(base_url, page_id, email, api_token):
    url = f"{base_url}/wiki/rest/api/content/{page_id}"
    resp = requests.get(
        url, params={"expand": "body.storage"},
        auth=(email, api_token), timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    return data["body"]["storage"]["value"], data["title"]


def parse_tables(html):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        return []
    dfs = []
    for table in tables:
        rows, headers = [], []
        header_row = table.find("tr")
        if header_row:
            headers = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])]
        for tr in table.find_all("tr")[1:]:
            cells = [td.get_text(separator=" ", strip=True) for td in tr.find_all(["td", "th"])]
            if any(cells):
                rows.append(cells)
        if rows:
            max_cols = max(len(r) for r in rows)
            if headers and len(headers) < max_cols:
                headers += [f"Column_{j}" for j in range(len(headers), max_cols)]
            elif not headers:
                headers = [f"Column_{j}" for j in range(max_cols)]
            dfs.append(pd.DataFrame(rows, columns=headers[:max_cols]))
    return dfs


def save_raw(dataframes, output_dir):
    today = datetime.now().strftime("%Y%m%d_%H%M")
    path = output_dir / f"{FILE_PREFIX_RAW}_{today}.xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for i, df in enumerate(dataframes):
            sheet = f"Table_{i+1}" if len(dataframes) > 1 else "Raw"
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for col_idx, col in enumerate(df.columns, 1):
                max_len = max(
                    df[col].astype(str).map(len).max() if not df.empty else 0,
                    len(str(col)),
                ) + 2
                ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = min(max_len, 50)
    return path


# ═════════════════════════════════════════════
# TEMPLATE-DRIVEN SAP TRANSFORM
# ═════════════════════════════════════════════

def read_template_headers(template_path):
    """Read first sheet header row from SAP template."""
    engine = "xlrd" if str(template_path).lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(template_path, sheet_name=0, header=0,
                       nrows=0, engine=engine, dtype=str)
    return [str(c).strip() for c in df.columns if str(c).strip() not in ("", "nan")]


def _col(src_row, *candidates):
    """Return first non-empty value matching any candidate column name (case-insensitive)."""
    lookup = {k.lower(): v for k, v in src_row.items()}
    for name in candidates:
        val = lookup.get(name.lower(), "")
        val = str(val).strip()
        if val and val not in ("nan", "None"):
            return val
    return ""


def _first_of_month(date_str):
    """Convert any date string to first-of-month in DD.MM.YYYY format."""
    if not date_str:
        return ""
    try:
        dt = pd.to_datetime(date_str, dayfirst=True)
        return dt.replace(day=1).strftime("%Y.%m.%d")
    except Exception:
        return date_str


# ─────────────────────────────────────────────
# AI HELPERS
# ─────────────────────────────────────────────

def _claude_call(system_prompt, user_msg, max_tokens=20):
    """Generic Claude API call. Returns response text."""
    hdrs = {
        "x-api-key": CLAUDE_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": CLAUDE_MODEL,
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_msg}],
    }
    resp = requests.post(CLAUDE_API_URL, json=payload, headers=hdrs, timeout=60)
    resp.raise_for_status()
    return resp.json()["content"][0]["text"].strip()


def classify_asset_class(row_data: dict) -> str:
    """Call Claude to determine Asset Class code (4-digit). Falls back to "2900"."""
    import re
    row_text = "\n".join(
        f"{k}: {v}" for k, v in row_data.items()
        if str(v).strip() not in ("", "nan", "None")
    )
    result = _claude_call(
        ASSET_CLASS_SYSTEM_PROMPT,
        f"아래 IT 구매 항목의 Asset Class 코드를 결정해주세요:\n\n{row_text}",
        max_tokens=10,
    )
    m = re.search(r"(1010|1011|1012|1020|2101|2900)", result)
    return m.group(1) if m else "2900"


# ─────────────────────────────────────────────
# SNIPE-IT HELPERS
# ─────────────────────────────────────────────

def _parse_asset_tag(raw_tag: str) -> str:
    """Strip surrounding underscores: "__PGAI000786__" → "PGAI000786"."""
    return raw_tag.strip().strip("_").strip()


def _fetch_snipeit_asset(asset_tag: str) -> dict:
    """
    GET /api/v1/hardware/bytag/{tag}
    Returns full asset dict, or {} on failure.
    """
    if not SNIPEIT_API_TOKEN or not asset_tag:
        return {}
    tag = _parse_asset_tag(asset_tag)
    if not tag:
        return {}
    url  = f"{SNIPEIT_BASE_URL.rstrip('/')}/api/v1/hardware/bytag/{tag}"
    hdrs = {
        "Authorization": f"Bearer {SNIPEIT_API_TOKEN}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.get(url, headers=hdrs, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except Exception:
        return {}


def fetch_snipeit_fields(asset_tag: str) -> dict:
    """
    Returns dict with:
      purchase_cost : 구매원가  (→ Acquisition Value)
      department    : 부서명    (→ Cost Center 판단)
    """
    data = _fetch_snipeit_asset(asset_tag)
    if not data:
        return {"purchase_cost": "", "department": ""}

    # purchase_cost
    cost = data.get("purchase_cost")
    cost = str(cost).strip() if cost is not None and str(cost).strip() not in ("", "None", "null") else ""

    # department — Snipe-IT returns location or department object
    dept = ""
    dept_obj = data.get("department") or data.get("location")
    if isinstance(dept_obj, dict):
        dept = str(dept_obj.get("name", "")).strip()
    elif isinstance(dept_obj, str):
        dept = dept_obj.strip()

    return {"purchase_cost": cost, "department": dept}


def _cost_center_from_dept(department: str) -> str:
    """140069 if "PBB" in department name, else 140006."""
    if "PBB" in str(department).upper():
        return "140069"
    return "140006"


# ─────────────────────────────────────────────
# VENDOR MASTER LOOKUP
# ─────────────────────────────────────────────

def load_vendor_master(vendor_master_path: str) -> pd.DataFrame:
    """
    Load SAP Vendor Master file.
    Expected columns (by position):
      Col B (idx 1) = Place of purchase code
      Col C (idx 2) = Place of purchase name
    Returns DataFrame with columns: code, name
    """
    path = str(vendor_master_path).strip()
    if not path or not Path(path).exists():
        return pd.DataFrame(columns=["code", "name"])
    engine = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
    try:
        df = pd.read_excel(path, sheet_name=0, header=0, engine=engine, dtype=str)
        df = df.fillna("")
        # Use positional columns B(1) and C(2) regardless of header name
        if df.shape[1] < 3:
            return pd.DataFrame(columns=["code", "name"])
        result = pd.DataFrame({
            "code": df.iloc[:, 1].str.strip(),   # Col B
            "name": df.iloc[:, 2].str.strip(),   # Col C
        })
        return result[result["code"] != ""]
    except Exception:
        return pd.DataFrame(columns=["code", "name"])


def load_team_members(team_members_path: str) -> pd.DataFrame:
    """
    Load IT_Team_Members.xlsx.
    Returns DataFrame with columns: code (E-prefixed), name, numeric_code (E stripped)
    Assumes: Col B = Employee code (e.g. E12345), Col C = Name (or similar)
    """
    path = str(team_members_path).strip()
    if not path or not Path(path).exists():
        return pd.DataFrame(columns=["code", "name", "numeric_code"])
    engine = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
    try:
        df = pd.read_excel(path, sheet_name=0, header=0, engine=engine, dtype=str)
        df = df.fillna("")
        if df.shape[1] < 3:
            return pd.DataFrame(columns=["code", "name", "numeric_code"])
        result = pd.DataFrame({
            "code": df.iloc[:, 1].str.strip(),   # Col B: E-prefixed code
            "name": df.iloc[:, 2].str.strip(),   # Col C: Name
        })
        result = result[result["code"] != ""]
        # Pre-compute numeric code (strip leading E)
        result["numeric_code"] = result["code"].str.lstrip("Ee").str.strip()
        return result
    except Exception:
        return pd.DataFrame(columns=["code", "name", "numeric_code"])


def _lookup_vendor_code(keyword: str, vendor_df: pd.DataFrame,
                        code_prefix: str = "") -> str:
    """
    Find the best matching Place of purchase code for a keyword.
    Matches keyword substring (case-insensitive) against name column.
    If code_prefix is set, only considers rows whose code starts with that prefix.

    Returns the matched code string, or "" if no match.
    """
    if vendor_df.empty or not keyword:
        return ""

    df = vendor_df.copy()
    if code_prefix:
        df = df[df["code"].str.upper().str.startswith(code_prefix.upper())]
    if df.empty:
        return ""

    kw = keyword.strip().lower()
    # 1. Exact match
    exact = df[df["name"].str.lower() == kw]
    if not exact.empty:
        return exact.iloc[0]["code"]

    # 2. Keyword contained in name
    contains = df[df["name"].str.lower().str.contains(kw, na=False)]
    if not contains.empty:
        return contains.iloc[0]["code"]

    # 3. Any word of keyword contained in name
    for word in kw.split():
        if len(word) < 3:
            continue
        partial = df[df["name"].str.lower().str.contains(word, na=False)]
        if not partial.empty:
            return partial.iloc[0]["code"]

    return ""


def _lookup_employee_code(keyword: str, team_df: pd.DataFrame) -> str:
    """
    Find employee numeric code (E stripped) by name keyword in IT_Team_Members.
    Returns numeric string (e.g. "12345" from "E12345"), or "" if not found.
    """
    if team_df.empty or not keyword:
        return ""
    kw = keyword.strip().lower()
    # 1. Exact name match
    exact = team_df[team_df["name"].str.lower() == kw]
    if not exact.empty:
        return exact.iloc[0]["numeric_code"]
    # 2. Keyword contained in name
    contains = team_df[team_df["name"].str.lower().str.contains(kw, na=False)]
    if not contains.empty:
        return contains.iloc[0]["numeric_code"]
    # 3. Any word of keyword in name
    for word in kw.split():
        if len(word) < 2:
            continue
        partial = team_df[team_df["name"].str.lower().str.contains(word, na=False)]
        if not partial.empty:
            return partial.iloc[0]["numeric_code"]
    return ""


def filter_by_period(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    """
    Filter dataframe rows where Purchase Date falls in the given year/month.
    Tries common column name variants (case-insensitive).
    Rows with unparseable dates are excluded.
    """
    date_col = None
    for col in df.columns:
        if col.strip().lower() in ("purchase date", "purchasedate", "date"):
            date_col = col
            break
    if date_col is None:
        return df  # no date column found — return all rows

    def _in_period(val):
        try:
            dt = pd.to_datetime(str(val).strip(), dayfirst=True)
            return dt.year == year and dt.month == month
        except Exception:
            return False

    mask = df[date_col].apply(_in_period)
    return df[mask].reset_index(drop=True)


def transform_to_sap(raw_df, template_headers, vendor_df=None, team_df=None, log_cb=None):
    """
    Build SAP upload DataFrame row-by-row.

    Asset Class      ← Claude AI (4-digit code)
    Asset Tag        ← Snipe-IT Asset Tag (Confluence column)
    Description      ← Item Description
    Acquisition Date ← Purchase Date → first of month (DD.MM.YYYY)
    Vendor           ← Vendor keyword → Vendor Master Col C→B (any code)
    Acquisition Value← Snipe-IT: 구매원가 (purchase_cost)
    Currency         ← EUR (default)
    Employee No.     ← Kissflow Raised By → IT_Team_Members name→code, strip E prefix
    Asset No.        ← (blank)
    Sub-No           ← (blank)
    Cost Center      ← Snipe-IT: 부서 → PBB=140069, else 140006
    """
    if vendor_df is None:
        vendor_df = pd.DataFrame(columns=["code", "name"])
    if team_df is None:
        team_df = pd.DataFrame(columns=["code", "name", "numeric_code"])
    rows  = []
    total = len(raw_df)

    for idx, (_, src) in enumerate(raw_df.iterrows()):
        row_dict = dict(src)
        desc = _col(src, "Item Description", "Description", "Item Name")
        if log_cb:
            log_cb(f"  [{idx+1}/{total}]  {desc[:60] or '(no description)'}", "DIM")

        # ── Asset Class (Claude AI) ───────────────────
        asset_class = classify_asset_class(row_dict)
        if log_cb:
            log_cb(f"         Asset Class  → {asset_class} ({ASSET_CLASSES.get(asset_class, '')})", "ACCENT")

        # ── Snipe-IT lookup: purchase_cost + department ──
        raw_tag    = _col(src, "Snipe-IT Asset Tag", "Asset Tag", "SnipeIT Asset Tag", "Snipe IT Asset Tag")
        tag_clean  = _parse_asset_tag(raw_tag) if raw_tag else ""
        snipe_data = fetch_snipeit_fields(raw_tag)
        acq_value  = snipe_data["purchase_cost"]
        department = snipe_data["department"]
        cost_center = _cost_center_from_dept(department)

        if log_cb:
            log_cb(
                f"         Snipe-IT [{tag_clean or 'no tag'}] "
                f"구매원가={acq_value or '(없음)'} | "
                f"부서={department or '(없음)'} → CC {cost_center}",
                "INFO",
            )

        # ── Build SAP row ─────────────────────────────
        sap_row = {}
        for th in template_headers:
            tl = th.lower()
            if tl == "asset class":
                sap_row[th] = asset_class
            elif tl == "asset tag":
                sap_row[th] = tag_clean
            elif tl == "description":
                sap_row[th] = desc
            elif tl == "acquisition date":
                raw_date = _col(src, "Purchase Date", "Acquisition Date", "Date")
                sap_row[th] = _first_of_month(raw_date)
            elif tl == "vendor":
                raw_vendor = _col(src, "Vendor", "Supplier", "Vendor Name")
                vendor_code = _lookup_vendor_code(raw_vendor, vendor_df, code_prefix="")
                sap_row[th] = vendor_code if vendor_code else raw_vendor
                if log_cb:
                    result_str = f"{vendor_code}" if vendor_code else f"(no match, kept: {raw_vendor})"
                    log_cb(f"         Vendor      → {result_str}", "INFO")
            elif tl == "acquisition value":
                sap_row[th] = acq_value
            elif tl == "currency":
                sap_row[th] = "EUR"
            elif tl in ("employee no.", "employee no", "employee number"):
                raw_emp = _col(src, "Kissflow Raised By", "Raised By", "Employee No.", "Employee No")
                emp_code = _lookup_employee_code(raw_emp, team_df)
                sap_row[th] = emp_code if emp_code else raw_emp
                if log_cb:
                    result_str = f"{emp_code}" if emp_code else f"(no match, kept: {raw_emp})"
                    log_cb(f"         Employee No → {result_str}", "INFO")
            elif tl == "asset no.":
                sap_row[th] = ""
            elif tl == "sub-no":
                sap_row[th] = ""
            elif tl == "cost center":
                sap_row[th] = cost_center
            else:
                sap_row[th] = ""
        rows.append(sap_row)

    sap_df = pd.DataFrame(rows, columns=template_headers)
    return sap_df

def save_sap(sap_df, output_dir):
    """Save SAP upload file as .xls using xlwt."""
    import xlwt
    today = datetime.now().strftime("%Y%m%d_%H%M")
    path = output_dir / f"{FILE_PREFIX_SAP}_{today}.xls"
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("SAP_Upload")

    # header style — orange bg, white bold centered
    header_style = xlwt.XFStyle()
    pat = xlwt.Pattern()
    pat.pattern = xlwt.Pattern.SOLID_PATTERN
    pat.pattern_fore_colour = 0x16   # orange in xlwt default palette
    header_style.pattern = pat
    fnt = xlwt.Font()
    fnt.bold = True
    fnt.colour_index = 0x01          # white
    header_style.font = fnt
    aln = xlwt.Alignment()
    aln.horz = xlwt.Alignment.HORZ_CENTER
    header_style.alignment = aln

    data_style = xlwt.XFStyle()

    # write header
    for col_idx, col_name in enumerate(sap_df.columns):
        ws.write(0, col_idx, col_name, header_style)
        ws.col(col_idx).width = 256 * 22

    # write data
    for row_idx, row in enumerate(sap_df.itertuples(index=False), start=1):
        for col_idx, val in enumerate(row):
            ws.write(row_idx, col_idx,
                     val if str(val) not in ("nan", "None") else "",
                     data_style)

    wb.save(str(path))
    return path


# ═════════════════════════════════════════════
# SESSION CACHE  (persist paths between runs)
# ═════════════════════════════════════════════

SESSION_KEYS = ["confluence_url", "output", "template", "vendor_master", "team_members"]

def load_session() -> dict:
    """Load last-used paths from JSON cache."""
    import json, os
    if os.path.exists(SESSION_CACHE_FILE):
        try:
            return json.loads(open(SESSION_CACHE_FILE, encoding="utf-8").read())
        except Exception:
            pass
    return {}


def save_session(data: dict):
    """Persist current paths to JSON cache."""
    import json
    try:
        open(SESSION_CACHE_FILE, "w", encoding="utf-8").write(
            json.dumps(data, ensure_ascii=False, indent=2)
        )
    except Exception:
        pass


def _extract_page_id(confluence_url: str) -> str:
    """
    Extract numeric page ID from a Confluence URL.
    e.g. .../pages/864695451/IT+Purchase+... → "864695451"
    Falls back to DEFAULT_PAGE_ID if not found.
    """
    import re
    m = re.search(r"/pages/(\d+)", confluence_url)
    return m.group(1) if m else DEFAULT_PAGE_ID


# ═════════════════════════════════════════════
# GUI
# ═════════════════════════════════════════════

class ConfluenceScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Confluence IT Purchase Scraper  v{VERSION}")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(720, 580)
        self._running = False
        self._session = load_session()   # last-used paths
        self._build_ui()
        self._center_window(780, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Save current paths to session cache before exit."""
        save_session({
            "confluence_url":  self._var_confluence_url.get().strip(),
            "output":          self._var_output.get().strip(),
            "template":        self._var_template.get().strip(),
            "vendor_master":   self._var_vendor_master.get().strip(),
            "team_members":    self._var_team_members.get().strip(),
        })
        self.destroy()

    def _center_window(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ════════════════════════════════════════
    # LAYOUT
    # ════════════════════════════════════════
    def _build_ui(self):
        self._build_header()
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")
        self._build_body()
        self._build_statusbar()

    # ── header ──────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x")
        left = tk.Frame(hdr, bg=BG)
        left.pack(side="left", padx=18, pady=13)
        cnv = tk.Canvas(left, bg=BG, highlightthickness=0, width=28, height=28)
        cnv.pack(side="left")
        cnv.create_rectangle(0, 0, 28, 28, fill=ACCENT, outline="")
        cnv.create_text(14, 14, text="K", fill=TEXT, font=("Segoe UI Black", 14))
        tk.Label(
            left, text="  Confluence IT Purchase Scraper",
            bg=BG, fg=TEXT, font=("Segoe UI Black", 12),
        ).pack(side="left")
        badge = tk.Frame(hdr, bg=ACCENT)
        badge.pack(side="right", padx=18, pady=16)
        tk.Label(badge, text=f" v{VERSION} ", bg=ACCENT, fg=TEXT,
                 font=("Segoe UI Semibold", 8)).pack()

    # ── body ────────────────────────────────
    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)
        self._build_confluence_card(body)
        self._build_input_card(body)
        self._build_output_card(body)
        self._build_run_row(body)
        self._build_log_card(body)

    # ── card factory ────────────────────────
    def _make_card(self, parent, title):
        wrapper = tk.Frame(parent, bg=PANEL,
                           highlightthickness=1, highlightbackground=BORDER)
        wrapper.pack(fill="x", pady=(0, 10))
        bar = tk.Frame(wrapper, bg=PANEL2, height=28)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Frame(bar, bg=ACCENT, width=3).pack(side="left", fill="y")
        tk.Label(bar, text=f"  {title}", bg=PANEL2, fg=TEXT,
                 font=("Segoe UI Semibold", 8)).pack(side="left", pady=6)
        inner = tk.Frame(wrapper, bg=PANEL, padx=14, pady=10)
        inner.pack(fill="x")
        return wrapper, inner

    def _path_row(self, parent, label, default, browse_cmd):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=PANEL, fg=TEXT_DIM,
                 font=FONT_LABEL, width=14, anchor="w").pack(side="left")
        var = tk.StringVar(value=default)
        tk.Entry(
            row, textvariable=var,
            bg=PANEL2, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_UI,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        ).pack(side="left", fill="x", expand=True, ipady=5, padx=(6, 6))
        tk.Button(
            row, text="Browse", command=browse_cmd,
            bg=BORDER, fg=TEXT,
            activebackground=ACCENT, activeforeground=TEXT,
            relief="flat", font=FONT_LABEL,
            padx=10, pady=4, cursor="hand2",
        ).pack(side="left")
        return var

    # ── confluence card ──────────────────────
    def _build_confluence_card(self, parent):
        from datetime import datetime as _dt
        _, inner = self._make_card(parent, "CONFLUENCE")
        default_url = self._session.get("confluence_url", DEFAULT_CONFLUENCE_URL)

        # ── Page URL row ──
        url_row = tk.Frame(inner, bg=PANEL)
        url_row.pack(fill="x", pady=3)
        tk.Label(url_row, text="Page URL", bg=PANEL, fg=TEXT_DIM,
                 font=FONT_LABEL, width=14, anchor="w").pack(side="left")
        self._var_confluence_url = tk.StringVar(value=default_url)
        tk.Entry(
            url_row, textvariable=self._var_confluence_url,
            bg=PANEL2, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_UI,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        ).pack(side="left", fill="x", expand=True, ipady=5, padx=(6, 0))

        # ── Period filter row ──
        now = _dt.now()
        period_row = tk.Frame(inner, bg=PANEL)
        period_row.pack(fill="x", pady=(6, 2))
        tk.Label(period_row, text="Period Filter", bg=PANEL, fg=TEXT_DIM,
                 font=FONT_LABEL, width=14, anchor="w").pack(side="left")

        years  = [str(y) for y in range(2024, 2100)]
        months = [f"{m:02d}" for m in range(1, 13)]
        MONTH_NAMES = ["01 - Jan","02 - Feb","03 - Mar","04 - Apr",
                       "05 - May","06 - Jun","07 - Jul","08 - Aug",
                       "09 - Sep","10 - Oct","11 - Nov","12 - Dec"]

        combo_style = dict(
            bg=PANEL2, fg=TEXT,
            activebackground=ACCENT, activeforeground=TEXT,
            selectcolor=PANEL2,
            relief="flat", font=FONT_UI,
            highlightthickness=1,
            highlightbackground=BORDER,
        )

        self._var_year  = tk.StringVar(value=str(now.year))
        self._var_month = tk.StringVar(value=f"{now.month:02d} - {_dt(now.year, now.month, 1).strftime('%b')}")

        # Year OptionMenu
        year_menu = tk.OptionMenu(period_row, self._var_year, *years)
        year_menu.config(bg=PANEL2, fg=TEXT, activebackground=ACCENT,
                         activeforeground=TEXT, relief="flat",
                         font=FONT_UI, highlightthickness=0,
                         indicatoron=True, width=6)
        year_menu["menu"].config(bg=PANEL2, fg=TEXT,
                                 activebackground=ACCENT, activeforeground=TEXT,
                                 relief="flat", font=FONT_UI)
        year_menu.pack(side="left", padx=(6, 4))

        tk.Label(period_row, text="년", bg=PANEL, fg=TEXT_DIM,
                 font=FONT_LABEL).pack(side="left")

        # Month OptionMenu
        month_menu = tk.OptionMenu(period_row, self._var_month, *MONTH_NAMES)
        month_menu.config(bg=PANEL2, fg=TEXT, activebackground=ACCENT,
                          activeforeground=TEXT, relief="flat",
                          font=FONT_UI, highlightthickness=0,
                          indicatoron=True, width=10)
        month_menu["menu"].config(bg=PANEL2, fg=TEXT,
                                  activebackground=ACCENT, activeforeground=TEXT,
                                  relief="flat", font=FONT_UI)
        month_menu.pack(side="left", padx=(4, 4))

        tk.Label(period_row, text="월  기준으로 필터링",
                 bg=PANEL, fg=TEXT_DIM, font=FONT_LABEL).pack(side="left")

    # ── input card ──────────────────────────
    def _build_input_card(self, parent):
        _, inner = self._make_card(parent, "INPUT")
        self._var_template = self._path_row(
            inner, "SAP Template",
            self._session.get("template", DEFAULT_TEMPLATE),
            self._browse_template,
        )
        self._var_vendor_master = self._path_row(
            inner, "Vendor Master",
            self._session.get("vendor_master", DEFAULT_VENDOR_MASTER),
            self._browse_vendor_master,
        )
        self._var_team_members = self._path_row(
            inner, "Team Members",
            self._session.get("team_members", DEFAULT_TEAM_MEMBERS),
            self._browse_team_members,
        )

    # ── output card ─────────────────────────
    def _build_output_card(self, parent):
        _, inner = self._make_card(parent, "OUTPUT")
        self._var_output = self._path_row(
            inner, "Output Folder",
            self._session.get("output", DEFAULT_OUTPUT),
            self._browse_folder,
        )

    def _browse_folder(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self._var_output.set(d)

    def _browse_template(self):
        f = filedialog.askopenfilename(
            title="Select SAP Template",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if f:
            self._var_template.set(f)

    def _browse_vendor_master(self):
        f = filedialog.askopenfilename(
            title="Select Vendor Master",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if f:
            self._var_vendor_master.set(f)

    def _browse_team_members(self):
        f = filedialog.askopenfilename(
            title="Select IT Team Members",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if f:
            self._var_team_members.set(f)

    # ── run row ─────────────────────────────
    def _build_run_row(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(4, 10))
        self._btn_run = tk.Button(
            row, text="▶  RUN",
            command=self._on_run,
            bg=ACCENT, fg=TEXT,
            activebackground=ACCENT_H, activeforeground=TEXT,
            relief="flat", font=("Segoe UI Black", 10),
            padx=24, pady=8, cursor="hand2",
        )
        self._btn_run.pack(side="left")
        self._btn_clear = tk.Button(
            row, text="Clear Log",
            command=self._clear_log,
            bg=PANEL, fg=TEXT_DIM,
            activebackground=BORDER, activeforeground=TEXT,
            relief="flat", font=FONT_LABEL,
            padx=14, pady=8, cursor="hand2",
        )
        self._btn_clear.pack(side="left", padx=(8, 0))
        style = ttk.Style()
        style.theme_use("default")
        style.configure("KR.Horizontal.TProgressbar",
                        troughcolor=PANEL, background=ACCENT, thickness=5)
        self._progress = ttk.Progressbar(
            row, mode="indeterminate", length=200,
            style="KR.Horizontal.TProgressbar",
        )
        self._progress.pack(side="right")

    # ── log card ────────────────────────────
    def _build_log_card(self, parent):
        wrapper = tk.Frame(parent, bg=PANEL,
                           highlightthickness=1, highlightbackground=BORDER)
        wrapper.pack(fill="both", expand=True)
        bar = tk.Frame(wrapper, bg=PANEL2, height=28)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Frame(bar, bg=ACCENT, width=3).pack(side="left", fill="y")
        tk.Label(bar, text="  LOG", bg=PANEL2, fg=TEXT,
                 font=("Segoe UI Semibold", 8)).pack(side="left", pady=6)
        self._log_box = scrolledtext.ScrolledText(
            wrapper,
            bg=LOG_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_MONO,
            state="disabled", wrap="word",
            padx=10, pady=8,
        )
        self._log_box.pack(fill="both", expand=True)
        self._log_box.tag_config("INFO",    foreground=TEXT)
        self._log_box.tag_config("SUCCESS", foreground=SUCCESS)
        self._log_box.tag_config("ERROR",   foreground=ERROR)
        self._log_box.tag_config("WARNING", foreground=WARNING)
        self._log_box.tag_config("DIM",     foreground=TEXT_DIM)
        self._log_box.tag_config("ACCENT",  foreground=ACCENT)

    # ── status bar ──────────────────────────
    def _build_statusbar(self):
        bar = tk.Frame(self, bg=PANEL2, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        dot = tk.Canvas(bar, bg=PANEL2, highlightthickness=0, width=12, height=12)
        dot.pack(side="left", padx=(12, 4), pady=7)
        self._dot_canvas = dot
        self._dot_id = dot.create_oval(1, 1, 11, 11, fill=TEXT_DIM, outline="")
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(bar, textvariable=self._status_var,
                 bg=PANEL2, fg=TEXT_DIM,
                 font=("Segoe UI", 8)).pack(side="left")

    def _set_status(self, msg, color=None):
        c = color or TEXT_DIM
        def _do():
            self._status_var.set(msg)
            self._dot_canvas.itemconfig(self._dot_id, fill=c)
        self.after(0, _do)

    # ════════════════════════════════════════
    # ACTIONS
    # ════════════════════════════════════════
    def _on_run(self):
        if self._running:
            return
        tmpl = Path(self._var_template.get().strip())
        if not tmpl.exists():
            messagebox.showwarning(
                "Template 없음",
                f"SAP Template 파일을 찾을 수 없습니다:\n{tmpl}",
            )
            return
        self._running = True
        self._btn_run.config(state="disabled", bg=BORDER, fg=TEXT_DIM)
        self._progress.start(10)
        self._set_status("Running...", ACCENT)
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        try:
            self._log("━" * 56, "DIM")
            self._log(
                f"  Pipeline START  [{datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}]",
                "ACCENT",
            )
            self._log("━" * 56, "DIM")

            out_dir  = Path(self._var_output.get().strip())
            tmpl_path = Path(self._var_template.get().strip())

            # ── Step 1: Scrape ────────────────────────
            conf_url  = self._var_confluence_url.get().strip()
            page_id   = _extract_page_id(conf_url)
            base_url  = DEFAULT_BASE_URL
            self._log(f"[ 1 / 3 ]  Connecting to Confluence (page {page_id})...", "DIM")
            html, title = fetch_page_html(
                base_url, page_id,
                DEFAULT_EMAIL, DEFAULT_API_TOKEN,
            )
            self._log(f"  ✔  Page fetched: {title}", "SUCCESS")

            self._log("[ 1 / 3 ]  Parsing tables...", "DIM")
            dataframes = parse_tables(html)
            if not dataframes:
                self._log("  ✘  No tables found!", "ERROR")
                self._set_status("Failed — no tables found", ERROR)
                return
            for i, df in enumerate(dataframes):
                self._log(f"  Table {i+1}: {len(df)} rows × {len(df.columns)} cols", "INFO")

            # ── Step 2: Save raw ──────────────────────
            self._log("[ 2 / 3 ]  Saving raw Excel...", "DIM")
            raw_path = save_raw(dataframes, out_dir)
            self._log(f"  ✔  Raw  → {raw_path.name}", "SUCCESS")

            # ── Step 3: Read template + vendor master ──
            self._log(f"[ 3 / 4 ]  Reading template: {tmpl_path.name}...", "DIM")
            template_headers = read_template_headers(tmpl_path)
            self._log(f"  ✔  Template headers ({len(template_headers)}): "
                      f"{', '.join(template_headers)}", "INFO")

            vm_path = self._var_vendor_master.get().strip()
            vendor_df = load_vendor_master(vm_path)
            if vendor_df.empty:
                self._log("  ⚠  Vendor Master not loaded (경로 확인 필요)", "WARNING")
            else:
                self._log(f"  ✔  Vendor Master loaded: {len(vendor_df)} entries", "SUCCESS")

            tm_path = self._var_team_members.get().strip()
            team_df = load_team_members(tm_path)
            if team_df.empty:
                self._log("  ⚠  IT Team Members not loaded (경로 확인 필요)", "WARNING")
            else:
                self._log(f"  ✔  Team Members loaded: {len(team_df)} entries", "SUCCESS")

            # ── Period filter ─────────────────────────
            sel_year  = int(self._var_year.get().strip())
            sel_month = int(self._var_month.get().strip().split(" ")[0])
            self._log(
                f"  Filtering by period: {sel_year}.{sel_month:02d}",
                "INFO",
            )
            raw_df    = dataframes[0]
            source_df = filter_by_period(raw_df, sel_year, sel_month)
            self._log(
                f"  ✔  {len(source_df)} / {len(raw_df)} row(s) match {sel_year}.{sel_month:02d}",
                "SUCCESS" if len(source_df) > 0 else "WARNING",
            )
            if len(source_df) == 0:
                self._log("  ⚠  No matching rows — check Purchase Date column or period setting.", "WARNING")
                self._set_status("No rows matched the period filter", WARNING)
                return

            self._log(f"  Confluence columns: {', '.join(source_df.columns)}", "DIM")

            # ── Step 4: AI classify + transform ───────
            self._log(f"[ 4 / 4 ]  Processing {len(source_df)} row(s) with AI...", "DIM")
            self._log(f"  Asset Classes: {', '.join(f'{k}' for k in ASSET_CLASSES.keys())}", "DIM")
            sap_df = transform_to_sap(
                source_df, template_headers, vendor_df=vendor_df, team_df=team_df, log_cb=self._log
            )
            self._log(f"  ✔  SAP rows: {len(sap_df)}", "SUCCESS")

            sap_path = save_sap(sap_df, out_dir)
            self._log(f"  ✔  SAP  → {sap_path.name}", "SUCCESS")

            self._log("━" * 56, "DIM")
            self._log(f"  Done!  2 files saved to: {out_dir}", "SUCCESS")
            self._log("━" * 56, "DIM")
            self._set_status(f"Done  —  {sap_path.name}", SUCCESS)

        except requests.exceptions.HTTPError as e:
            code = e.response.status_code
            self._log(f"✘  HTTP {code}: {e.response.text[:300]}", "ERROR")
            self._set_status(f"HTTP Error {code}", ERROR)
        except Exception as e:
            import traceback
            self._log(f"✘  {type(e).__name__}: {e}", "ERROR")
            self._log(traceback.format_exc(), "DIM")
            self._set_status("Error — check log", ERROR)
        finally:
            self._running = False
            self.after(0, self._reset_ui)

    def _reset_ui(self):
        self._progress.stop()
        self._btn_run.config(state="normal", bg=ACCENT, fg=TEXT)

    def _log(self, msg, tag="INFO"):
        def _insert():
            self._log_box.config(state="normal")
            self._log_box.insert("end", msg + "\n", tag)
            self._log_box.see("end")
            self._log_box.config(state="disabled")
        self.after(0, _insert)

    def _clear_log(self):
        self._log_box.config(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.config(state="disabled")
        self._set_status("Ready", TEXT_DIM)


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = ConfluenceScraperApp()
    app.mainloop()