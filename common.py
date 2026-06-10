"""
shared_iq/common.py — Shared utilities for AccountIQ and ChurnIQ
"""
import re
from datetime import datetime, date

# ── Numeric/string helpers ────────────────────────────────────────────────────
def _num(row, col):
    try:
        v = row.get(col, 0)
        import pandas as pd
        if pd.isna(v): return 0
        return float(str(v).replace(",","").replace("%","").strip() or 0)
    except: return 0

def _str(row, col):
    v = str(row.get(col,"")).strip()
    return "" if v.lower() in ("nan","none","") else v

# ── Date parser ───────────────────────────────────────────────────────────────
_EXCEL_EPOCH = datetime(1899, 12, 30).date()

def _parse_date_robust(val):
    """Parse a date value from virtually any format into a Python date."""
    if val is None: return None
    if isinstance(val, date) and not isinstance(val, datetime): return val
    if isinstance(val, datetime): return val.date()
    try:
        import pandas as _pd
        if isinstance(val, _pd.Timestamp): return val.date()
    except Exception: pass
    try:
        n = float(val)
        if 20000 < n < 80000:
            from datetime import timedelta
            return _EXCEL_EPOCH + timedelta(days=int(n))
    except (TypeError, ValueError): pass
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'n/a', '-', ''): return None
    s = re.sub(r'T\d{2}:\d{2}.*$', '', s)
    s = re.sub(r'\s+\d{2}:\d{2}.*$', '', s)
    s = s.strip()
    for fmt in [
        "%Y-%m-%d", "%Y/%m/%d",
        "%m/%d/%Y", "%m-%d-%Y",
        "%m/%d/%y", "%m-%d-%y",
        "%d/%m/%Y", "%d-%m-%Y",
        "%d/%m/%y", "%d-%m-%y",
        "%b %d, %Y", "%B %d, %Y",
        "%b %d %Y",  "%B %d %Y",
        "%d-%b-%Y",  "%d-%B-%Y",
        "%d %b %Y",  "%d %B %Y",
        "%b-%d-%Y",  "%B-%d-%Y",
        "%Y%m%d",
    ]:
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    try:
        import pandas as _pd
        return _pd.to_datetime(s, dayfirst=False, yearfirst=False).date()
    except Exception: pass
    return None

# ── File type detection ───────────────────────────────────────────────────────
def _detect_filetype(content: bytes, fname: str) -> str:
    """Detect file type by magic bytes first, then filename extension."""
    if content[:4] == b'\xd0\xcf\x11\xe0': return 'xls'
    if content[:4] == b'PK\x03\x04':       return 'xlsx'
    fname_lower = (fname or '').lower()
    if fname_lower.endswith('.xlsx'): return 'xlsx'
    if fname_lower.endswith('.xls'):  return 'xls'
    if fname_lower.endswith('.csv'):  return 'csv'
    try:
        content[:1024].decode('utf-8')
        return 'csv'
    except Exception:
        return 'xlsx'

# ── Column alias map (AccountIQ's richer version) ────────────────────────────
COLUMN_ALIASES = {
    "Account":               ["account name","company","customer","organization"],
    "Account Owner":         ["ae","ae name","account executive","owner","sales rep"],
    "SE Name":               ["systems engineer","solutions engineer","se"],
    "TSM Name":              ["tsm","territory sales manager","sales manager"],
    "Industry":              ["vertical","sector"],
    "ARR":                   ["annual recurring revenue","arr $","revenue","contract value","acv"],
    "HQ Renewal":            ["renewal","renewal date","contract end","end date","expiry"],
    "Upcoming Renewal Risk": ["renewal risk","risk flag"],
    "# of Meetings L90D":    ["meetings","total meetings","meetings l90","mtgs l90d","# of meetings l90d"],
    "AE Meeting L90D":       ["ae meeting l90d","ae meetings l90"],
    "SE Meeting L90D":       ["se meeting l90d","se meetings l90"],
    "CSM Meeting L90D":      ["csm meeting l90d"],
    "# Exec/CTM Meeting L90D": ["exec meeting l90d","# exec/ctm meeting l90d"],
    "# CISO CTO CXO L90D":   ["ciso cto cxo l90d","# ciso cto cxo l90d"],
    "Partner Meeting L90D":  ["partner meeting l90d"],
    "TA Meeting L90D":       ["ta meeting l90d"],
    "#Sales/SE Leader L90D": ["sales/se leader l90d","#sales/se leader l90d"],
    "Average Resilience Score": ["resilience","resilience score"],
    "Customer Sentiment":    ["sentiment","csat sentiment"],
    "NPS":                   ["net promoter score","nps score"],
    "Score%":                ["score%","health score","arr customer score%"],
    "Employees":             ["employees","# employees","headcount"],
    "G2K Rank":              ["g2k","g2k rank","g2000"],
    "Prospect Score":        ["prospect score"],
    "MCI L180D":             ["mci l180d","mci"],
    "MQL L180D":             ["mql l180d","mql"],
    "PAI Engmnt Lvl":        ["pai engmnt lvl","pai engmnt level","pai level","pai engagement level"],
    "Relationship Map (Yes/No)": ["relationship map","relationship map (yes/no)"],
    "#DR APPROVED L90D":     ["dr approved l90d","#dr approved l90d"],
    "DR Submitted L90D":     ["dr submitted l90d"],
    "SDR Completed DMs L6Ms":["sdr completed dms l6ms","sdr dms"],
    "$Open Pipeline Amount L90D": ["open pipeline amount l90d","$open pipeline amount l90d"],
    "#Open Pipeline L90D":   ["open pipeline l90d","#open pipeline l90d"],
    "Next 90D Pipeline Amount": ["next 90d pipeline amount","next 90d pipeline","pipeline amount"],
    "# N+U Open Pipeline CFY": ["# n+u open pipeline cfy","n+u pipeline count"],
    "$ N+U Open Pipeline CFY": ["$ n+u open pipeline cfy","n+u open pipeline","$ n+u open pipeline - all"],
    "Partner":               ["channel partner","reseller","var"],
    "Incumbent":             ["competitor","legacy vendor","displacement"],
}

# ── Key normalizer ────────────────────────────────────────────────────────────
def _norm_key(s):
    """Normalize account name for dedup matching."""
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+(inc|llc|ltd|corp|co|plc|gmbh|ag|sa|nv|bv|pty|limited)$", "", s)
    return re.sub(r"\s+", " ", s).strip()

# ── ARR formatter ─────────────────────────────────────────────────────────────
def _fmt_arr(arr):
    """Format ARR as $1.2M / $450K / $5K."""
    try:
        arr = float(arr or 0)
        if arr >= 1_000_000: return f"${arr/1_000_000:.1f}M"
        if arr >= 1_000:     return f"${arr/1_000:.0f}K"
        return f"${arr:.0f}"
    except: return "$0"
