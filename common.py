"""
shared_iq/common.py — Shared utilities for AccountIQ and ChurnIQ
"""
import re
import json
import sqlite3
from datetime import datetime, date
from pathlib import Path

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
        "%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y",
        "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
        "%d-%b-%Y", "%d-%B-%Y", "%d %b %Y", "%d %B %Y",
        "%b-%d-%Y", "%B-%d-%Y", "%Y%m%d",
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

# ── Column alias map ──────────────────────────────────────────────────────────
COLUMN_ALIASES = {
    "Account":               ["account name","company","customer","organization","account_name"],
    "Salesforce ID":         ["salesforce id","sfdc id","sf id","account id","salesforce account id",
                              "sfid","crm id","account_id","salesforceid"],
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
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+(inc|llc|ltd|corp|co|plc|gmbh|ag|sa|nv|bv|pty|limited)$", "", s)
    return re.sub(r"\s+", " ", s).strip()

def _norm_sfid(s):
    """Normalize Salesforce ID — strip whitespace, uppercase."""
    return str(s).strip().upper() if s and str(s).strip().lower() not in ("nan","none","") else ""

# ── ARR formatter ─────────────────────────────────────────────────────────────
def _fmt_arr(arr):
    try:
        arr = float(arr or 0)
        if arr >= 1_000_000: return f"${arr/1_000_000:.1f}M"
        if arr >= 1_000:     return f"${arr/1_000:.0f}K"
        return f"${arr:.0f}"
    except: return "$0"


# ══════════════════════════════════════════════════════════════════════════════
# SOURCE MODEL — multi-source join engine
# ══════════════════════════════════════════════════════════════════════════════

SOURCE_TYPES = ["primary", "supplement", "override"]

# Fields that identify an account across sources
JOIN_KEY_FIELDS = ["Salesforce ID", "Account"]

def _row_sfid(row):
    """Extract normalised Salesforce ID from a row dict."""
    for f in ("Salesforce ID", "Salesforce Account ID", "Account ID", "SFDC ID", "SF ID"):
        v = _norm_sfid(row.get(f, ""))
        if v: return v
    return ""

def _row_name_key(row):
    """Extract normalised account name key from a row dict."""
    for f in ("Account", "Account Name", "Company", "Customer"):
        v = str(row.get(f, "")).strip()
        if v and v.lower() not in ("nan","none",""):
            return _norm_key(v)
    return ""


def join_sources(sources: list) -> dict:
    """
    Merge multiple data sources into a unified dict of account rows.

    sources: list of dicts, each:
        {
          "id": str,
          "name": str,
          "type": "primary" | "supplement" | "override",
          "priority": int,          # lower = higher priority
          "rows": list[dict],       # already alias-mapped rows
        }

    Returns:
        {
          "accounts": {account_key: merged_row},   # key = sfid or norm_name
          "report": {
            "total": int,
            "matched_by_id": int,
            "matched_by_name": int,
            "new_accounts": int,
            "unmatched_rows": list,   # rows from supplements that didn't join
            "source_contributions": {source_id: {matched, new, unmatched}},
          }
        }
    """
    # Sort: primary first, then by priority asc
    type_order = {"primary": 0, "supplement": 1, "override": 2}
    sorted_sources = sorted(sources, key=lambda s: (type_order.get(s["type"], 1), s.get("priority", 99)))

    unified = {}          # key -> merged row
    sfid_index = {}       # sfid -> key (for fast lookup)
    name_index = {}       # norm_name -> key

    report = {
        "total": 0,
        "matched_by_id": 0,
        "matched_by_name": 0,
        "new_accounts": 0,
        "unmatched_rows": [],
        "source_contributions": {},
    }

    for src in sorted_sources:
        src_id = src["id"]
        src_type = src["type"]
        contrib = {"matched_by_id": 0, "matched_by_name": 0, "new": 0, "unmatched": 0}

        for row in src.get("rows", []):
            sfid = _row_sfid(row)
            name_k = _row_name_key(row)
            if not name_k and not sfid:
                continue  # no key at all — skip

            # Find existing record
            existing_key = None
            match_type = None

            if sfid and sfid in sfid_index:
                existing_key = sfid_index[sfid]
                match_type = "id"
            elif name_k and name_k in name_index:
                existing_key = name_index[name_k]
                match_type = "name"

            if existing_key is not None:
                # Merge into existing record
                base = unified[existing_key]
                if src_type == "override":
                    # Override: replace all non-empty fields
                    for k, v in row.items():
                        sv = str(v).strip() if v is not None else ""
                        if sv and sv.lower() not in ("nan", "none", ""):
                            base[k] = v
                            base[f"_src_{k}"] = src_id
                elif src_type == "supplement":
                    # Supplement: fill missing fields only
                    for k, v in row.items():
                        if k.startswith("_"): continue
                        sv = str(v).strip() if v is not None else ""
                        existing = str(base.get(k, "")).strip()
                        existing_empty = not existing or existing.lower() in ("nan","none","0","")
                        if existing_empty and sv and sv.lower() not in ("nan","none",""):
                            base[k] = v
                            base[f"_src_{k}"] = src_id
                else:
                    # Primary: update all non-empty fields (refresh)
                    for k, v in row.items():
                        sv = str(v).strip() if v is not None else ""
                        if sv and sv.lower() not in ("nan","none",""):
                            base[k] = v
                            base[f"_src_{k}"] = src_id

                # Update sfid index if we now have an ID
                if sfid and sfid not in sfid_index:
                    sfid_index[sfid] = existing_key

                if match_type == "id":
                    contrib["matched_by_id"] += 1
                    report["matched_by_id"] += 1
                else:
                    contrib["matched_by_name"] += 1
                    report["matched_by_name"] += 1

            else:
                # New account
                if src_type == "supplement" and unified:
                    # Supplement with no primary match — flag as unmatched
                    contrib["unmatched"] += 1
                    report["unmatched_rows"].append({
                        "source": src["name"],
                        "account": row.get("Account", row.get("Account Name", "?")),
                        "sfid": sfid or "",
                    })
                    continue

                # Add as new account
                canonical_key = sfid if sfid else name_k
                row["_primary_source"] = src_id
                unified[canonical_key] = dict(row)
                # tag all fields with source
                for k in list(row.keys()):
                    if not k.startswith("_"):
                        row[f"_src_{k}"] = src_id

                if sfid: sfid_index[sfid] = canonical_key
                if name_k: name_index[name_k] = canonical_key

                contrib["new"] += 1
                report["new_accounts"] += 1

        report["source_contributions"][src_id] = contrib

    report["total"] = len(unified)
    return {"accounts": unified, "report": report}


# ══════════════════════════════════════════════════════════════════════════════
# SOURCE DB HELPERS — SQLite CRUD for data_sources + source_rows
# ══════════════════════════════════════════════════════════════════════════════

def sources_db_init(db_file: Path):
    """Create data_sources and source_rows tables if not exist."""
    con = sqlite3.connect(str(db_file))
    con.execute("""
        CREATE TABLE IF NOT EXISTS data_sources (
            id          TEXT PRIMARY KEY,
            name        TEXT NOT NULL,
            type        TEXT NOT NULL DEFAULT 'supplement',
            priority    INTEGER NOT NULL DEFAULT 10,
            col_map     TEXT DEFAULT '{}',
            last_uploaded TEXT,
            row_count   INTEGER DEFAULT 0,
            file_name   TEXT,
            uploaded_by TEXT
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS source_rows (
            source_id   TEXT NOT NULL,
            row_idx     INTEGER NOT NULL,
            account_key TEXT,
            sfid        TEXT,
            rows_json   TEXT,
            PRIMARY KEY (source_id, row_idx)
        )
    """)
    con.execute("CREATE INDEX IF NOT EXISTS idx_source_rows_sid ON source_rows(source_id)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_source_rows_sfid ON source_rows(sfid)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_source_rows_key ON source_rows(account_key)")
    con.commit()
    con.close()


def sources_save(db_file: Path, source: dict):
    """Upsert a data source record (no rows)."""
    con = sqlite3.connect(str(db_file))
    con.execute("""
        INSERT INTO data_sources (id, name, type, priority, col_map, last_uploaded, row_count, file_name, uploaded_by)
        VALUES (:id, :name, :type, :priority, :col_map, :last_uploaded, :row_count, :file_name, :uploaded_by)
        ON CONFLICT(id) DO UPDATE SET
            name=excluded.name, type=excluded.type, priority=excluded.priority,
            col_map=excluded.col_map, last_uploaded=excluded.last_uploaded,
            row_count=excluded.row_count, file_name=excluded.file_name,
            uploaded_by=excluded.uploaded_by
    """, {
        "id": source["id"],
        "name": source["name"],
        "type": source.get("type", "supplement"),
        "priority": source.get("priority", 10),
        "col_map": json.dumps(source.get("col_map", {})),
        "last_uploaded": source.get("last_uploaded", ""),
        "row_count": source.get("row_count", 0),
        "file_name": source.get("file_name", ""),
        "uploaded_by": source.get("uploaded_by", ""),
    })
    con.commit()
    con.close()


def sources_save_rows(db_file: Path, source_id: str, rows: list):
    """Replace all rows for a source."""
    con = sqlite3.connect(str(db_file))
    con.execute("DELETE FROM source_rows WHERE source_id=?", (source_id,))
    for i, row in enumerate(rows):
        sfid = _row_sfid(row)
        name_k = _row_name_key(row)
        con.execute(
            "INSERT INTO source_rows (source_id, row_idx, account_key, sfid, rows_json) VALUES (?,?,?,?,?)",
            (source_id, i, name_k, sfid, json.dumps(row))
        )
    con.commit()
    con.close()


def sources_load_all(db_file: Path) -> list:
    """Load all data sources with their rows."""
    if not db_file.exists():
        return []
    con = sqlite3.connect(str(db_file))
    con.row_factory = sqlite3.Row
    srcs = [dict(r) for r in con.execute("SELECT * FROM data_sources ORDER BY priority ASC, name ASC").fetchall()]
    for src in srcs:
        src["col_map"] = json.loads(src.get("col_map") or "{}")
        rows_raw = con.execute(
            "SELECT rows_json FROM source_rows WHERE source_id=? ORDER BY row_idx ASC",
            (src["id"],)
        ).fetchall()
        src["rows"] = [json.loads(r["rows_json"]) for r in rows_raw]
    con.close()
    return srcs


def sources_load_meta(db_file: Path) -> list:
    """Load source metadata only (no rows) — fast."""
    if not db_file.exists():
        return []
    con = sqlite3.connect(str(db_file))
    con.row_factory = sqlite3.Row
    srcs = [dict(r) for r in con.execute("SELECT * FROM data_sources ORDER BY priority ASC, name ASC").fetchall()]
    for src in srcs:
        src["col_map"] = json.loads(src.get("col_map") or "{}")
    con.close()
    return srcs


def sources_delete(db_file: Path, source_id: str):
    """Delete a source and all its rows."""
    con = sqlite3.connect(str(db_file))
    con.execute("DELETE FROM data_sources WHERE id=?", (source_id,))
    con.execute("DELETE FROM source_rows WHERE source_id=?", (source_id,))
    con.commit()
    con.close()


def sources_update_priority(db_file: Path, source_id: str, priority: int):
    con = sqlite3.connect(str(db_file))
    con.execute("UPDATE data_sources SET priority=? WHERE id=?", (priority, source_id))
    con.commit()
    con.close()


def sources_update_col_map(db_file: Path, source_id: str, col_map: dict):
    con = sqlite3.connect(str(db_file))
    con.execute("UPDATE data_sources SET col_map=? WHERE id=?", (json.dumps(col_map), source_id))
    con.commit()
    con.close()


def sources_get_join_report(db_file: Path) -> dict:
    """Return a quick join health summary across all sources."""
    if not db_file.exists():
        return {}
    con = sqlite3.connect(str(db_file))
    con.row_factory = sqlite3.Row
    srcs = [dict(r) for r in con.execute("SELECT id, name, type, priority, row_count, last_uploaded FROM data_sources ORDER BY priority").fetchall()]
    total_rows = sum(s["row_count"] for s in srcs)
    # Count rows with sfid
    with_id = con.execute("SELECT COUNT(*) FROM source_rows WHERE sfid != ''").fetchone()[0]
    con.close()
    return {
        "sources": srcs,
        "total_source_rows": total_rows,
        "rows_with_sfid": with_id,
        "rows_name_only": total_rows - with_id,
    }
