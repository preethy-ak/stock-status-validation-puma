"""
PUMA Marketplace Stock & Status Validation Automation
Enterprise-grade Streamlit Application
Version 3.2 — Auto-dependency installer
"""

# ─────────────────────────────────────────────────────────────
# AUTO-INSTALL MISSING DEPENDENCIES
# Runs silently before anything else loads
# ─────────────────────────────────────────────────────────────
import subprocess, sys, importlib

def _ensure(package: str, import_name: str = None):
    """Install package if not already available."""
    name = import_name or package
    try:
        importlib.import_module(name)
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", package, "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

_ensure("openpyxl")
_ensure("plotly")
_ensure("xlsxwriter")
_ensure("chardet")          # robust CSV encoding detection

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import warnings
from typing import Dict, List, Optional, Tuple, Any
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PUMA Marketplace Validator",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
REGIONS      = ["SG", "MY", "PH"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

REGION_MP_MAP = {
    "SG": ["Lazada", "Shopee", "Zalora"],
    "MY": ["Lazada", "Shopee", "Zalora", "TikTok"],
    "PH": ["Lazada", "Shopee", "Zalora", "TikTok"],
}

TIKTOK_VALIDATION_ONLY = {"PH"}          # no action files for TikTok PH

# Product Master Stock columns by region
PM_STOCK_COL = {
    "SG": "MyStock-YCH-SG quantity",
    "MY": "MyStock-YCH-MY quantity",
    "PH": "MyStock-PH quantity",
}
PM_RESERVED_COL = {
    "SG": "MyStock-YCH-SG reservedQuantity",
    "MY": "MyStock-YCH-MY reservedQuantity",
    "PH": "MyStock-PH reservedQuantity",
}

# MP SKU column names → normalised to EAN
MP_SKU_COLS = {
    "Lazada":  ["SellerSKU", "Seller SKU", "seller_sku", "sellsku"],
    "Shopee":  ["SKU", "sku"],
    "Zalora":  ["SellerSku", "Seller Sku", "seller_sku", "SellerSKU"],
    "TikTok":  ["Seller sku", "Seller SKU", "seller_sku", "SellerSku"],
}

# Status column aliases
STATUS_COLS  = ["Status", "status", "Product Status", "Listing Status", "item_status"]
STOCK_COLS   = ["Stock", "stock", "Quantity", "quantity", "Available", "available",
                "Available Stock", "Price", "price"]

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "results": None,
        "run_done": False,
        "logs": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────
def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{ts}] {msg}")

def _is_csv(f) -> bool:
    """Detect CSV from filename extension."""
    name = getattr(f, "name", "") or ""
    return name.lower().endswith(".csv")

def _detect_encoding(raw: bytes) -> str:
    """Detect encoding using chardet if available, else return utf-8."""
    try:
        import chardet
        result = chardet.detect(raw[:10000])
        return result.get("encoding") or "utf-8"
    except ImportError:
        return "utf-8"

def read_excel(f, sheet_name=0) -> Optional[pd.DataFrame]:
    """
    Read an uploaded Excel OR CSV file safely.
    Explicitly uses openpyxl engine to avoid ImportError.
    Tries multiple encodings for CSV files.
    """
    if f is None:
        return None
    try:
        if _is_csv(f):
            f.seek(0)
            raw = f.read()
            detected = _detect_encoding(raw)
            encodings = list(dict.fromkeys(
                [detected, "utf-8", "utf-8-sig", "latin-1", "cp1252"]
            ))
            for enc in encodings:
                try:
                    text = raw.decode(enc)
                    df = pd.read_csv(
                        io.StringIO(text), dtype=str,
                        low_memory=False, on_bad_lines="skip",
                    )
                    df.columns = df.columns.str.strip()
                    df = df.dropna(how="all")
                    log(f"📄 CSV ({enc}): {getattr(f,'name','')} → {len(df)} rows")
                    return df
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue
            log(f"⚠️  Could not decode CSV: {getattr(f,'name','')}")
            return None
        else:
            f.seek(0)
            df = pd.read_excel(f, sheet_name=sheet_name, dtype=str,
                               engine="openpyxl")
            df.columns = df.columns.str.strip()
            df = df.dropna(how="all")
            return df
    except Exception as e:
        log(f"⚠️  Could not read {getattr(f,'name','file')}: {e}")
        return None

def read_all_sheets(f) -> Dict[str, pd.DataFrame]:
    """
    Read all sheets from an Excel file (openpyxl explicit).
    If file is CSV, returns single-sheet dict keyed by filename stem.
    """
    if f is None:
        return {}
    try:
        if _is_csv(f):
            df = read_excel(f)
            if df is not None:
                stem = getattr(f, "name", "Sheet1").rsplit(".", 1)[0]
                return {stem: df}
            return {}
        else:
            f.seek(0)
            xl = pd.ExcelFile(f, engine="openpyxl")
            sheets = {}
            for name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=name, dtype=str,
                                   engine="openpyxl")
                df.columns = df.columns.str.strip()
                df = df.dropna(how="all")
                sheets[name] = df
            return sheets
    except Exception as e:
        log(f"⚠️  Could not read sheets from {getattr(f,'name','file')}: {e}")
        return {}

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Find first matching column (case-insensitive)."""
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in low:
            return low[c.lower()]
    return None

def safe_float(v, default=0.0) -> float:
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return default

def norm_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()

def norm_ean(v) -> str:
    s = norm_str(v)
    # Try to strip .0 from numeric EANs
    if s.endswith(".0"):
        s = s[:-2]
    return s

def norm_status(v) -> str:
    s = norm_str(v).upper()
    if s in ("YES", "ACTIVE", "1", "TRUE", "LIVE"):
        return "ACTIVE"
    if s in ("NO", "OFF", "INACTIVE", "0", "FALSE", "DELISTED", "DEACTIVATED"):
        return "INACTIVE"
    return s

# ─────────────────────────────────────────────────────────────
# FILE LOADERS
# ─────────────────────────────────────────────────────────────

def load_content_file(f) -> Dict[str, List[str]]:
    """Returns {article_no: [ean1, ean2, ...]}"""
    df = read_excel(f)
    if df is None:
        return {}
    art_col = find_col(df, ["Color_No", "PIM Article#", "Article No", "ArticleNo",
                             "Article Number", "color_no", "article_no"])
    ean_col = find_col(df, ["EAN", "ean", "EAN Code"])
    if not art_col or not ean_col:
        log(f"❌ Content file missing Article/EAN columns. Found: {list(df.columns)}")
        return {}
    result: Dict[str, List[str]] = {}
    for _, row in df.iterrows():
        art = norm_str(row[art_col])
        ean = norm_ean(row[ean_col])
        if art and ean:
            result.setdefault(art, [])
            if ean not in result[art]:
                result[art].append(ean)
    log(f"✅ Content file: {len(result)} articles, "
        f"{sum(len(v) for v in result.values())} EANs")
    return result

def load_zecom(f) -> Dict[str, Dict[str, Dict]]:
    """
    Returns {region: {article_no: {mp: status, 'launch_date': date}}}
    ZeCom has sub-sheets per region AND marketplace columns.
    """
    sheets = read_all_sheets(f)
    if not sheets:
        return {}

    result: Dict[str, Dict[str, Dict]] = {}

    # Try to match sheet names to regions
    region_sheets: Dict[str, str] = {}
    for sname in sheets:
        su = sname.upper().strip()
        for r in REGIONS:
            if r in su:
                region_sheets[r] = sname
                break

    # If no region sheets found, use first sheet for all regions
    if not region_sheets:
        first = list(sheets.keys())[0]
        for r in REGIONS:
            region_sheets[r] = first

    for region, sname in region_sheets.items():
        df = sheets[sname]
        art_col = find_col(df, ["Color_No", "PIM Article#", "Article No",
                                 "ArticleNo", "Article Number", "color_no"])
        if not art_col:
            log(f"⚠️  ZeCom sheet '{sname}': no Article column found")
            continue

        # Find launch date column
        launch_col = find_col(df, ["Launch Date", "Go Live Date", "launch_date",
                                    "LaunchDate", "Go-Live Date", "Live Date"])

        # Find marketplace columns
        mp_cols: Dict[str, str] = {}
        for mp in MARKETPLACES:
            col = find_col(df, [mp, mp.lower()])
            if col:
                mp_cols[mp] = col

        region_data: Dict[str, Dict] = {}
        for _, row in df.iterrows():
            art = norm_str(row[art_col])
            if not art:
                continue
            entry: Dict[str, Any] = {}
            # Launch date
            ld = None
            if launch_col:
                try:
                    raw = row[launch_col]
                    if pd.notna(raw) and str(raw).strip():
                        ld = pd.to_datetime(raw, dayfirst=False, errors="coerce")
                        if pd.isna(ld):
                            ld = None
                        else:
                            ld = ld.date()
                except:
                    ld = None
            entry["launch_date"] = ld
            # MP statuses
            for mp, col in mp_cols.items():
                entry[mp] = norm_status(row.get(col, ""))
            region_data[art] = entry

        result[region] = region_data
        log(f"✅ ZeCom '{sname}' → {region}: {len(region_data)} articles")

    return result

def load_product_master(f) -> Dict[str, Dict[str, float]]:
    """
    Returns {ean: {region: quantity}}
    e.g. {ean: {"SG": 10, "MY": 5, "PH": 0}}
    """
    df = read_excel(f)
    if df is None:
        return {}
    ean_col = find_col(df, ["EAN", "ean", "EAN Code", "Ean"])
    if not ean_col:
        log(f"❌ Product Master: no EAN column. Found: {list(df.columns)}")
        return {}
    result: Dict[str, Dict[str, float]] = {}
    for _, row in df.iterrows():
        ean = norm_ean(row[ean_col])
        if not ean:
            continue
        stocks: Dict[str, float] = {}
        for region, col in PM_STOCK_COL.items():
            if col in df.columns:
                stocks[region] = safe_float(row.get(col, 0))
            else:
                # Try partial match
                matched = next((c for c in df.columns if col.lower() in c.lower()), None)
                stocks[region] = safe_float(row.get(matched, 0)) if matched else 0.0
        result[ean] = stocks
    log(f"✅ Product Master: {len(result)} EANs")
    return result

def load_regional_inventory(f, region: str) -> Dict[str, float]:
    """Returns {ean: stock}"""
    df = read_excel(f)
    if df is None:
        return {}
    ean_col = find_col(df, ["EAN", "ean", "EAN Code", "Ean", "SellerSKU",
                              "Seller SKU", "SKU", "sku"])
    stock_col = find_col(df, ["Stock", "stock", "Quantity", "quantity",
                                "Available", "Inventory", "inventory",
                                "Available Stock", "B2C Stock"])
    if not ean_col or not stock_col:
        log(f"⚠️  Inventory {region}: cannot find EAN/Stock columns. "
            f"Found: {list(df.columns)}")
        return {}
    result: Dict[str, float] = {}
    for _, row in df.iterrows():
        ean = norm_ean(row[ean_col])
        if not ean:
            continue
        result[ean] = safe_float(row.get(stock_col, 0))
    log(f"✅ Inventory {region}: {len(result)} EANs")
    return result

def load_marketplace_file(f, marketplace: str) -> Dict[str, Dict]:
    """Returns {ean: {status, stock}}"""
    df = read_excel(f)
    if df is None:
        return {}
    # Find EAN column
    candidates = MP_SKU_COLS.get(marketplace, [])
    ean_col = find_col(df, candidates + ["EAN", "ean"])
    if not ean_col:
        log(f"⚠️  {marketplace} file: no SKU column. Found: {list(df.columns)}")
        return {}
    status_col = find_col(df, STATUS_COLS)
    stock_col  = find_col(df, STOCK_COLS)
    result: Dict[str, Dict] = {}
    for _, row in df.iterrows():
        ean = norm_ean(row[ean_col])
        if not ean:
            continue
        result[ean] = {
            "status": norm_status(row.get(status_col, "")) if status_col else "UNKNOWN",
            "stock":  safe_float(row.get(stock_col, 0)) if stock_col else 0.0,
        }
    log(f"✅ {marketplace} file: {len(result)} EANs")
    return result

def load_zalora_combined(status_f, stock_f) -> Dict[str, Dict]:
    """Merge Zalora status + stock files."""
    status_df = read_excel(status_f)
    stock_df  = read_excel(stock_f)
    result: Dict[str, Dict] = {}

    def _extract(df, label):
        if df is None:
            return {}
        ean_col = find_col(df, ["SellerSku", "SellerSKU", "Seller Sku",
                                  "seller_sku", "SKU", "EAN"])
        if not ean_col:
            log(f"⚠️  Zalora {label}: no SKU column")
            return {}
        return df.set_index(ean_col.strip()) if ean_col else {}

    s_df = status_df.copy() if status_df is not None else pd.DataFrame()
    k_df = stock_df.copy()  if stock_df  is not None else pd.DataFrame()

    ean_col_s = find_col(s_df, ["SellerSku","SellerSKU","Seller Sku","seller_sku","SKU","EAN"]) if not s_df.empty else None
    ean_col_k = find_col(k_df, ["SellerSku","SellerSKU","Seller Sku","seller_sku","SKU","EAN"]) if not k_df.empty else None

    status_map: Dict[str, str] = {}
    stock_map:  Dict[str, float] = {}

    if ean_col_s and not s_df.empty:
        status_c = find_col(s_df, STATUS_COLS)
        for _, row in s_df.iterrows():
            ean = norm_ean(row[ean_col_s])
            if ean:
                status_map[ean] = norm_status(row.get(status_c, "")) if status_c else "UNKNOWN"

    if ean_col_k and not k_df.empty:
        stock_c = find_col(k_df, STOCK_COLS)
        for _, row in k_df.iterrows():
            ean = norm_ean(row[ean_col_k])
            if ean:
                stock_map[ean] = safe_float(row.get(stock_c, 0)) if stock_c else 0.0

    all_eans = set(status_map) | set(stock_map)
    for ean in all_eans:
        result[ean] = {
            "status": status_map.get(ean, "UNKNOWN"),
            "stock":  stock_map.get(ean, 0.0),
        }
    log(f"✅ Zalora combined: {len(result)} EANs")
    return result

def load_special_override(f) -> Dict[str, str]:
    """Returns {article_no: ACTIVE/INACTIVE}"""
    df = read_excel(f)
    if df is None:
        return {}
    art_col    = find_col(df, ["Article No","ArticleNo","article_no","Color_No",
                                 "PIM Article#","Article Number"])
    status_col = find_col(df, ["Final Status","Status","final_status","status"])
    if not art_col or not status_col:
        log(f"❌ Override file missing columns. Found: {list(df.columns)}")
        return {}
    result: Dict[str, str] = {}
    for _, row in df.iterrows():
        art = norm_str(row[art_col])
        st  = norm_status(row.get(status_col, ""))
        if art and st:
            result[art] = st
    log(f"✅ Special Override: {len(result)} articles")
    return result

def load_exclusion_file(f) -> List[Dict]:
    """Returns list of exclusion records."""
    df = read_excel(f)
    if df is None:
        return []
    records = []
    for _, row in df.iterrows():
        mp       = norm_str(row.get(find_col(df, ["Marketplace","marketplace"]) or "", ""))
        art      = norm_str(row.get(find_col(df, ["Article No","ArticleNo","article_no",
                                                     "Color_No"]) or "", ""))
        col_name = norm_str(row.get(find_col(df, ["Collection Name","Collection",
                                                     "collection_name"]) or "", ""))
        reason   = norm_str(row.get(find_col(df, ["Reason","reason"]) or "", ""))
        comments = norm_str(row.get(find_col(df, ["Comments","comments","Comment"]) or "", ""))
        if art:
            records.append({
                "marketplace": mp.upper(),
                "article_no":  art,
                "collection":  col_name,
                "reason":      reason,
                "comments":    comments,
            })
    log(f"✅ Exclusion file: {len(records)} records")
    return records

def load_hero_sku(f) -> List[Dict]:
    """Returns list of hero SKU records {article_no, marketplace, region}"""
    df = read_excel(f)
    if df is None:
        return []
    art_col = find_col(df, ["Article No","ArticleNo","article_no","Color_No",
                              "PIM Article#"])
    mp_col  = find_col(df, ["Marketplace","marketplace","MP"])
    reg_col = find_col(df, ["Region","region"])
    if not art_col:
        log(f"❌ Hero SKU file missing Article column. Found: {list(df.columns)}")
        return []
    records = []
    for _, row in df.iterrows():
        art = norm_str(row[art_col])
        mp  = norm_str(row.get(mp_col, "")) if mp_col else ""
        reg = norm_str(row.get(reg_col, "")).upper() if reg_col else ""
        if art:
            records.append({"article_no": art,
                             "marketplace": mp.upper(),
                             "region": reg})
    log(f"✅ Hero SKU file: {len(records)} records")
    return records

# ─────────────────────────────────────────────────────────────
# VALIDATION ENGINE
# ─────────────────────────────────────────────────────────────

class ValidationEngine:
    def __init__(
        self,
        content:       Dict[str, List[str]],
        zecom:         Dict[str, Dict[str, Dict]],
        product_master: Dict[str, Dict[str, float]],
        inventories:   Dict[str, Dict[str, float]],
        mp_data:       Dict[str, Dict[str, Dict[str, Dict]]],
        overrides:     Dict[str, str],
        exclusions:    List[Dict],
        heroes:        List[Dict],
    ):
        self.content        = content          # {art: [ean]}
        self.zecom          = zecom            # {region: {art: {mp: status, launch_date}}}
        self.pm             = product_master   # {ean: {region: qty}}
        self.inv            = inventories      # {region: {ean: qty}}
        self.mp_data        = mp_data          # {region: {mp: {ean: {status, stock}}}}
        self.overrides      = overrides        # {art: ACTIVE/INACTIVE}
        self.exclusion_set  = self._build_excl_set(exclusions)
        self.hero_set       = self._build_hero_set(heroes)

        # reverse map EAN → article
        self.ean_to_art: Dict[str, str] = {}
        for art, eans in content.items():
            for ean in eans:
                self.ean_to_art[ean] = art

    def _build_excl_set(self, exclusions: List[Dict]) -> set:
        """Set of (article_no, marketplace) tuples that are excluded."""
        s = set()
        for rec in exclusions:
            art = rec["article_no"]
            mp  = rec["marketplace"]   # could be "" = all MPs
            if mp:
                s.add((art, mp))
            else:
                for m in MARKETPLACES:
                    s.add((art, m.upper()))
        return s

    def _build_hero_set(self, heroes: List[Dict]) -> set:
        """Set of (article_no, marketplace, region) tuples."""
        s = set()
        for rec in heroes:
            art = rec["article_no"]
            mp  = rec["marketplace"].upper()
            reg = rec["region"].upper()
            if mp and reg:
                s.add((art, mp, reg))
            elif mp:
                for r in REGIONS:
                    s.add((art, mp, r))
            elif reg:
                for m in MARKETPLACES:
                    s.add((art, m.upper(), reg))
            else:
                for r in REGIONS:
                    for m in MARKETPLACES:
                        s.add((art, m.upper(), r))
        return s

    def _is_hero(self, art: str, mp: str, region: str) -> bool:
        return (art, mp.upper(), region.upper()) in self.hero_set

    def _is_excluded(self, art: str, mp: str) -> bool:
        return (art, mp.upper()) in self.exclusion_set

    def _launch_passed(self, launch_date) -> bool:
        if launch_date is None:
            return True
        try:
            return date.today() >= launch_date
        except:
            return True

    def _pm_stock(self, ean: str, region: str) -> float:
        return self.pm.get(ean, {}).get(region, 0.0)

    def _inv_stock(self, ean: str, region: str) -> float:
        return self.inv.get(region, {}).get(ean, 0.0)

    def _apply_buffer(self, stock: float, region: str, mp: str) -> float:
        """PH Lazada only: stock - 1."""
        if region == "PH" and mp == "Lazada":
            return max(0.0, stock - 1)
        return stock

    def _expected_status(
        self, art: str, ean: str, mp: str, region: str,
        zecom_entry: Optional[Dict]
    ) -> Tuple[str, str]:
        """
        Returns (expected_status, comment).
        Priority: Override → Exclusion → Hero → ZeCom → No info
        """
        # 1. Special Override
        if art in self.overrides:
            return self.overrides[art], "Special override applied"

        # 2. Exclusion
        if self._is_excluded(art, mp):
            return "INACTIVE", "Excluded SKU"

        # 3. Hero SKU — keep ACTIVE regardless of stock
        if self._is_hero(art, mp, region):
            return "ACTIVE", f"Hero SKU - due to {mp} Hero"

        # 4. ZeCom
        if zecom_entry is None:
            return "INACTIVE", "Not in ZeCom Tracker"

        zecom_status = zecom_entry.get(mp, "")
        launch_date  = zecom_entry.get("launch_date")

        if zecom_status == "ACTIVE":
            if not self._launch_passed(launch_date):
                return "INACTIVE", f"Future launch date ({launch_date})"
            pm_stock  = self._pm_stock(ean, region)
            inv_stock = self._inv_stock(ean, region)
            buf_stock = self._apply_buffer(pm_stock, region, mp)
            if buf_stock <= 0 and inv_stock <= 0:
                return "INACTIVE", "No inventory available"
            return "ACTIVE", "ZeCom ACTIVE, inventory available"
        elif zecom_status == "INACTIVE":
            return "INACTIVE", "ZeCom INACTIVE"
        else:
            return "INACTIVE", f"ZeCom status unknown ({zecom_status})"

    # ── STATUS VALIDATION ────────────────────────────────────

    def run_status_validation(
        self,
        region: str,
        mp: str,
        progress_cb=None,
    ) -> List[Dict]:
        rows = []
        zecom_region = self.zecom.get(region, {})
        mp_eans      = self.mp_data.get(region, {}).get(mp, {})

        # All articles to check = union of ZeCom + Content
        all_articles = set(zecom_region.keys()) | set(self.content.keys())
        total        = len(all_articles)

        for i, art in enumerate(sorted(all_articles)):
            if progress_cb and i % 50 == 0:
                progress_cb(i / max(total, 1))

            zecom_entry = zecom_region.get(art)
            expected_mp_status = zecom_entry.get(mp, "") if zecom_entry else ""

            eans = self.content.get(art, [])
            if not eans:
                # Article in ZeCom but not in Content → flag
                rows.append({
                    "Region": region, "Marketplace": mp,
                    "Article No": art, "EAN": "",
                    "MP Status": "NOT LISTED",
                    "Expected Status": "INACTIVE" if expected_mp_status != "ACTIVE" else "ACTIVE",
                    "Status Check": "MISSING ARTICLE",
                    "Action Required": "Check Content File",
                    "Comments": "Article in ZeCom but no EANs in Content File",
                })
                continue

            for ean in eans:
                mp_info      = mp_eans.get(ean, {})
                mp_status    = mp_info.get("status", "NOT LISTED") if mp_info else "NOT LISTED"
                exp_status, comment = self._expected_status(
                    art, ean, mp, region, zecom_entry
                )

                if mp_status == "NOT LISTED":
                    if exp_status == "ACTIVE":
                        check  = "MISSING - SHOULD BE ACTIVE"
                        action = "Import SKU / Create Listing"
                    else:
                        check  = "NOT LISTED (OK)"
                        action = "No Action"
                elif mp_status == exp_status:
                    check  = "✅ MATCH"
                    action = "No Action"
                elif mp_status == "ACTIVE" and exp_status == "INACTIVE":
                    check  = "❌ WRONGLY ACTIVE"
                    action = "Inactivate"
                elif mp_status == "INACTIVE" and exp_status == "ACTIVE":
                    check  = "❌ WRONGLY INACTIVE"
                    action = "Activate"
                else:
                    check  = f"MISMATCH ({mp_status} vs {exp_status})"
                    action = "Review"

                rows.append({
                    "Region": region, "Marketplace": mp,
                    "Article No": art, "EAN": ean,
                    "MP Status": mp_status,
                    "Expected Status": exp_status,
                    "Status Check": check,
                    "Action Required": action,
                    "Comments": comment,
                })

        return rows

    # ── STOCK VALIDATION ─────────────────────────────────────

    def run_stock_validation(
        self,
        region: str,
        mp: str,
        progress_cb=None,
    ) -> List[Dict]:
        rows = []
        zecom_region = self.zecom.get(region, {})
        mp_eans      = self.mp_data.get(region, {}).get(mp, {})
        all_articles = set(zecom_region.keys()) | set(self.content.keys())
        total        = len(all_articles)

        for i, art in enumerate(sorted(all_articles)):
            if progress_cb and i % 50 == 0:
                progress_cb(i / max(total, 1))

            zecom_entry  = zecom_region.get(art)
            exp_status, comment = ("INACTIVE", "No ZeCom entry")
            eans = self.content.get(art, [])

            for ean in eans:
                exp_status, comment = self._expected_status(
                    art, ean, mp, region, zecom_entry
                )
                mp_info    = mp_eans.get(ean, {})
                mp_stock   = mp_info.get("stock", 0.0) if mp_info else 0.0
                pm_stock   = self._pm_stock(ean, region)
                inv_stock  = self._inv_stock(ean, region)
                buf_stock  = self._apply_buffer(pm_stock, region, mp)

                # Corrected stock = PM quantity (available stock)
                corrected  = buf_stock
                # If inactive → max = 0
                if exp_status == "INACTIVE":
                    corrected = 0.0

                diff = corrected - mp_stock

                if exp_status == "INACTIVE":
                    if mp_stock > 0:
                        check  = "❌ SHOULD BE 0"
                        action = "Set Max 0"
                        remark = "Status inactive — set stock to 0"
                    else:
                        check  = "✅ OK (0)"
                        action = "No Action"
                        remark = "Inactive, stock already 0"
                else:
                    if abs(diff) < 0.01:
                        check  = "✅ MATCH"
                        action = "No Action"
                        remark = "Stock in sync"
                    elif diff > 0:
                        check  = "⬆️ INCREASE NEEDED"
                        action = "Update Stock"
                        remark = f"Increase by {diff:.0f} units"
                    else:
                        check  = "⬇️ REDUCE NEEDED"
                        action = "Update Stock"
                        remark = f"Reduce by {abs(diff):.0f} units"

                    if mp_stock == 0 and pm_stock > 0:
                        remark += " | MP stock outdated"
                    if pm_stock == 0 and inv_stock > 0:
                        remark += " | Product Master missing, check inventory"

                rows.append({
                    "Region": region, "Marketplace": mp,
                    "Article No": art, "EAN": ean,
                    "MP Stock": mp_stock,
                    "TC Stock (PM)": pm_stock,
                    "Regional Inventory": inv_stock,
                    "Buffer Applied": buf_stock,
                    "Corrected Stock": corrected,
                    "Difference": round(diff, 2),
                    "Stock Check": check,
                    "Action Required": action,
                    "Remarks": remark,
                })

        return rows

    # ── TO BE LISTED ─────────────────────────────────────────

    def run_to_be_listed(
        self,
        region: str,
        mp: str,
    ) -> List[Dict]:
        rows = []
        zecom_region = self.zecom.get(region, {})
        mp_eans      = self.mp_data.get(region, {}).get(mp, {})

        for art, eans in self.content.items():
            zecom_entry = zecom_region.get(art)
            if not zecom_entry:
                continue
            zecom_status = zecom_entry.get(mp, "")
            if zecom_status != "ACTIVE":
                continue

            for ean in eans:
                pm_stock  = self._pm_stock(ean, region)
                inv_stock = self._inv_stock(ean, region)
                has_stock = (pm_stock > 0 or inv_stock > 0)
                on_mp     = ean in mp_eans

                # Trigger: ZeCom=YES + has stock + NOT on marketplace
                if has_stock and not on_mp:
                    rows.append({
                        "Region": region, "Marketplace": mp,
                        "Article No": art, "EAN": ean,
                        "PM Stock": pm_stock,
                        "Regional Inventory": inv_stock,
                        "Zecom Status": zecom_status,
                        "Missing Reason": "EAN not listed on marketplace",
                        "Action Required": "Import SKU / Create Listing",
                    })
        return rows

    # ── VARIANT GAPS ─────────────────────────────────────────

    def run_variant_gaps(
        self,
        region: str,
        mp: str,
    ) -> List[Dict]:
        rows = []
        mp_eans      = self.mp_data.get(region, {}).get(mp, {})

        for art, eans in self.content.items():
            expected   = len(eans)
            live       = [e for e in eans if e in mp_eans]
            missing    = [e for e in eans if e not in mp_eans]
            live_count = len(live)
            miss_count = len(missing)

            if miss_count == 0:
                continue

            if live_count == 0:
                comment = "Entire article missing from marketplace"
                gap_type = "ARTICLE MISSING"
            else:
                comment = f"{miss_count} of {expected} variants missing"
                gap_type = "PARTIAL LISTING"

            rows.append({
                "Region": region, "Marketplace": mp,
                "Article No": art,
                "Expected Variants": expected,
                "Live Variants": live_count,
                "Missing Variants": miss_count,
                "Missing EAN List": ", ".join(missing),
                "Gap Type": gap_type,
                "Comments": comment,
            })

        return rows

    # ── MASTER RUN ───────────────────────────────────────────

    def run_all(
        self,
        regions_filter: List[str],
        mp_filter: List[str],
        val_types: List[str],
        progress_bar=None,
    ) -> Dict:
        results = {
            "status":      [],
            "stock":       [],
            "to_be_listed": [],
            "variant_gaps": [],
        }

        tasks = [
            (r, m)
            for r in regions_filter
            for m in REGION_MP_MAP.get(r, [])
            if m in mp_filter
        ]
        total = len(tasks)

        for idx, (region, mp) in enumerate(tasks):
            pct = idx / max(total, 1)
            if progress_bar:
                progress_bar.progress(pct, text=f"Processing {region} – {mp}…")

            if "Status" in val_types:
                rows = self.run_status_validation(region, mp)
                for r in rows:
                    r["_region"] = region
                    r["_mp"]     = mp
                results["status"].extend(rows)

            if "Stock" in val_types:
                rows = self.run_stock_validation(region, mp)
                for r in rows:
                    r["_region"] = region
                    r["_mp"]     = mp
                results["stock"].extend(rows)

            rows = self.run_to_be_listed(region, mp)
            for r in rows:
                r["_region"] = region
                r["_mp"]     = mp
            results["to_be_listed"].extend(rows)

            rows = self.run_variant_gaps(region, mp)
            for r in rows:
                r["_region"] = region
                r["_mp"]     = mp
            results["variant_gaps"].extend(rows)

        if progress_bar:
            progress_bar.progress(1.0, text="✅ Validation complete!")

        return results


# ─────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────

def _write_sheet(writer, df: pd.DataFrame, sheet_name: str):
    """Write df to sheet, truncating name to 31 chars."""
    if df.empty:
        return
    name = sheet_name[:31]
    df.to_excel(writer, sheet_name=name, index=False)
    # Basic formatting
    ws = writer.sheets[name]
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col if cell.value is not None),
            default=10
        )
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

def build_excel(results: Dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        def to_df(lst, drop_internal=True):
            df = pd.DataFrame(lst)
            if drop_internal:
                df = df.drop(columns=[c for c in df.columns if c.startswith("_")],
                             errors="ignore")
            return df

        # ── SUMMARY ──────────────────────────────────────────
        summary_rows = []
        s_df  = to_df(results["status"])
        st_df = to_df(results["stock"])

        for region in REGIONS:
            for mp in REGION_MP_MAP.get(region, []):
                sf = s_df[(s_df["Region"] == region) & (s_df["Marketplace"] == mp)] if not s_df.empty else pd.DataFrame()
                kf = st_df[(st_df["Region"] == region) & (st_df["Marketplace"] == mp)] if not st_df.empty else pd.DataFrame()
                tbl = to_df(results["to_be_listed"])
                tf  = tbl[(tbl["Region"] == region) & (tbl["Marketplace"] == mp)] if not tbl.empty else pd.DataFrame()
                vdf = to_df(results["variant_gaps"])
                vf  = vdf[(vdf["Region"] == region) & (vdf["Marketplace"] == mp)] if not vdf.empty else pd.DataFrame()

                summary_rows.append({
                    "Region": region,
                    "Marketplace": mp,
                    "Total EANs Checked": len(sf),
                    "Active (MP)": len(sf[sf["MP Status"] == "ACTIVE"]) if not sf.empty else 0,
                    "Inactive (MP)": len(sf[sf["MP Status"] == "INACTIVE"]) if not sf.empty else 0,
                    "Status Match ✅": len(sf[sf["Status Check"] == "✅ MATCH"]) if not sf.empty else 0,
                    "Wrongly Active ❌": len(sf[sf["Status Check"] == "❌ WRONGLY ACTIVE"]) if not sf.empty else 0,
                    "Wrongly Inactive ❌": len(sf[sf["Status Check"] == "❌ WRONGLY INACTIVE"]) if not sf.empty else 0,
                    "Missing Listings": len(sf[sf["Status Check"].str.contains("MISSING", na=False)]) if not sf.empty else 0,
                    "Stock Mismatches": len(kf[~kf["Stock Check"].str.startswith("✅", na=False)]) if not kf.empty else 0,
                    "To Be Listed": len(tf),
                    "Variant Gaps": len(vf),
                })

        _write_sheet(writer, pd.DataFrame(summary_rows), "SUMMARY")

        # ── MAIN VALIDATION SHEETS ───────────────────────────
        if not s_df.empty:
            _write_sheet(writer, s_df, "FINAL_STATUS_UPLOAD")

        if not st_df.empty:
            _write_sheet(writer, st_df, "FINAL_STOCK_UPLOAD")

        tbl = to_df(results["to_be_listed"])
        if not tbl.empty:
            _write_sheet(writer, tbl, "FINAL_TO_BE_LISTED")

        vdf = to_df(results["variant_gaps"])
        if not vdf.empty:
            _write_sheet(writer, vdf, "FINAL_VARIANT_GAPS")

        # ── PER-MARKETPLACE SHEETS ───────────────────────────
        for mp in MARKETPLACES:
            sf = s_df[s_df["Marketplace"] == mp] if not s_df.empty else pd.DataFrame()
            kf = st_df[st_df["Marketplace"] == mp] if not st_df.empty else pd.DataFrame()

            if not sf.empty:
                _write_sheet(writer, sf[sf["MP Status"] == "ACTIVE"],
                             f"{mp}_Active")
                _write_sheet(writer, sf[sf["MP Status"] == "INACTIVE"],
                             f"{mp}_Inactive")
                _write_sheet(writer, sf[sf["Status Check"].str.contains("MISMATCH|WRONG",
                             na=False, regex=True)], f"{mp}_Status_Mismatch")
                _write_sheet(writer, sf[sf["Status Check"].str.contains("MISSING",
                             na=False)], f"{mp}_Missing")

            if not kf.empty:
                _write_sheet(writer, kf[~kf["Stock Check"].str.startswith("✅", na=False)],
                             f"{mp}_Stock_Mismatch")

            tf = tbl[tbl["Marketplace"] == mp] if not tbl.empty else pd.DataFrame()
            if not tf.empty:
                _write_sheet(writer, tf, f"{mp}_To_Be_Listed")

        # ── ACTION-ONLY SHEETS ───────────────────────────────
        if not s_df.empty:
            _write_sheet(writer,
                s_df[s_df["Action Required"].isin(["Activate","Inactivate",
                                                    "Import SKU / Create Listing"])],
                "ACTION_STATUS_REQUIRED")

        if not st_df.empty:
            _write_sheet(writer,
                st_df[st_df["Action Required"] != "No Action"],
                "ACTION_STOCK_REQUIRED")

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────

def show_kpis(results: Dict):
    s_df = pd.DataFrame(results["status"]) if results["status"] else pd.DataFrame()
    k_df = pd.DataFrame(results["stock"])  if results["stock"]  else pd.DataFrame()
    t_df = pd.DataFrame(results["to_be_listed"]) if results["to_be_listed"] else pd.DataFrame()
    v_df = pd.DataFrame(results["variant_gaps"]) if results["variant_gaps"] else pd.DataFrame()

    r1 = st.columns(4)
    r2 = st.columns(4)
    r3 = st.columns(3)

    def kpi(col, label, val, color="normal"):
        col.metric(label, val)

    kpi(r1[0], "📦 Total EANs",       len(s_df))
    kpi(r1[1], "✅ Status Match",
        len(s_df[s_df["Status Check"] == "✅ MATCH"]) if not s_df.empty else 0)
    kpi(r1[2], "❌ Wrongly Active",
        len(s_df[s_df["Status Check"] == "❌ WRONGLY ACTIVE"]) if not s_df.empty else 0)
    kpi(r1[3], "❌ Wrongly Inactive",
        len(s_df[s_df["Status Check"] == "❌ WRONGLY INACTIVE"]) if not s_df.empty else 0)

    kpi(r2[0], "📋 Missing Listings",
        len(s_df[s_df["Status Check"].str.contains("MISSING", na=False)]) if not s_df.empty else 0)
    kpi(r2[1], "📊 Stock Mismatches",
        len(k_df[~k_df["Stock Check"].str.startswith("✅", na=False)]) if not k_df.empty else 0)
    kpi(r2[2], "🆕 To Be Listed",      len(t_df))
    kpi(r2[3], "🔍 Variant Gaps",      len(v_df))

    kpi(r3[0], "🚀 Actions Required",
        len(s_df[s_df["Action Required"] != "No Action"]) +
        len(k_df[k_df.get("Action Required", pd.Series()) != "No Action"]) if not k_df.empty else
        len(s_df[s_df["Action Required"] != "No Action"]) if not s_df.empty else 0)
    kpi(r3[1], "🏆 Hero SKUs",
        len(s_df[s_df["Comments"].str.contains("Hero", na=False)]) if not s_df.empty else 0)
    kpi(r3[2], "🚫 Excluded SKUs",
        len(s_df[s_df["Comments"].str.contains("Excluded", na=False)]) if not s_df.empty else 0)


# ─────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────

# ── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/6/6d/Puma_logo.svg/200px-Puma_logo.svg.png",
             width=120)
    st.title("🐾 PUMA Validator")
    st.markdown("---")

    st.subheader("🌍 Region Filter")
    sel_regions = []
    for r in REGIONS:
        if st.checkbox(r, value=True, key=f"reg_{r}"):
            sel_regions.append(r)

    st.markdown("---")
    st.subheader("🏪 Marketplace Filter")
    sel_mps = []
    for m in MARKETPLACES:
        if st.checkbox(m, value=True, key=f"mp_{m}"):
            sel_mps.append(m)

    st.markdown("---")
    st.subheader("🔍 Validation Type")
    val_status = st.checkbox("Status Validation", value=True)
    val_stock  = st.checkbox("Stock Validation",  value=True)
    val_types  = []
    if val_status: val_types.append("Status")
    if val_stock:  val_types.append("Stock")

    st.markdown("---")
    if st.session_state.run_done:
        st.success("✅ Analysis ready")
        st.caption(f"Last run: {datetime.now().strftime('%H:%M')}")

# ── MAIN AREA ────────────────────────────────────────────────
st.title("🐾 PUMA Marketplace Stock & Status Validator")
st.caption("Enterprise-grade validation automation | v3.1 — Excel & CSV supported")

tabs = st.tabs([
    "📁 Upload Files",
    "📊 Dashboard",
    "📋 Status Results",
    "📦 Stock Results",
    "🆕 To Be Listed",
    "🔍 Variant Gaps",
    "📥 Export",
    "📝 Logs",
])

# ── TAB 0: UPLOAD ─────────────────────────────────────────────
with tabs[0]:
    st.header("📁 File Uploads")
    st.info("Upload all required files. Both **Excel (.xlsx / .xls)** and **CSV (.csv)** formats are supported for all files **except ZeCom Tracker** (Excel only). Marketplace files are optional per region — upload only what you have.")

    # ── Core Files ────────────────────────────────────────────
    st.subheader("🗂️ Core Master Files")
    c1, c2, c3 = st.columns(3)
    with c1:
        uf_content = st.file_uploader(
            "📚 Content Master File",
            type=["xlsx","xls","csv"], key="content_file",
            help="Super master — Article No → EAN mapping")
    with c2:
        uf_zecom = st.file_uploader(
            "📋 ZeCom Tracker  ⚠️ Excel only",
            type=["xlsx","xls"], key="zecom_file",
            help="Excel only (.xlsx/.xls). Has sub-sheets per region (SG/MY/PH) + Lazada/Shopee/Zalora/TikTok columns")
    with c3:
        uf_pm = st.file_uploader(
            "📊 Product Master Stock ('ALL')",
            type=["xlsx","xls","csv"], key="pm_file",
            help="filename contains ALL. Final stock source of truth")

    st.markdown("---")

    # ── Override / Exclusion / Hero ───────────────────────────
    st.subheader("⚙️ Override & Rule Files")
    o1, o2, o3 = st.columns(3)
    with o1:
        uf_override = st.file_uploader(
            "🎯 Special Override File",
            type=["xlsx","xls","csv"], key="override_file",
            help="Article No + Final Status — highest priority")
    with o2:
        uf_exclusion = st.file_uploader(
            "🚫 Exclusion File",
            type=["xlsx","xls","csv"], key="exclusion_file",
            help="Force INACTIVE for specific articles")
    with o3:
        uf_hero = st.file_uploader(
            "🏆 Hero SKU File",
            type=["xlsx","xls","csv"], key="hero_file",
            help="Article No + Marketplace + Region — keeps ACTIVE override")

    st.markdown("---")

    # ── Regional Inventory ────────────────────────────────────
    st.subheader("🏭 Regional Inventory Files")
    i1, i2, i3 = st.columns(3)
    with i1:
        uf_inv_sg = st.file_uploader(
            "📦 Inventory — SG",
            type=["xlsx","xls","csv"], key="inv_sg",
            help="Starts with: SG_PUMA SG B2C Inventory Rpt_New_")
    with i2:
        uf_inv_my = st.file_uploader(
            "📦 Inventory — MY",
            type=["xlsx","xls","csv"], key="inv_my",
            help="Starts with: PUMA_MY_B2C_Channel_Inventory_")
    with i3:
        uf_inv_ph = st.file_uploader(
            "📦 Inventory — PH",
            type=["xlsx","xls","csv"], key="inv_ph",
            help="Starts with: Inventory_")

    st.markdown("---")

    # ── Marketplace Files ─────────────────────────────────────
    for region in REGIONS:
        st.subheader(f"🛒 Marketplace Files — {region}")
        mps = REGION_MP_MAP[region]
        cols = st.columns(len(mps) + (1 if "Zalora" in mps else 0))
        ci = 0
        for mp in mps:
            if mp == "Zalora":
                with cols[ci]:
                    st.file_uploader(
                        f"Zalora {region} — STATUS (SellerStatusTemplate)",
                        type=["xlsx","xls","csv"],
                        key=f"mp_{region}_Zalora_status")
                ci += 1
                with cols[ci]:
                    st.file_uploader(
                        f"Zalora {region} — STOCK (SellerStockTemplate)",
                        type=["xlsx","xls","csv"],
                        key=f"mp_{region}_Zalora_stock")
            else:
                label_hints = {
                    "Lazada": "pricestock…",
                    "Shopee": "Shopee PH Masterfile…",
                    "TikTok": "Tiktoksellercenter_batchedit…",
                }
                with cols[ci]:
                    st.file_uploader(
                        f"{mp} {region} ({label_hints.get(mp,'')})",
                        type=["xlsx","xls","csv"],
                        key=f"mp_{region}_{mp}")
            ci += 1

    st.markdown("---")

    # ── RUN BUTTON ────────────────────────────────────────────
    col_btn = st.columns([2, 1, 2])
    with col_btn[1]:
        run_btn = st.button("🚀 Run Validation", type="primary",
                            use_container_width=True)

    if run_btn:
        if not uf_content or not uf_zecom:
            st.error("❌ Content Master and ZeCom Tracker are required.")
        elif not sel_regions:
            st.error("❌ Select at least one region.")
        elif not sel_mps:
            st.error("❌ Select at least one marketplace.")
        elif not val_types:
            st.error("❌ Select at least one validation type.")
        else:
            st.session_state.logs = []
            progress = st.progress(0, text="⏳ Loading files…")

            with st.spinner("Loading and processing files…"):
                # Load core files
                progress.progress(0.05, "Loading Content Master…")
                content = load_content_file(uf_content)

                progress.progress(0.10, "Loading ZeCom Tracker…")
                zecom = load_zecom(uf_zecom)

                progress.progress(0.15, "Loading Product Master Stock…")
                pm = load_product_master(uf_pm) if uf_pm else {}

                progress.progress(0.20, "Loading Override / Exclusion / Hero files…")
                overrides  = load_special_override(uf_override)  if uf_override  else {}
                exclusions = load_exclusion_file(uf_exclusion)   if uf_exclusion else []
                heroes     = load_hero_sku(uf_hero)              if uf_hero      else []

                # Regional inventory
                progress.progress(0.25, "Loading inventory files…")
                inventories = {}
                if uf_inv_sg: inventories["SG"] = load_regional_inventory(uf_inv_sg, "SG")
                if uf_inv_my: inventories["MY"] = load_regional_inventory(uf_inv_my, "MY")
                if uf_inv_ph: inventories["PH"] = load_regional_inventory(uf_inv_ph, "PH")

                # Marketplace files
                progress.progress(0.30, "Loading marketplace files…")
                mp_data: Dict[str, Dict[str, Dict[str, Dict]]] = {}

                for region in REGIONS:
                    mp_data[region] = {}
                    for mp in REGION_MP_MAP[region]:
                        if mp == "Zalora":
                            sf_key = f"mp_{region}_Zalora_status"
                            sk_key = f"mp_{region}_Zalora_stock"
                            sf = st.session_state.get(sf_key)
                            sk = st.session_state.get(sk_key)
                            # Streamlit stores uploaded files in session via widget keys
                            # Access from widget states
                            sf2 = st.session_state.get(sf_key)
                            sk2 = st.session_state.get(sk_key)
                            if sf2 and sk2:
                                mp_data[region][mp] = load_zalora_combined(sf2, sk2)
                            else:
                                log(f"⊘ Zalora {region}: missing status/stock file")
                        else:
                            fkey = f"mp_{region}_{mp}"
                            uf = st.session_state.get(fkey)
                            if uf:
                                mp_data[region][mp] = load_marketplace_file(uf, mp)
                            else:
                                log(f"⊘ {mp} {region}: file not uploaded")

                # Run engine
                progress.progress(0.40, "Running validation engine…")
                engine = ValidationEngine(
                    content=content,
                    zecom=zecom,
                    product_master=pm,
                    inventories=inventories,
                    mp_data=mp_data,
                    overrides=overrides,
                    exclusions=exclusions,
                    heroes=heroes,
                )

                results = engine.run_all(
                    regions_filter=sel_regions,
                    mp_filter=sel_mps,
                    val_types=val_types,
                    progress_bar=progress,
                )

            st.session_state.results   = results
            st.session_state.run_done  = True
            progress.progress(1.0, "✅ Done!")
            st.success("✅ Validation complete! Switch to Dashboard tab.")

# ── TAB 1: DASHBOARD ─────────────────────────────────────────
with tabs[1]:
    st.header("📊 Validation Dashboard")
    if not st.session_state.run_done:
        st.info("Run validation first (Upload Files tab).")
    else:
        results = st.session_state.results
        show_kpis(results)
        st.markdown("---")

        # Charts
        try:
            import plotly.express as px

            s_df = pd.DataFrame(results["status"]) if results["status"] else pd.DataFrame()
            k_df = pd.DataFrame(results["stock"])  if results["stock"]  else pd.DataFrame()

            if not s_df.empty:
                cc1, cc2 = st.columns(2)
                with cc1:
                    st.subheader("Status Check Distribution")
                    vc = s_df["Status Check"].value_counts().reset_index()
                    vc.columns = ["Status Check", "Count"]
                    fig = px.bar(vc, x="Status Check", y="Count",
                                 color="Status Check",
                                 color_discrete_sequence=px.colors.qualitative.Set2)
                    fig.update_layout(showlegend=False, height=350)
                    st.plotly_chart(fig, use_container_width=True)

                with cc2:
                    st.subheader("Issues by Marketplace")
                    bad = s_df[s_df["Status Check"] != "✅ MATCH"]
                    if not bad.empty:
                        vc2 = bad.groupby("Marketplace").size().reset_index(name="Issues")
                        fig2 = px.pie(vc2, names="Marketplace", values="Issues",
                                      color_discrete_sequence=px.colors.qualitative.Pastel)
                        fig2.update_layout(height=350)
                        st.plotly_chart(fig2, use_container_width=True)

                cc3, cc4 = st.columns(2)
                with cc3:
                    st.subheader("Issues by Region")
                    bad2 = s_df[s_df["Status Check"] != "✅ MATCH"]
                    if not bad2.empty:
                        vc3 = bad2.groupby("Region").size().reset_index(name="Issues")
                        fig3 = px.bar(vc3, x="Region", y="Issues",
                                      color="Region",
                                      color_discrete_sequence=px.colors.qualitative.Bold)
                        fig3.update_layout(showlegend=False, height=300)
                        st.plotly_chart(fig3, use_container_width=True)

                with cc4:
                    if not k_df.empty:
                        st.subheader("Stock Action Required")
                        vc4 = k_df["Action Required"].value_counts().reset_index()
                        vc4.columns = ["Action", "Count"]
                        fig4 = px.bar(vc4, x="Action", y="Count",
                                      color="Action",
                                      color_discrete_sequence=px.colors.qualitative.Vivid)
                        fig4.update_layout(showlegend=False, height=300)
                        st.plotly_chart(fig4, use_container_width=True)

        except ImportError:
            st.info("Install plotly for charts: pip install plotly")

# ── TAB 2: STATUS RESULTS ─────────────────────────────────────
with tabs[2]:
    st.header("📋 Status Validation Results")
    if not st.session_state.run_done:
        st.info("Run validation first.")
    else:
        s_df = pd.DataFrame(st.session_state.results["status"]) if st.session_state.results["status"] else pd.DataFrame()
        if s_df.empty:
            st.warning("No status validation data.")
        else:
            # Filters
            f1, f2, f3, f4 = st.columns(4)
            with f1:
                fr = st.multiselect("Region", REGIONS, default=REGIONS, key="sf_reg")
            with f2:
                fm = st.multiselect("Marketplace", MARKETPLACES, default=MARKETPLACES, key="sf_mp")
            with f3:
                checks = s_df["Status Check"].unique().tolist()
                fc = st.multiselect("Status Check", checks, default=checks, key="sf_chk")
            with f4:
                actions = s_df["Action Required"].unique().tolist()
                fa = st.multiselect("Action", actions, default=actions, key="sf_act")

            mask = (
                s_df["Region"].isin(fr) &
                s_df["Marketplace"].isin(fm) &
                s_df["Status Check"].isin(fc) &
                s_df["Action Required"].isin(fa)
            )
            filtered = s_df[mask].drop(columns=["_region","_mp"], errors="ignore")
            st.caption(f"Showing {len(filtered):,} of {len(s_df):,} rows")
            st.dataframe(filtered, use_container_width=True, height=500)

# ── TAB 3: STOCK RESULTS ──────────────────────────────────────
with tabs[3]:
    st.header("📦 Stock Validation Results")
    if not st.session_state.run_done:
        st.info("Run validation first.")
    else:
        k_df = pd.DataFrame(st.session_state.results["stock"]) if st.session_state.results["stock"] else pd.DataFrame()
        if k_df.empty:
            st.warning("No stock validation data.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                fr = st.multiselect("Region", REGIONS, default=REGIONS, key="kf_reg")
            with f2:
                fm = st.multiselect("Marketplace", MARKETPLACES, default=MARKETPLACES, key="kf_mp")
            with f3:
                actions = k_df["Action Required"].unique().tolist()
                fa = st.multiselect("Action", actions, default=actions, key="kf_act")

            mask = (
                k_df["Region"].isin(fr) &
                k_df["Marketplace"].isin(fm) &
                k_df["Action Required"].isin(fa)
            )
            filtered = k_df[mask].drop(columns=["_region","_mp"], errors="ignore")
            st.caption(f"Showing {len(filtered):,} of {len(k_df):,} rows")
            st.dataframe(filtered, use_container_width=True, height=500)

# ── TAB 4: TO BE LISTED ───────────────────────────────────────
with tabs[4]:
    st.header("🆕 To Be Listed")
    if not st.session_state.run_done:
        st.info("Run validation first.")
    else:
        t_df = pd.DataFrame(st.session_state.results["to_be_listed"]) if st.session_state.results["to_be_listed"] else pd.DataFrame()
        if t_df.empty:
            st.success("✅ No to-be-listed gaps found.")
        else:
            f1, f2 = st.columns(2)
            with f1:
                fr = st.multiselect("Region", REGIONS, default=REGIONS, key="tf_reg")
            with f2:
                fm = st.multiselect("Marketplace", MARKETPLACES, default=MARKETPLACES, key="tf_mp")
            filtered = t_df[t_df["Region"].isin(fr) & t_df["Marketplace"].isin(fm)]
            filtered = filtered.drop(columns=["_region","_mp"], errors="ignore")
            st.metric("To Be Listed Count", len(filtered))
            st.dataframe(filtered, use_container_width=True, height=450)

# ── TAB 5: VARIANT GAPS ───────────────────────────────────────
with tabs[5]:
    st.header("🔍 Variant Gaps")
    if not st.session_state.run_done:
        st.info("Run validation first.")
    else:
        v_df = pd.DataFrame(st.session_state.results["variant_gaps"]) if st.session_state.results["variant_gaps"] else pd.DataFrame()
        if v_df.empty:
            st.success("✅ No variant gaps found.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                fr = st.multiselect("Region", REGIONS, default=REGIONS, key="vf_reg")
            with f2:
                fm = st.multiselect("Marketplace", MARKETPLACES, default=MARKETPLACES, key="vf_mp")
            with f3:
                gap_types = v_df["Gap Type"].unique().tolist()
                fg = st.multiselect("Gap Type", gap_types, default=gap_types, key="vf_gap")

            filtered = v_df[
                v_df["Region"].isin(fr) &
                v_df["Marketplace"].isin(fm) &
                v_df["Gap Type"].isin(fg)
            ]
            filtered = filtered.drop(columns=["_region","_mp"], errors="ignore")
            col1, col2 = st.columns(2)
            col1.metric("Articles with Gaps", len(filtered))
            col2.metric("Total Missing Variants",
                        int(filtered["Missing Variants"].astype(float).sum()) if not filtered.empty else 0)
            st.dataframe(filtered, use_container_width=True, height=450)

# ── TAB 6: EXPORT ─────────────────────────────────────────────
with tabs[6]:
    st.header("📥 Export Results")
    if not st.session_state.run_done:
        st.info("Run validation first.")
    else:
        st.markdown("""
        The exported Excel file contains:
        - **SUMMARY** — High-level KPIs per region/marketplace
        - **FINAL_STATUS_UPLOAD** — All status validation rows
        - **FINAL_STOCK_UPLOAD** — All stock validation rows
        - **FINAL_TO_BE_LISTED** — EANs to be listed
        - **FINAL_VARIANT_GAPS** — Missing variant analysis
        - **Per-marketplace sheets** — Active, Inactive, Mismatches, Missing
        - **ACTION sheets** — Only rows requiring action
        """)

        if st.button("📊 Generate Excel Report", type="primary"):
            with st.spinner("Generating Excel…"):
                xlsx_bytes = build_excel(st.session_state.results)
            fname = f"PUMA_Validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.download_button(
                label="⬇️ Download Excel Report",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.success(f"✅ Ready: {fname}")

# ── TAB 7: LOGS ───────────────────────────────────────────────
with tabs[7]:
    st.header("📝 Processing Logs")
    if st.session_state.logs:
        log_text = "\n".join(st.session_state.logs)
        st.text_area("Logs", log_text, height=400)
        st.download_button("⬇️ Download Logs", log_text,
                           file_name="validation_logs.txt")
    else:
        st.info("No logs yet. Run validation to see processing details.")
