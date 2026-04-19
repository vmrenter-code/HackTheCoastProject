import argparse
from pathlib import Path
import re
import sys
from typing import Optional

import numpy as np
import pandas as pd

# ============================================================
# DEFAULT CONFIG
# ============================================================

DEFAULT_BASE_DIR = Path(__file__).resolve().parent

DEFAULT_SALES_FILE = "POP_SalesTransactionHistory.csv"
DEFAULT_CHARGEBACK_FILE = "POP_ChargeBack_Deductions_Penalties_Freight.xlsx"
DEFAULT_CHANNEL_KEY_FILE = "SLPRSNID_SALESCHANNEL_KEY.xlsx"  # optional
DEFAULT_CHARGEBACK_SHEET = "Data - Deductions & Cause Code"

CLEANED_DEMAND_OUT = "cleaned_demand_summary.csv"
RAW_VS_CLEANED_OUT = "raw_vs_cleaned_comparison.csv"
STOCKOUT_OUT = "stockout_flags.csv"
SKU_IMPACT_OUT = "sku_impact_summary.csv"
CHANNEL_SUMMARY_OUT = "channel_summary.csv"
CHANNEL_SIGNAL_OUT = "channel_demand_signal.csv"

PROMO_CAUSE_CODES = {
    "CRED02",    # Retailer TPR / Insertion / Admin Fee
    "CRED03",    # Retailer TPR Scan Down
    "CRED04",    # Distributor Promotion Support / TPR Fees
    "CRED05",    # Distributor Promotion Support / TPR Scan Down
    "CRED15",    # Markdown Charges (%)
    "CRED-SDT",  # Short-Dated Items
    "CRED-PRO",  # Promotional such as End Cap
}

FILE_SEARCH_HINTS = {
    "sales": ("sales", "transaction", "history"),
    "chargebacks": ("charge", "deduction", "penalt", "freight", "cause"),
    "channel_key": ("slprsnid", "saleschannel", "channel"),
}

REQUIRED_SALES_COLUMNS = [
    "LOCNCODE",
    "SOP TYPE",
    "DOCDATE",
    "ITEMNMBR",
    "ITEMDESC",
    "CUSTNMBR",
    "QUANTITY_adj",
    "XTNDPRCE_adj",
]

REQUIRED_CHARGEBACK_COLUMNS = [
    "Document Date",
    "Customer Number",
    "Item Description",
    "Cause Code",
]

CHANNEL_KEY_COLUMNS = ["SLPRSNID", "SALESCHANNEL", "SALESCHANNEL_DESC"]

LOCATION_NAME_MAP = {
    "1": "SF",
    "2": "NJ",
    "3": "LA",
}

CHANNEL_PLANNING_RULES = {
    "American Market": {
        "demand_pattern": "planogram-driven",
        "lookback_months": 3,
        "planning_note": "Use recent trend plus seasonal resets and merchandising windows.",
    },
    "Health Food": {
        "demand_pattern": "steady-distributor",
        "lookback_months": 6,
        "planning_note": "Use smoother history because distributor demand is steadier but promo influenced.",
    },
    "Asian Market": {
        "demand_pattern": "opportunistic-spiky",
        "lookback_months": 3,
        "planning_note": "Protect against spikes and short gaps because demand is reactive and lumpy.",
    },
    "eCom": {
        "demand_pattern": "long-tail",
        "lookback_months": 6,
        "planning_note": "Use a longer tail because order sizes are small and evolving.",
    },
    "Other": {
        "demand_pattern": "mixed",
        "lookback_months": 4,
        "planning_note": "Use blended rules and review manually for unusual signals.",
    },
}

# Fallback mapping if SLPRSNID_SALESCHANNEL_KEY.xlsx is missing
AMERICAN_MARKET = {
    "MASS", "SUPERMARKET", "DRUG", "FOOD", "GROCERY",
    "DOLLAR", "DIST DRUG", "DIST FOOD", "AM OTHERS", "HARDWARE",
    "MILITARY", "CLUB",
}
HEALTH_FOOD = {"DIST HF", "HEALTH FOOD VMS", "HERBAL"}
ECOM = {"ONLINE", "ONLINE MAIL ORD", "E COMM", "E COMM", "E-COMM", "E COMM ", "E-COM"}
ASIAN_MARKET = {"CHAIN", "SUB D", "TREASURE DISC", "GIFT SHOP", "TREASURE DISC"}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_text(value) -> str:
    """Normalize free-text fields for safer joins and comparisons."""
    if pd.isna(value):
        return ""
    text = str(value).upper().strip()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_filename(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def label_location(locncode) -> str:
    code = str(locncode).strip()
    return LOCATION_NAME_MAP.get(code, "UNMAPPED")


def build_path(base_dir: Path, raw_path: Optional[str]) -> Optional[Path]:
    if not raw_path:
        return None
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def score_candidate(path: Path, search_terms, suffix: str) -> int:
    if suffix and path.suffix.lower() != suffix.lower():
        return 0

    normalized_name = normalize_filename(path.stem)
    score = 0
    for term in search_terms:
        if normalize_filename(term) in normalized_name:
            score += 1
    return score


def find_candidate_files(base_dir: Path, search_terms, suffix: str):
    candidates = []
    if not base_dir.exists():
        return candidates

    for path in base_dir.iterdir():
        if not path.is_file():
            continue
        score = score_candidate(path, search_terms, suffix)
        if score > 0:
            candidates.append((score, path))

    return sorted(candidates, key=lambda item: (-item[0], item[1].name.lower()))


def resolve_input_file(base_dir: Path, raw_path: Optional[str], label: str, search_terms, required=True):
    path = build_path(base_dir, raw_path)
    if path is None:
        return None

    if path.exists():
        return path

    candidates = find_candidate_files(path.parent, search_terms, path.suffix)
    if len(candidates) == 1 and candidates[0][0] >= 2:
        chosen = candidates[0][1]
        print(f"Warning: {label} not found at {path}. Using {chosen.name} instead.")
        return chosen

    if not required:
        print(f"{label} not found. Using fallback logic.")
        if candidates:
            print("  Possible nearby matches:")
            for _, candidate in candidates[:5]:
                print(f"    - {candidate.name}")
        return None

    message = [f"{label} not found: {path}"]
    if candidates:
        message.append("Potential matches in the same folder:")
        for _, candidate in candidates[:5]:
            message.append(f"  - {candidate.name}")
    message.append("Update the filename in the script or pass a file path with command-line arguments.")
    raise SystemExit("\n".join(message))


def ensure_excel_support(required_excel_paths):
    needs_openpyxl = any(
        path is not None and path.suffix.lower() in {".xlsx", ".xlsm", ".xltx", ".xltm"}
        for path in required_excel_paths
    )
    if not needs_openpyxl:
        return

    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        raise SystemExit(
            "Missing Python dependency 'openpyxl'. "
            "Install it with `python3 -m pip install openpyxl`, "
            "or run this script with a Python environment that already has openpyxl."
        ) from exc


def safe_read_excel(path: Path, sheet_name=None, label="Excel file"):
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except ValueError as exc:
        if sheet_name is None or "Worksheet named" not in str(exc):
            raise

        workbook = pd.ExcelFile(path)
        chosen_sheet = pick_sheet_name(workbook.sheet_names, sheet_name)
        if chosen_sheet is not None:
            print(
                f"Warning: sheet {sheet_name!r} not found in {path.name}. "
                f"Using {chosen_sheet!r} instead."
            )
            return workbook.parse(chosen_sheet)

        available = ", ".join(workbook.sheet_names)
        raise SystemExit(
            f"{label} is missing sheet {sheet_name!r}. "
            f"Available sheets: {available}"
        ) from exc


def pick_sheet_name(sheet_names, target_sheet):
    target_norm = normalize_filename(target_sheet)
    tokens = [token for token in re.findall(r"[A-Za-z0-9]+", target_sheet.lower()) if len(token) >= 4]

    scored = []
    for name in sheet_names:
        name_norm = normalize_filename(name)
        score = 0
        if name_norm == target_norm:
            score += 100
        if target_norm in name_norm or name_norm in target_norm:
            score += 5
        for token in tokens:
            if token in name.lower():
                score += 1
        if score > 0:
            scored.append((score, name))

    if not scored:
        return None

    scored.sort(reverse=True)
    top_score, top_name = scored[0]
    if len(scored) == 1 or top_score > scored[1][0]:
        return top_name
    return None


def ensure_required_columns(df: pd.DataFrame, required_columns, dataset_name: str):
    missing = [column for column in required_columns if column not in df.columns]
    if not missing:
        return

    preview = ", ".join(df.columns[:20])
    if len(df.columns) > 20:
        preview = f"{preview}, ..."

    raise SystemExit(
        f"{dataset_name} is missing required columns: {', '.join(missing)}\n"
        f"Available columns: {preview}"
    )


def map_super_channel(customer_type) -> str:
    """Fallback channel grouping if no external channel key exists."""
    ct = normalize_text(customer_type)
    if ct in AMERICAN_MARKET:
        return "American Market"
    if ct in HEALTH_FOOD:
        return "Health Food"
    if ct in ECOM:
        return "eCom"
    if ct in ASIAN_MARKET:
        return "Asian Market"
    return "Other"


def load_sales(sales_file: Path) -> pd.DataFrame:
    print("Loading sales data...")
    sales = pd.read_csv(sales_file, low_memory=False)
    print(f"  Loaded {len(sales):,} raw sales rows")

    ensure_required_columns(sales, REQUIRED_SALES_COLUMNS, "Sales file")

    if "Customer Type" not in sales.columns:
        print("Warning: 'Customer Type' column not found. Channel fallback will use UNKNOWN.")
        sales["Customer Type"] = "UNKNOWN"
    else:
        sales["Customer Type"] = sales["Customer Type"].fillna("UNKNOWN")

    sales = sales[sales["SOP TYPE"].astype(str).str.upper().eq("INVOICE")].copy()
    print(f"  After filtering to invoices: {len(sales):,} rows")

    sales["DOCDATE"] = pd.to_datetime(sales["DOCDATE"], errors="coerce")
    sales = sales[sales["DOCDATE"].notna()].copy()

    sales["month"] = sales["DOCDATE"].dt.to_period("M")
    sales["month_start"] = sales["month"].dt.to_timestamp()

    for col in ["QUANTITY_adj", "XTNDPRCE_adj", "EXTDCOST_adj", "Unit_Price_adj"]:
        if col in sales.columns:
            sales[col] = pd.to_numeric(sales[col], errors="coerce").fillna(0)

    sales["LOCNCODE"] = sales["LOCNCODE"].fillna("").astype(str).str.strip()
    sales["LOCATION_NAME"] = sales["LOCNCODE"].apply(label_location)
    sales["ITEMNMBR"] = sales["ITEMNMBR"].fillna("").astype(str).str.strip()
    sales["ITEMDESC"] = sales["ITEMDESC"].fillna("").astype(str).str.strip()
    sales["item_desc_key"] = sales["ITEMDESC"].apply(normalize_text)
    sales["CUSTNMBR"] = sales["CUSTNMBR"].fillna("").astype(str).str.strip()
    sales["Customer Type"] = sales["Customer Type"].fillna("UNKNOWN").astype(str).str.strip()

    return sales


def load_channel_key(channel_key_file: Optional[Path]):
    if channel_key_file is None:
        print("Channel key file not found. Using fallback channel logic from Customer Type.")
        return None

    print(f"Loading channel key: {channel_key_file.name}")
    channel_key = safe_read_excel(channel_key_file, sheet_name=0, label="Channel key file")
    if not set(CHANNEL_KEY_COLUMNS).issubset(channel_key.columns):
        print("Warning: channel key file exists but columns are not as expected. Using fallback channel logic.")
        return None

    return channel_key[CHANNEL_KEY_COLUMNS].drop_duplicates()


def add_channel_information(sales: pd.DataFrame, channel_key: Optional[pd.DataFrame]) -> pd.DataFrame:
    sales = sales.copy()

    if channel_key is not None and "SLPRSNID" in sales.columns:
        sales["SLPRSNID"] = sales["SLPRSNID"].fillna("").astype(str).str.strip()
        channel_key = channel_key.copy()
        channel_key["SLPRSNID"] = channel_key["SLPRSNID"].fillna("").astype(str).str.strip()

        sales = sales.merge(channel_key, on="SLPRSNID", how="left")
        sales["SALESCHANNEL"] = sales["SALESCHANNEL"].fillna("UNKNOWN")
        sales["SALESCHANNEL_DESC"] = sales["SALESCHANNEL_DESC"].fillna("Unknown Channel")
    else:
        if channel_key is not None and "SLPRSNID" not in sales.columns:
            print("Warning: channel key is available but sales data is missing 'SLPRSNID'. Using fallback channel logic.")
        sales["SALESCHANNEL"] = sales["Customer Type"].replace("", pd.NA).fillna("UNKNOWN")
        sales["SALESCHANNEL_DESC"] = sales["Customer Type"].replace("", pd.NA).fillna("Unknown Channel")

    sales["SUPER_CHANNEL"] = sales["Customer Type"].apply(map_super_channel)

    print("\nSales by super channel:")
    channel_check = (
        sales.groupby("SUPER_CHANNEL")["QUANTITY_adj"]
        .sum()
        .sort_values(ascending=False)
    )
    print(channel_check)

    return sales


def load_chargebacks(chargeback_file: Path, sheet_name: str) -> pd.DataFrame:
    print("\nLoading chargebacks...")
    chargebacks = safe_read_excel(
        chargeback_file,
        sheet_name=sheet_name,
        label="Chargeback file",
    )
    print(f"  Loaded {len(chargebacks):,} chargeback rows")

    ensure_required_columns(chargebacks, REQUIRED_CHARGEBACK_COLUMNS, "Chargeback file")

    chargebacks["Document Date"] = pd.to_datetime(chargebacks["Document Date"], errors="coerce")
    chargebacks = chargebacks[chargebacks["Document Date"].notna()].copy()

    chargebacks["month"] = chargebacks["Document Date"].dt.to_period("M")
    chargebacks["month_start"] = chargebacks["month"].dt.to_timestamp()
    chargebacks["Customer Number"] = chargebacks["Customer Number"].fillna("").astype(str).str.strip()
    chargebacks["item_desc_key"] = chargebacks["Item Description"].apply(normalize_text)
    chargebacks["Cause Code"] = chargebacks["Cause Code"].fillna("").astype(str).str.strip().str.upper()

    return chargebacks


def build_promo_flags(sales: pd.DataFrame, chargebacks: pd.DataFrame) -> pd.DataFrame:
    promo_cb = chargebacks[chargebacks["Cause Code"].isin(PROMO_CAUSE_CODES)].copy()
    print(f"  Promo / markdown chargeback rows: {len(promo_cb):,}")

    promo_item_level = (
        promo_cb[
            promo_cb["item_desc_key"].ne("") &
            promo_cb["Customer Number"].ne("")
        ][["Customer Number", "month", "month_start", "item_desc_key"]]
        .drop_duplicates()
        .rename(columns={"Customer Number": "CUSTNMBR"})
    )
    promo_item_level["promo_match_item"] = True

    promo_customer_month = (
        promo_cb[["Customer Number", "month", "month_start"]]
        .drop_duplicates()
        .rename(columns={"Customer Number": "CUSTNMBR"})
    )
    promo_customer_month["promo_match_customer_month"] = True

    sales = sales.merge(
        promo_item_level,
        on=["CUSTNMBR", "month", "month_start", "item_desc_key"],
        how="left",
    )

    sales = sales.merge(
        promo_customer_month,
        on=["CUSTNMBR", "month", "month_start"],
        how="left",
    )

    sales["promo_match_item"] = sales["promo_match_item"].eq(True)
    sales["promo_match_customer_month"] = sales["promo_match_customer_month"].eq(True)

    sales["promo_match_level"] = np.select(
        [
            sales["promo_match_item"],
            (~sales["promo_match_item"]) & sales["promo_match_customer_month"],
        ],
        [
            "item_customer_month",
            "customer_month_proxy",
        ],
        default="none",
    )

    sales["is_promo"] = sales["promo_match_level"].ne("none")

    promo_count = int(sales["is_promo"].sum())
    total_count = int(len(sales))
    print(
        f"\nFlagged {promo_count:,} sales rows as PROMOTIONAL "
        f"({promo_count / max(total_count, 1) * 100:.1f}% of all invoice rows)"
    )
    print("\nPromo match level breakdown:")
    print(sales["promo_match_level"].value_counts(dropna=False))

    return sales


def build_channel_signal(sales: pd.DataFrame) -> pd.DataFrame:
    sales = sales.copy()
    sales["organic_units_txn"] = np.where(sales["is_promo"], 0, sales["QUANTITY_adj"])
    sales["organic_revenue_txn"] = np.where(sales["is_promo"], 0, sales["XTNDPRCE_adj"])

    channel_signal = (
        sales.groupby(
            [
                "ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME",
                "month", "month_start", "SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC",
            ],
            as_index=False,
        )
        .agg(
            raw_units=("QUANTITY_adj", "sum"),
            raw_revenue=("XTNDPRCE_adj", "sum"),
            organic_units=("organic_units_txn", "sum"),
            organic_revenue=("organic_revenue_txn", "sum"),
        )
    )

    channel_signal["promo_units"] = channel_signal["raw_units"] - channel_signal["organic_units"]
    channel_signal["promo_revenue"] = channel_signal["raw_revenue"] - channel_signal["organic_revenue"]
    channel_signal["promo_pct_units"] = (
        channel_signal["promo_units"] / channel_signal["raw_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    channel_signal["item_location_key"] = (
        channel_signal["ITEMNMBR"].astype(str) + "|" + channel_signal["LOCNCODE"].astype(str)
    )

    return channel_signal


def build_location_totals(channel_signal: pd.DataFrame) -> pd.DataFrame:
    location_monthly = (
        channel_signal.groupby(
            ["ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME", "month", "month_start"],
            as_index=False,
        )
        .agg(
            raw_units=("raw_units", "sum"),
            raw_revenue=("raw_revenue", "sum"),
            organic_units=("organic_units", "sum"),
            organic_revenue=("organic_revenue", "sum"),
        )
    )

    location_monthly["promo_units"] = location_monthly["raw_units"] - location_monthly["organic_units"]
    location_monthly["promo_revenue"] = location_monthly["raw_revenue"] - location_monthly["organic_revenue"]

    return location_monthly


def positive_median(series) -> float:
    values = pd.Series(series)
    values = values[pd.notna(values) & (values > 0)]
    if values.empty:
        return 0.0
    return float(values.median())


def estimate_suppressed_demand(location_monthly: pd.DataFrame):
    if location_monthly.empty:
        empty_cols = [
            "ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME", "month", "month_start",
            "raw_units", "raw_revenue", "organic_units", "organic_revenue",
            "promo_units", "promo_revenue", "baseline_units", "within_active_span",
            "zero_sales_gap", "deep_drop_gap", "likely_stockout", "stockout_reason",
            "suppressed_units_est", "cleaned_units", "est_unit_price", "cleaned_revenue_est",
            "prev_raw_units", "next_raw_units",
        ]
        return pd.DataFrame(columns=empty_cols), pd.DataFrame(columns=empty_cols)

    full_rows = []

    group_cols = ["ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME"]
    for (item, itemdesc, locncode, location_name), group in location_monthly.groupby(group_cols):
        group = group.sort_values("month").reset_index(drop=True)

        full_months = pd.period_range(group["month"].min(), group["month"].max(), freq="M")
        full = pd.DataFrame({"month": full_months})
        full["month_start"] = full["month"].dt.to_timestamp()
        full["ITEMNMBR"] = item
        full["ITEMDESC"] = itemdesc
        full["LOCNCODE"] = locncode
        full["LOCATION_NAME"] = location_name

        full = full.merge(
            group,
            on=["ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME", "month", "month_start"],
            how="left",
        )

        for col in [
            "raw_units", "raw_revenue", "organic_units", "organic_revenue",
            "promo_units", "promo_revenue",
        ]:
            full[col] = full[col].fillna(0.0)

        has_sales = full["raw_units"] > 0
        if has_sales.any():
            first_pos = int(np.argmax(has_sales.to_numpy()))
            last_pos = int(len(full) - 1 - np.argmax(has_sales.to_numpy()[::-1]))
            full["within_active_span"] = False
            full.loc[first_pos:last_pos, "within_active_span"] = True
        else:
            full["within_active_span"] = False

        baselines = []
        for idx in range(len(full)):
            start_idx = max(0, idx - 2)
            end_idx = min(len(full), idx + 3)
            surrounding = full.iloc[start_idx:end_idx].copy()
            surrounding = surrounding[surrounding.index != idx]

            clean_baseline = positive_median(surrounding["organic_units"])
            raw_baseline = positive_median(surrounding["raw_units"])

            neighbor_values = []
            if idx > 0 and full.at[idx - 1, "raw_units"] > 0:
                neighbor_values.append(full.at[idx - 1, "raw_units"])
            if idx < len(full) - 1 and full.at[idx + 1, "raw_units"] > 0:
                neighbor_values.append(full.at[idx + 1, "raw_units"])
            neighbor_avg = float(np.mean(neighbor_values)) if len(neighbor_values) == 2 else 0.0

            baseline = clean_baseline if clean_baseline > 0 else neighbor_avg if neighbor_avg > 0 else raw_baseline
            baselines.append(float(baseline))

        full["baseline_units"] = baselines
        full["prev_raw_units"] = full["raw_units"].shift(1).fillna(0)
        full["next_raw_units"] = full["raw_units"].shift(-1).fillna(0)

        observed_unit_price = np.where(
            full["organic_units"] > 0,
            full["organic_revenue"] / full["organic_units"],
            np.where(
                full["raw_units"] > 0,
                full["raw_revenue"] / full["raw_units"],
                np.nan,
            ),
        )
        observed_unit_price = pd.Series(observed_unit_price).replace([np.inf, -np.inf], np.nan)
        median_price = observed_unit_price.dropna()
        median_price = float(median_price.median()) if not median_price.empty else 0.0
        full["est_unit_price"] = observed_unit_price.ffill().bfill().fillna(median_price).fillna(0)

        full["zero_sales_gap"] = (
            full["within_active_span"] &
            full["raw_units"].eq(0) &
            full["baseline_units"].gt(0)
        )
        full["deep_drop_gap"] = (
            full["within_active_span"] &
            full["raw_units"].gt(0) &
            full["baseline_units"].gt(0) &
            (full["raw_units"] < full["baseline_units"] * 0.4) &
            full["prev_raw_units"].gt(0) &
            full["next_raw_units"].gt(0)
        )
        full["likely_stockout"] = full["zero_sales_gap"] | full["deep_drop_gap"]
        full["stockout_reason"] = np.select(
            [full["zero_sales_gap"], full["deep_drop_gap"]],
            ["zero_sales_gap", "deep_drop_gap"],
            default="none",
        )
        full["suppressed_units_est"] = np.where(
            full["likely_stockout"],
            np.maximum(full["baseline_units"] - full["organic_units"], 0),
            0,
        ).round(1)
        full["cleaned_units"] = full["organic_units"] + full["suppressed_units_est"]
        full["cleaned_revenue_est"] = (
            full["organic_revenue"] + full["suppressed_units_est"] * full["est_unit_price"]
        ).round(2)

        full_rows.append(full)

    location_signal = pd.concat(full_rows, ignore_index=True)
    stockout_flags = location_signal[location_signal["likely_stockout"]].copy()

    print("\nBuilding stockout and suppressed-demand flags...")
    print(f"Detected {len(stockout_flags):,} likely constrained months")
    if not stockout_flags.empty:
        print(f"Across {stockout_flags[['ITEMNMBR', 'LOCNCODE']].drop_duplicates().shape[0]:,} SKU-location pairs")
        print(
            f"Estimated suppressed units: {stockout_flags['suppressed_units_est'].sum():,.0f}"
        )

    return location_signal, stockout_flags


def build_output_tables(location_signal: pd.DataFrame):
    comparison = location_signal.copy()
    comparison["promo_pct_units"] = (
        comparison["promo_units"] / comparison["raw_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    comparison["suppressed_pct_vs_cleaned"] = (
        comparison["suppressed_units_est"] / comparison["cleaned_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    comparison["cleaned_vs_raw_pct"] = (
        comparison["cleaned_units"] / comparison["raw_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)

    cleaned_summary = comparison[
        [
            "ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME", "month", "month_start",
            "raw_units", "organic_units", "suppressed_units_est", "cleaned_units",
            "raw_revenue", "organic_revenue", "cleaned_revenue_est",
            "promo_units", "promo_revenue", "likely_stockout", "stockout_reason",
        ]
    ].copy()

    return cleaned_summary, comparison


def build_sku_impact_summary(comparison: pd.DataFrame) -> pd.DataFrame:
    sku_impact = (
        comparison.groupby(["ITEMNMBR", "ITEMDESC"], as_index=False)
        .agg(
            raw_units=("raw_units", "sum"),
            organic_units=("organic_units", "sum"),
            suppressed_units_est=("suppressed_units_est", "sum"),
            cleaned_units=("cleaned_units", "sum"),
            raw_revenue=("raw_revenue", "sum"),
            organic_revenue=("organic_revenue", "sum"),
            cleaned_revenue_est=("cleaned_revenue_est", "sum"),
            stockout_months=("likely_stockout", "sum"),
            active_locations=("LOCNCODE", "nunique"),
        )
    )

    sku_impact["promo_units"] = sku_impact["raw_units"] - sku_impact["organic_units"]
    sku_impact["promo_revenue"] = sku_impact["raw_revenue"] - sku_impact["organic_revenue"]
    sku_impact["promo_pct_units"] = (
        sku_impact["promo_units"] / sku_impact["raw_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    sku_impact["suppressed_pct_vs_cleaned"] = (
        sku_impact["suppressed_units_est"] / sku_impact["cleaned_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    sku_impact["cleaned_lift_vs_organic_pct"] = (
        (sku_impact["cleaned_units"] - sku_impact["organic_units"]) /
        sku_impact["organic_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)

    sku_impact = sku_impact.sort_values(
        ["suppressed_units_est", "promo_units", "raw_units"],
        ascending=[False, False, False],
    )

    return sku_impact


def build_channel_summary(channel_signal: pd.DataFrame) -> pd.DataFrame:
    latest_month = channel_signal["month"].max()
    recent_3_cutoff = latest_month - 2
    recent_6_cutoff = latest_month - 5

    channel_summary = (
        channel_signal.groupby(
            ["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"],
            as_index=False,
        )
        .agg(
            raw_units=("raw_units", "sum"),
            organic_units=("organic_units", "sum"),
            raw_revenue=("raw_revenue", "sum"),
            organic_revenue=("organic_revenue", "sum"),
            active_skus=("ITEMNMBR", "nunique"),
            active_item_locations=("item_location_key", "nunique"),
            active_months=("month", "nunique"),
        )
    )

    recent_3 = (
        channel_signal[channel_signal["month"] >= recent_3_cutoff]
        .groupby(["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"], as_index=False)
        .agg(recent_3m_organic_units=("organic_units", "sum"))
    )
    recent_6 = (
        channel_signal[channel_signal["month"] >= recent_6_cutoff]
        .groupby(["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"], as_index=False)
        .agg(recent_6m_organic_units=("organic_units", "sum"))
    )
    channel_monthly = (
        channel_signal.groupby(["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC", "month"], as_index=False)
        .agg(monthly_organic_units=("organic_units", "sum"))
    )
    channel_volatility = (
        channel_monthly.groupby(["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"], as_index=False)
        .agg(
            monthly_avg_units=("monthly_organic_units", "mean"),
            monthly_std_units=("monthly_organic_units", "std"),
        )
    )

    channel_summary = channel_summary.merge(
        recent_3,
        on=["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"],
        how="left",
    )
    channel_summary = channel_summary.merge(
        recent_6,
        on=["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"],
        how="left",
    )
    channel_summary = channel_summary.merge(
        channel_volatility,
        on=["SUPER_CHANNEL", "SALESCHANNEL", "SALESCHANNEL_DESC"],
        how="left",
    )

    for col in ["recent_3m_organic_units", "recent_6m_organic_units", "monthly_avg_units", "monthly_std_units"]:
        channel_summary[col] = channel_summary[col].fillna(0).round(1)

    channel_summary["promo_units"] = channel_summary["raw_units"] - channel_summary["organic_units"]
    channel_summary["promo_revenue"] = channel_summary["raw_revenue"] - channel_summary["organic_revenue"]
    channel_summary["promo_pct_units"] = (
        channel_summary["promo_units"] / channel_summary["raw_units"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    channel_summary["monthly_cv"] = (
        channel_summary["monthly_std_units"] / channel_summary["monthly_avg_units"].replace(0, np.nan)
    ).fillna(0).round(2)

    channel_summary["demand_pattern"] = channel_summary["SUPER_CHANNEL"].map(
        lambda x: CHANNEL_PLANNING_RULES.get(x, CHANNEL_PLANNING_RULES["Other"])["demand_pattern"]
    )
    channel_summary["forecast_lookback_months"] = channel_summary["SUPER_CHANNEL"].map(
        lambda x: CHANNEL_PLANNING_RULES.get(x, CHANNEL_PLANNING_RULES["Other"])["lookback_months"]
    )
    channel_summary["planning_note"] = channel_summary["SUPER_CHANNEL"].map(
        lambda x: CHANNEL_PLANNING_RULES.get(x, CHANNEL_PLANNING_RULES["Other"])["planning_note"]
    )

    channel_summary = channel_summary.sort_values(
        ["organic_units", "promo_pct_units"],
        ascending=[False, False],
    )

    return channel_summary


def convert_period_columns_for_export(dataframes):
    for df in dataframes:
        if "month" in df.columns:
            df["month"] = df["month"].astype(str)


def save_outputs(
    base_dir: Path,
    cleaned_summary: pd.DataFrame,
    comparison: pd.DataFrame,
    stockout_flags: pd.DataFrame,
    sku_impact: pd.DataFrame,
    channel_summary: pd.DataFrame,
    channel_signal: pd.DataFrame,
):
    convert_period_columns_for_export(
        [cleaned_summary, comparison, stockout_flags, sku_impact, channel_summary, channel_signal]
    )

    output_map = {
        CLEANED_DEMAND_OUT: cleaned_summary,
        RAW_VS_CLEANED_OUT: comparison,
        STOCKOUT_OUT: stockout_flags,
        SKU_IMPACT_OUT: sku_impact,
        CHANNEL_SUMMARY_OUT: channel_summary,
        CHANNEL_SIGNAL_OUT: channel_signal,
    }

    for filename, df in output_map.items():
        path = base_dir / filename
        df.to_csv(path, index=False)
        print(f"Saved: {path.name}")


def print_qa_checks(comparison: pd.DataFrame, sku_impact: pd.DataFrame, channel_summary: pd.DataFrame):
    print("\n=== QA CHECKS ===")
    print(f"Total raw units:           {comparison['raw_units'].sum():,.0f}")
    print(f"Total organic units:       {comparison['organic_units'].sum():,.0f}")
    print(f"Total suppressed units:    {comparison['suppressed_units_est'].sum():,.0f}")
    print(f"Total cleaned units:       {comparison['cleaned_units'].sum():,.0f}")

    print("\n=== TOP 10 PROMO-AFFECTED SKUs ===")
    print(
        sku_impact[
            [
                "ITEMNMBR", "ITEMDESC", "raw_units", "organic_units",
                "promo_units", "promo_pct_units",
            ]
        ]
        .sort_values(["promo_units", "promo_pct_units"], ascending=[False, False])
        .head(10)
        .to_string(index=False)
    )

    print("\n=== TOP 10 STOCKOUT-AFFECTED SKUs ===")
    print(
        sku_impact[
            [
                "ITEMNMBR", "ITEMDESC", "organic_units", "suppressed_units_est",
                "cleaned_units", "stockout_months",
            ]
        ]
        .sort_values(["suppressed_units_est", "stockout_months"], ascending=[False, False])
        .head(10)
        .to_string(index=False)
    )

    print("\n=== TOP 10 CHANNELS BY ORGANIC UNITS ===")
    print(
        channel_summary[
            [
                "SUPER_CHANNEL", "SALESCHANNEL", "organic_units", "promo_pct_units",
                "monthly_cv", "forecast_lookback_months",
            ]
        ]
        .head(10)
        .to_string(index=False)
    )


def parse_args():
    parser = argparse.ArgumentParser(
        description="Clean demand signal by removing promo volume and estimating suppressed demand."
    )
    parser.add_argument("--base-dir", default=str(DEFAULT_BASE_DIR), help="Folder containing the input files.")
    parser.add_argument("--sales-file", default=DEFAULT_SALES_FILE, help="Sales CSV filename or full path.")
    parser.add_argument(
        "--chargeback-file",
        default=DEFAULT_CHARGEBACK_FILE,
        help="Chargeback Excel filename or full path.",
    )
    parser.add_argument(
        "--channel-key-file",
        default=DEFAULT_CHANNEL_KEY_FILE,
        help="Optional channel-key Excel filename or full path.",
    )
    parser.add_argument(
        "--chargeback-sheet",
        default=DEFAULT_CHARGEBACK_SHEET,
        help="Worksheet name inside the chargeback workbook.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    base_dir = Path(args.base_dir).expanduser().resolve()

    sales_file = resolve_input_file(
        base_dir,
        args.sales_file,
        label="Sales file",
        search_terms=FILE_SEARCH_HINTS["sales"],
        required=True,
    )
    chargeback_file = resolve_input_file(
        base_dir,
        args.chargeback_file,
        label="Chargeback file",
        search_terms=FILE_SEARCH_HINTS["chargebacks"],
        required=True,
    )
    channel_key_file = resolve_input_file(
        base_dir,
        args.channel_key_file,
        label="Channel key file",
        search_terms=FILE_SEARCH_HINTS["channel_key"],
        required=False,
    )

    ensure_excel_support([chargeback_file, channel_key_file])

    print(f"Using Python: {sys.executable}")
    print(f"Base directory: {base_dir}")
    print(f"Sales file: {sales_file.name}")
    print(f"Chargeback file: {chargeback_file.name}")
    if channel_key_file is not None:
        print(f"Channel key file: {channel_key_file.name}")

    sales = load_sales(sales_file)
    channel_key = load_channel_key(channel_key_file)
    sales = add_channel_information(sales, channel_key)

    chargebacks = load_chargebacks(chargeback_file, args.chargeback_sheet)
    sales = build_promo_flags(sales, chargebacks)

    channel_signal = build_channel_signal(sales)
    location_monthly = build_location_totals(channel_signal)
    location_signal, stockout_flags = estimate_suppressed_demand(location_monthly)

    cleaned_summary, comparison = build_output_tables(location_signal)
    sku_impact = build_sku_impact_summary(comparison)
    channel_summary = build_channel_summary(channel_signal)

    save_outputs(
        base_dir,
        cleaned_summary,
        comparison,
        stockout_flags,
        sku_impact,
        channel_summary,
        channel_signal,
    )
    print_qa_checks(comparison, sku_impact, channel_summary)

    print("\nDone!")
    print("Next step: run reorder_logic.py to generate reorder alerts and draft PO recommendations.")


if __name__ == "__main__":
    main()
