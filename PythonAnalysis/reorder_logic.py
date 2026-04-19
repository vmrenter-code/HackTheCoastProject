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

DEFAULT_CLEANED_DEMAND_FILE = "cleaned_demand_summary.csv"
DEFAULT_CHANNEL_SIGNAL_FILE = "channel_demand_signal.csv"
DEFAULT_INVENTORY_FILE = "POP_InventorySnapshot.xlsx"
DEFAULT_ITEM_SPEC_FILE = "POP_ItemSpecMaster.xlsx"
DEFAULT_PO_FILE = "POP_PurchaseOrderHistory.XLSX"
DEFAULT_VENDOR_FILE = "POP_VendorMaster.xlsx"

REORDER_ALERTS_OUT = "reorder_alerts.csv"
DRAFT_PO_OUT = "draft_po_recommendations.csv"
BUYER_SUMMARY_OUT = "buyer_attention_summary.csv"
MAX_PROJECTED_DATE_OFFSET_DAYS = 3650

FILE_SEARCH_HINTS = {
    "cleaned_demand": ("cleaned", "demand", "summary"),
    "channel_signal": ("channel", "demand", "signal"),
    "inventory": ("inventory", "snapshot"),
    "item_spec": ("item", "spec", "master"),
    "po_history": ("purchase", "order", "history"),
    "vendor": ("vendor", "supplier", "master"),
}

REQUIRED_CLEANED_COLUMNS = [
    "ITEMNMBR",
    "LOCNCODE",
    "month",
    "organic_units",
    "suppressed_units_est",
    "cleaned_units",
]

REQUIRED_CHANNEL_COLUMNS = [
    "ITEMNMBR",
    "LOCNCODE",
    "month",
    "SUPER_CHANNEL",
    "organic_units",
]

LOCATION_NAME_MAP = {
    "1": "SF",
    "2": "NJ",
    "3": "LA",
}

CHANNEL_RULES = {
    "American Market": {
        "lookback_months": 3,
        "service_z": 1.40,
        "review_days": 28,
        "planning_note": "Blend recent trend with seasonal planogram behavior.",
    },
    "Health Food": {
        "lookback_months": 6,
        "service_z": 1.20,
        "review_days": 28,
        "planning_note": "Use smoother trailing history due to distributor lag.",
    },
    "Asian Market": {
        "lookback_months": 3,
        "service_z": 1.65,
        "review_days": 21,
        "planning_note": "Bias upward because demand can spike opportunistically.",
    },
    "eCom": {
        "lookback_months": 6,
        "service_z": 1.05,
        "review_days": 14,
        "planning_note": "Use longer trailing history for long-tail demand.",
    },
    "Other": {
        "lookback_months": 4,
        "service_z": 1.25,
        "review_days": 21,
        "planning_note": "Use blended defaults and review unusual items manually.",
    },
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

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
        print(f"{label} not found. Continuing without it.")
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
        available = ", ".join(workbook.sheet_names)
        raise SystemExit(
            f"{label} is missing sheet {sheet_name!r}. "
            f"Available sheets: {available}"
        ) from exc


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


def extract_numbers(value):
    if pd.isna(value):
        return []
    text = str(value).replace(",", "")
    return [float(match) for match in re.findall(r"\d+(?:\.\d+)?", text)]


def parse_lead_time_days(value):
    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()
    if not text or text == "nan":
        return np.nan

    if "half a year" in text:
        return 182.0

    numbers = extract_numbers(text)
    if not numbers:
        return np.nan

    amount = numbers[0]
    if len(numbers) >= 2 and any(token in text for token in ["-", "~", " to "]):
        amount = float(np.mean(numbers[:2]))

    if "day" in text:
        return float(amount)
    if "week" in text or "wk" in text:
        return float(amount * 7)
    if "month" in text or "mth" in text or "mths" in text or "mon" in text:
        return float(amount * 30.4)
    if "year" in text or "yr" in text:
        return float(amount * 365)

    if amount <= 8:
        return float(amount * 30.4)
    return float(amount)


def parse_case_pack_units(value):
    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()
    numbers = extract_numbers(text)
    if not numbers:
        return np.nan

    if "/" in text or "x" in text:
        product = 1.0
        for number in numbers[:3]:
            product *= number
        if product <= 5000:
            return float(product)

    return float(numbers[-1])


def parse_cases_per_pallet(value):
    numbers = extract_numbers(value)
    if not numbers:
        return np.nan
    return float(numbers[0])


def parse_moq_units(value, case_pack_units, cases_per_pallet):
    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()
    numbers = extract_numbers(text)
    if not numbers:
        return np.nan

    amount = numbers[0]

    if "container" in text or "40hq" in text or "truck load" in text or "truckload" in text:
        return np.nan
    if "lb" in text:
        return np.nan
    if "dz" in text:
        units = amount * 12
    elif "pallet" in text:
        if pd.notna(cases_per_pallet) and pd.notna(case_pack_units):
            units = amount * cases_per_pallet * case_pack_units
        else:
            return np.nan
    elif "case" in text or "cs" in text or "box" in text or "bag" in text:
        if pd.notna(case_pack_units):
            units = amount * case_pack_units
        else:
            units = amount
    else:
        units = amount

    if units <= 0 or units > 500000:
        return np.nan
    return float(units)


def round_up_to_multiple(value, multiple):
    if value <= 0:
        return 0
    if multiple and multiple > 0:
        return int(np.ceil(value / multiple) * multiple)
    return int(np.ceil(value))


def normalize_order_quantity(recommended_units, case_pack_units, moq_units):
    quantity = max(float(recommended_units), 0.0)
    if quantity <= 0:
        return 0

    quantity = round_up_to_multiple(quantity, case_pack_units)

    if pd.notna(moq_units) and moq_units > 0:
        quantity = max(quantity, int(np.ceil(moq_units)))
        quantity = round_up_to_multiple(quantity, case_pack_units)

    return int(quantity)


def get_channel_rule(super_channel):
    return CHANNEL_RULES.get(super_channel, CHANNEL_RULES["Other"])


def safe_project_date(base_date: pd.Timestamp, offset_days):
    if not np.isfinite(offset_days):
        return pd.NaT

    bounded_days = min(int(np.floor(offset_days)), MAX_PROJECTED_DATE_OFFSET_DAYS)
    if bounded_days < 0:
        bounded_days = 0
    return base_date + pd.to_timedelta(bounded_days, unit="D")


def mean_last(series: pd.Series, periods: int) -> float:
    if series.empty:
        return 0.0
    tail = series.tail(periods)
    return float(tail.mean()) if not tail.empty else 0.0


def build_full_month_series(history: pd.DataFrame, value_col: str, latest_month: pd.Period) -> pd.Series:
    if history.empty:
        return pd.Series(dtype=float)

    monthly = history.groupby("month", as_index=True)[value_col].sum().sort_index()
    start_month = monthly.index.min()
    full_index = pd.period_range(start_month, latest_month, freq="M")
    return monthly.reindex(full_index, fill_value=0.0).astype(float)


def forecast_channel(series: pd.Series, super_channel: str) -> float:
    if series.empty:
        return 0.0

    rule = get_channel_rule(super_channel)
    recent_avg = mean_last(series, rule["lookback_months"])
    trailing6_avg = mean_last(series, 6)
    trailing12 = series.tail(12)
    trailing12_median = float(trailing12.median()) if not trailing12.empty else recent_avg
    same_month_last_year = float(series.iloc[-12]) if len(series) >= 12 else np.nan

    if super_channel == "American Market":
        forecast = recent_avg
        if pd.notna(same_month_last_year) and same_month_last_year > 0:
            forecast = 0.65 * recent_avg + 0.35 * same_month_last_year
    elif super_channel == "Health Food":
        forecast = 0.75 * trailing6_avg + 0.25 * trailing12_median
    elif super_channel == "Asian Market":
        forecast = max(mean_last(series, 3), trailing12_median)
    elif super_channel == "eCom":
        forecast = 0.60 * mean_last(series, 3) + 0.40 * trailing6_avg
    else:
        forecast = 0.70 * recent_avg + 0.30 * trailing6_avg

    return float(max(forecast, 0.0))


# ============================================================
# LOADERS
# ============================================================

def load_cleaned_demand(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, low_memory=False)
    ensure_required_columns(df, REQUIRED_CLEANED_COLUMNS, "Cleaned demand file")

    df["ITEMNMBR"] = df["ITEMNMBR"].fillna("").astype(str).str.strip()
    df["LOCNCODE"] = df["LOCNCODE"].fillna("").astype(str).str.strip()
    df["month"] = pd.PeriodIndex(df["month"].astype(str), freq="M")

    for col in ["raw_units", "organic_units", "suppressed_units_est", "cleaned_units"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df


def load_channel_signal(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, low_memory=False)
    ensure_required_columns(df, REQUIRED_CHANNEL_COLUMNS, "Channel demand signal file")

    df["ITEMNMBR"] = df["ITEMNMBR"].fillna("").astype(str).str.strip()
    df["LOCNCODE"] = df["LOCNCODE"].fillna("").astype(str).str.strip()
    df["month"] = pd.PeriodIndex(df["month"].astype(str), freq="M")
    df["SUPER_CHANNEL"] = df["SUPER_CHANNEL"].fillna("Other").astype(str).str.strip()
    df["organic_units"] = pd.to_numeric(df["organic_units"], errors="coerce").fillna(0.0)

    return df


def load_inventory_snapshot(path: Path) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    frames = []

    for sheet_name in workbook.sheet_names:
        df = workbook.parse(sheet_name)
        required = {"Item Number", "Description", "Available", "On Hand"}
        if not required.issubset(df.columns):
            continue

        match = re.search(r"Site\s*(\d+)\s*-\s*([A-Za-z]+)", sheet_name, flags=re.IGNORECASE)
        loc_code = match.group(1) if match else ""

        frame = df[["Item Number", "Description", "Available", "On Hand"]].copy()
        frame["LOCNCODE"] = loc_code
        frame["LOCATION_NAME"] = label_location(loc_code)
        frame = frame.rename(
            columns={
                "Item Number": "ITEMNMBR",
                "Description": "ITEMDESC",
                "Available": "available_units",
                "On Hand": "on_hand_units",
            }
        )

        frame["ITEMNMBR"] = frame["ITEMNMBR"].fillna("").astype(str).str.strip()
        frame["ITEMDESC"] = frame["ITEMDESC"].fillna("").astype(str).str.strip()
        frame["available_units"] = pd.to_numeric(frame["available_units"], errors="coerce").fillna(0.0)
        frame["on_hand_units"] = pd.to_numeric(frame["on_hand_units"], errors="coerce").fillna(0.0)
        frames.append(frame)

    if not frames:
        raise SystemExit("Inventory snapshot did not contain the expected site sheets.")

    inventory = pd.concat(frames, ignore_index=True)
    inventory = inventory.groupby(
        ["ITEMNMBR", "ITEMDESC", "LOCNCODE", "LOCATION_NAME"],
        as_index=False,
    ).agg(
        available_units=("available_units", "sum"),
        on_hand_units=("on_hand_units", "sum"),
    )

    return inventory


def load_item_specs(path: Path) -> pd.DataFrame:
    df = safe_read_excel(path, sheet_name="Item Spec Master", label="Item spec file")

    df = df.rename(
        columns={
            "Item Number": "ITEMNMBR",
            "Description": "ITEMDESC_SPEC",
            "Case Pack": "case_pack_raw",
            "Case/ Pallet": "cases_per_pallet_raw",
            "Lead Time": "lead_time_raw",
            "MOQ": "moq_raw",
            "Maufactuer/ CoPacker": "manufacturer",
        }
    )

    if "ITEMNMBR" not in df.columns:
        raise SystemExit("Item spec file is missing 'Item Number'.")

    df["ITEMNMBR"] = df["ITEMNMBR"].fillna("").astype(str).str.strip()
    df["lead_time_days_spec"] = df["lead_time_raw"].apply(parse_lead_time_days)
    df["case_pack_units"] = df["case_pack_raw"].apply(parse_case_pack_units)
    df["cases_per_pallet"] = df["cases_per_pallet_raw"].apply(parse_cases_per_pallet)
    df["moq_units"] = df.apply(
        lambda row: parse_moq_units(row.get("moq_raw"), row.get("case_pack_units"), row.get("cases_per_pallet")),
        axis=1,
    )

    keep_cols = [
        "ITEMNMBR",
        "ITEMDESC_SPEC",
        "manufacturer",
        "lead_time_raw",
        "lead_time_days_spec",
        "case_pack_raw",
        "case_pack_units",
        "cases_per_pallet_raw",
        "cases_per_pallet",
        "moq_raw",
        "moq_units",
    ]
    keep_cols = [col for col in keep_cols if col in df.columns]
    return df[keep_cols].drop_duplicates(subset=["ITEMNMBR"])


def load_po_history(path: Path):
    po = safe_read_excel(path, sheet_name="PO Order History 2023-2025", label="PO history file")
    ensure_required_columns(
        po,
        ["PO Number", "PO Date", "Receipt Date", "Item Number", "Vendor ID", "Location Code", "Unit Cost"],
        "PO history file",
    )

    po = po.rename(
        columns={
            "Item Number": "ITEMNMBR",
            "Vendor ID": "vendor_id",
            "Location Code": "LOCNCODE",
            "Unit Cost": "unit_cost",
        }
    )
    po["ITEMNMBR"] = po["ITEMNMBR"].fillna("").astype(str).str.strip()
    po["vendor_id"] = po["vendor_id"].fillna("").astype(str).str.strip()
    po["LOCNCODE"] = po["LOCNCODE"].fillna("").astype(str).str.strip()
    po["PO Date"] = pd.to_datetime(po["PO Date"], errors="coerce")
    po["Receipt Date"] = pd.to_datetime(po["Receipt Date"], errors="coerce")
    po["unit_cost"] = pd.to_numeric(po["unit_cost"], errors="coerce")
    po["actual_lead_days"] = (po["Receipt Date"] - po["PO Date"]).dt.days

    po_valid = po[(po["ITEMNMBR"] != "") & (po["vendor_id"] != "")]

    item_loc_vendor = (
        po_valid.groupby(["ITEMNMBR", "LOCNCODE", "vendor_id"], as_index=False)
        .agg(
            po_lines=("PO Number", "count"),
            median_lead_days=("actual_lead_days", "median"),
            last_po_date=("PO Date", "max"),
            median_unit_cost=("unit_cost", "median"),
        )
        .sort_values(
            ["ITEMNMBR", "LOCNCODE", "po_lines", "last_po_date"],
            ascending=[True, True, False, False],
        )
    )
    item_loc_primary = item_loc_vendor.groupby(["ITEMNMBR", "LOCNCODE"], as_index=False).first()

    item_vendor = (
        po_valid.groupby(["ITEMNMBR", "vendor_id"], as_index=False)
        .agg(
            po_lines=("PO Number", "count"),
            median_lead_days=("actual_lead_days", "median"),
            last_po_date=("PO Date", "max"),
            median_unit_cost=("unit_cost", "median"),
        )
        .sort_values(["ITEMNMBR", "po_lines", "last_po_date"], ascending=[True, False, False])
    )
    item_primary = item_vendor.groupby("ITEMNMBR", as_index=False).first()

    return item_loc_primary, item_primary


def load_vendor_master(path: Optional[Path]) -> Optional[pd.DataFrame]:
    if path is None:
        return None

    df = safe_read_excel(path, sheet_name="Supplier Master", label="Vendor master file")
    if "Vendor ID" not in df.columns:
        print("Warning: vendor master file is missing 'Vendor ID'. Skipping vendor enrichment.")
        return None

    df = df.rename(columns={"Vendor ID": "vendor_id"})
    df["vendor_id"] = df["vendor_id"].fillna("").astype(str).str.strip()

    keep_cols = [
        "vendor_id", "Brand", "Product Line", "Category",
        "Vendor Status", "Country", "Shipment Terms", "Payment Terms",
    ]
    keep_cols = [col for col in keep_cols if col in df.columns]
    return df[keep_cols].drop_duplicates(subset=["vendor_id"])


# ============================================================
# REORDER LOGIC
# ============================================================

def select_lead_time_days(spec_row, po_loc_row, po_item_row):
    if po_loc_row is not None and pd.notna(po_loc_row.get("median_lead_days")) and po_loc_row.get("po_lines", 0) >= 3:
        return float(po_loc_row["median_lead_days"]), "po_history_item_location"

    if spec_row is not None and pd.notna(spec_row.get("lead_time_days_spec")) and spec_row["lead_time_days_spec"] > 0:
        return float(spec_row["lead_time_days_spec"]), "item_spec"

    if po_item_row is not None and pd.notna(po_item_row.get("median_lead_days")) and po_item_row.get("po_lines", 0) >= 3:
        return float(po_item_row["median_lead_days"]), "po_history_item"

    if po_loc_row is not None and pd.notna(po_loc_row.get("median_lead_days")) and po_loc_row["median_lead_days"] > 0:
        return float(po_loc_row["median_lead_days"]), "po_history_item_location_low_sample"

    if po_item_row is not None and pd.notna(po_item_row.get("median_lead_days")) and po_item_row["median_lead_days"] > 0:
        return float(po_item_row["median_lead_days"]), "po_history_item_low_sample"

    return 90.0, "default_90d"


def build_channel_breakdown(channel_history: pd.DataFrame, latest_month: pd.Period):
    breakdown = []

    for super_channel, subset in channel_history.groupby("SUPER_CHANNEL"):
        series = build_full_month_series(subset, "organic_units", latest_month)
        forecast_units = forecast_channel(series, super_channel)
        recent_units = float(series.tail(12).sum()) if not series.empty else 0.0
        breakdown.append(
            {
                "super_channel": super_channel,
                "forecast_units_monthly": forecast_units,
                "recent_12m_units": recent_units,
            }
        )

    return breakdown


def format_channel_breakdown(channel_breakdown):
    if not channel_breakdown:
        return ""

    ordered = sorted(channel_breakdown, key=lambda row: row["forecast_units_monthly"], reverse=True)
    return "; ".join(
        f"{row['super_channel']}:{row['forecast_units_monthly']:.1f}"
        for row in ordered if row["forecast_units_monthly"] > 0
    )


def build_reorder_alerts(
    inventory: pd.DataFrame,
    cleaned_demand: pd.DataFrame,
    channel_signal: pd.DataFrame,
    item_specs: pd.DataFrame,
    po_item_loc: pd.DataFrame,
    po_item: pd.DataFrame,
    vendor_master: Optional[pd.DataFrame],
    as_of_date: pd.Timestamp,
):
    latest_month = cleaned_demand["month"].max()

    spec_map = item_specs.set_index("ITEMNMBR").to_dict("index")
    po_loc_map = po_item_loc.set_index(["ITEMNMBR", "LOCNCODE"]).to_dict("index")
    po_item_map = po_item.set_index("ITEMNMBR").to_dict("index")
    vendor_map = vendor_master.set_index("vendor_id").to_dict("index") if vendor_master is not None else {}

    alert_rows = []

    for row in inventory.itertuples(index=False):
        item = row.ITEMNMBR
        loc = str(row.LOCNCODE).strip()
        location_name = row.LOCATION_NAME

        demand_hist = cleaned_demand[
            (cleaned_demand["ITEMNMBR"] == item) &
            (cleaned_demand["LOCNCODE"] == loc)
        ].copy()
        channel_hist = channel_signal[
            (channel_signal["ITEMNMBR"] == item) &
            (channel_signal["LOCNCODE"] == loc)
        ].copy()

        cleaned_series = build_full_month_series(demand_hist, "cleaned_units", latest_month)
        suppressed_series = build_full_month_series(demand_hist, "suppressed_units_est", latest_month)

        channel_breakdown = build_channel_breakdown(channel_hist, latest_month)
        channel_forecast_units = sum(entry["forecast_units_monthly"] for entry in channel_breakdown)
        suppressed_recent_avg = mean_last(suppressed_series, 6)
        recent_cleaned_avg = mean_last(cleaned_series, 3)

        if channel_forecast_units > 0:
            forecast_units_monthly = channel_forecast_units + suppressed_recent_avg
        else:
            forecast_units_monthly = max(recent_cleaned_avg, mean_last(cleaned_series, 6))

        forecast_units_monthly = float(max(forecast_units_monthly, 0.0))
        daily_forecast_units = forecast_units_monthly / 30.4 if forecast_units_monthly > 0 else 0.0

        recent_12m_suppressed = float(suppressed_series.tail(12).sum()) if not suppressed_series.empty else 0.0
        stockout_months_last12 = int(
            demand_hist.sort_values("month").tail(12)["likely_stockout"].fillna(False).sum()
        ) if "likely_stockout" in demand_hist.columns else 0

        if channel_breakdown:
            dominant_channel_row = max(channel_breakdown, key=lambda entry: entry["recent_12m_units"])
            dominant_super_channel = dominant_channel_row["super_channel"]
        else:
            dominant_super_channel = "Other"

        channel_rule = get_channel_rule(dominant_super_channel)
        demand_std_units = float(cleaned_series.tail(6).std(ddof=0)) if len(cleaned_series) > 1 else 0.0

        spec_row = spec_map.get(item)
        po_loc_row = po_loc_map.get((item, loc))
        po_item_row = po_item_map.get(item)
        manufacturer = spec_row.get("manufacturer", "") if spec_row is not None else ""

        lead_time_days, lead_time_source = select_lead_time_days(spec_row, po_loc_row, po_item_row)
        service_z = channel_rule["service_z"]
        review_days = channel_rule["review_days"]

        safety_stock_units = service_z * demand_std_units * np.sqrt(max(lead_time_days, 1) / 30.4)
        reorder_point_units = daily_forecast_units * lead_time_days + safety_stock_units
        target_cover_days = lead_time_days + review_days + max(7, int(round(lead_time_days * 0.2)))
        target_inventory_units = daily_forecast_units * target_cover_days + safety_stock_units

        available_units = float(row.available_units)
        on_hand_units = float(row.on_hand_units)
        inventory_gap_units = reorder_point_units - available_units
        recommended_order_units = max(target_inventory_units - available_units, 0.0)

        case_pack_units = spec_row.get("case_pack_units") if spec_row is not None else np.nan
        moq_units = spec_row.get("moq_units") if spec_row is not None else np.nan
        recommended_order_units_rounded = normalize_order_quantity(recommended_order_units, case_pack_units, moq_units)
        recommended_order_cases = (
            recommended_order_units_rounded / case_pack_units
            if pd.notna(case_pack_units) and case_pack_units > 0 else np.nan
        )

        vendor_id = ""
        unit_cost_est = np.nan
        if po_loc_row is not None:
            vendor_id = str(po_loc_row.get("vendor_id", "")).strip()
            unit_cost_est = po_loc_row.get("median_unit_cost", np.nan)
        elif po_item_row is not None:
            vendor_id = str(po_item_row.get("vendor_id", "")).strip()
            unit_cost_est = po_item_row.get("median_unit_cost", np.nan)

        estimated_order_cost = (
            recommended_order_units_rounded * unit_cost_est
            if pd.notna(unit_cost_est) and recommended_order_units_rounded > 0 else np.nan
        )

        if daily_forecast_units > 0:
            days_of_supply = available_units / daily_forecast_units
            days_until_reorder = max((available_units - reorder_point_units) / daily_forecast_units, 0.0)
            projected_runout_date = safe_project_date(as_of_date, days_of_supply)
        else:
            days_of_supply = np.inf
            days_until_reorder = np.inf
            projected_runout_date = pd.NaT

        if forecast_units_monthly <= 0.1:
            priority = "Low"
            alert_status = "No recent demand"
        elif available_units <= 0:
            priority = "Critical"
            alert_status = "Order now"
        elif available_units <= reorder_point_units:
            priority = "High"
            alert_status = "Order now"
        elif days_of_supply <= lead_time_days + 7:
            priority = "Medium"
            alert_status = "Watch"
        else:
            priority = "Low"
            alert_status = "Healthy"

        if priority in {"Critical", "High"}:
            suggested_order_date = as_of_date
        elif np.isfinite(days_until_reorder):
            suggested_order_date = safe_project_date(as_of_date, days_until_reorder)
        else:
            suggested_order_date = pd.NaT

        notes = []
        if lead_time_source == "default_90d":
            notes.append("Lead time fallback used")
        if recent_12m_suppressed > 0:
            notes.append("Suppressed demand uplift included")
        if not vendor_id:
            notes.append("Primary vendor missing")
        if forecast_units_monthly <= 0.1:
            notes.append("Demand history too light for standard reorder rule")

        vendor_attrs = vendor_map.get(vendor_id, {})
        alert_rows.append(
            {
                "ITEMNMBR": item,
                "ITEMDESC": row.ITEMDESC,
                "LOCNCODE": loc,
                "LOCATION_NAME": location_name,
                "available_units": round(available_units, 1),
                "on_hand_units": round(on_hand_units, 1),
                "forecast_units_monthly": round(forecast_units_monthly, 1),
                "daily_forecast_units": round(daily_forecast_units, 2),
                "recent_3m_cleaned_avg": round(recent_cleaned_avg, 1),
                "recent_12m_suppressed_units": round(recent_12m_suppressed, 1),
                "stockout_months_last12": stockout_months_last12,
                "dominant_super_channel": dominant_super_channel,
                "channel_forecast_breakdown": format_channel_breakdown(channel_breakdown),
                "lead_time_days": round(lead_time_days, 1),
                "lead_time_source": lead_time_source,
                "service_z": service_z,
                "safety_stock_units": round(float(safety_stock_units), 1),
                "reorder_point_units": round(float(reorder_point_units), 1),
                "target_inventory_units": round(float(target_inventory_units), 1),
                "inventory_gap_units": round(float(inventory_gap_units), 1),
                "days_of_supply": round(float(days_of_supply), 1) if np.isfinite(days_of_supply) else np.nan,
                "projected_runout_date": projected_runout_date.date().isoformat() if pd.notna(projected_runout_date) else "",
                "alert_status": alert_status,
                "priority": priority,
                "suggested_order_date": suggested_order_date.date().isoformat() if pd.notna(suggested_order_date) else "",
                "recommended_order_units": round(float(recommended_order_units), 1),
                "recommended_order_units_rounded": recommended_order_units_rounded,
                "recommended_order_cases": round(float(recommended_order_cases), 1) if pd.notna(recommended_order_cases) else np.nan,
                "case_pack_units": round(float(case_pack_units), 1) if pd.notna(case_pack_units) else np.nan,
                "moq_units": round(float(moq_units), 1) if pd.notna(moq_units) else np.nan,
                "vendor_id": vendor_id,
                "manufacturer": manufacturer,
                "estimated_unit_cost": round(float(unit_cost_est), 2) if pd.notna(unit_cost_est) else np.nan,
                "estimated_order_cost": round(float(estimated_order_cost), 2) if pd.notna(estimated_order_cost) else np.nan,
                "vendor_brand": vendor_attrs.get("Brand", ""),
                "vendor_product_line": vendor_attrs.get("Product Line", ""),
                "vendor_category": vendor_attrs.get("Category", ""),
                "vendor_status": vendor_attrs.get("Vendor Status", ""),
                "shipment_terms": vendor_attrs.get("Shipment Terms", ""),
                "planning_note": channel_rule["planning_note"],
                "assumption_notes": "; ".join(notes),
            }
        )

    alerts = pd.DataFrame(alert_rows)
    priority_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    alerts["priority_rank"] = alerts["priority"].map(priority_order).fillna(9)
    alerts = alerts.sort_values(
        ["priority_rank", "inventory_gap_units", "forecast_units_monthly"],
        ascending=[True, False, False],
    ).drop(columns=["priority_rank"])

    return alerts


def build_draft_po_recommendations(alerts: pd.DataFrame) -> pd.DataFrame:
    po_recs = alerts[
        (alerts["recommended_order_units_rounded"] > 0) &
        (alerts["priority"].isin(["Critical", "High", "Medium"]))
    ].copy()

    po_recs = po_recs[
        [
            "priority", "suggested_order_date", "LOCNCODE", "LOCATION_NAME",
            "vendor_id", "manufacturer", "vendor_brand", "ITEMNMBR", "ITEMDESC",
            "forecast_units_monthly", "lead_time_days", "days_of_supply",
            "recommended_order_units_rounded", "recommended_order_cases",
            "case_pack_units", "moq_units", "estimated_unit_cost",
            "estimated_order_cost", "assumption_notes",
        ]
    ].copy()

    po_recs = po_recs.sort_values(
        ["priority", "suggested_order_date", "vendor_id", "LOCATION_NAME", "ITEMNMBR"],
        ascending=[True, True, True, True, True],
    )
    return po_recs


def build_buyer_summary(alerts: pd.DataFrame) -> pd.DataFrame:
    summary_source = alerts.copy()
    actionable_mask = summary_source["priority"].isin(["Critical", "High", "Medium"])
    summary_source["actionable_units"] = np.where(
        actionable_mask,
        summary_source["recommended_order_units_rounded"],
        0,
    )
    summary_source["actionable_cost"] = np.where(
        actionable_mask,
        summary_source["estimated_order_cost"].fillna(0),
        0,
    )

    summary = (
        summary_source.groupby(["LOCATION_NAME", "priority"], as_index=False)
        .agg(
            sku_count=("ITEMNMBR", "nunique"),
            action_lines=("actionable_units", lambda s: int((s > 0).sum())),
            total_recommended_units=("actionable_units", "sum"),
            total_estimated_order_cost=("actionable_cost", "sum"),
        )
    )

    priority_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    summary["priority_rank"] = summary["priority"].map(priority_order).fillna(9)
    summary = summary.sort_values(["LOCATION_NAME", "priority_rank"]).drop(columns=["priority_rank"])
    return summary


def save_outputs(base_dir: Path, alerts: pd.DataFrame, po_recs: pd.DataFrame, buyer_summary: pd.DataFrame):
    output_map = {
        REORDER_ALERTS_OUT: alerts,
        DRAFT_PO_OUT: po_recs,
        BUYER_SUMMARY_OUT: buyer_summary,
    }

    for filename, df in output_map.items():
        path = base_dir / filename
        df.to_csv(path, index=False)
        print(f"Saved: {path.name}")


def print_qa_checks(alerts: pd.DataFrame, po_recs: pd.DataFrame, buyer_summary: pd.DataFrame, as_of_date: pd.Timestamp):
    print("\n=== QA CHECKS ===")
    print(f"As-of date:                 {as_of_date.date().isoformat()}")
    print(f"Alert rows:                 {len(alerts):,}")
    print(f"Critical / High rows:       {len(alerts[alerts['priority'].isin(['Critical', 'High'])]):,}")
    print(f"Draft PO lines:             {len(po_recs):,}")
    print(f"Estimated draft PO cost:    ${po_recs['estimated_order_cost'].fillna(0).sum():,.0f}")

    print("\n=== TOP 10 ACTION ITEMS ===")
    print(
        alerts[
            [
                "priority", "LOCATION_NAME", "ITEMNMBR", "available_units",
                "forecast_units_monthly", "reorder_point_units",
                "recommended_order_units_rounded", "vendor_id",
            ]
        ]
        .head(10)
        .to_string(index=False)
    )

    print("\n=== BUYER SUMMARY ===")
    print(buyer_summary.to_string(index=False))


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate reorder alerts and draft PO recommendations from cleaned demand and inventory."
    )
    parser.add_argument("--base-dir", default=str(DEFAULT_BASE_DIR), help="Folder containing the input files.")
    parser.add_argument(
        "--cleaned-demand-file",
        default=DEFAULT_CLEANED_DEMAND_FILE,
        help="cleaned_demand_summary.csv filename or full path.",
    )
    parser.add_argument(
        "--channel-signal-file",
        default=DEFAULT_CHANNEL_SIGNAL_FILE,
        help="channel_demand_signal.csv filename or full path.",
    )
    parser.add_argument(
        "--inventory-file",
        default=DEFAULT_INVENTORY_FILE,
        help="Inventory snapshot Excel filename or full path.",
    )
    parser.add_argument(
        "--item-spec-file",
        default=DEFAULT_ITEM_SPEC_FILE,
        help="Item spec master Excel filename or full path.",
    )
    parser.add_argument(
        "--po-file",
        default=DEFAULT_PO_FILE,
        help="Purchase order history Excel filename or full path.",
    )
    parser.add_argument(
        "--vendor-file",
        default=DEFAULT_VENDOR_FILE,
        help="Optional vendor master Excel filename or full path.",
    )
    parser.add_argument(
        "--as-of-date",
        default=None,
        help="Override the planning date in YYYY-MM-DD format. Defaults to today.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    base_dir = Path(args.base_dir).expanduser().resolve()

    cleaned_demand_file = resolve_input_file(
        base_dir,
        args.cleaned_demand_file,
        label="Cleaned demand file",
        search_terms=FILE_SEARCH_HINTS["cleaned_demand"],
        required=True,
    )
    channel_signal_file = resolve_input_file(
        base_dir,
        args.channel_signal_file,
        label="Channel signal file",
        search_terms=FILE_SEARCH_HINTS["channel_signal"],
        required=True,
    )
    inventory_file = resolve_input_file(
        base_dir,
        args.inventory_file,
        label="Inventory snapshot file",
        search_terms=FILE_SEARCH_HINTS["inventory"],
        required=True,
    )
    item_spec_file = resolve_input_file(
        base_dir,
        args.item_spec_file,
        label="Item spec file",
        search_terms=FILE_SEARCH_HINTS["item_spec"],
        required=True,
    )
    po_file = resolve_input_file(
        base_dir,
        args.po_file,
        label="PO history file",
        search_terms=FILE_SEARCH_HINTS["po_history"],
        required=True,
    )
    vendor_file = resolve_input_file(
        base_dir,
        args.vendor_file,
        label="Vendor master file",
        search_terms=FILE_SEARCH_HINTS["vendor"],
        required=False,
    )

    ensure_excel_support([inventory_file, item_spec_file, po_file, vendor_file])

    as_of_date = pd.Timestamp(args.as_of_date).normalize() if args.as_of_date else pd.Timestamp.today().normalize()

    print(f"Using Python: {sys.executable}")
    print(f"Base directory: {base_dir}")
    print(f"As-of date: {as_of_date.date().isoformat()}")
    print(f"Cleaned demand file: {cleaned_demand_file.name}")
    print(f"Inventory file: {inventory_file.name}")
    print(f"Item spec file: {item_spec_file.name}")
    print(f"PO history file: {po_file.name}")
    if vendor_file is not None:
        print(f"Vendor master file: {vendor_file.name}")

    cleaned_demand = load_cleaned_demand(cleaned_demand_file)
    channel_signal = load_channel_signal(channel_signal_file)
    inventory = load_inventory_snapshot(inventory_file)
    item_specs = load_item_specs(item_spec_file)
    po_item_loc, po_item = load_po_history(po_file)
    vendor_master = load_vendor_master(vendor_file)

    alerts = build_reorder_alerts(
        inventory,
        cleaned_demand,
        channel_signal,
        item_specs,
        po_item_loc,
        po_item,
        vendor_master,
        as_of_date,
    )
    po_recs = build_draft_po_recommendations(alerts)
    buyer_summary = build_buyer_summary(alerts)

    save_outputs(base_dir, alerts, po_recs, buyer_summary)
    print_qa_checks(alerts, po_recs, buyer_summary, as_of_date)

    print("\nDone!")


if __name__ == "__main__":
    main()
