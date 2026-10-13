"""
parking_analysis.py
-------------------
Core analysis utilities for ParkSmart City Services – Round 2 EDA & Feature Engineering.
All functions are designed to be imported by the notebook and reused.
"""

import os
import pandas as pd
import numpy as np
from pathlib import Path


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_RAW      = PROJECT_ROOT / "data" / "raw"
DATA_PROC     = PROJECT_ROOT / "data" / "processed"
CHARTS_DIR    = PROJECT_ROOT / "outputs" / "charts"
REPORTS_DIR   = PROJECT_ROOT / "outputs" / "reports"

# Source CSVs live in the original dataset folder (never moved / overwritten)
SOURCE_DIR = PROJECT_ROOT.parent / "ParkSmart_Round2_Practice_Dataset"


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_raw_datasets(source_dir: Path = SOURCE_DIR) -> dict[str, pd.DataFrame]:
    """Load all five raw CSV files and return a dict keyed by short name."""
    files = {
        "sessions": "parking_sessions.csv",
        "zones":    "zones.csv",
        "payments": "payments.csv",
        "calendar": "calendar.csv",
        "weather":  "weather.csv",
    }
    datasets: dict[str, pd.DataFrame] = {}
    missing: list[str] = []
    for key, fname in files.items():
        fpath = source_dir / fname
        if fpath.exists():
            datasets[key] = pd.read_csv(fpath)
            print(f"  ✓ Loaded {fname:30s} → {datasets[key].shape[0]:,} rows × {datasets[key].shape[1]} cols")
        else:
            missing.append(fname)
            print(f"  ✗ MISSING: {fname}")
    if missing:
        print(f"\n  WARNING: {len(missing)} file(s) not found: {missing}")
    return datasets


# ---------------------------------------------------------------------------
# Occupancy helper
# ---------------------------------------------------------------------------

def compute_occupancy(sessions_valid: pd.DataFrame,
                      zones: pd.DataFrame,
                      freq: str = "1h") -> pd.DataFrame:
    """
    Estimate hourly occupancy per zone using entry/exit interval overlap.

    Parameters
    ----------
    sessions_valid : rows with valid (non-null, non-negative) entry & exit
    zones          : zones master (zone_id, capacity)
    freq           : time bucket width (default '1h')

    Returns
    -------
    DataFrame with columns: time_bucket, zone_id, occupancy_count, capacity, occupancy_rate
    """
    if sessions_valid.empty:
        return pd.DataFrame()

    min_ts = sessions_valid["entry"].min().floor(freq)
    max_ts = sessions_valid["exit"].max().ceil(freq)
    time_index = pd.date_range(min_ts, max_ts, freq=freq)

    zone_ids = sessions_valid["zone_id"].unique()
    records = []

    for zone_id in zone_ids:
        zone_sess = sessions_valid[sessions_valid["zone_id"] == zone_id]
        cap_row = zones[zones["zone_id"] == zone_id]
        capacity = int(cap_row["capacity"].iloc[0]) if not cap_row.empty else np.nan

        for t in time_index:
            t_end = t + pd.Timedelta(freq)
            # count sessions whose interval overlaps [t, t_end)
            count = int(((zone_sess["entry"] < t_end) & (zone_sess["exit"] >= t)).sum())
            occ_rate = (count / capacity * 100) if (pd.notna(capacity) and capacity > 0) else np.nan
            records.append({
                "time_bucket": t,
                "zone_id": zone_id,
                "occupancy_count": count,
                "capacity": capacity,
                "occupancy_rate": occ_rate,
            })

    occ_df = pd.DataFrame(records)
    return occ_df


# ---------------------------------------------------------------------------
# Thresholds
# ---------------------------------------------------------------------------

def assign_utilization_category(occ_rate: float | None) -> str:
    """
    Classify occupancy rate into Low / Moderate / High.
    Thresholds follow typical urban parking management practice
    (ITDP Parking Management Guide):
      Low      < 50 %
      Moderate 50–85 %
      High     > 85 %
    """
    if occ_rate is None or pd.isna(occ_rate):
        return "Unknown"
    if occ_rate < 50:
        return "Low"
    if occ_rate <= 85:
        return "Moderate"
    return "High"


def assign_time_period(hour: int) -> str:
    """Map hour-of-day (0-23) to named time period."""
    if 0 <= hour < 6:
        return "Night"
    if 6 <= hour < 12:
        return "Morning"
    if 12 <= hour < 18:
        return "Afternoon"
    return "Evening"
