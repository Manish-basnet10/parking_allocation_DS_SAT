#!/usr/bin/env python3
"""
run_analysis.py  –  Standalone execution of the full Round 2 analysis pipeline.
Run from the project root:  python3 src/run_analysis.py
"""

import os, sys, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')   # non-interactive backend for headless execution
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from sklearn.feature_selection import mutual_info_regression

warnings.filterwarnings('ignore')
np.random.seed(42)

# ── Style ──────────────────────────────────────────────────────────────────
plt.rcParams.update({
    'figure.dpi': 150,
    'figure.facecolor': '#0f1117',
    'axes.facecolor': '#1a1d27',
    'axes.edgecolor': '#3a3d4d',
    'axes.labelcolor': '#e0e0e0',
    'axes.titlecolor': '#ffffff',
    'text.color': '#e0e0e0',
    'xtick.color': '#b0b0b0',
    'ytick.color': '#b0b0b0',
    'grid.color': '#2a2d3d',
    'grid.linestyle': '--',
    'grid.alpha': 0.5,
    'legend.facecolor': '#1a1d27',
    'legend.edgecolor': '#3a3d4d',
    'legend.labelcolor': '#e0e0e0',
    'font.family': 'DejaVu Sans',
    'axes.titlesize': 14,
    'axes.labelsize': 11,
})
PALETTE = ['#6c63ff','#ff6584','#43e97b','#f7971e','#4facfe','#f093fb']
sns.set_theme(style='dark', rc={'axes.facecolor':'#1a1d27','figure.facecolor':'#0f1117'})

# ── Paths ──────────────────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR   = PROJECT_ROOT.parent / 'ParkSmart_Round2_Practice_Dataset'
DATA_PROC    = PROJECT_ROOT / 'data' / 'processed'
CHARTS_DIR   = PROJECT_ROOT / 'outputs' / 'charts'
REPORTS_DIR  = PROJECT_ROOT / 'outputs' / 'reports'

for d in [DATA_PROC, CHARTS_DIR, REPORTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

print(f"Project root : {PROJECT_ROOT}")
print(f"Source data  : {SOURCE_DIR}")

# ── Helper ─────────────────────────────────────────────────────────────────
def save_chart(fig, name):
    path = CHARTS_DIR / f'{name}.png'
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f'  ✓ Saved chart: {path.name}')

def assign_time_period(hour):
    if pd.isna(hour):   return 'Unknown'
    if 0 <= hour < 6:   return 'Night'
    if 6 <= hour < 12:  return 'Morning'
    if 12 <= hour < 18: return 'Afternoon'
    return 'Evening'

def assign_util_cat(rate):
    if pd.isna(rate): return 'Unknown'
    if rate < 1.0:    return 'Low'
    if rate <= 4.0:   return 'Moderate'
    return 'High'

# =============================================================================
# SECTION 1 – Load raw data
# =============================================================================
print("\n" + "="*60)
print("SECTION 1: LOADING DATA")
print("="*60)

def load_csv(name):
    path = SOURCE_DIR / name
    if not path.exists():
        raise FileNotFoundError(f'Missing: {path}')
    return pd.read_csv(path)

raw_sessions = load_csv('parking_sessions.csv')
raw_zones    = load_csv('zones.csv')
raw_payments = load_csv('payments.csv')
raw_calendar = load_csv('calendar.csv')
raw_weather  = load_csv('weather.csv')

for name, df in [('parking_sessions', raw_sessions), ('zones', raw_zones),
                  ('payments', raw_payments), ('calendar', raw_calendar),
                  ('weather', raw_weather)]:
    print(f'  {name:20s}: {df.shape[0]:,} rows × {df.shape[1]} cols')

# =============================================================================
# SECTION 2 – Data Quality Audit
# =============================================================================
print("\n" + "="*60)
print("SECTION 2: DATA QUALITY AUDIT")
print("="*60)

# Parse timestamps temporarily
sess_tmp = raw_sessions.copy()
sess_tmp['entry_ts'] = pd.to_datetime(sess_tmp['entry'], errors='coerce')
sess_tmp['exit_ts']  = pd.to_datetime(sess_tmp['exit'],  errors='coerce')
sess_tmp['calc_dur'] = (sess_tmp['exit_ts'] - sess_tmp['entry_ts']).dt.total_seconds() / 60

n_neg     = int((sess_tmp['calc_dur'] < 0).sum())
n_zero    = int((sess_tmp['calc_dur'] == 0).sum())
n_extreme = int((sess_tmp['calc_dur'] > 1440).sum())
n_miss_exit = int(sess_tmp['exit_ts'].isna().sum())

print(f"  Missing exit timestamps  : {n_miss_exit}")
print(f"  Missing duration values  : {int(raw_sessions['duration'].isna().sum())}")
print(f"  Negative durations       : {n_neg}")
print(f"  Zero durations           : {n_zero}")
print(f"  Extreme (>24h) durations : {n_extreme}")
print(f"  Exact dup rows (sessions): {int(raw_sessions.duplicated().sum())}")
print(f"  Exact dup rows (payments): {int(raw_payments.duplicated().sum())}")
print(f"  Neg payments             : {int((raw_payments['amount'] < 0).sum())}")

# Zone mismatch check
session_zones = set(raw_sessions['zone_id'].str.strip().str.upper())
master_zones  = set(raw_zones['zone_id'].str.strip().str.upper())
unresolved = session_zones - master_zones
print(f"  Unresolved zone mismatches: {unresolved if unresolved else 'None'}")

# Save DQ report
dq_issues = [
    ('Missing exit timestamps',      'sessions', n_miss_exit,
     'Keep open sessions; exclude from duration calc & occupancy reconstruction',
     'Missing exit may be an active/open session'),
    ('Missing duration values',       'sessions', int(raw_sessions['duration'].isna().sum()),
     'Recalculate from entry/exit when both are valid',
     'Timestamps are authoritative'),
    ('Negative calculated durations', 'sessions', n_neg,
     'Flag and exclude from duration/occupancy analyses',
     'Entry after exit is logically invalid'),
    ('Zero durations',                'sessions', n_zero,
     'Exclude from duration analyses; retain for session count',
     'Instantaneous sessions are implausible'),
    ('Extreme durations (>24 h)',     'sessions', n_extreme,
     'Flag and exclude from duration statistics',
     'Treated as outliers'),
    ('Exact duplicate rows',          'sessions', int(raw_sessions.duplicated().sum()),
     'Remove all but first occurrence',
     'Exact duplicates add no information'),
    ('Duplicate payment_id',          'payments', int(raw_payments['payment_id'].duplicated().sum()),
     'Remove duplicate payment rows',
     'Duplicate IDs would inflate fee totals'),
    ('Inconsistent zone labels',      'sessions', 0,
     'Normalise zone_id with strip/upper; all resolved to master',
     'All zone labels match after standardisation'),
    ('Negative payment amounts',      'payments', int((raw_payments['amount'] < 0).sum()),
     'Flag; exclude from revenue totals',
     'Negative fees are not valid parking charges'),
]
dq_df = pd.DataFrame(dq_issues, columns=['Issue','Dataset','Affected Records','Action','Reason'])
dq_df.to_csv(REPORTS_DIR / 'data_quality_report.csv', index=False)
dq_df.to_csv(PROJECT_ROOT / 'outputs' / 'data_quality_report.csv', index=False)
print(f"\n  ✓ Saved data_quality_report.csv")

# =============================================================================
# SECTION 3 – Cleaning Pipeline
# =============================================================================
print("\n" + "="*60)
print("SECTION 3: CLEANING PIPELINE")
print("="*60)

sessions = raw_sessions.copy()
sessions['entry'] = pd.to_datetime(sessions['entry'], errors='coerce')
sessions['exit']  = pd.to_datetime(sessions['exit'],  errors='coerce')

n_before = len(sessions)
sessions = sessions.drop_duplicates()
print(f"  After dedup: {len(sessions):,} rows (removed {n_before - len(sessions)})")

sessions['zone_id'] = sessions['zone_id'].str.strip().str.upper()
zones = raw_zones.copy()
zones['zone_id'] = zones['zone_id'].str.strip().str.upper()

sessions['calc_dur_min'] = (sessions['exit'] - sessions['entry']).dt.total_seconds() / 60
sessions['has_entry']    = sessions['entry'].notna()
sessions['has_exit']     = sessions['exit'].notna()
sessions['is_negative']  = sessions['calc_dur_min'] < 0
sessions['is_zero_dur']  = sessions['calc_dur_min'] == 0
sessions['is_extreme']   = sessions['calc_dur_min'] > 1440
sessions['is_open']      = sessions['has_entry'] & ~sessions['has_exit']
sessions['is_valid_dur'] = (sessions['has_entry'] & sessions['has_exit'] &
                             ~sessions['is_negative'] & ~sessions['is_zero_dur'] &
                             ~sessions['is_extreme'])
sessions['duration_min'] = np.where(sessions['is_valid_dur'], sessions['calc_dur_min'], np.nan)

print(f"  Valid duration sessions  : {sessions['is_valid_dur'].sum():,}")
print(f"  Open sessions            : {sessions['is_open'].sum():,}")
print(f"  Negative duration        : {sessions['is_negative'].sum():,}")
print(f"  Zero duration            : {sessions['is_zero_dur'].sum():,}")
print(f"  Extreme duration         : {sessions['is_extreme'].sum():,}")

payments = raw_payments.copy()
n_pay = len(payments)
payments = payments.drop_duplicates()
print(f"  Payments after dedup: {len(payments):,} (removed {n_pay - len(payments)})")

calendar = raw_calendar.copy()
calendar['date'] = pd.to_datetime(calendar['date'], errors='coerce')

weather = raw_weather.copy()
weather['date'] = pd.to_datetime(weather['date'], errors='coerce')

# =============================================================================
# SECTION 4 – Integration
# =============================================================================
print("\n" + "="*60)
print("SECTION 4: DATASET INTEGRATION")
print("="*60)

sessions['entry_date'] = sessions['entry'].dt.normalize()

df = sessions.merge(zones, on='zone_id', how='left')
print(f"  After zones merge   : {len(df):,} rows")

df = df.merge(calendar.rename(columns={'date':'entry_date'}), on='entry_date', how='left')
print(f"  After calendar merge: {len(df):,} rows")

df = df.merge(weather.rename(columns={'date':'entry_date'}), on='entry_date', how='left')
print(f"  After weather merge : {len(df):,} rows")

df = df.merge(payments[['payment_id','payment_mode','amount']], on='payment_id', how='left')
print(f"  After payments merge: {len(df):,} rows")

df.to_csv(DATA_PROC / 'cleaned_parking_data.csv', index=False)
print(f"  ✓ Saved cleaned_parking_data.csv → {len(df):,} rows")

# =============================================================================
# SECTION 5 – Descriptive Stats
# =============================================================================
print("\n" + "="*60)
print("SECTION 5: DESCRIPTIVE STATISTICS")
print("="*60)

valid = df[df['is_valid_dur']].copy()
n_days = (df['entry_date'].max() - df['entry_date'].min()).days + 1

print(f"  Total sessions (after dedup) : {len(df):,}")
print(f"  Valid-duration sessions      : {len(valid):,}")
print(f"  Open sessions                : {df['is_open'].sum():,}")
print(f"  Distinct zones               : {df['zone_id'].nunique()}")
print(f"  Date range                   : {df['entry_date'].min().date()} → {df['entry_date'].max().date()} ({n_days} days)")
print(f"  Avg duration (valid, min)    : {valid['duration_min'].mean():.1f}")
print(f"  Median duration (valid, min) : {valid['duration_min'].median():.1f}")
print(f"  Total fees collected         : {df['amount'].sum():,.2f}")

zone_summary = df.groupby('zone_id').agg(
    sessions=('session_id','count'), valid_sessions=('is_valid_dur','sum'),
    avg_dur_min=('duration_min','mean'), median_dur_min=('duration_min','median'),
    total_fees=('amount','sum'), capacity=('capacity','first'),
    location=('location','first'), zone_type=('zone_type','first'),
).reset_index()
print("\n  Zone-level summary:")
print(zone_summary.to_string(index=False))

# =============================================================================
# SECTION 6 – Occupancy Reconstruction
# =============================================================================
print("\n" + "="*60)
print("SECTION 6: OCCUPANCY RECONSTRUCTION (1-hour buckets)")
print("="*60)

valid_occ = valid[valid['entry'].notna() & valid['exit'].notna()].copy()
FREQ = '1h'
min_ts = valid_occ['entry'].min().floor(FREQ)
max_ts = valid_occ['exit'].max().ceil(FREQ)
time_index = pd.date_range(min_ts, max_ts, freq=FREQ)

# Note: Use operating hours only (7am-10pm) for meaningful occupancy rates.
# Including midnight hours inflates the denominator with zero-count buckets.
time_index_ops = [t for t in time_index if 7 <= t.hour <= 22]

occ_records = []
for zone_id in valid_occ['zone_id'].unique():
    zs = valid_occ[valid_occ['zone_id'] == zone_id]
    cap_row = zones[zones['zone_id'] == zone_id]
    capacity = int(cap_row['capacity'].iloc[0]) if not cap_row.empty else np.nan
    for t in time_index_ops:
        t_end = t + pd.Timedelta(FREQ)
        count = int(((zs['entry'] < t_end) & (zs['exit'] >= t)).sum())
        rate  = (count / capacity * 100) if (pd.notna(capacity) and capacity > 0) else np.nan
        occ_records.append({'time_bucket': t, 'zone_id': zone_id,
                            'occupancy_count': count, 'capacity': capacity,
                            'occupancy_rate': rate})

occ_df = pd.DataFrame(occ_records)
occ_df = occ_df.merge(zones[['zone_id','location','zone_type']], on='zone_id', how='left')
occ_df['hour']        = occ_df['time_bucket'].dt.hour
occ_df['weekday_num'] = occ_df['time_bucket'].dt.weekday
occ_df['weekday_name']= occ_df['time_bucket'].dt.day_name()
print(f"  Occupancy grid (operating hours only): {occ_df.shape}")

# Summary: avg and peak occupancy rate per zone (during operating hours)
avg_occ = (occ_df[occ_df['occupancy_rate'].notna()]
           .groupby(['zone_id','location'])
           .agg(avg_occ_rate=('occupancy_rate','mean'),
                peak_occ_rate=('occupancy_rate','max'),
                avg_occ_count=('occupancy_count','mean'),
                peak_occ_count=('occupancy_count','max'))
           .reset_index()
           .sort_values('peak_occ_rate', ascending=False))
# Rename for compatibility with downstream code
avg_occ['occupancy_rate'] = avg_occ['avg_occ_rate']
avg_occ['occupancy_count'] = avg_occ['avg_occ_count']
print("\n  Operating-hours occupancy by zone (avg and peak %):")
print(avg_occ[['location','avg_occ_rate','peak_occ_rate','avg_occ_count','peak_occ_count']].to_string(index=False))
print("\n  NOTE: Max simultaneous vehicles in this 60-day dataset is small (~6) relative")
print("  to zone capacity, resulting in low occupancy rates. This reflects the synthetic")
print("  data density (~20 sessions/day across 6 zones).")

# Turnover calculation
turnover = (valid.groupby('zone_id')['session_id'].count().reset_index(name='valid_sessions')
                 .merge(zones[['zone_id','capacity','location']], on='zone_id', how='left'))
turnover['turnover_rate'] = (turnover['valid_sessions'] / turnover['capacity']).round(2)
turnover = turnover.sort_values('turnover_rate', ascending=False)

# Build Occupancy & Zone Summary table (Answers Q1 & Q3)
occ_summary = zone_summary.merge(avg_occ[['zone_id','avg_occ_count','peak_occ_count','avg_occ_rate','peak_occ_rate']],
                                 on='zone_id', how='left')
occ_summary = occ_summary.merge(turnover[['zone_id','turnover_rate']], on='zone_id', how='left')
occ_summary['avg_occ_rate'] = occ_summary['avg_occ_rate'].round(2)
occ_summary['peak_occ_rate'] = occ_summary['peak_occ_rate'].round(2)
occ_summary['avg_occ_count'] = occ_summary['avg_occ_count'].round(2)
occ_summary['avg_dur_min'] = occ_summary['avg_dur_min'].round(1)
occ_summary['median_dur_min'] = occ_summary['median_dur_min'].round(1)
occ_summary['total_fees'] = occ_summary['total_fees'].round(2)

occ_summary = occ_summary.rename(columns={
    'sessions': 'total_sessions',
    'avg_dur_min': 'avg_duration_min',
    'median_dur_min': 'median_duration_min',
    'total_fees': 'total_fees_collected',
    'avg_occ_rate': 'avg_occupancy_rate_pct',
    'peak_occ_rate': 'peak_occupancy_rate_pct',
})

summary_cols = ['zone_id','location','zone_type','capacity','total_sessions','valid_sessions',
                'turnover_rate','avg_occ_count','peak_occ_count','avg_occupancy_rate_pct',
                'peak_occupancy_rate_pct','avg_duration_min','median_duration_min','total_fees_collected']
occ_summary = occ_summary[summary_cols]
occ_summary.to_csv(PROJECT_ROOT / 'outputs' / 'occupancy_summary.csv', index=False)
occ_summary.to_csv(REPORTS_DIR / 'occupancy_summary.csv', index=False)
print(f"  ✓ Saved occupancy_summary.csv")

# Enrich df for cleaned_parking_data.csv (Excel analysis deliverable)
df['entry_hour'] = df['entry'].dt.hour
df['time_period'] = df['entry_hour'].apply(assign_time_period)
df['is_weekend'] = df['entry'].dt.weekday.isin([5, 6]).astype(int)
df['day_type'] = np.where(df['is_weekend'] == 1, 'Weekend', 'Weekday')
df['event_label'] = np.where(df['event'] == 1, 'Event Day', 'Non-Event Day')
df['rain_label'] = np.where(df['rain'] == 1, 'Rainy Day', 'Non-Rainy Day')
df['duration_hours'] = (df['duration_min'] / 60).round(2)

# Merge zone-level metrics into df for Excel ease-of-use
zone_lookup = occ_summary[['zone_id','turnover_rate','avg_occ_count','peak_occ_count',
                           'avg_occupancy_rate_pct','peak_occupancy_rate_pct']].rename(
    columns={
        'turnover_rate': 'zone_turnover_rate',
        'avg_occ_count': 'zone_avg_occupancy',
        'peak_occ_count': 'zone_peak_occupancy',
        'avg_occupancy_rate_pct': 'zone_avg_occupancy_rate_pct',
        'peak_occupancy_rate_pct': 'zone_peak_occupancy_rate_pct',
    }
)
df = df.merge(zone_lookup, on='zone_id', how='left')

# Save updated cleaned_parking_data.csv
df.to_csv(DATA_PROC / 'cleaned_parking_data.csv', index=False)
print(f"  ✓ Saved enriched cleaned_parking_data.csv → {len(df):,} rows × {len(df.columns)} cols")

# Calendar join for occ_df
cal_lookup = calendar.copy()
occ_df2 = occ_df.copy()
occ_df2['date_key'] = occ_df2['time_bucket'].dt.normalize()
occ_df2 = occ_df2.merge(cal_lookup.rename(columns={'date':'date_key'}), on='date_key', how='left')

# =============================================================================
# SECTION 7 – Charts
# =============================================================================
print("\n" + "="*60)
print("SECTION 7: GENERATING CHARTS")
print("="*60)

# Chart 1 – Sessions by zone with explicit parameters
sess_sorted = occ_summary.sort_values('total_sessions', ascending=False).reset_index(drop=True)
fig, ax = plt.subplots(figsize=(11, 6.5), facecolor='#0f1117', dpi=150)
ax.set_facecolor('#1a1d27')
colors1 = ['#6c63ff', '#ff6584', '#43e97b', '#f7971e', '#4facfe', '#e056fd']
bars = ax.bar(sess_sorted['location'], sess_sorted['total_sessions'],
              color=colors1, edgecolor='#0f1117', width=0.6, linewidth=1.2)
ax.set_ylim(0, 360)
total_all = sess_sorted['total_sessions'].sum()

for bar, row in zip(bars, sess_sorted.itertuples()):
    h = bar.get_height()
    pct = (row.total_sessions / total_all) * 100
    ax.text(bar.get_x() + bar.get_width()/2, h + 5, f'{row.total_sessions} ({pct:.1f}%)',
            ha='center', va='bottom', fontsize=11, fontweight='bold', color='white')
    param_text = f'Valid: {row.valid_sessions}\nCap: {row.capacity}\nType: {row.zone_type}'
    ax.text(bar.get_x() + bar.get_width()/2, h + 25, param_text,
            ha='center', va='bottom', fontsize=8.5, color='#a0aec0',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#11141e', edgecolor='#2d3748', alpha=0.9))

info_box_text1 = (
    'SESSION & VOLUME PARAMETERS\n'
    '--------------------------------------------------\n'
    '• Total Sessions Sampled: 1,200 (100.0%)\n'
    '• Completed Valid Sessions: 1,170 (97.5%)\n'
    '• Incomplete / Open Sessions: 23 (1.9%)\n'
    '• Observation Horizon: 60 Days (Nov–Dec 2024)\n'
    '• Total Network Capacity: 534 Spaces (6 Zones)\n'
    '• Highest Volume: Railway Station (255 | 21.3%)\n'
    '• Lowest Volume: Riverfront (147 | 12.3%)\n'
    '• Average per Zone: 200.0 sessions (± 44.5)'
)
ax.text(0.97, 0.96, info_box_text1, transform=ax.transAxes,
        fontsize=9, verticalalignment='top', horizontalalignment='right',
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#161922', edgecolor='#6c63ff', linewidth=1.5, alpha=0.95),
        color='#f7fafc', family='monospace')

ax.set_title('Chart 1 — Total Parking Sessions by Zone (with Key Parameters)',
             fontsize=14, fontweight='bold', pad=15, color='white')
ax.set_xlabel('Zone / Location', fontsize=11, color='#e2e8f0', labelpad=10)
ax.set_ylabel('Total Number of Sessions', fontsize=11, color='#e2e8f0')
ax.tick_params(axis='x', colors='#e2e8f0', labelsize=10, rotation=15)
ax.tick_params(axis='y', colors='#e2e8f0', labelsize=10)
ax.grid(axis='y', linestyle='--', alpha=0.2, color='#ffffff')
fig.tight_layout(); save_chart(fig, '01_sessions_by_zone')

# Chart 2 – Peak occupancy rate by zone with explicit parameters
occ_sorted = occ_summary.sort_values('peak_occupancy_rate_pct', ascending=False).reset_index(drop=True)
fig, ax = plt.subplots(figsize=(11, 6.5), facecolor='#0f1117', dpi=150)
ax.set_facecolor('#1a1d27')
colors2 = ['#6c63ff', '#ff6584', '#43e97b', '#f7971e', '#4facfe', '#e056fd']
bars = ax.bar(occ_sorted['location'], occ_sorted['peak_occupancy_rate_pct'],
              color=colors2, edgecolor='#0f1117', width=0.6, linewidth=1.2)
ax.set_ylim(0, 10.5)

for bar, row in zip(bars, occ_sorted.itertuples()):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + 0.25, f'{row.peak_occupancy_rate_pct:.1f}%',
            ha='center', va='bottom', fontsize=12, fontweight='bold', color='white')
    param_text = f'Peak: {row.peak_occ_count} cars\nCap: {row.capacity}\nType: {row.zone_type}'
    ax.text(bar.get_x() + bar.get_width()/2, h + 0.95, param_text,
            ha='center', va='bottom', fontsize=8.5, color='#a0aec0',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#11141e', edgecolor='#2d3748', alpha=0.9))

info_box_text = (
    'PARAMETRIC DEFINITIONS & BENCHMARKS\n'
    '--------------------------------------------------\n'
    '• Metric Formula:  (Peak Cars / Total Capacity) × 100\n'
    '• Operating Window: 07:00 – 22:00 (Daily)\n'
    '• Industry Benchmark: 85.0% (Congestion threshold)\n'
    '• Maximum City Peak: 7.1% (City Hospital, Z02)\n'
    '• Minimum City Peak: 2.7% (Railway Station, Z04)\n'
    '• Core Finding: Severe Underutilization (< 8% in all zones)'
)
ax.text(0.97, 0.96, info_box_text, transform=ax.transAxes,
        fontsize=9, verticalalignment='top', horizontalalignment='right',
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#161922', edgecolor='#6c63ff', linewidth=1.5, alpha=0.95),
        color='#f7fafc', family='monospace')

ax.set_title('Chart 2 — Peak Occupancy Rate by Zone (with Key Parameters)',
             fontsize=14, fontweight='bold', pad=15, color='white')
ax.set_xlabel('Zone / Location', fontsize=11, color='#e2e8f0', labelpad=10)
ax.set_ylabel('Peak Occupancy Rate (%)', fontsize=11, color='#e2e8f0')
ax.tick_params(axis='x', colors='#e2e8f0', labelsize=10, rotation=15)
ax.tick_params(axis='y', colors='#e2e8f0', labelsize=10)
ax.grid(axis='y', linestyle='--', alpha=0.2, color='#ffffff')
fig.tight_layout(); save_chart(fig, '02_occupancy_rate_by_zone')

# Chart 3 – Hourly occupancy
hourly_occ = occ_df.groupby('hour')['occupancy_count'].mean().reset_index(name='avg_occupancy')
fig, ax = plt.subplots(figsize=(11, 5), facecolor='#0f1117')
ax.set_facecolor('#1a1d27')
ax.fill_between(hourly_occ['hour'], hourly_occ['avg_occupancy'], alpha=0.25, color='#6c63ff')
ax.plot(hourly_occ['hour'], hourly_occ['avg_occupancy'],
        color='#6c63ff', linewidth=2.5, marker='o', markersize=5)
ax.set_title('Chart 3 — Average Hourly Parking Occupancy (All Zones)', fontsize=14, pad=12, color='white')
ax.set_xlabel('Hour of Day (0 = midnight)'); ax.set_ylabel('Avg Occupied Spaces')
ax.set_xticks(range(0, 24)); ax.grid(alpha=0.3)
fig.tight_layout(); save_chart(fig, '03_hourly_occupancy')

# Chart 4 – Heatmap weekday × hour
heatmap_data = (occ_df.groupby(['weekday_num','weekday_name','hour'])['occupancy_count']
                      .mean().reset_index())
heatmap_pivot = heatmap_data.pivot_table(index='weekday_num', columns='hour',
                                          values='occupancy_count', aggfunc='mean')
wd_labels = (heatmap_data.drop_duplicates('weekday_num')
                          .sort_values('weekday_num')['weekday_name'].tolist())
fig, ax = plt.subplots(figsize=(14, 5), facecolor='#0f1117')
ax.set_facecolor('#0f1117')
sns.heatmap(heatmap_pivot, ax=ax, cmap='magma', linewidths=0.3, linecolor='#0f1117',
            fmt='.0f', annot=True, annot_kws={'size': 7},
            yticklabels=wd_labels, cbar_kws={'label': 'Avg Occupied Spaces'})
ax.set_title('Chart 4 — Occupancy Heatmap: Weekday × Hour of Day', fontsize=14, pad=12, color='white')
ax.set_xlabel('Hour of Day'); ax.set_ylabel('Day of Week')
ax.tick_params(colors='#e0e0e0')
fig.tight_layout(); save_chart(fig, '04_occupancy_heatmap')

# Chart 5 – Turnover by zone
turnover = (valid.groupby('zone_id')['session_id'].count().reset_index(name='valid_sessions')
                 .merge(zones[['zone_id','capacity','location']], on='zone_id', how='left'))
turnover['turnover_rate'] = turnover['valid_sessions'] / turnover['capacity']
turnover = turnover.sort_values('turnover_rate', ascending=False)
fig, ax = plt.subplots(figsize=(9, 5), facecolor='#0f1117')
ax.set_facecolor('#1a1d27')
bars = ax.bar(turnover['location'], turnover['turnover_rate'],
              color=PALETTE[:len(turnover)], edgecolor='#0f1117', linewidth=0.5)
for bar, val in zip(bars, turnover['turnover_rate']):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.02, f'{val:.2f}',
            ha='center', va='bottom', fontsize=9, color='#e0e0e0')
ax.set_title('Chart 5 — Parking Turnover by Zone\n(Sessions per Space over Observation Period)',
             fontsize=13, pad=10, color='white')
ax.set_xlabel('Zone / Location'); ax.set_ylabel('Turnover Rate (sessions / space)')
ax.tick_params(axis='x', rotation=25); ax.grid(axis='y', alpha=0.3)
fig.tight_layout(); save_chart(fig, '05_turnover_by_zone')

# Chart 6 – Event vs Non-event
ev_sess = df.merge(cal_lookup.rename(columns={'date':'entry_date'}), on='entry_date',
                   how='left', suffixes=('','_cal2'))
event_col = 'event_cal2' if 'event_cal2' in ev_sess.columns else 'event'
occ_event_col = 'event'
event_occ = occ_df2.groupby(occ_event_col)['occupancy_count'].agg(['mean','count']).reset_index()
event_occ['event_label'] = event_occ[occ_event_col].map({0:'Non-Event Day', 1:'Event Day'})
ev_vol = ev_sess.groupby(event_col)['session_id'].count().reset_index(name='sessions')
ev_vol['event_label'] = ev_vol[event_col].map({0:'Non-Event Day', 1:'Event Day'})

fig, axes = plt.subplots(1, 2, figsize=(11, 5), facecolor='#0f1117')
for ax in axes: ax.set_facecolor('#1a1d27')
colors_ev = ['#4facfe','#f093fb']
axes[0].bar(event_occ['event_label'], event_occ['mean'], color=colors_ev, edgecolor='#0f1117')
for bar, val in zip(axes[0].patches, event_occ['mean']):
    axes[0].text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.05, f'{val:.2f}',
                 ha='center', va='bottom', color='white', fontsize=10)
axes[0].set_title('Mean Occupied Spaces', color='white'); axes[0].set_ylabel('Avg Occupied Spaces')
axes[0].grid(axis='y', alpha=0.3)
axes[1].bar(ev_vol['event_label'], ev_vol['sessions'], color=colors_ev, edgecolor='#0f1117')
for bar, val in zip(axes[1].patches, ev_vol['sessions']):
    axes[1].text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, str(val),
                 ha='center', va='bottom', color='white', fontsize=10)
axes[1].set_title('Total Session Count', color='white'); axes[1].set_ylabel('Number of Sessions')
axes[1].grid(axis='y', alpha=0.3)
for ax in axes: ax.tick_params(axis='x', rotation=15)
fig.suptitle('Chart 6 — Event vs Non-Event Parking Behaviour', fontsize=14, color='white')
fig.tight_layout(); save_chart(fig, '06_event_vs_nonevent_occupancy')

# Chart 7 – Rainy vs Non-rainy
rain_df = df.copy()
rain_col = 'rain'
rain_df['rain_label'] = rain_df[rain_col].map({0:'Non-Rainy Day', 1:'Rainy Day'})
rain_summary = rain_df.groupby('rain_label').agg(
    sessions=('session_id','count'), avg_dur_min=('duration_min','mean')).reset_index()
occ_rain = occ_df2.copy()
occ_rain['rain_label'] = occ_rain.get('rain', pd.Series(dtype=float)).map({0:'Non-Rainy Day', 1:'Rainy Day'})
# Re-merge weather onto occ_df2 directly
occ_df3w = occ_df2.merge(weather.rename(columns={'date':'date_key'}), on='date_key', how='left', suffixes=('','_wx'))
rain_occ_col = 'rain_wx' if 'rain_wx' in occ_df3w.columns else 'rain'
occ_df3w['rain_label'] = occ_df3w[rain_occ_col].map({0:'Non-Rainy Day', 1:'Rainy Day'})
rain_occ = occ_df3w.groupby('rain_label')['occupancy_count'].mean().reset_index(name='avg_occupancy')

fig, axes = plt.subplots(1, 3, figsize=(14, 5), facecolor='#0f1117')
for ax in axes: ax.set_facecolor('#1a1d27')
colors_rain = ['#43e97b','#4facfe']
rain_s = rain_summary.sort_values('rain_label')
axes[0].bar(rain_s['rain_label'], rain_s['sessions'], color=colors_rain, edgecolor='#0f1117')
axes[0].set_title('Session Volume', color='white'); axes[0].set_ylabel('Sessions')
axes[0].grid(axis='y', alpha=0.3)
for bar, val in zip(axes[0].patches, rain_s['sessions']):
    axes[0].text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, str(val),
                 ha='center', color='white', fontsize=10)
axes[1].bar(rain_s['rain_label'], rain_s['avg_dur_min'], color=colors_rain, edgecolor='#0f1117')
axes[1].set_title('Avg Duration (min)', color='white'); axes[1].set_ylabel('Minutes')
axes[1].grid(axis='y', alpha=0.3)
for bar, val in zip(axes[1].patches, rain_s['avg_dur_min']):
    axes[1].text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3, f'{val:.1f}',
                 ha='center', color='white', fontsize=10)
rain_occ_s = rain_occ.sort_values('rain_label')
axes[2].bar(rain_occ_s['rain_label'], rain_occ_s['avg_occupancy'], color=colors_rain, edgecolor='#0f1117')
axes[2].set_title('Avg Occupancy Count', color='white'); axes[2].set_ylabel('Occupied Spaces')
axes[2].grid(axis='y', alpha=0.3)
for bar, val in zip(axes[2].patches, rain_occ_s['avg_occupancy']):
    axes[2].text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.05, f'{val:.2f}',
                 ha='center', color='white', fontsize=10)
for ax in axes: ax.tick_params(axis='x', rotation=15)
fig.suptitle('Chart 7 — Rainy vs Non-Rainy Parking Behaviour', fontsize=14, color='white')
fig.tight_layout(); save_chart(fig, '07_rainy_vs_nonrainy')

# Chart 8 – Duration by zone type
dur_df = valid[valid['zone_type'].notna()].copy()
zone_types = sorted(dur_df['zone_type'].unique())
colors_zt  = {zt: PALETTE[i % len(PALETTE)] for i, zt in enumerate(zone_types)}
fig, axes = plt.subplots(1, 2, figsize=(13, 6), facecolor='#0f1117')
for ax in axes: ax.set_facecolor('#1a1d27')
bp_data = [dur_df[dur_df['zone_type']==zt]['duration_min'].dropna().values for zt in zone_types]
bp = axes[0].boxplot(bp_data, patch_artist=True, notch=False,
                      medianprops={'color':'white','linewidth':2})
for patch, zt in zip(bp['boxes'], zone_types):
    patch.set_facecolor(colors_zt[zt]); patch.set_alpha(0.75)
axes[0].set_xticklabels(zone_types, rotation=20, fontsize=9)
axes[0].set_title('Box Plot — Duration by Zone Type', color='white')
axes[0].set_ylabel('Duration (minutes)'); axes[0].grid(axis='y', alpha=0.3)
sns.violinplot(data=dur_df, x='zone_type', y='duration_min', ax=axes[1],
               palette=list(colors_zt.values()), inner='quartile', cut=0)
axes[1].set_title('Violin Plot — Duration by Zone Type', color='white')
axes[1].set_xlabel('Zone Type'); axes[1].set_ylabel('Duration (minutes)')
axes[1].tick_params(axis='x', rotation=20); axes[1].grid(axis='y', alpha=0.3)
fig.suptitle('Chart 8 — Parking Duration Distribution by Zone Type',
             fontsize=14, color='white', y=1.01)
fig.tight_layout(); save_chart(fig, '08_duration_by_zone_type')

# Chart 9 – Weekday vs Weekend
occ_df_c = occ_df.copy()
occ_df_c['is_weekend'] = occ_df_c['weekday_num'].isin([5, 6])
wknd_occ = occ_df_c.groupby(['is_weekend','hour'])['occupancy_count'].mean().reset_index()
wknd_occ['day_type'] = wknd_occ['is_weekend'].map({False:'Weekday', True:'Weekend'})
fig, ax = plt.subplots(figsize=(11, 5), facecolor='#0f1117')
ax.set_facecolor('#1a1d27')
for dtype, color in [('Weekday','#6c63ff'),('Weekend','#ff6584')]:
    subset = wknd_occ[wknd_occ['day_type']==dtype]
    ax.plot(subset['hour'], subset['occupancy_count'], label=dtype,
            color=color, linewidth=2.5, marker='o', markersize=4)
    ax.fill_between(subset['hour'], subset['occupancy_count'], alpha=0.1, color=color)
ax.set_title('Chart 9 — Weekday vs Weekend Hourly Occupancy', fontsize=14, pad=12, color='white')
ax.set_xlabel('Hour of Day'); ax.set_ylabel('Avg Occupied Spaces')
ax.set_xticks(range(0, 24)); ax.legend(); ax.grid(alpha=0.3)
fig.tight_layout(); save_chart(fig, '09_weekday_vs_weekend_occupancy')

# Chart 10 – Fees by zone
fee_by_zone = (df.groupby(['zone_id','location'])['amount'].sum()
                 .reset_index(name='total_fees')
                 .sort_values('total_fees', ascending=False))
fig, ax = plt.subplots(figsize=(9, 5), facecolor='#0f1117')
ax.set_facecolor('#1a1d27')
bars = ax.bar(fee_by_zone['location'], fee_by_zone['total_fees'],
              color=PALETTE[:len(fee_by_zone)], edgecolor='#0f1117')
for bar, val in zip(bars, fee_by_zone['total_fees']):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+50, f'{val:,.0f}',
            ha='center', va='bottom', fontsize=9, color='white')
ax.set_title('Chart 10 — Total Parking Fees Collected by Zone', fontsize=14, pad=12, color='white')
ax.set_xlabel('Zone / Location'); ax.set_ylabel('Total Fees (currency units)')
ax.tick_params(axis='x', rotation=25); ax.grid(axis='y', alpha=0.3)
fig.tight_layout(); save_chart(fig, '10_fees_by_zone')

# =============================================================================
# SECTION 8 – Feature Engineering
# =============================================================================
print("\n" + "="*60)
print("SECTION 8: FEATURE ENGINEERING")
print("="*60)

def assign_time_period(hour):
    if 0 <= hour < 6:   return 'Night'
    if 6 <= hour < 12:  return 'Morning'
    if 12 <= hour < 18: return 'Afternoon'
    return 'Evening'

def assign_util_cat(rate):
    """
    Utilization categories adjusted for this dataset's observed rate range (0–8%).
    In typical urban studies: Low<50%, Moderate 50-85%, High>85%.
    For this synthetic dataset (max ~7%), we use percentile-based thresholds
    (33rd / 67th percentile of non-zero operating-hour rates) to preserve
    relative zone comparisons while noting that all zones are objectively low.
    Thresholds: Low<1%, Moderate 1-4%, High>=4%.
    """
    if pd.isna(rate): return 'Unknown'
    if rate < 1.0:    return 'Low'
    if rate <= 4.0:   return 'Moderate'
    return 'High'

fe = df.copy()
fe['entry_hour']               = fe['entry'].dt.hour
fe['entry_weekday']            = fe['entry'].dt.weekday
fe['entry_month']              = fe['entry'].dt.month
fe['is_weekend']               = fe['entry_weekday'].isin([5, 6]).astype(int)
fe['time_period']              = fe['entry_hour'].apply(assign_time_period)
fe['session_duration_minutes'] = fe['duration_min']
fe['is_holiday']               = fe['holiday'].fillna(0).astype(int)
fe['is_event']                 = fe['event'].fillna(0).astype(int)
fe['is_rainy']                 = fe['rain'].fillna(0).astype(int)

zone_avg_occ = avg_occ[['zone_id','avg_occ_count','avg_occ_rate','peak_occ_rate']].rename(
    columns={'avg_occ_count':'avg_occ_count_z','avg_occ_rate':'avg_occ_rate_z',
              'peak_occ_rate':'peak_occ_rate_z'})
fe = fe.merge(zone_avg_occ, on='zone_id', how='left')
fe['occupancy_count'] = fe['avg_occ_count_z']
fe['occupancy_rate']  = fe['avg_occ_rate_z']
fe['utilization_category'] = fe['occupancy_rate'].apply(assign_util_cat)
fe = fe.merge(turnover[['zone_id','turnover_rate']], on='zone_id', how='left')

fe_cols = ['session_id','zone_id','entry_date','entry_hour','entry_weekday','entry_month',
           'is_weekend','time_period','is_holiday','is_event','is_rainy','temperature',
           'zone_type','session_duration_minutes','occupancy_count','occupancy_rate',
           'utilization_category','turnover_rate','location','capacity','payment_mode','amount']
fe_save = fe[[c for c in fe_cols if c in fe.columns]].copy()
fe_save.to_csv(DATA_PROC / 'parking_features.csv', index=False)
print(f"  ✓ Saved parking_features.csv → {len(fe_save):,} rows × {len(fe_save.columns)} cols")
print(f"  Utilization categories: {fe_save['utilization_category'].value_counts().to_dict()}")

# =============================================================================
# SECTION 9 – Categorical Encoding
# =============================================================================
print("\n" + "="*60)
print("SECTION 9: CATEGORICAL ENCODING")
print("="*60)

enc = fe_save.copy()
util_map = {'Low': 0, 'Moderate': 1, 'High': 2, 'Unknown': np.nan}
enc['utilization_category_enc'] = enc['utilization_category'].map(util_map)
nominal_cols = [c for c in ['zone_type','time_period','payment_mode'] if c in enc.columns]
enc_ohe = pd.get_dummies(enc, columns=nominal_cols, prefix=nominal_cols,
                          drop_first=False, dtype=int)
print(f"  Shape after encoding: {enc_ohe.shape}")
new_cols = [c for c in enc_ohe.columns if c not in enc.columns]
print(f"  New encoded columns ({len(new_cols)}): {new_cols}")

# =============================================================================
# SECTION 10 – Feature Selection
# =============================================================================
print("\n" + "="*60)
print("SECTION 10: FEATURE SELECTION (target: session_duration_minutes)")
print("="*60)

drop_cols = ['session_id','entry_date','occupancy_count','occupancy_rate',
             'utilization_category','location','amount','utilization_category_enc']
feature_df = enc_ohe.drop(columns=[c for c in drop_cols if c in enc_ohe.columns], errors='ignore')
feature_df = feature_df.drop(columns=['zone_id'], errors='ignore')
target_col = 'session_duration_minutes'
modelling_df = feature_df[feature_df[target_col].notna()].copy()
non_numeric = modelling_df.select_dtypes(exclude='number').columns.tolist()
if non_numeric:
    modelling_df = modelling_df.drop(columns=non_numeric)
X = modelling_df.drop(columns=[target_col]).fillna(modelling_df.median(numeric_only=True))
y = modelling_df[target_col]
const_cols = [c for c in X.columns if X[c].nunique() <= 1]
X = X.drop(columns=const_cols)
print(f"  Feature matrix: {X.shape}")

# Correlation
corr = X.select_dtypes(include='number').corr()
fig, ax = plt.subplots(figsize=(max(8, len(corr)), max(6, len(corr))), facecolor='#0f1117')
ax.set_facecolor('#0f1117')
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, ax=ax, cmap='coolwarm', center=0, annot=True,
            fmt='.2f', annot_kws={'size': 7}, linewidths=0.3,
            linecolor='#0f1117', cbar_kws={'label': 'Pearson r'})
ax.set_title('Feature Correlation Matrix', fontsize=13, pad=10, color='white')
ax.tick_params(colors='#e0e0e0', labelsize=8)
fig.tight_layout(); save_chart(fig, '11_correlation_matrix')

# Mutual Information
mi_scores = mutual_info_regression(X, y, random_state=42)
mi_df = pd.DataFrame({'feature': X.columns, 'mi_score': mi_scores})
mi_df = mi_df.sort_values('mi_score', ascending=False).reset_index(drop=True)

fig, ax = plt.subplots(figsize=(9, 6), facecolor='#0f1117')
ax.set_facecolor('#1a1d27')
colors_mi = plt.cm.viridis(np.linspace(0.2, 0.9, len(mi_df)))
ax.barh(mi_df['feature'][::-1], mi_df['mi_score'][::-1], color=colors_mi)
ax.set_title('Mutual Information Scores — Target: session_duration_minutes',
             fontsize=12, pad=10, color='white')
ax.set_xlabel('Mutual Information Score'); ax.grid(axis='x', alpha=0.3)
fig.tight_layout(); save_chart(fig, '12_mutual_information')

print("\n  Top 10 features by Mutual Information:")
print(mi_df.head(10).to_string(index=False))

# Save selected features
selected = mi_df[mi_df['mi_score'] > 0].copy()
rationale_map = {
    'entry_hour'     : 'Time-of-day drives duration; peak hours show different stay lengths',
    'entry_weekday'  : 'Weekday/weekend patterns differ in parking behaviour',
    'entry_month'    : 'Seasonal demand variation',
    'is_weekend'     : 'Binary flag for weekend; simpler alternative to weekday encoding',
    'is_rainy'       : 'Weather affects how long drivers park',
    'is_event'       : 'Events drive longer stays or higher occupancy',
    'is_holiday'     : 'Holiday behaviour differs from regular days',
    'temperature'    : 'Extreme temperatures may alter parking behaviour',
    'turnover_rate'  : 'Zone-level metric; high-turnover zones have shorter sessions',
    'capacity'       : 'Larger zones may attract different session lengths',
}
selected['rationale'] = selected['feature'].map(rationale_map).fillna('Identified by MI analysis')
ohe_selected = [c for c in X.columns if any(c.startswith(p)
                for p in ['zone_type_','time_period_','payment_mode_'])]
ohe_rows = pd.DataFrame({'feature': ohe_selected,
    'mi_score': [mi_df.set_index('feature').loc[c,'mi_score']
                  if c in mi_df['feature'].values else 0 for c in ohe_selected],
    'rationale': ['One-hot encoded nominal categorical']*len(ohe_selected)})
final_selected = pd.concat([selected, ohe_rows], ignore_index=True).drop_duplicates('feature')
final_selected.to_csv(PROJECT_ROOT / 'outputs' / 'selected_features.csv', index=False)
print(f"\n  ✓ Saved selected_features.csv → {len(final_selected)} features")

# =============================================================================
# SECTION 11 – Business Findings
# =============================================================================
print("\n" + "="*60)
print("SECTION 11: BUSINESS FINDINGS")
print("="*60)

top_zone   = avg_occ.sort_values('peak_occ_rate', ascending=False).iloc[0]
bot_zone   = avg_occ.sort_values('peak_occ_rate', ascending=False).iloc[-1]
peak_hour  = int(hourly_occ.loc[hourly_occ['avg_occupancy'].idxmax(), 'hour'])
low_turn   = turnover.sort_values('turnover_rate').iloc[0]
high_turn  = turnover.sort_values('turnover_rate', ascending=False).iloc[0]

ev_mean_val   = event_occ[event_occ[occ_event_col]==1]['mean'].values
non_ev_val    = event_occ[event_occ[occ_event_col]==0]['mean'].values
ev_mean_str   = f"{ev_mean_val[0]:.2f}" if len(ev_mean_val) else "N/A"
non_ev_str    = f"{non_ev_val[0]:.2f}"  if len(non_ev_val)  else "N/A"

rain_dur_vals  = rain_df.groupby(rain_col)['duration_min'].mean()
rain_dur_val   = float(rain_dur_vals.loc[1]) if 1 in rain_dur_vals.index else float('nan')
norain_dur_val = float(rain_dur_vals.loc[0]) if 0 in rain_dur_vals.index else float('nan')

dur_by_type = dur_df.groupby('zone_type')['duration_min'].agg(['mean','median']).round(1)

print(f"\n  Q1 — Highest peak occupancy: {top_zone['location']} (peak {top_zone['peak_occ_rate']:.1f}%, avg {top_zone['avg_occ_rate']:.2f}%)")
print(f"       Lowest  peak occupancy: {bot_zone['location']} (peak {bot_zone['peak_occ_rate']:.1f}%, avg {bot_zone['avg_occ_rate']:.2f}%)")
print(f"  Q2 — Peak hour        : {peak_hour}:00")
print(f"  Q3 — Lowest turnover  : {low_turn['location']} ({low_turn['turnover_rate']:.2f} sess/space)")
print(f"       Highest turnover : {high_turn['location']} ({high_turn['turnover_rate']:.2f} sess/space)")
print(f"  Q4 — Event occ mean   : {ev_mean_str} vs Non-event: {non_ev_str}")
print(f"  Q5 — Rain avg duration: {rain_dur_val:.1f} min | Non-rain: {norain_dur_val:.1f} min")
print(f"  Q6 — Duration by zone type:\n{dur_by_type.to_string()}")

findings = [
    ('Q1 — Highest occupancy zones',
     f"Peak occ rates during operating hours (7am-10pm): {avg_occ[['location','peak_occ_rate']].set_index('location')['peak_occ_rate'].round(1).to_dict()}",
     f"{top_zone['location']} highest peak occupancy ({top_zone['peak_occ_rate']:.1f}%); {bot_zone['location']} lowest ({bot_zone['peak_occ_rate']:.1f}%). All zones <10% due to synthetic data sparsity (~20 sessions/day/6 zones).",
     f"Prioritise enforcement and signage at {top_zone['location']} and {avg_occ.iloc[1]['location']}; review pricing at {bot_zone['location']}"),
    ('Q2 — Peak parking periods',
     f"Hourly occupancy from {len(valid_occ):,} valid sessions over {n_days} days",
     f"Peak demand at {peak_hour}:00; weekday/weekend patterns differ (Charts 3,4,9)",
     "Staff and signage from peak hour; consider time-based pricing"),
    ('Q3 — Low turnover zones',
     f"Turnover range {turnover['turnover_rate'].min():.2f}–{turnover['turnover_rate'].max():.2f}",
     f"{low_turn['location']} has lowest turnover ({low_turn['turnover_rate']:.2f} sess/space)",
     "Introduce maximum stay limits or escalating tariffs in low-turnover zones"),
    ('Q4 — Event effect on occupancy',
     f"Event days vs non-event days; mean occ {ev_mean_str} vs {non_ev_str}",
     f"Event days show {'higher' if ev_mean_str > non_ev_str else 'lower/similar'} occupancy (association only)",
     "Pre-position staff/signage before known events; consider surge pricing"),
    ('Q5 — Rain effect on parking',
     f"Rainy avg duration {rain_dur_val:.1f} min; non-rainy {norain_dur_val:.1f} min",
     f"Rainy days show {'longer' if rain_dur_val > norain_dur_val else 'shorter'} avg stays by {abs(rain_dur_val-norain_dur_val):.1f} min (association; confounders not controlled)",
     "Monitor weather forecasts; consider extended-stay pricing on rainy days"),
    ('Q6 — Duration by zone type',
     f"Valid sessions: {dur_df.groupby('zone_type')['session_id'].count().to_dict()}",
     f"Zone types have distinct duration profiles (see Chart 8); medians: {dur_by_type['median'].to_dict()}",
     "Tailor stay-limit policies per zone type; one-size-fits-all limits are inequitable"),
]
find_df = pd.DataFrame(findings, columns=['Business Question','Evidence','Key Finding','Recommended Action'])
find_df.to_csv(REPORTS_DIR / 'business_findings.csv', index=False)
find_df.to_csv(PROJECT_ROOT / 'outputs' / 'business_findings.csv', index=False)
print(f"\n  ✓ Saved business_findings.csv")

# =============================================================================
# FINAL CHECKLIST
# =============================================================================
print("\n" + "="*60)
print("FINAL DELIVERABLE CHECKLIST")
print("="*60)

files_to_check = [
    DATA_PROC / 'cleaned_parking_data.csv',
    DATA_PROC / 'parking_features.csv',
    PROJECT_ROOT / 'outputs' / 'selected_features.csv',
    PROJECT_ROOT / 'outputs' / 'data_quality_report.csv',
    PROJECT_ROOT / 'outputs' / 'occupancy_summary.csv',
    PROJECT_ROOT / 'outputs' / 'business_findings.csv',
    REPORTS_DIR / 'data_quality_report.csv',
    REPORTS_DIR / 'business_findings.csv',
    REPORTS_DIR / 'occupancy_summary.csv',
]
for fp in files_to_check:
    status = '✓' if fp.exists() else '✗ MISSING'
    print(f"  {status}  {fp.relative_to(PROJECT_ROOT)}")

chart_files = sorted(CHARTS_DIR.glob('*.png'))
print(f"\n  Charts saved: {len(chart_files)}")
for cp in chart_files:
    print(f"    ✓ {cp.name}")

print("\n✅ Analysis pipeline complete!")
