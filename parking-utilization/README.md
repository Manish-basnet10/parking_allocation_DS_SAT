# ParkSmart City Services — Round 2: EDA & Feature Engineering
**Use Case 17: Parking Utilization Insight**  
**Organisation:** ParkSmart City Services | **Sector:** Smart City  

---

## 1. Project Overview & Deliverables Structure

```
parking-utilization/
├── data/
│   ├── raw/
│   └── processed/
│       ├── cleaned_parking_data.csv        # Primary deliverable for Round 3 Excel analysis (1,200 rows × 40 cols)
│       └── parking_features.csv            # Feature-engineered dataset (1,200 rows × 22 cols)
├── notebooks/
│   └── round2_eda_feature_engineering.ipynb# Fully executed notebook with all inline outputs & figures
├── outputs/
│   ├── charts/                             # 12 high-resolution visualisations covering all business questions
│   │   ├── 01_sessions_by_zone.png
│   │   ├── 02_occupancy_rate_by_zone.png
│   │   ├── 03_hourly_occupancy.png
│   │   ├── 04_occupancy_heatmap.png
│   │   ├── 05_turnover_by_zone.png
│   │   ├── 06_event_vs_nonevent_occupancy.png
│   │   ├── 07_rainy_vs_nonrainy.png
│   │   ├── 08_duration_by_zone_type.png
│   │   ├── 09_weekday_vs_weekend_occupancy.png
│   │   ├── 10_fees_by_zone.png
│   │   ├── 11_correlation_matrix.png
│   │   └── 12_mutual_information.png
│   ├── data_quality_report.csv             # Data quality audit issues, counts, actions & rationale
│   ├── occupancy_summary.csv               # Reliable zone-level occupancy & turnover statistics
│   ├── selected_features.csv               # Selected features with MI scores and modeling rationales
│   ├── business_findings.csv               # Evidence-based answers to all six business questions
│   └── reports/                            # Duplicate copies of reports for archive integrity
│       ├── data_quality_report.csv
│       ├── occupancy_summary.csv
│       └── business_findings.csv
├── src/
│   ├── parking_analysis.py                 # Reusable utility functions
│   └── run_analysis.py                     # Standalone end-to-end execution script
├── requirements.txt
└── README.md
```

Raw data source: `ParkSmart_Round2_Practice_Dataset/` (read-only, never modified).

---

## 2. Data Cleaning Rules & Quality Treatments

| Issue Identified | Raw Count | Treatment Action | Business & Technical Rationale |
|---|---|---|---|
| **Exact Duplicate Rows** | 12 session rows, 5 payment rows | Deduplicated (`drop_duplicates()`) | Duplicate rows inflate session counts and fee totals without providing new information. |
| **Missing Exit Timestamps** | 23 sessions | Preserved in dataset; `is_open = True`; excluded from duration & interval occupancy | Open/ongoing sessions are legitimate parking events; deleting them introduces sampling bias into traffic volume counts. |
| **Missing Duration Values** | 10 sessions | Recomputed from authoritative entry/exit timestamps | Sensor duration field had nulls; timestamp calculation is reliable and authoritative. |
| **Negative Calculated Durations** | 8 sessions | Flagged `is_negative = True`; `duration_min = NaN` | Entry timestamp after exit timestamp is physically impossible; excluded from duration stats. |
| **Extreme & Zero Durations** | 0 detected | Filter rules implemented (>24h or == 0 flagged) | Guards against data collection anomalies. |
| **Inconsistent Zone Labels** | 0 unresolved | Stripped whitespace, converted to uppercase | All 6 zones in session data (`Z01`–`Z06`) cleanly match master `zones.csv`. |

Full audit details: [`outputs/data_quality_report.csv`](outputs/data_quality_report.csv).

---

## 3. Findings for the Six Business Questions

### Q1: Which zones experience the highest occupancy?
* **Evidence:** Reconstructed hourly interval overlap across 60 days (7:00 AM – 10:00 PM operating hours).
* **Highest Peak Occupancy:** **City Hospital (`Z02`)** reached **7.06% peak occupancy** (6 vehicles parked concurrently; capacity 85) and highest average occupancy (**0.74%**).
* **Second Highest:** **Market Street (`Z05`)** reached **6.32% peak occupancy** (6 vehicles concurrently; capacity 95).
* **Lowest Occupancy:** **Railway Station (`Z04`)** reached only **2.67% peak occupancy** (4 vehicles concurrently; capacity 150).
* **Context Note:** All absolute occupancy rates remain under 10% because the synthetic sample consists of ~20 sessions per day spread across 6 zones. Relative differences between zones remain consistent.
* **Chart:** [`outputs/charts/02_occupancy_rate_by_zone.png`](outputs/charts/02_occupancy_rate_by_zone.png) & Summary: [`outputs/occupancy_summary.csv`](outputs/occupancy_summary.csv).

### Q2: When do peak parking periods occur?
* **Evidence:** Aggregated average vehicle counts by hour of day (0–23) and weekday heatmap.
* **Peak Hour:** **16:00 (4:00 PM)** is the daily peak demand hour with the highest simultaneous vehicle volume and 104 session arrivals. Secondary morning peak occurs at **08:00 (8:00 AM)** with 100 arrivals.
* **Operating Profile:** Parking demand is concentrated between **07:00 and 18:00**, with sharp drop-offs after 19:00.
* **Charts:** [`outputs/charts/03_hourly_occupancy.png`](outputs/charts/03_hourly_occupancy.png), [`outputs/charts/04_occupancy_heatmap.png`](outputs/charts/04_occupancy_heatmap.png), [`outputs/charts/09_weekday_vs_weekend_occupancy.png`](outputs/charts/09_weekday_vs_weekend_occupancy.png).

### Q3: Which zones have low turnover?
* **Evidence:** Turnover calculated as `Valid Sessions ÷ Zone Capacity` over the 60-day observation window.
* **Lowest Turnover:** **University (`Z06`)** with **1.55 sessions/space** (171 valid sessions / 110 capacity). This aligns with University having the longest stay durations (median 180 min).
* **Second Lowest:** **Railway Station (`Z04`)** with **1.65 sessions/space** (248 sessions / 150 capacity).
* **Highest Turnover:** **Market Street (`Z05`)** with **2.29 sessions/space** (218 sessions / 95 capacity).
* **Chart:** [`outputs/charts/05_turnover_by_zone.png`](outputs/charts/05_turnover_by_zone.png).

### Q4: How do events affect parking occupancy?
* **Evidence:** 4 Event days (77 sessions) vs 56 Non-Event days (1,123 sessions).
* **Findings:** Mean occupied spaces on Event days (**0.66 vehicles**) is virtually identical to Non-Event days (**0.66 vehicles**). Median duration increases slightly from **119.0 min** (non-event) to **127.5 min** (event). Daily session volume is ~19.3/day on event days vs ~20.1/day on non-event days.
* **Limitation:** In this synthetic dataset, events do not show a statistically significant surge in utilization.
* **Chart:** [`outputs/charts/06_event_vs_nonevent_occupancy.png`](outputs/charts/06_event_vs_nonevent_occupancy.png).

### Q5: Does rain change parking behaviour?
* **Evidence:** 10 Rainy days (208 sessions) vs 50 Non-Rainy days (992 sessions).
* **Findings:** Rainy days show slightly higher session volume (**20.8 sessions/day** vs **19.8 sessions/day** on non-rainy days). Average stay duration is nearly identical: **147.1 minutes** on rainy days vs **148.7 minutes** on non-rainy days (median 122.0 min vs 119.0 min).
* **Chart:** [`outputs/charts/07_rainy_vs_nonrainy.png`](outputs/charts/07_rainy_vs_nonrainy.png).

### Q6: How does parking duration vary by zone type?
* **Evidence:** Distribution of valid durations across zone types:
  1. **Educational (`University`):** Median = **180.0 min**, Mean = **224.3 min** (longest stays, students/staff lectures)
  2. **Hospital (`City Hospital`):** Median = **154.0 min**, Mean = **176.6 min** (appointments and visiting hours)
  3. **Recreational (`Riverfront`):** Median = **133.0 min**, Mean = **157.2 min** (leisure visits)
  4. **Commercial (`Central Plaza`, `Market Street`):** Median = **108.0 min**, Mean = **130.2 min** (shopping, dining)
  5. **Transit (`Railway Station`):** Median = **87.5 min**, Mean = **106.0 min** (shortest stays, drop-offs and short commutes)
* **Chart:** [`outputs/charts/08_duration_by_zone_type.png`](outputs/charts/08_duration_by_zone_type.png).

---

## 4. Excel Guide for Round 3 Analysis (Microsoft Excel for the Web)

The primary dataset [`data/processed/cleaned_parking_data.csv`](data/processed/cleaned_parking_data.csv) has been enriched specifically for effortless Pivot Table and formula analysis in Excel.

### Key Column Directory

| Analysis Goal | Recommended Excel Column(s) | Data Type | Notes for Excel |
|---|---|---|---|
| **1. Zone Occupancy** | `location`, `zone_type`, `capacity`, `zone_avg_occupancy_rate_pct`, `zone_peak_occupancy_rate_pct`, `zone_avg_occupancy` | Text / Decimal | Summarized from hourly reconstruction; use `AVERAGE` or Pivot Table Rows: `location`. |
| **2. Peak Periods** | `entry_hour`, `time_period`, `weekday`, `day_type` | Integer (0-23) / Text | Place `entry_hour` in Pivot Table Rows and `session_id` (Count) in Values to immediately plot peak curve. |
| **3. Turnover** | `zone_turnover_rate`, `valid_sessions`, `capacity` | Decimal | Calculated as sessions per space. Direct column `zone_turnover_rate` available. |
| **4. Event Comparison** | `event_label`, `event` | Text (`Event Day` vs `Non-Event Day`) | Put `event_label` in Pivot Table Rows, `duration_min` (Average) and `session_id` (Count) in Values. |
| **5. Weather Comparison**| `rain_label`, `rain`, `temperature` | Text (`Rainy Day` vs `Non-Rainy Day`) | Put `rain_label` in Pivot Table Rows, `duration_min` (Average) and `session_id` (Count) in Values. |
| **6. Duration Stats** | `duration_min`, `duration_hours`, `zone_type`, `is_valid_dur` | Numeric Minutes / Hours | Filter `is_valid_dur = TRUE` to exclude open/negative sessions. Put `zone_type` in Rows, `duration_min` in Values (`AVERAGE`, `MIN`, `MAX`). |

### Quick Excel Pivot Table Recipes

1. **Peak Hours Table & Chart:**
   - **Rows:** `entry_hour`
   - **Values:** `session_id` (Summarize by Count)
   - **Filter:** `is_valid_dur` = `TRUE`
   - **Chart:** 2D Line or Column chart. Peak will show at hour 16 (104 sessions) and hour 8 (100 sessions).

2. **Duration by Zone Type Table:**
   - **Rows:** `zone_type`
   - **Values:** `duration_min` (Average), `duration_min` (Count), `amount` (Sum)
   - **Filter:** `is_valid_dur` = `TRUE`

3. **Event & Weather Impact Comparison:**
   - **Rows:** `event_label` or `rain_label`
   - **Values:** `session_id` (Count), `duration_min` (Average), `amount` (Average)

---

## 5. Summary Table: `outputs/occupancy_summary.csv`

| zone_id | location | zone_type | capacity | total_sessions | valid_sessions | turnover_rate | avg_occ_count | peak_occ_count | avg_occ_rate_% | peak_occ_rate_% | avg_duration_min | median_duration_min | total_fees |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Z01 | Central Plaza | Commercial | 120 | 238 | 233 | 1.94 | 0.70 | 4 | 0.58% | 3.33% | 125.2 | 103.0 | 13,728.21 |
| Z02 | City Hospital | Hospital | 85 | 163 | 161 | 1.89 | 0.63 | 6 | 0.74% | 7.06% | 176.6 | 154.0 | 12,811.56 |
| Z03 | Riverfront | Recreational | 74 | 147 | 139 | 1.88 | 0.49 | 4 | 0.66% | 5.41% | 157.2 | 133.0 | 10,491.14 |
| Z04 | Railway Station | Transit | 150 | 255 | 248 | 1.65 | 0.66 | 4 | 0.44% | 2.67% | 106.0 | 87.5 | 11,881.03 |
| Z05 | Market Street | Commercial | 95 | 225 | 218 | 2.29 | 0.69 | 6 | 0.72% | 6.32% | 135.5 | 116.0 | 14,097.39 |
| Z06 | University | Educational | 110 | 172 | 171 | 1.55 | 0.77 | 6 | 0.70% | 5.45% | 224.3 | 180.0 | 17,873.11 |

---

## 6. Execution Instructions

To re-run the entire pipeline from terminal:
```bash
# 1. Activate environment
source .venv/bin/activate

# 2. Run analysis pipeline
python3 src/run_analysis.py

# 3. Or execute Jupyter Notebook
jupyter nbconvert --to notebook --execute --inplace notebooks/round2_eda_feature_engineering.ipynb
```
