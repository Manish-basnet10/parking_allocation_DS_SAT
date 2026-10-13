# Round 3: Data Analysis Using Microsoft Excel
**Use Case 17: Parking Utilization Insight**  
**Organization:** ParkSmart City Services | **Sector:** Smart City  
**Workbook Deliverable:** `outputs/Round3_Parking_Analysis.xlsx` (also accessible as `outputs/ParkSmart_Round3_Analysis.xlsx`)

---

## 1. Overview & Technical Transparency

This workbook was developed as a clean, structured college-level Excel project demonstrating practical data analysis on smart-city parking utilization.

### Important Note on PivotTables in Headless Python
As requested, here is the honest technical explanation regarding PivotTables:
* Headless Python libraries (`openpyxl`, `xlsxwriter`, `pandas`) run without the Microsoft Excel desktop GUI and **cannot synthesize the internal proprietary binary OLAP PivotCache engine** required for interactive PivotTables.
* To solve this for **Excel for the Web**:
  1. The master dataset on `Cleaned_Data` has been officially registered as an Excel Table named **`CleanedParkingData`**.
  2. The `Pivot_Analysis` sheet provides the **exact 5 core assessment summaries** with full formulas and charts.
  3. Clear, 1-click recipes are provided so that when you open the file in **Excel for the Web**, you can click **Insert > PivotTable** and select `CleanedParkingData` to create native PivotTables in seconds.

---

## 2. Workbook Structure & Sheet Guide

| Sheet Name | Content & Purpose | Key Excel Formulas & Features Used |
|---|---|---|
| **1. Cleaned_Data** | Full master dataset (1,200 rows × 40 columns). | Registered Table (`CleanedParkingData`), auto-filters enabled, frozen header row (`A2`), formatted currency, durations, and integers. |
| **2. Zone_Analysis** | Comprehensive breakdown of all 6 parking zones covering capacity, session counts, turnover rate, stay duration, revenue, and occupancy. | `=COUNTIF()`, `=COUNTIFS()`, `=ROUND()`, `=AVERAGEIFS()`, `=SUMIF()`, `=SUM()`. Includes a Column Chart comparing turnover rates across zones. |
| **3. Peak_Time** | Hourly arrival distribution across all 24 hours of the day (00:00 to 23:00) with volume percentages and peak period callouts. | `=COUNTIF()`, `=INDEX(MATCH(MAX()))`, `=MAX()`, `=SUM()`. Includes a 24-hour arrival column chart highlighting peak arrival windows. |
| **4. Event_Analysis** | Comparison between 4 Event days and 56 Non-Event days covering daily volume, average/median durations, and daily revenue. | `=COUNTIF()`, `=COUNTIFS()`, `=AVERAGEIFS()`, `=SUMIF()`. Includes a comparative bar chart. |
| **5. Weather_Analysis**| Comparison between 10 Rainy days and 50 Non-Rainy days covering daily arrival rates, stay lengths, and revenue impact. | `=COUNTIF()`, `=COUNTIFS()`, `=AVERAGEIFS()`, `=SUMIF()`. Includes a comparative bar chart. |
| **6. Duration_Analysis**| Stay duration profile across 5 zone land-use types (Educational, Hospital, Recreational, Commercial, Transit) with min, max, average, and median. | `=COUNTIFS()`, `=AVERAGEIFS()`, `=MINIFS()`, `=MAXIFS()`. Excludes open and invalid records. Includes duration bar chart. |
| **7. Pivot_Analysis** | The 5 required assessment summaries: 1. Sessions by Zone, 2. Arrivals by Hour, 3. Event vs Non-Event, 4. Rainy vs Non-Rainy, 5. Duration by Zone Type. | Pre-calculated tables, 5 embedded column charts, plus exact PivotTable recipes for Excel for the Web. |
| **8. Summary** | Executive overview summarizing key numerical takeaways, data quality metrics, and a sheet-by-sheet navigation guide. | Plain white layout, cross-sheet formula references, key findings table. |

---

## 3. How to Create Native PivotTables in Microsoft Excel for the Web

When you open `Round3_Parking_Analysis.xlsx` in **Microsoft Excel for the Web**:
1. Click the **`Cleaned_Data`** sheet.
2. Click anywhere inside the data (it is already recognized as Table **`CleanedParkingData`**).
3. On the top ribbon, click **Insert > PivotTable**.
4. In the dialog, leave the source as **`CleanedParkingData`** and choose **New Sheet** (or place it next to the tables on `Pivot_Analysis`).
5. Drag and drop the fields according to the 5 recipes:
   * **Sessions by Zone:** Rows = `location`, Values = `session_id` (Count).
   * **Arrivals by Hour:** Rows = `entry_hour`, Values = `session_id` (Count).
   * **Event Comparison:** Rows = `event_label`, Values = `session_id` (Count), `duration_min` (Average).
   * **Weather Comparison:** Rows = `rain_label`, Values = `session_id` (Count), `duration_min` (Average).
   * **Duration by Zone Type:** Rows = `zone_type`, Values = `duration_min` (Average), Filter = `is_valid_dur` (TRUE).
6. To add a PivotChart: With the PivotTable selected, click **Insert > PivotChart > Clustered Column**.

---

## 4. Verification

Calculations were verified against the underlying raw and processed Python data pipelines. The verification report is saved at:
[`outputs/Round3_Verification_Report.csv`](Round3_Verification_Report.csv).
