# Customer Transactions Analysis (Excel)

A full analysis and reporting solution in Excel. It analyses 250 customer transactions, finds insights, and presents them in a dashboard and a written report.

**Workbook:** `Customer_Transactions_Analysis.xlsx`

---

## Dataset

Sheet: **Final Project Dataset** (250 rows, calendar year 2025, 79 customers, 10 products, 5 regions)

| Column | Description |
|---|---|
| `transaction_id` | Unique ID (T0001 to T0250) |
| `date` | Transaction date |
| `product` | Laptop, Smartphone, Monitor, Tablet, Smartwatch, Printer, Headphones, Webcam, Keyboard, Mouse |
| `quantity`, `unit_price`, `sales` | Units, price per unit, and quantity × price (₹) |
| `region` | North, South, East, West, Central |
| `profit` | Profit on the transaction (₹) |
| `customer_id`, `customer_name` | Customer identifiers (names are generated sample data) |
| `payment_method` | Credit Card, Debit Card, UPI, Cash, Net Banking |

Columns L to T are formula columns: Month End, Months Ago, Days Ago, Profit Margin (with arrow icons), Customer Type, Txn Code, Initials, Filter Flag and Filter Match #.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Dashboard** | Filters, KPIs with indicators, Top 5 customers, and 4 charts |
| **Report** | Final report with insights, storytelling and recommendations (numbers are live formulas) |
| **Final Project Dataset** | Source data plus helper formula columns |
| **Pivot_Tables** | Pivot-style summaries by product, region, payment method and month |
| **FILTER_Demo** | Multi-value return based on Region, Product and minimum sales |
| **High_Value_Customers** | GroupBy by customer, tiers, and top 10 |
| **Frequency_Match** | Most frequent product, repeating names, and list comparison |
| **Text_Abbrev** | Abbreviations built with text functions |
| **Date_Time** | TODAY, NOW, DATEDIF, EOMONTH and the auto-timestamp |
| **WhatIf** | Scenario Manager scenarios and Goal Seek |
| **Regression** | Linear regression (ToolPak-style output) and scatter chart |
| **HowTo** | Steps for the native Excel versions of the interactive features |

---

## Requirements Coverage

| Requirement | Implementation |
|---|---|
| TODAY, NOW, DATEDIF, EOMONTH | `Date_Time` sheet and dataset columns L to N |
| FILTER for multi-value returns | `FILTER_Demo`: helper columns plus INDEX/MATCH (works in all Excel versions). The Excel 365 `FILTER` formula is shown beside it |
| Conditional formatting with icons and arrows | Profit margin arrows (dataset), KPI variance arrows (dashboard), tier colours and sales arrows (customers) |
| Timestamp with NOW and iterative calculation | `Date_Time`: `=IF(A<>"",IF(OR(B="",B=0),NOW(),B),"")`. Iterative calculation is saved on in the file |
| Scenario Manager and Goal Seek | `WhatIf`: Base, Optimistic and Pessimistic scenarios saved on B11:B13, and a closed-form Goal Seek answer in B26 |
| Linear regression (Data Analysis ToolPak) | `Regression`: Profit on Sales, with regression statistics, ANOVA, coefficients, 95% limits and a trendline chart |
| High-value customers (GroupBy) | `High_Value_Customers`: COUNTIF/SUMIF per customer, percentile threshold (adjustable), tiers |
| Dashboard: pivot tables, pivot charts, slicers, timeline, KPI indicators | `Pivot_Tables` + `Dashboard`: SUMIFS tables, bar/line/pie/column charts, drop-down filters, From/To dates |
| Most frequently purchased product, repeating names | `Frequency_Match` |
| Compare two lists and extract matches | `Frequency_Match`: a mailing list is matched against customer names, with a compact extracted list |
| Abbreviations with TEXT functions | `Text_Abbrev` and the Txn Code and Initials columns |
| Final report with storytelling | `Report` sheet |

---

## How to Use

1. **Dashboard filters:** choose Region and Product from the drop-downs and set the From/To dates (cells B3 to B6). All KPIs, tables and charts update.
2. **KPI targets:** edit the blue-on-yellow target cells (C15:C18). The arrows and status labels update.
3. **High-value threshold:** change the percentile in `High_Value_Customers!B3` (default 80%).
4. **FILTER demo:** change Region, Product and minimum sales in `FILTER_Demo!B3:B5`.
5. **Timestamps:** type in `Date_Time!A19:A26`. A timestamp is stamped once and stays fixed. Clear the entry and retype it to re-stamp.
6. **Scenarios:** Data > What-If Analysis > Scenario Manager, select a scenario and click **Show**.
7. **Goal Seek:** Data > What-If Analysis > Goal Seek. Set cell `B18`, to the value in `B25`, by changing `B12`.
8. **Prediction:** enter a sales amount in `Regression!B21` to get a predicted profit.

Inputs are shown as blue text on a yellow fill.

---

## Key Findings

- Total sales are **₹52.6 lakh** and profit is **₹8.0 lakh**, a **15.2%** margin.
- **Laptops** lead on revenue (26% of sales). **Smartphones and Tablets** have the thinnest margins (about 12%). Mouse, Keyboard and Headphones have the best (26% to 29%).
- Monitor and Headphones tie as the most frequently purchased products (30 transactions each).
- **West** is the strongest region and **North** the weakest. **March** is the peak month and **September** the low.
- The top 20% of customers (16 people) generate about **46%** of sales.
- Regression: each extra ₹1,000 of sales adds about **₹137** of profit (R² = 0.893).

The recommendations in the Report sheet are written for the data as delivered. Re-read them if you change the data.

---

## Limitations and Notes

- **Native Excel objects:** real PivotTables, PivotCharts, Slicers, Timelines and the Analysis ToolPak output can only be created inside Excel. This workbook provides formula-based equivalents, and the **HowTo** sheet lists the steps for the native versions.
- **Filters** on the dashboard are drop-downs and date cells, not Slicer objects.
- **Compatibility:** built with functions available in Excel 2010 and later (`MAXIFS` needs Excel 2019 or Microsoft 365). Opening in Excel recalculates all formulas.
- **Circular reference warning:** if Excel warns about it, enable *File > Options > Formulas > Enable iterative calculation*.
- **Dates:** `TODAY()` and `NOW()` are volatile, so Months Ago, Days Ago and the "as of" date change each day.
- **Customer names** are generated sample data. The earlier `transactions.csv` does not include them.
