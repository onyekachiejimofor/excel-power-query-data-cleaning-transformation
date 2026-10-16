# Excel Power Query | Data Cleaning, Consolidation & Transformation

**Three Excel Power Query portfolio projects | LumaWorks Retail (fictional business case)**

A practical, refreshable Excel Power Query portfolio demonstrating how messy and fragmented business records can be prepared for reliable reporting. The work uses **customized synthetic training data** organized around LumaWorks Retail, a fictional business operating marketing campaigns, a puzzle-products division, and a specialty-coffee line.

The focus is on **data operations and repeatable data preparation**: source ingestion, data-quality improvements, folder-based consolidation, reshaping, and verification. These are educational portfolio exercises, **not work performed for an actual client**.

## At a glance

| Project | Business challenge | Power Query solution | Verified outcome |
|---|---|---|---|
| **01 — Marketing Data Cleaning** | Grouped labels, inconsistent text, sparse metrics | Promote headers, fill down, standardize text, preserve unknowns, calculate a ratio | **154** marketing records retained |
| **02 — Multi-Year Sales Consolidation** | Annual product sales stored in separate CSV files | Combine all CSVs from a folder, standardize types, add a quarter field, refresh | **480** records across **2023–2026**, 120 per year |
| **03 — Monthly Coffee Sales Restructuring** | Product-by-month sales laid out as a wide matrix | Import a workbook named range, unpivot monthly columns, assign reporting dates, summarize in PivotTable | **240** product-month records; **59,773** total units sold |

## Project details

### 01 — Marketing Data Cleaning

**Business need:** Prepare campaign performance records for consistent operational review, despite blank group labels, formatting inconsistencies, and missing units-sold values.

**Workflow:** Imported an Excel sheet into Power Query, removed excess header rows and separator rows, promoted proper headers, filled down grouping information, normalized text values, set appropriate data types, and computed `Conversion = Units Sold / Clicks` where inputs were present.

**Data-quality checks:** 154 observations retained; campaign names and channel locations populated; three missing `Units Sold` values intentionally kept as unknown rather than interpreted as zero. The `Conversion` metric is a **units-per-click ratio**, not a verified order-level/customer-level conversion rate.

**Deliverable:** [`LumaWorks_Retail_Marketing_Data_Cleaned.xlsx`](LumaWorks_Retail_Marketing_Data_Cleaned.xlsx)

### 02 — Multi-Year Sales Consolidation

**Business need:** Replace manual annual workbook updates with a reliable consolidation process.

**Workflow:** Power Query reads yearly CSVs from one source folder, applies common transformations to incoming files, creates a `Quarter` field from each date, and loads the combined table. A 2026 CSV was added after configuring the initial years to demonstrate that **Refresh All** incorporates newly added files.

**Data-quality checks:** 120 observations per year for 2023, 2024, 2025, and 2026; **480** records overall; no missing or duplicate product-month keys in the validated set.

**Deliverable:** [`LumaWorks_Retail_Multi_Year_Sales_Consolidation.xlsx`](LumaWorks_Retail_Multi_Year_Sales_Consolidation.xlsx)

### 03 — Monthly Coffee Sales Restructuring

**Business need:** Turn 12 month columns across 20 products into a consistent row-based reporting dataset.

**Workflow:** Reads the Excel named range `SalesData` with `Excel.CurrentWorkbook()`, promotes headers, unpivots months to rows, standardizes the month-year date, types `Units Sold`, and supports quarterly reporting through a PivotTable. Subsequent months were incorporated into the source range and the query was refreshed.

**Data-quality checks:** 20 products × 12 months = **240** records; quarterly totals are 15,568, 14,775, 14,131, and 15,299 units, respectively, totaling **59,773**. The workbook's saved result includes the quarterly PivotTable; a separate PivotChart is not claimed here.

**Deliverable:** [`LumaWorks_Retail_Monthly_Coffee_Sales_Unpivoted.xlsx`](LumaWorks_Retail_Monthly_Coffee_Sales_Unpivoted.xlsx)

## Repository contents

```text
excel-power-query-data-cleaning-transformation/
├── README.md
├── LumaWorks_Retail_Marketing_Data_Cleaned.xlsx
├── LumaWorks_Retail_Multi_Year_Sales_Consolidation.xlsx
├── LumaWorks_Retail_Monthly_Coffee_Sales_Unpivoted.xlsx
└── source_data/
    ├── Marketing_Data.xlsx
    └── Sales/
        ├── RubiksCubeSales_2023.csv
        ├── RubiksCubeSales_2024.csv
        ├── RubiksCubeSales_2025.csv
        └── RubiksCubeSales_2026.csv
```

*The `source_data` directory is needed for external-file refresh. Include only source files you have the right to redistribute publicly.*

## Reproduce the transformations

**Requirements:** Microsoft Excel for Windows with Power Query. Download/clone the **entire repository**, preserving the folder layout shown above, then open the `.xlsx` workbooks in Excel desktop.

- **Marketing:** The workbook's `Config` sheet contains `tblSourcePath`, which derives the location of `source_data\Marketing_Data.xlsx` from the workbook's current saved directory. Check the path, then choose **Data → Refresh All**. Expected output: **154** rows.
- **Sales:** Its `Config` table similarly derives the path to `source_data\Sales`. Both the `Sales` and `Sample File` helper queries must use this location. Choose **Data → Refresh All**. Expected output: **480** rows.
- **Coffee:** The query uses the internal `SalesData` named range; no external directory is required. Refresh to reproduce **240** records and a **59,773-unit** quarterly grand total.

The two external-source workbooks must sit **in the repository root beside `source_data/`**. Moving only the `.xlsx` files, without their source folder, breaks refresh. The dynamic path formulas use `CELL("filename")`, so save/reopen the workbooks and recalculate formulas if the directory changes. Excel may ask you to authorize data connections or adjust privacy permissions; enable connections only when you trust the files.

## Verification and limitations

- Source files and results use fictional, synthetic/adapted training records; figures are not real LumaWorks transactions.
- The marketing project retains truly unknown units-sold values as missing; they were not imputed as zero.
- The coffee workbook uses a **fixed-size named range**; adding further rows beyond its boundaries requires expanding that range.
- `.xlsx` files contain the Power Query transformations. GitHub may not show interactive Excel connections in a browser preview: download and open in Excel desktop to inspect **Data → Queries & Connections → Applied Steps**.
- The portfolio folder was moved to another location and its refreshable connection setup was tested in Excel by the project author.

## Learning-source attribution

The original instructional workflow and exercise structure were adapted from **XelPlus / Leila Gharani**, *Learn to Automate Everything with Power Query in Excel*. The LumaWorks framing, customized synthetic records, execution, validation, reproducibility configuration, and portfolio documentation are presented as portfolio work. Any third-party training materials remain the property of their respective owners; source files should be published only if redistribution is authorized.

## Skills demonstrated

**Excel Power Query · Data Cleaning · Data Transformation · Data Quality Assurance · Folder-Based Data Consolidation · Refresh Automation · Unpivot · PivotTables · Data Operations**
