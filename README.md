# Enterprise Audit Pipeline & Automated Fiscal Analytics Engine

## 📌 Project Overview
This project solves a critical scalability and technical friction issue in financial auditing: the inability of audit teams to process 100% of transaction logs due to software row limits (e.g., Excel constraints) and database access barriers.

I designed and implemented an end-to-end ETL and automated reporting infrastructure that integrates heterogeneous data sources, applies automated risk-filtering, and distributes institutional audit working papers at high speed.

*   **Business Impact:** Reduced processing and reporting timeline from **5–7 business days to under 2 hours** (a 95% efficiency boost).
*   **Data Scale:** Eliminated traditional sample-based auditing, enabling **100% data coverage** over datasets exceeding **20 Million records** per fiscal year.
*   **User Empowerment:** Removed technical friction for non-technical tax auditors by delivering one-click, fully formatted compliance reports.

---

## 🛠️ Tech Stack & Architecture
*   **Data Integration & ETL:** KNIME Analytics Platform
*   **Database Management:** Oracle SQL (Query Optimization & Aggregation)
*   **Automation Engine:** VBA (Excel Advanced Macros & Memory Optimization)
*   **Data Silos:** Oracle DB, Large-Scale CSVs, and Excel Sheets


[Oracle DB / CSV / Excel] ➔ [KNIME ETL & Materiality Filter] ➔ [Oracle SQL Integrity Anti-Join] ➔ [In-Memory VBA Reporting Engine] ➔ [135 Institutional Reports]


---

## 🚀 Technical Deep-Dive

### 1. Ingestion & Materiality Filtering (KNIME Workflow)
Audit information originated from three distinct silos, frequently causing memory crashes on auditors' workstations. 
*   **Heterogeneous ETL:** Configured multi-source ingestion nodes in KNIME to standardize schemas across Oracle DB, CSV, and Excel tables.
*   **Risk Isolation:** Implemented logical data pipelines using materiality thresholds and tax risk parameters. Instead of forcing manual filtering, the workflow isolates high-risk sub-ledgers, ensuring data integrity while reducing file sizes for output stages.

### 2. Revenue Leakage & Fraud Audit (Oracle SQL)
To guarantee that the audit team analyzes only *realized revenue*, the extraction layer uses an advanced **Left Join + Anti-Join pattern** to scrub the data.

*   **The Problem:** Traditional queries pull raw invoice totals, masking revenue distortions caused by credit notes.
*   **The Logic:** The query consolidates all credit note references (`invoice_id` and `origin_doc_id`) via a unified subquery. It then isolates active transactions by filtering out inverted documents (`cn.doc_reference IS NULL`).
*   **Data Defense:** Applied `UPPER(TRIM())` functions across join keys to mitigate trailing spaces and casing anomalies caused by manual system entries.

👉 *View the full script in:* [`/sql_queries/clean_sales_metrics.sql`](./sql_queries/clean_sales_metrics.sql)

### 3. High-Performance Reporting Engine (VBA)
Once filtered data is exported into a master Excel workbook, a custom-engineered VBA macro handles distribution, cell formatting, and reporting layouts.

*   **UI Overhead Deactivation:** Bypasses Excel’s graphical rendering and automatic calculation pipelines using `Application.ScreenUpdating = False` and `xlCalculationManual`. This drastically reduces CPU cycles.
*   **In-Memory Exception Logging:** To maximize hard drive I/O speed, missing or zero-transaction account codes are buffered directly into a dynamic array memory structure (`errorLogArray`). A separate exception workbook is created and saved asynchronously *only* if errors are caught.
*   **Dynamic Layout Processing:** Calculates spatial constraints (`TargetRowOffset`, `HeaderRowOffset`) on the fly based on the variable length of the filtered general ledger dataset (`VisibleRowCount`). This preserves structural corporate styling and mathematical sum accuracy without rigid templates.
*   **Performance Metric:** Generates **135 distinct, fully-formatted institutional reports** (with monthly accounting breakdowns and automated sums) in **70 seconds**.

👉 *View the source code in:* [`/vba_macros/automated_ledger_reporting.bas`](./vba_macros/automated_ledger_reporting.bas)

---

## 📊 Quantifiable Business Outcomes
*   **Risk Mitigation:** 100% transaction coverage completely eliminated auditing blind spots caused by random sampling.
*   **Operational Velocity:** The audit department repurposed **38+ hours of manual work per month** from report generation into strategic fiscal analysis.
*   **Zero-Friction Adoption:** Provided non-technical tax inspectors with an abstract, single-click solution, bypassing the need for them to learn SQL or ETL architecture.
