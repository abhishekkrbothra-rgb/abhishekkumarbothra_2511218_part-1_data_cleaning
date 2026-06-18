Business Data Cleaning, Validation & Excel Reporting (Part 1)

Student Name: Abhishek Kumar Bothra 
Student ID: 2511218  
Repository Assignment: Part 1 - Individual Submission

---

1. Business Problem Summary
The company’s internal ordering database handles large legacy exports from multiple regional ERP systems. The raw system download (`raw_orders.xlsx`) arrived with severe operational anomalies: compromised date entries, corrupted regional structures, blank-space, duplicated inputs, and mathematical discrepancies across sales records. 

As the Business Analyst, my objective was to engineer an authenticated, production-ready master sheet (`cleaned_orders.xlsx`), reconcile systemic transaction errors, execute data validation audits against company policy, and provide high-level visual summaries to corporate stakeholders.

---

2. Dataset Description
The dataset contains transaction-level historical data monitoring **912 total operational order lines**. The tracking data spans multiple columns including:
- Identifiers:`order_id`, `customer_id`, `product_name`
- Chronology:`order_date`, `ship_date`
- Demographics:`customer_name`, `segment`, `region`, `state`, `city`
- Product Hierarchy:`category`, `sub_category`
- Financial Metrics: `quantity`, `unit_price`, `discount`, `sales`, `cost`, `profit`
- System Flags:`payment_status`, `order_status`

---

3. Tools Used
- Microsoft Excel: Selected as the primary data engine. Utilized logical criteria text functions (`PROPER`, `TRIM`), date re-serialization utilities, conditional filtering blocks, financial arithmetic columns, and executive Pivot Tables.
- GitHub Desktop & Git Markdown: Leveraged to construct the standalone version-controlled documentation and folder pathways.

---

4. Cleaning Steps Performed
To transition the un-vetted database records into audit-ready datasets, the following processes were applied:
1. Text Trailing & Case Uniformity:** Built a temporary workspace column using `=PROPER(TRIM(cell))` across 10 text fields (`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`). This permanently scrubbed double spaces and unified mismatched cases before locking variables as text values.
2. Date De-Serialization:** Applied the **Text to Columns Wizard** on `order_date` and `ship_date` to normalize inconsistent system strings into uniform chronological dates.
3. Duplicate Cleanse:** Executed an automated rows check. Identical entries were isolated, extracting duplicate copies safely while leaving unique core rows untouched.
4. Calculated Field Engineering:** Generated specialized columns with live mathematical formulas:
   - `shipping_delay_days`: `=ship_date - order_date`
   - `order_month`: `=TEXT(order_date, "mmmm")`
   - `order_year`: `=YEAR(order_date)`
   - `calculated_sales`: `=quantity * unit_price * (1 - cleaned_discount)`
   - `calculated_profit`: `=calculated_sales - cost`
   - `profit_margin`: `=calculated_profit / calculated_sales`

---

5. Business Rules Applied
- Rule 1 (Data Segregation):Preserved `raw_orders.xlsx` intact inside `data/` to provide full data lineage transparency. 
- Rule 2 (Structural Fallbacks):For rows missing critical `region` or `ship_mode` attributes, the cell was manually populated with `"Unknown"` to keep data queries operational.
- Rule 3 (Financial Audit Reconciliation):Identified corrupted figures in the raw `sales` column. Overrode legacy data cells with correct arithmetic formulas to prevent a ₹71,167.63 financial error.
- Rule 4 (Revenue Integrity Filter):Enforced a strict status filter on final dashboards, excluding transaction lines marked as `Cancelled`, `Returned`, or `Failed` from active completed sales metrics.

---

6. Summary of Data Quality Issues Found
A complete audit profile was generated inside 'data_quality_report.xlsx' outlining the following system issues:
- Clean Records:"761 rows" passed all strict data validation layers perfectly.
- Warning Records:"45 rows" were flagged with structural gaps (24 missing regions, 20 missing ship modes, and 1 record missing both fields simultaneously) and successfully resolved to "Unknown".
- Invalid Records:"106 rows" contained fatal chronological exceptions (86 instances where products shipped before the order date) or corrupted out-of-bounds promotional discounts (20 instances).
- Duplicates Handled:"20 exact duplicate rows" were completely removed, and "12 conflicting Order IDs" were flagged for further system review.

---

7. Summary of Final Pivot Reports
The workbook 'outputs/pivot_summary.xlsx' structures regional performance via interactive reporting sections:
1. Regional Sales & Profit Matrix: Displays regional metrics showing that the "South" and "East" regions lead gross activity.
2. Category Performance Hierarchy: Breaks down sales volume by product category and sub-category, highlighting that "Furniture and Technology" generates the largest revenue streams (driven by Chairs and Copiers).
3. Order Status Anomaly Map: Isolates operational disruptions by region (Failed payments and Cancellations) to evaluate supply chain risk.
4. Segment Profitability Review: Details average profit margins grouped across customer segments (`Consumer`, `Corporate`, `Home Office`, `Small Business`).
5. Year-Wise Monthly Sales Trend: Tracks month-over-month growth trajectories across financial years.

---

8. Key Business Insights
- The True Financial Footprint: Legacy systems over-reported revenue. By executing audited calculated formulas, the actual sales total was corrected to ₹9,008,422.84 (resolving a calculation variance of ₹71,167.63).
- High-Risk Operations: The "North" region exhibits severe transactional vulnerability, accounting for a high concentration of failed payment processing instances and cancelled orders. Security protocols and payment gateway stability should be investigated here immediately.
- Profitability Dynamics: "Furniture" items generate large top-line volumes and their margins are also higher compared with other segment, which functions as the primary engine for net corporate margins.

---

9. Assumptions and Limitations
- Assumptions: Rows containing missing regions or shipping methods were assumed to be valid business transactions rather than data dropouts, allowing them to be retained under the `"Unknown"` designation. Blank discount values were treated as intentional zero-discount sales.
- Limitations: Because this optimization project relies entirely on a flat-file database export, true root-cause tracking for why 86 orders backwards-shipped cannot be determined without analyzing live server application logs.

---

10. Repository Screenshots Included
All visual evidence is saved directly inside the 'screenshots/' directory:
- 'raw_data_preview.png': Captured initial messy states before transformation.
- 'cleaned_data_preview.png': Displays new engineered columns and flags.
- 'pivot_summary_1.png': Demonstrates regional sales breakdown and category hierarchies.
- 'pivot_summary_2.png': Documents anomaly matrices and historical revenue trends.