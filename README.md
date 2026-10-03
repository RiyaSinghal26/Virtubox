# VirtuBox Data Analyst Assessment Test - Submission Package

**Candidate Role:** Data Analyst  
**Business Domain:** B2B Technology & E-Commerce Retail  
**Dataset:** Kaggle E-Commerce Online Retail Dataset (541,909 raw records)  
**Primary Deliverables Package:** 
1. `VirtuBox_Data_Analyst_Assessment_updated.xlsx` (11 Worksheet Tabs including Raw & Processed Data)
2. `VirtuBox_Analysis_Notebook.ipynb` (Jupyter Notebook with Full EDA, Calculations & Charts)
3. `Executive_Dashboard.html` (Interactive Single-Page HTML Executive Dashboard)
4. `Presentation_Management.pptx` (7-Slide Management Presentation)
5. `processed_data.csv` (Analysis-Ready Cleaned Dataset: 392,692 rows)
6. `process_data.py` & `generate_presentation.py` (Automated Data Pipeline & Slide Generator)
7. `README.md` (Methodology, Looker Studio & HTML Dashboard Architecture, AI Disclosure)

---

## 🔗 Project & Deliverables Directory Structure

```
c:/Users/RIYA/Downloads/Assessment/
│
├── Assessment-Test-(Data-Analyst)-v1 (1) (1).pdf # VirtuBox Assessment Brief
├── data.csv                                     # Raw Kaggle Dataset (541,909 rows × 8 cols)
├── processed_data.csv                            # Analysis-Ready Dataset (392,692 rows × 13 cols)
│
├── VirtuBox_Data_Analyst_Assessment_updated.xlsx # Excel Workbook (11 Tabs: Data, Q1-Q8, Q10)
├── VirtuBox_Analysis_Notebook.ipynb              # Executed Jupyter Notebook with Full EDA & Plots
├── Executive_Dashboard.html                      # Interactive HTML Executive Dashboard
├── Presentation_Management.pptx                  # 7-Slide Executive PowerPoint Presentation
│
├── process_data.py                               # Automated Data Pipeline & Excel Builder
├── generate_presentation.py                      # Script Generating PowerPoint Presentation
├── generate_notebook.py                          # Script Generating Executed Jupyter Notebook
└── README.md                                     # Full Documentation & Assessment Methodology
```

---

## 1. Executive Summary & Business Problem Statement (Questions 1 & 2)

Management of our B2B e-commerce platform seeks to transition from unsegmented mass marketing to a data-driven growth engine. Using 541,909 transactional records covering global sales from December 2010 to December 2011, this project identifies high-leverage growth opportunities, analyzes customer concentration risks, evaluates international expansion channels, and provides actionable operational recommendations.

### Key Data Benchmarks (Cleaned Baseline):
* **Net Revenue:** **$8,887,208.89** across 12 months
* **Cleaned Transactions:** **392,692 order items** (from 541,909 raw records)
* **Unique Active Accounts:** **4,338 buyers**
* **Completed Orders:** **18,532 invoices**
* **Average Order Value (AOV):** **$480.86**
* **Top Revenue Region:** United Kingdom ($7.31M) followed by Netherlands ($285.4K), Ireland ($265.5K), Germany ($228.9K), and France ($209.0K).

---

## 2. Dataset Selection & Quality Profiling (Questions 1 & 6)

* **Dataset Name:** E-Commerce Online Retail Transaction Dataset
* **Source & URL:** [Kaggle E-Commerce Data](https://www.kaggle.com/datasets/carrie1/ecommerce-data)
* **Row/Column Count:** 541,909 rows × 8 raw columns → 392,692 rows × 13 cleaned columns

### Data Quality Issues & Remediation (Question 6):
1. **Missing Customer IDs (135,080 rows / 24.9%):** Dropped for customer cohort/RFM metrics to prevent distorted account calculations. Retained in macro validation.
2. **Order Cancellations (10,624 rows):** Negative quantities starting with invoice 'C' were isolated to calculate clean net revenue and track return loss ratios.
3. **Non-transactional Zero Prices (2,517 rows):** Filtered items with `UnitPrice <= $0` caused by system adjustments or samples.
4. **Duplicate Logs (5,268 rows):** Purged identical duplicate lines to avoid overcounting revenue.

---

## 3. Data Cleaning Pipeline & Transformations (Question 3)

The raw dataset was cleaned and transformed via Python (`process_data.py`) using the following rules:
* `InvoiceDate` converted to standardized DateTime format.
* `CustomerID` cast from float to string integer (`17850`).
* Engineered Features:
  * `TotalAmount` = `Quantity * UnitPrice`
  * `YearMonth` = `YYYY-MM` string for temporal trend analysis
  * `DayOfWeek` = Day name (Monday–Sunday)
  * `Hour` = Integer hour of transaction (0–23)

---

## 4. Key Business Insights & Tested Hypotheses (Questions 2, 4, 5)

### Tested Hypotheses (Question 2):
* **Hypothesis 1 (International B2B Wholesale):** *International customers generate significantly higher Average Revenue Per User (ARPU) than UK domestic customers because overseas buyers operate primarily as wholesale B2B distributors.*  
  👉 **VALIDATED:** Netherlands ARPU is **$31,716** (9 buyers = $285k) and Ireland ARPU is **$88,515** (3 buyers = $265k), compared to UK domestic ARPU of **$1,864** (3,920 buyers = $7.31M).
* **Hypothesis 2 (Q4 Pre-Holiday Bulk Stocking):** *Q4 revenue growth is driven by bulk stocking transactions ahead of Christmas.*  
  👉 **VALIDATED:** November revenue surged by **95% to $1.16M** (vs July's $600k) with average unit quantity per invoice spiking significantly.

### 5 Meaningful Business Insights (Question 4):
1. **Customer Concentration Risk:** Top 10% of customers (433 buyers) generate **61.33% ($5.46M)** of total revenue.
2. **High-Yield Overseas Channels:** Netherlands and Ireland accounts generate 17x to 47x higher revenue per customer than domestic UK buyers.
3. **Q4 Demand Explosion:** November is the peak sales month ($1.16M), requiring advance inventory positioning by September.
4. **Mid-Day Order Window:** 55% of transactions occur between 10 AM and 1 PM (peak hour: 12 PM at $1.38M total revenue).
5. **Top SKU Revenue Reliance:** Top 3 products ('Paper Craft Little Birdie', 'Regency Cakestand 3 Tier', 'White Hanging Heart') generated $411,510 combined.

---

## 5. Dashboard Architecture (Question 8)

This submission includes **two** executive visualization solutions:
1. **Interactive HTML Dashboard (`Executive_Dashboard.html`):** Built with Chart.js & Tailwind CSS for immediate executive presentation and dynamic filtering.
2. **Google Looker Studio Blueprint (Worksheet Q8):** Designed for direct connection to `processed_data.csv`.

### Dashboard Components:
* **Executive Scorecards:** Net Revenue ($8.89M), Orders (18,532), B2B Accounts (4,338), AOV ($480.86).
* **Time-Series Chart:** Monthly Revenue & Order Volume Trend (2010–2011).
* **Concentration Donut Chart:** Revenue split (Top 10% B2B accounts vs Remaining 90%).
* **International ARPU Bar Chart:** Highlighting wholesale performance in Netherlands, Ireland, Australia, Germany, France.
* **Flagship SKUs Chart:** Top 5 revenue-generating items.

---

## 6. Actionable Recommendations (Question 7)

1. **Launch European B2B Wholesale Portals (High Impact / High Feasibility):** Target Netherlands, Germany, and Ireland with localized wholesale portals and dedicated account managers. (Owner: Head of International Growth; KPI: +$1.5M Year 1 revenue).
2. **Establish VIP Key Account Loyalty Tier (High Impact / High Feasibility):** Assign dedicated account management to the top 400 B2B buyers driving 61% of sales. (Owner: Director of Customer Success; KPI: 50% churn reduction).
3. **Implement Q4 Inventory & Logistics Surge Planning (Medium Impact / High Feasibility):** Advance inventory POs by 45 days for top 20 SKUs before October. (Owner: Supply Chain Operations Lead; KPI: <2% stockout rate).

---

## 7. AI Disclosure Statement (Question 10)

* **AI Tools Used:** Gemini 3.6 Flash (High), Python Data Science Stack (Pandas, OpenPyXL, PPTX, NbConvert).
* **Applications:** Automating data pipeline generation, creating interactive HTML executive dashboard, structuring PowerPoint slides, and generating executed Jupyter notebook cells.
* **Verification Example:** Verified that order cancellations (`InvoiceNo` starting with 'C') were isolated rather than dropped globally to preserve accurate net revenue math and calculate return ratios cleanly.
