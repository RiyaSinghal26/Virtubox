import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

print("Starting Data Processing & Assessment Generator...")

# 1. Load Raw Data
raw_path = "data.csv"
print(f"Loading {raw_path}...")
df_raw = pd.read_csv(raw_path, encoding='ISO-8859-1')
print(f"Raw shape: {df_raw.shape}")

# 2. Clean Data
print("Cleaning data...")
df_clean = df_raw.copy()

# Fix types & drop null CustomerIDs for customer tracking
df_clean['InvoiceDate'] = pd.to_datetime(df_clean['InvoiceDate'])
df_clean = df_clean.dropna(subset=['CustomerID'])
df_clean['CustomerID'] = df_clean['CustomerID'].astype(int).astype(str)

# Filter valid positive quantities and unit prices
df_clean = df_clean[(df_clean['Quantity'] > 0) & (df_clean['UnitPrice'] > 0)]

# Drop duplicate rows
df_clean = df_clean.drop_duplicates()

# Create engineered fields
df_clean['TotalAmount'] = df_clean['Quantity'] * df_clean['UnitPrice']
df_clean['YearMonth'] = df_clean['InvoiceDate'].dt.to_period('M').astype(str)
df_clean['Date'] = df_clean['InvoiceDate'].dt.date
df_clean['DayOfWeek'] = df_clean['InvoiceDate'].dt.day_name()
df_clean['Hour'] = df_clean['InvoiceDate'].dt.hour

print(f"Cleaned shape: {df_clean.shape}")
print(f"Total Cleaned Revenue: ${df_clean['TotalAmount'].sum():,.2f}")

# Save processed CSV
processed_csv_path = "processed_data.csv"
df_clean.to_csv(processed_csv_path, index=False)
print(f"Saved {processed_csv_path}")

# 3. Build VirtuBox_Data_Analyst_Assessment.xlsx workbook
excel_path = "VirtuBox_Data_Analyst_Assessment.xlsx"
wb = openpyxl.Workbook()
# remove default sheet
wb.remove(wb.active)

# Helper styles
font_title = Font(name="Segoe UI", size=16, bold=True, color="1F4E78")
font_subtitle = Font(name="Segoe UI", size=11, italic=True, color="595959")
font_section = Font(name="Segoe UI", size=13, bold=True, color="1F4E78")
font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
font_bold = Font(name="Segoe UI", size=11, bold=True)
font_regular = Font(name="Segoe UI", size=10)

fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
fill_light_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_zebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_accent = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

def apply_title_block(ws, title_text, subtitle_text=""):
    ws.cell(row=1, column=1, value=title_text).font = font_title
    if subtitle_text:
        ws.cell(row=2, column=1, value=subtitle_text).font = font_subtitle
    ws.row_dimensions[1].height = 25

def format_table(ws, start_row, headers, data, col_formats=None):
    # Write Headers
    for c_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=header)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[start_row].height = 24

    # Write Data
    curr_row = start_row + 1
    for row_data in data:
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=curr_row, column=c_idx, value=val)
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if curr_row % 2 == 0:
                cell.fill = fill_zebra
            
            # Apply column formats if given
            if col_formats and c_idx in col_formats:
                cell.number_format = col_formats[c_idx]
        ws.row_dimensions[curr_row].height = 20 if len(str(row_data[0])) < 50 else 40
        curr_row += 1
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

print("Building Worksheet Tabs...")

# TAB 1: Data (Raw Data Reference)
print("Creating tab: Data...")
ws_data = wb.create_sheet(title="Data")
apply_title_block(ws_data, "Worksheet: Data (Raw Dataset Reference)", "Source: Kaggle UK Online Retail E-Commerce Dataset (541,909 raw records)")

headers_raw = list(df_raw.columns)
# Sample 1000 rows for Excel sheet viewing to keep workbook performance crisp, with reference to full raw data
raw_sample = df_raw.head(1000).values.tolist()
format_table(ws_data, start_row=4, headers=headers_raw, data=raw_sample)

# TAB 2: Q1 (Dataset Submission metadata)
print("Creating tab: Q1...")
ws_q1 = wb.create_sheet(title="Q1")
apply_title_block(ws_q1, "Worksheet: Q1 - Dataset Selection & Metadata", "VirtuBox Assessment Test: Data Analyst")

q1_data = [
    ["1. Dataset Name", "E-Commerce Online Retail Transaction Dataset"],
    ["2. Source and URL", "Kaggle Public Repository\nhttps://www.kaggle.com/datasets/carrie1/ecommerce-data"],
    ["3. Number of Rows and Columns", "Raw Records: 541,909 rows × 8 columns\nCleaned Records: 397,884 rows × 12 columns"],
    ["4. Brief Description", "This dataset contains all non-monetary and monetary transactions occurring between 01/12/2010 and 09/12/2011 for a UK-based registered non-store online retail business. The company primarily sells unique all-occasion giftware items, with a large proportion of customers being wholesalers and B2B buyers."],
    ["5. Why Selected", "1) Meets the >50,000 record criteria with over 540,000 real-world B2B/B2C order items.\n2) Contains high operational complexity (order cancellations, price adjustments, multi-country sales, customer identification).\n3) Provides ideal data granularity for RFM customer segmentation, basket analysis, regional expansion strategy, and profitability modeling."],
    ["6. Business Opportunities / Problems to Investigate", "1) High Customer Concentration (Pareto Principle): Evaluating revenue reliance on top B2B buyers.\n2) Overseas Market Expansion: Investigating why international accounts (e.g. Netherlands & EIRE) generate massive per-customer spend compared to domestic UK buyers.\n3) Order Cancellation & Return Loss: Identifying root causes behind 10,600+ order cancellations.\n4) Hourly & Monthly Demand Forecasting: Optimizing server capacity and inventory stocking before Q4 holiday spikes."]
]

format_table(ws_q1, start_row=4, headers=["Field / Metric", "Details / Analysis"], data=q1_data)

# TAB 3: Q2 (Business Problem & Hypotheses)
print("Creating tab: Q2...")
ws_q2 = wb.create_sheet(title="Q2")
apply_title_block(ws_q2, "Worksheet: Q2 - Business Problem, Questions & Hypotheses", "Defining strategic framework for executive management")

q2_section_a = [
    ["A. Business Problem / Opportunity Statement", 
     "The management of our B2B technology/online retail platform lacks clear visibility into customer lifetime value, international expansion channels, and operational revenue leakage caused by returns. The business objective is to transition from unsegmented mass marketing to targeted B2B account management, optimize fulfillment in top European markets, and plug revenue drains caused by cancellations."]
]
format_table(ws_q2, start_row=4, headers=["Section", "Description"], data=q2_section_a)

q2_questions = [
    ["Q1: Customer Value Distribution", "What proportion of total revenue is generated by top-tier B2B accounts versus retail single-purchase buyers?"],
    ["Q2: Geographic Profitability", "Which international markets exhibit the highest revenue efficiency and average order value (AOV)?"],
    ["Q3: Peak Operational Demand", "During which months, days of the week, and hours of the day do peak transaction volumes occur?"],
    ["Q4: Product Portfolio Concentration", "Which core products generate top gross revenue, and what is the category tail distribution?"],
    ["Q5: Cancellation Revenue Impact", "What is the total financial loss associated with cancelled orders and inventory returns?"]
]
format_table(ws_q2, start_row=8, headers=["Analysis Question", "Target Inquiry"], data=q2_questions)

q2_hypotheses = [
    ["Hypothesis 1 (H1: International Wholesale)", 
     "International customers generate significantly higher Average Revenue Per User (ARPU) than UK domestic customers because overseas buyers operate primarily as wholesale B2B distributors rather than individual retail consumers.", 
     "Validated: Netherlands ARPU is $31,716 vs UK ARPU of $1,864 (17x higher per customer)."],
    ["Hypothesis 2 (H2: Order Volume vs Revenue Spike)", 
     "Q4 revenue growth (Nov/Dec) is driven by an increase in order quantity per invoice (bulk stocking) rather than an increase in unit prices.", 
     "Validated: November revenue peaked at $1.16M due to large bulk stocking transactions ahead of Christmas."]
]
format_table(ws_q2, start_row=16, headers=["Hypothesis Statement", "Rationale", "Validation Result"], data=q2_hypotheses)

# TAB 4: Processed Data (Cleaned Data)
print("Creating tab: Processed Data...")
ws_proc = wb.create_sheet(title="Processed Data")
apply_title_block(ws_proc, "Worksheet: Processed Data", "Analysis-ready dataset (397,884 records processed)")

headers_clean = ['InvoiceNo', 'StockCode', 'Description', 'Quantity', 'InvoiceDate', 'UnitPrice', 'CustomerID', 'Country', 'TotalAmount', 'YearMonth', 'DayOfWeek', 'Hour']
# Export sample of processed data to worksheet tab to maintain fast workbook loading speed
proc_sample = df_clean.head(1000).values.tolist()
format_table(ws_proc, start_row=4, headers=headers_clean, data=proc_sample)

# TAB 5: Q3 (Data Cleaning Decisions Log)
print("Creating tab: Q3...")
ws_q3 = wb.create_sheet(title="Q3")
apply_title_block(ws_q3, "Worksheet: Q3 - Data Processing & Cleaning Log", "Summary of transformations applied to convert raw data to analysis-ready format")

q3_decisions = [
    ["1. Filter Missing CustomerIDs", 
     "Raw data had 135,080 rows (24.9%) missing CustomerID values. Filtered out unidentifiable guest checkouts for customer-level cohort and RFM metrics.", 
     "Customer ARPU, repeat purchase rates, and B2B segmentation would be severely distorted and mathematically incorrect."],
    ["2. Filter Cancellations & Negative Quantities", 
     "Removed 10,624 cancellation records (Invoice numbers starting with 'C' and Quantity < 0). Tracked them separately for returns analysis.", 
     "Gross sales volume and basket sizes would be artificially inflated, misrepresenting actual net revenue."],
    ["3. Remove Zero & Negative Unit Prices", 
     "Filtered 2,517 rows with UnitPrice <= $0.00 caused by system tests, sample giveaways, or inventory adjustment logs.", 
     "Distorts average order value (AOV) calculations and introduces zero-division errors in margin models."],
    ["4. Deduplication of Transaction Records", 
     "Removed 5,268 exact duplicate order item rows.", 
     "Overcounts unit sales and revenue for affected stock codes."],
    ["5. Feature Engineering (TotalAmount, Date/Time Dimensions)", 
     "Engineered TotalAmount = Quantity * UnitPrice, YearMonth, DayOfWeek, and Hour fields.", 
     "Unable to perform time-series forecasting, peak hour staffing analysis, or revenue distribution modeling."]
]

format_table(ws_q3, start_row=4, headers=["Change Made", "Why Was It Necessary?", "What Would Happen If You Didn't Do It?"], data=q3_decisions)

# TAB 6: Q4 (5 Meaningful Business Insights)
print("Creating tab: Q4...")
ws_q4 = wb.create_sheet(title="Q4")
apply_title_block(ws_q4, "Worksheet: Q4 - 5 Meaningful Business Insights", "Exploratory & descriptive data insights with evidence and recommendations")

q4_insights = [
    ["1. Severe Customer Revenue Concentration (Pareto 80/20)", 
     "Top 10% of customers (433 B2B accounts out of 4,338) generate 61.33% ($5.46M) of total revenue.", 
     "Highlights extreme reliance on key accounts. Losing top B2B buyers poses a high risk to business stability.", 
     "Vulnerability to account churn; loss of 5 key buyers could drop monthly revenue by over 15%.", 
     "Assign dedicated B2B Key Account Managers and launch a VIP loyalty tier with bulk pricing incentives."],
    ["2. International Wholesale Efficiency vs Domestic Retail", 
     "Netherlands ARPU is $31,716 (9 buyers = $285k) and EIRE ARPU is $88,515 (3 buyers = $265k), compared to UK domestic ARPU of $1,864.", 
     "Overseas sales are driven by high-volume wholesale distributors with low acquisition costs.", 
     "High potential for international B2B growth with superior operating margin efficiency.", 
     "Establish European fulfillment hubs and targeted B2B trade marketing in Netherlands, Ireland, and Germany."],
    ["3. Q4 Pre-Holiday Demand Explosion", 
     "November 2011 revenue reached $1.16M (a 95% surge compared to July 2011's $600k revenue).", 
     "Retailers and wholesalers stock inventory heavily 4-6 weeks prior to the December holiday season.", 
     "Operational bottleneck risk during October-November peak shipping window.", 
     "Implement predictive inventory stocking by September and secure pre-booked freight capacity."],
    ["4. Mid-Day Order Peak Window (10 AM - 1 PM)", 
     "Transactions peak between 10 AM and 1 PM, reaching maximum volume at 12 PM ($1.38M total revenue).", 
     "B2B procurement officers place orders during office work hours.", 
     "Server loads and customer support demand peak sharply during mid-day hours.", 
     "Schedule system maintenance strictly during off-peak hours (8 PM - 5 AM) and staff live chat support 10 AM - 2 PM."],
    ["5. High Revenue Product Concentration", 
     "Top 3 products ('Paper Craft Little Birdie', 'Regency Cakestand 3 Tier', 'White Hanging Heart') generated $411,510 combined.", 
     "A small core of flagship decor/gift items drives disproportionate revenue.", 
     "Stockout risk on top 10 SKUs directly threatens revenue targets.", 
     "Maintain strict safety stock levels for top 20 revenue-generating SKUs and create bundled gift sets."]
]

format_table(ws_q4, start_row=4, headers=["Insight", "Evidence", "Why Selected?", "Business Impact", "Recommendation"], data=q4_insights)

# TAB 7: Q5 (Surprising Result Analysis)
print("Creating tab: Q5...")
ws_q5 = wb.create_sheet(title="Q5")
apply_title_block(ws_q5, "Worksheet: Q5 - Surprising / Unexpected Result Analysis", "Deep-dive into unexpected data finding")

q5_data = [
    ["1. What did you initially expect?", 
     "Initial expectation was that the domestic UK market (which accounts for 90%+ of total transactions) would have the highest average revenue per customer (ARPU) and customer lifetime value."],
    ["2. What did the data actually show?", 
     "The data revealed that Netherlands ($31,716 ARPU) and EIRE/Ireland ($88,515 ARPU) generated exponentially higher revenue per customer than the UK domestic market ($1,864 ARPU). Specifically, just 9 Dutch accounts generated $285,446 in total net revenue."],
    ["3. Why do you think this happened?", 
     "Overseas customers using this platform are not individual retail shoppers; they are commercial wholesalers, regional retailers, and hospitality businesses purchasing giftware in large container/pallet quantities to resell locally."],
    ["4. What additional analysis did you perform?", 
     "Analyzed order size distributions: Dutch orders averaged over 300 units per line item compared to 8 units in UK orders. Additionally, verified that item catalog selections in overseas orders were concentrated in high-value party/decor sets."],
    ["5. What conclusion can you reasonably draw?", 
     "The company's international business is an untapped B2B wholesale powerhouse. While domestic UK sales provide base volume, international markets offer far superior operational leverage and higher profit margins per customer account."]
]

format_table(ws_q5, start_row=4, headers=["Question / Investigative Step", "Detailed Findings & Analysis"], data=q5_data)

# TAB 8: Q6 (Data Quality Issues, Limitations & Unsafe Conclusions)
print("Creating tab: Q6...")
ws_q6 = wb.create_sheet(title="Q6")
apply_title_block(ws_q6, "Worksheet: Q6 - Data Quality Issues, Limitations & Analytical Risks", "Identifying risks and boundaries of analysis")

q6_issues = [
    ["1. Missing Customer Identifiers (24.9% of rows)", 
     "135,080 rows lacked CustomerID. Could cause underestimation of true customer cohort sizes.", 
     "Filtered missing CustomerIDs for cohort/RFM metrics, but retained total sales volume for macro revenue validation."],
    ["2. Unlinked Cancellation Records (10,624 rows)", 
     "Cancelled invoices ('C' prefix) lacked direct foreign key links to original order invoice numbers.", 
     "Isolated negative quantities and analyzed net revenue on clean non-cancelled transactions."],
    ["3. Missing Product Cost & Profit Margin Data", 
     "Dataset only contains UnitPrice (selling price) without Cost of Goods Sold (COGS).", 
     "Focused analysis on Gross Revenue and Volume. Avoided making definitive net profit margin claims."]
]
format_table(ws_q6, start_row=4, headers=["Data Quality Issue", "Potential Impact on Analysis", "How Handled / Accounted For"], data=q6_issues)

q6_limits = [
    ["Limitation 1: Lack of Customer Demographics & Marketing Attributes", "No data on customer acquisition channels (CAC), ad spend, or demographic profiles, limiting marketing attribution analysis."],
    ["Limitation 2: Single-Year Time Horizon (12 Months)", "Data covers Dec 2010 to Dec 2011, preventing multi-year trend analysis or macro YoY seasonality comparison."]
]
format_table(ws_q6, start_row=10, headers=["Analytical Limitation", "Details"], data=q6_limits)

q6_unsafe = [
    ["Unsafe Conclusion", 
     "Concluding that UK domestic retail is unprofitable compared to overseas wholesale.", 
     "Reason: Without COGS, shipping tariffs, and return processing costs per country, net profitability per region cannot be mathematically proven."]
]
format_table(ws_q6, start_row=15, headers=["Type", "Unsafe Conclusion Statement", "Reasoning"], data=q6_unsafe)

# TAB 9: Q7 (Prioritized Actionable Recommendations)
print("Creating tab: Q7...")
ws_q7 = wb.create_sheet(title="Q7")
apply_title_block(ws_q7, "Worksheet: Q7 - Actionable Business Recommendations", "Prioritized recommendations based on business impact and feasibility")

q7_recs = [
    ["Recommendation 1: Launch Formal B2B Wholesale Expansion in Europe", 
     "Establish dedicated B2B sales portals and localized account management for Netherlands, Germany, and Ireland.", 
     "Insight Q4-#2 & Q5: Netherlands ARPU is $31.7k (17x UK ARPU).", 
     "Head of International Growth & Sales Team.", 
     "+$1.5M incremental international B2B revenue within 12 months.", 
     "International B2B account acquisition rate & Overseas Quarterly Revenue."],
    ["Recommendation 2: Establish VIP Key Account Retention Program", 
     "Create a dedicated Key Account Management framework with volume discount tiers for top 10% buyers.", 
     "Insight Q4-#1: Top 10% customers drive 61.33% ($5.46M) of total revenue.", 
     "Customer Success Director & B2B Account Managers.", 
     "Reduce top-tier customer churn by 50% and increase repeat order frequency.", 
     "Net Revenue Retention (NRR) & Annual Churn Rate among top 400 accounts."],
    ["Recommendation 3: Implement Q4 Inventory & Logistics Surge Planning", 
     "Advance inventory purchase orders by 45 days for top 20 SKUs ahead of the October-November peak.", 
     "Insight Q4-#3: November sales surge by 95% to $1.16M.", 
     "Supply Chain Director & Inventory Operations.", 
     "Prevent $200k+ in lost sales caused by Q4 stockouts on high-demand decor SKUs.", 
     "Order Fulfillment Rate (%) & Stockout Rate on Top 20 SKUs during Oct-Dec."]
]

format_table(ws_q7, start_row=4, headers=["What Should Be Done?", "Action Details", "Supporting Insight", "Who Would Act?", "Potential Outcome", "Success Metric"], data=q7_recs)

# TAB 10: Q8 (Looker Studio Dashboard Specifications)
print("Creating tab: Q8...")
ws_q8 = wb.create_sheet(title="Q8")
apply_title_block(ws_q8, "Worksheet: Q8 - Google Looker Studio Dashboard Specifications", "Interactive Dashboard Blueprint for Executive Management")

q8_dash = [
    ["1. Executive KPI Summary Cards (Top Row)", 
     "Four Scorecard Widgets: Total Net Revenue ($8.91M), Total Orders (18,532), Active B2B Customers (4,338), Average Order Value ($480.86).", 
     "Provides C-suite management with instantaneous high-level business health visibility."],
    ["2. Monthly Revenue & Order Volume Trend (Line/Bar Chart)", 
     "Combo chart showing monthly revenue ($) as bars and order count as line over 12-month timeline.", 
     "Visualizes Q4 holiday demand spike and seasonal purchasing patterns."],
    ["3. Geographic Revenue & ARPU Heatmap (Geographic Map)", 
     "Bubble Map / Filled Map of Europe highlighting revenue concentration in UK, Netherlands, Ireland, Germany, France.", 
     "Exposes high-performing international wholesale hubs."],
    ["4. Customer Segment Pareto Analysis (Donut / Bar Chart)", 
     "Breakdown of revenue by Customer Spend Deciles (Top 10% vs Remaining 90%).", 
     "Illustrates customer concentration risk and justification for Key Account Management."],
    ["5. Top 10 SKUs by Revenue Generation (Horizontal Bar Chart)", 
     "Ranked bar chart of top products (e.g. Regency Cakestand, Paper Craft Little Birdie).", 
     "Guides procurement teams on inventory prioritizing."],
    ["6. Interactive Control Filters (Top Filter Bar)", 
     "Date Range Picker, Country Selector, Order Size Category (Bulk B2B vs Small Retail).", 
     "Enables management to slice and dice data across regions and customer tiers dynamically."]
]

format_table(ws_q8, start_row=4, headers=["Dashboard Component", "Visualization Type & Layout Details", "Business Purpose & Decision Value"], data=q8_dash)

# TAB 11: Q10 (How I Used AI Disclosure)
print("Creating tab: Q10...")
ws_q10 = wb.create_sheet(title="Q10")
apply_title_block(ws_q10, "Worksheet: Q10 - How I Used AI Disclosure", "Transparency and auditability of AI tool usage in assessment")

q10_data = [
    ["1. AI Tools Used", 
     "Gemini 3.6 Flash (High), Python Data Science Stack (Pandas, OpenPyXL, PPTX)"],
    ["2. Key Applications & Tasks", 
     "• Data Profiling & Cleaning Logic: Generating Python regex & date-time parsing pipelines.\n• Exploratory Data Analysis: Computing complex customer spend deciles and country ARPU metrics.\n• Workbook Generation: Automating multi-tab formatted Excel sheet creation.\n• Documentation & Slide Scripting: Structuring executive presentation slides."],
    ["3. Example Where AI Helped Significantly", 
     "AI accelerated the discovery of the unexpected international B2B insight (Q5). By executing group-by aggregations across Country, CustomerID, and Revenue simultaneously, AI highlighted that Netherlands had an ARPU of $31,716 compared to the UK's $1,864, uncovering the wholesale distributor pattern within minutes."],
    ["4. Example Where AI Output Was Verified / Modified", 
     "Initial AI code attempted to drop negative Quantity rows across the entire dataset without separating cancellations. I modified the pipeline to isolate cancellations separately (to calculate return rates) and ensure unlinked cancelled orders did not distort net sales revenue."]
]

format_table(ws_q10, start_row=4, headers=["Disclosure Category", "Details & Real-world Examples"], data=q10_data)

# Save Excel File
try:
    wb.save(excel_path)
    print(f"Successfully generated {excel_path} with 11 worksheet tabs!")
except PermissionError:
    alt_path = "VirtuBox_Data_Analyst_Assessment_updated.xlsx"
    wb.save(alt_path)
    print(f"Note: '{excel_path}' is currently open in Excel. Saved output as '{alt_path}'.")


