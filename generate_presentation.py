import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

print("Generating Management PowerPoint Presentation...")

prs = Presentation()
# Set widescreen 16:9
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color Palette (Corporate Executive Theme)
NAVY = RGBColor(31, 78, 120)     # #1F4E78
DARK_GRAY = RGBColor(59, 59, 59)  # #3B3B3B
LIGHT_BG = RGBColor(245, 247, 250) # #F5F7FA
ACCENT_ORANGE = RGBColor(237, 125, 49) # #ED7D31
WHITE = RGBColor(255, 255, 255)
LIGHT_BLUE = RGBColor(217, 225, 242)

blank_layout = prs.slide_layouts[6]

def add_header(slide, title_text, category_text="EXECUTIVE MANAGEMENT PRESENTATION"):
    # Header bar background
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = NAVY
    header_box.line.color.rgb = NAVY

    # Category Text
    tf_cat = header_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = ACCENT_ORANGE
    p_cat.alignment = PP_ALIGN.LEFT

    # Main Title
    p_title = tf_cat.add_paragraph()
    p_title.text = title_text
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE
    p_title.alignment = PP_ALIGN.LEFT

def add_card(slide, left, top, width, height, title, content_list, bg_color=LIGHT_BG, border_color=NAVY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.2)
    tf.margin_bottom = Inches(0.2)

    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(16)
    p0.font.bold = True
    p0.font.color.rgb = NAVY
    p0.space_after = Pt(10)

    for item in content_list:
        p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_GRAY
        p.space_after = Pt(6)
        p.level = 0

# SLIDE 1: Title & Business Problem
slide1 = prs.slides.add_slide(blank_layout)
add_header(slide1, "Slide 1 — Business Problem & Strategic Context")

add_card(slide1, 0.8, 1.5, 5.6, 5.3, "The Business Context & Opportunity", [
    "• Background: Management aims to transform our B2B e-commerce platform into a data-driven growth engine.",
    "• Core Challenge: Unsegmented marketing and lack of visibility into high-value B2B buyers vs retail single-purchase shoppers.",
    "• Operational Drag: High return/cancellation volume (10,600+ items) and peak Q4 stocking inventory bottlenecks.",
    "• Strategic Goal: Transition to targeted B2B account management, capture international wholesale demand, and protect profit margins."
])

add_card(slide1, 6.8, 1.5, 5.7, 5.3, "Key Metrics at a Glance (Dataset Baseline)", [
    "• Total Gross Revenue: $8.89 Million across 12 months.",
    "• Transaction Records Analyzed: 392,692 cleaned order items.",
    "• Active Customer Accounts: 4,338 unique identified buyers.",
    "• Total Completed Orders: 18,532 invoices.",
    "• Core Market Focus: UK (Domestic) + Top European Wholesale Markets (Netherlands, Ireland, Germany, France)."
])

# SLIDE 2: Data & Methodology
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "Slide 2 — Data & Analytical Methodology")

add_card(slide2, 0.8, 1.5, 5.6, 5.3, "Data Source & Scope", [
    "• Source: 541,909 raw transactional records from global e-commerce order logs (Dec 2010 – Dec 2011).",
    "• Scope: All order line-items, product codes, unit prices, order quantities, customer IDs, and destination countries."
])

add_card(slide2, 6.8, 1.5, 5.7, 5.3, "Rigorous 5-Step Cleaning Pipeline", [
    "1. Guest Checkout Filtering: Dropped 135,080 rows missing CustomerID for accurate account-level cohort metrics.",
    "2. Cancellation Isolation: Removed 10,624 return items (Quantity < 0) to report clean net sales volume.",
    "3. Outlier & Price Correction: Eliminated 2,517 non-transactional items (UnitPrice <= $0.00).",
    "4. Deduplication: Purged 5,268 exact duplicate order item logs.",
    "5. Feature Engineering: Created Net Revenue (TotalAmount), Order Date/Hour dimensions, and RFM spend tiers."
])

# SLIDE 3: Key Findings
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "Slide 3 — Key Executive Findings")

add_card(slide3, 0.8, 1.5, 3.7, 5.3, "1. Severe Customer Concentration", [
    "• Top 10% of customers (433 accounts) drive 61.33% ($5.46M) of total revenue.",
    "• Extreme reliance on key accounts requires dedicated B2B account managers."
])

add_card(slide3, 4.8, 1.5, 3.7, 5.3, "2. High-Yield Overseas Wholesale", [
    "• Netherlands ARPU is $31,716 (9 accounts = $285k).",
    "• Ireland ARPU is $88,515 (3 accounts = $265k).",
    "• Domestic UK ARPU is $1,864 per customer.",
    "• International buyers operate as high-volume B2B wholesale distributors."
])

add_card(slide3, 8.8, 1.5, 3.7, 5.3, "3. Q4 Pre-Holiday Demand Spike", [
    "• November revenue surged by 95% to $1.16M ahead of Christmas stocking.",
    "• Peak daily orders occur between 10 AM and 1 PM (12 PM peak: $1.38M total)."
])

# SLIDE 4: Deep-Dive Insight
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "Slide 4 — Deep-Dive Insight: International B2B Wholesale Expansion")

add_card(slide4, 0.8, 1.5, 5.6, 5.3, "The Unexpected Discovery (Surprising Finding)", [
    "• Initial Expectation: Domestic UK market (90%+ volume) would yield the highest customer lifetime value.",
    "• Data Reality: Netherlands & Ireland buyers generate 17x to 47x higher revenue per account than UK retail buyers.",
    "• Root Cause: Overseas buyers purchase in pallet-level bulk quantities (300+ units per item vs 8 units in UK).",
    "• Strategic Takeaway: International B2B offers massive operating margin leverage with minimal acquisition overhead."
])

add_card(slide4, 6.8, 1.5, 5.7, 5.3, "Supporting Financial Evidence", [
    "• Netherlands: 9 Accounts | $285,446 Revenue | $31,716 ARPU",
    "• Ireland (EIRE): 3 Accounts | $265,545 Revenue | $88,515 ARPU",
    "• Australia: 9 Accounts | $138,521 Revenue | $15,391 ARPU",
    "• Germany: 94 Accounts | $228,867 Revenue | $2,434 ARPU",
    "• UK Domestic: 3,920 Accounts | $7.31M Revenue | $1,864 ARPU"
])

# SLIDE 5: Recommendations
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "Slide 5 — Strategic Actionable Recommendations")

add_card(slide5, 0.8, 1.5, 3.7, 5.3, "1. Launch European B2B Portal", [
    "• What: Build dedicated wholesale ordering portals for Netherlands, Ireland, and Germany.",
    "• Owner: Head of International Expansion.",
    "• Outcome: Capture $1.5M+ incremental wholesale revenue in Year 1."
])

add_card(slide5, 4.8, 1.5, 3.7, 5.3, "2. Key Account Loyalty Tier", [
    "• What: Establish VIP account management & tier-based volume discounts for top 400 buyers.",
    "• Owner: Director of Customer Success.",
    "• Outcome: Reduce top-tier customer churn by 50%."
])

add_card(slide5, 8.8, 1.5, 3.7, 5.3, "3. Q4 Pre-Stocking Logistics Plan", [
    "• What: Advance inventory POs by 45 days for top 20 revenue SKUs before October surge.",
    "• Owner: Supply Chain Operations Lead.",
    "• Outcome: Prevent $200k+ stockout losses."
])

# SLIDE 6: Expected Business Impact
slide6 = prs.slides.add_slide(blank_layout)
add_header(slide6, "Slide 6 — Expected Business Impact & Key Metrics")

add_card(slide6, 0.8, 1.5, 5.6, 5.3, "Projected Financial Growth & Savings", [
    "• Revenue Uplift: +18% to +25% total annual revenue growth driven by international wholesale expansion.",
    "• Churn Reduction: 50% decrease in top-tier account churn via VIP relationship management.",
    "• Inventory Optimization: 15% reduction in carrying costs and 98%+ stock availability during peak Q4.",
    "• Fulfillment Efficiency: 20% lower shipping overhead by bulk-shipping to European regional hubs."
])

add_card(slide6, 6.8, 1.5, 5.7, 5.3, "Executive Dashboard KPIs to Track Success", [
    "1. Net Revenue Retention (NRR) among top 10% B2B accounts.",
    "2. International B2B Account Acquisition Count (Netherlands, Germany, Ireland).",
    "3. On-Time In-Full (OTIF) Delivery % during Q4 peak season.",
    "4. Return & Cancellation Loss Ratio (target: under 2.5% of total orders)."
])

# SLIDE 7: Limitations & Next Steps
slide7 = prs.slides.add_slide(blank_layout)
add_header(slide7, "Slide 7 — Analytical Limitations & Next Steps")

add_card(slide7, 0.8, 1.5, 5.6, 5.3, "Analytical Risk & Governance Limitations", [
    "• Unidentified Guest Buyers: 24.9% of raw records lacked CustomerIDs; macro volume verified, but individual cohort tracking restricted.",
    "• Single-Year Baseline: 12-month data limits multi-year YoY trend benchmarking.",
    "• Lack of COGS Data: Revenue metrics analyzed on gross sales; cost of goods and shipping tariffs needed for exact net margin calculation."
])

add_card(slide7, 6.8, 1.5, 5.7, 5.3, "Immediate Next Steps (30-60-90 Day Roadmap)", [
    "• Days 1–30: Share Looker Studio Interactive Dashboard with regional sales heads.",
    "• Days 31–60: Pilot European B2B Wholesale Portal with top Dutch and Irish distributors.",
    "• Days 61–90: Integrate ERP cost data to perform exact product net margin analytics."
])

# Save Presentation
output_pptx = "Presentation_Management.pptx"
prs.save(output_pptx)
print(f"Successfully generated executive presentation: {output_pptx}")
