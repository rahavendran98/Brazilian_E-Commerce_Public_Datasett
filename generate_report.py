"""Generate the Data Understanding Report as a Word document."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Styles helpers ────────────────────────────────────────────────────────────
def set_font(run, bold=False, italic=False, size=11, color=None):
    run.bold   = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    for run in p.runs:
        run.font.color.rgb = RGBColor(31, 73, 125)
    return p

def add_bullet(doc, text, bold_prefix=None, indent_level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Inches(0.25 * (indent_level + 1))
    if bold_prefix:
        r = p.add_run(bold_prefix + " ")
        r.bold = True
        r.font.size = Pt(10)
        r2 = p.add_run(text)
        r2.font.size = Pt(10)
    else:
        r = p.add_run(text)
        r.font.size = Pt(10)
    return p

def add_para(doc, text, bold=False, italic=False, size=10, align=None, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(2)
    if align:
        p.alignment = align
    r = p.add_run(text)
    set_font(r, bold=bold, italic=italic, size=size)
    return p

def shade_cell(cell, hex_color="1F497D"):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        shade_cell(cell, "1F497D")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
    # Data rows
    for ri, row in enumerate(rows):
        tr = table.rows[ri + 1]
        bg = "DEEAF1" if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate(row):
            cell = tr.cells[ci]
            shade_cell(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val))
            r.font.size = Pt(9)
    # Column widths
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)
    doc.add_paragraph()
    return table

def page_break(doc):
    doc.add_page_break()


# ══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Brazilian E-Commerce Public Dataset")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = RGBColor(31, 73, 125)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Data Understanding Report")
r.bold = True
r.font.size = Pt(16)
r.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Prepared by")
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(89, 89, 89)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Ragavendra")
r.bold = True
r.font.size = Pt(14)
r.font.color.rgb = RGBColor(31, 73, 125)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("AI / ML Engineer")
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("June 2026")
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(89, 89, 89)

page_break(doc)


# ══════════════════════════════════════════════════════════════════════════════
# 1. PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "1. Project Overview", level=1)

add_bullet(doc, "Source dataset is the Olist Brazilian E-Commerce Public Dataset — a real-world marketplace transaction record covering orders placed between 2016 and 2018.")
add_bullet(doc, "The dataset was released by Olist, the largest department store in Brazilian marketplaces, and made publicly available on Kaggle.")
add_bullet(doc, "The objective of this project is to perform end-to-end data engineering and exploratory analysis to derive actionable business insights on customer behaviour, delivery performance, product performance, and seller distribution.")
add_bullet(doc, "The analysis covers nine source tables ingested as raw CSVs, sampled down to 10,000 orders for tractable processing, integrated into two analytical tables (order-level and item-level), cleaned, and explored through targeted EDA.")
add_bullet(doc, "Deliverables from this phase include clean datasets, a validated data quality report, EDA charts, and this data understanding document.")


# ══════════════════════════════════════════════════════════════════════════════
# 2. DATA SOURCE & STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "2. Data Source & Structure", level=1)
add_heading(doc, "2.1 Raw Dataset Inventory", level=2)
add_bullet(doc, "Nine CSV files form the raw source layer, covering the full order lifecycle from placement to customer review.")

headers = ["Table", "Rows (Raw)", "Columns", "Primary Entity"]
rows = [
    ("olist_orders_dataset",         "99,441",    "8",  "Order header & timestamps"),
    ("olist_customers_dataset",      "99,441",    "5",  "Customer identity & location"),
    ("olist_order_items_dataset",    "112,650",   "7",  "Line items per order"),
    ("olist_order_payments_dataset", "103,886",   "5",  "Payment method & value"),
    ("olist_order_reviews_dataset",  "99,224",    "7",  "Customer ratings & comments"),
    ("olist_products_dataset",       "32,951",    "9",  "Product catalogue & dimensions"),
    ("olist_sellers_dataset",        "3,095",     "4",  "Seller location"),
    ("olist_geolocation_dataset",    "1,000,163", "5",  "Zip-code lat/lng lookup"),
    ("product_category_name_tr",     "71",        "2",  "PT → EN category mapping"),
]
add_table(doc, headers, rows, col_widths=[2.3, 1.1, 0.8, 2.3])

add_heading(doc, "2.2 Sampling Strategy", level=2)
add_bullet(doc, "A stratified random sample of 10,000 orders (random_state=42) was drawn from the full 99,441-row orders table to keep computation manageable while preserving distribution shape.")
add_bullet(doc, "All dependent tables (customers, items, payments, reviews, products, sellers) were filtered to match only the sampled order IDs — no data leakage from outside the sample boundary.")
add_bullet(doc, "The geolocation table collapsed from 650,746 sample-matched rows to 6,486 unique zip codes after deduplication (mean lat/lng per zip, mode for city/state).")

headers2 = ["Table", "Rows After Sampling"]
rows2 = [
    ("orders",       "10,000"),
    ("customers",    "10,000"),
    ("order_items",  "11,383"),
    ("payments",     "10,476"),
    ("reviews",       "9,959"),
    ("products",      "6,747"),
    ("sellers",       "1,654"),
    ("category_tr",      "71"),
    ("geolocation",   "6,486"),
]
add_table(doc, headers2, rows2, col_widths=[2.5, 2.0])

add_heading(doc, "2.3 Analytical Table Design", level=2)
add_bullet(doc, "Two wide analytical tables were built from the sampled data to serve as the foundation for all downstream analysis.")
add_bullet(doc, "order_level (10,000 rows × 23 cols) — one row per order. Built by joining orders → customers → payments (aggregated) → reviews (aggregated) → geolocation. Captures the full customer and payment picture at order grain.")
add_bullet(doc, "item_level (11,383 rows × 22 cols) — one row per order-item. Built by joining order_items → products → category_tr → sellers → orders. Enables product- and seller-level revenue analysis.")
add_bullet(doc, "item_revenue was engineered as price + freight_value to represent the true per-line cost borne by the customer.")


# ══════════════════════════════════════════════════════════════════════════════
# 3. DATA DICTIONARY (KEY COLUMNS)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "3. Key Column Reference", level=1)
add_bullet(doc, "The full data dictionary covers all 52 raw columns across 9 tables. Key columns used in analysis are listed below.")

headers3 = ["Column", "Table", "Type", "Description"]
rows3 = [
    ("order_id",                      "orders",        "str",      "Unique order identifier (PK)"),
    ("customer_unique_id",            "customers",     "str",      "True customer identity — persists across repeat purchases"),
    ("order_status",                  "orders",        "str",      "8 states: delivered, shipped, canceled, etc."),
    ("order_purchase_timestamp",      "orders",        "datetime", "When the customer placed the order"),
    ("order_delivered_customer_date", "orders",        "datetime", "Actual delivery date (null if not yet delivered)"),
    ("order_estimated_delivery_date", "orders",        "datetime", "Promised delivery date used for lateness check"),
    ("price",                         "order_items",   "float",    "Unit product price (BRL)"),
    ("freight_value",                 "order_items",   "float",    "Shipping fee for that line item (BRL)"),
    ("payment_type",                  "payments",      "str",      "credit_card / boleto / voucher / debit_card"),
    ("payment_installments",          "payments",      "int",      "Number of installment splits chosen by customer"),
    ("payment_value",                 "payments",      "float",    "Amount paid per payment record (BRL)"),
    ("review_score",                  "reviews",       "int",      "Customer rating 1 (worst) to 5 (best)"),
    ("product_category_name_english", "products/tr",   "str",      "Translated English product category"),
    ("product_weight_g",              "products",      "float",    "Item weight in grams"),
    ("seller_state",                  "sellers",       "str",      "2-letter Brazilian state code for the seller"),
    ("geolocation_lat/lng",           "geolocation",   "float",    "Mean lat/lng coordinates per zip code"),
]
add_table(doc, headers3, rows3, col_widths=[1.8, 1.2, 0.8, 2.7])


page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# 4. DATA QUALITY ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "4. Data Quality Assessment", level=1)

add_heading(doc, "4.1 Missing Value Analysis", level=2)
add_bullet(doc, "order_level overall completeness: 99.63% (229,157 of 230,000 cells filled).")
add_bullet(doc, "item_level overall completeness: 99.70% (249,668 of 250,426 cells filled).")
add_bullet(doc, "Delivery timestamps account for the largest gap — 2.85% of orders have no delivered-customer-date and 1.79% no carrier-handoff date. These are structurally absent for non-delivered orders, not data entry errors.")
add_bullet(doc, "Review score and count are missing for 88 orders (0.88%) — orders that completed but never received any customer feedback.")
add_bullet(doc, "Geolocation fields (lat, lng, city, state) are missing for 35 orders (0.35%) where the customer zip code had no match in the geo lookup table.")
add_bullet(doc, "In item_level, 151 items (1.33%) carry no product category name even in Portuguese — these were back-filled with the string 'unknown' to preserve row count.")

headers4 = ["Field", "Missing Count", "Missing %", "Treatment"]
rows4 = [
    ("order_delivered_customer_date", "285", "2.85%", "Flagged via is_delivered binary column"),
    ("order_delivered_carrier_date",  "179", "1.79%", "Left as NaN; not used in core metrics"),
    ("avg_review_score",               "88", "0.88%", "Flagged via has_review; n_reviews filled with 0"),
    ("geolocation fields (5 cols)",    "35", "0.35%", "city/state → 'unknown'; lat/lng left as NaN"),
    ("order_approved_at",              "24", "0.24%", "Flagged via is_approved binary column"),
    ("payment fields (4 cols)",         "1", "0.01%", "Filled with 0 / 'unknown' / 1 by type"),
    ("product_category (item_level)", "151", "1.33%", "Back-filled from PT name, else 'unknown'"),
]
add_table(doc, headers4, rows4, col_widths=[2.0, 1.1, 0.9, 2.5])

add_heading(doc, "4.2 Duplicate Analysis", level=2)
add_bullet(doc, "Zero fully duplicate rows found in all eight tables checked.")
add_bullet(doc, "Zero duplicate primary keys across order_id, customer_id, product_id, and seller_id.")
add_bullet(doc, "The reviews table contained 10 duplicate review_ids in the raw data — these were not present in the sampled subset and required no action.")

add_heading(doc, "4.3 Referential Integrity", level=2)
add_bullet(doc, "All four core foreign key relationships were verified against the sampled tables:")
add_bullet(doc, "orders.customer_id → customers.customer_id: 0 orphan keys", indent_level=1)
add_bullet(doc, "order_items.order_id → orders.order_id: 0 orphan keys", indent_level=1)
add_bullet(doc, "order_items.product_id → products.product_id: 0 orphan keys", indent_level=1)
add_bullet(doc, "order_items.seller_id → sellers.seller_id: 0 orphan keys", indent_level=1)
add_bullet(doc, "The sample extraction logic correctly cascaded from order IDs outward — no child records referencing non-existent parents.")

add_heading(doc, "4.4 Validity & Consistency Checks", level=2)
add_bullet(doc, "Price ≤ 0: 0 records — all item prices are strictly positive.")
add_bullet(doc, "Freight < 0: 0 records — no negative shipping fees.")
add_bullet(doc, "Review score outside 1–5 range: 0 records.")
add_bullet(doc, "Delivered date before purchase date: 0 records — temporal ordering is logically sound throughout.")
add_bullet(doc, "Installments = 0: fixed — 0 records remain after replacing zero-installment orders with 1.")
add_bullet(doc, "City and state fields were standardised to lowercase and stripped of leading/trailing whitespace before analysis.")

add_heading(doc, "4.5 Data Validation Summary", level=2)

headers5 = ["Check", "Result", "Status"]
rows5 = [
    ("Schema completeness (order_level)",     "All expected columns present",   "PASS"),
    ("Schema completeness (item_level)",      "All expected columns present",   "PASS"),
    ("Missing values — critical cols",        "No missing in key join columns", "PASS"),
    ("Business rules (price, dates, scores)", "All rules satisfied",            "PASS"),
    ("Referential integrity (4 FK links)",    "Zero orphan keys",               "PASS"),
    ("Consistency — date ordering",           "No delivery before purchase",    "PASS"),
    ("Data type alignment",                   "order_id loaded as str not int", "NOTE"),
]
add_table(doc, headers5, rows5, col_widths=[2.5, 2.2, 0.8])

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# 5. EXPLORATORY DATA ANALYSIS FINDINGS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "5. Exploratory Data Analysis Findings", level=1)

add_heading(doc, "5.1 Order Status & Volume", level=2)
add_bullet(doc, "97.1% of orders in the sample carry 'delivered' status — the dataset is skewed toward completed transactions.")
add_bullet(doc, "Shipped (1.0%), canceled (0.7%), and unavailable (0.6%) statuses are marginal and will not materially distort revenue or review metrics.")
add_bullet(doc, "Monthly order volume shows a clear upward trend through 2017–2018, with visible peaks in Q4 (November Black Friday season) and a dip in late 2018 as the dataset window closes.")

add_heading(doc, "5.2 Geographic Distribution", level=2)
add_bullet(doc, "São Paulo city alone generates 15.8% of all customer orders, with Rio de Janeiro adding 6.7% — together nearly a quarter of total volume.")
add_bullet(doc, "The top 5 cities (São Paulo, Rio de Janeiro, Belo Horizonte, Brasília, Campinas) collectively represent ~23% of orders, confirming that demand is concentrated in major urban centres rather than spread evenly.")
add_bullet(doc, "At state level, SP (São Paulo state) leads by a wide margin, followed by RJ and MG — a pattern that mirrors Brazil's GDP distribution.")
add_bullet(doc, "Any logistics disruption in São Paulo state would have an outsized impact on platform-wide revenue and delivery performance.")

add_heading(doc, "5.3 Delivery Performance", level=2)
add_bullet(doc, "Mean delivery time: 12.0 days. Median: 10.0 days. The 2-day gap indicates a right-skewed distribution inflated by a small number of extreme delays (worst case: 186 days).")
add_bullet(doc, "On-time delivery rate: 92.3% — orders arrived on or before the estimated date.")
add_bullet(doc, "Late deliveries (7.7%) are disproportionately likely to produce 1-star reviews, given the bimodal review distribution observed.")
add_bullet(doc, "The delivery time histogram is right-skewed, not normally distributed — median is the more representative central tendency for reporting.")

add_heading(doc, "5.4 Product & Category Analysis", level=2)
add_bullet(doc, "Health & Beauty is the top revenue category at R$141,728, followed by Watches & Gifts (R$136,689) and Bed/Bath/Table (R$122,963).")
add_bullet(doc, "The top 3 categories account for roughly 25% of total item revenue (R$1,581,151) — a reasonably balanced spread rather than extreme concentration.")
add_bullet(doc, "Computers & Accessories (R$106,869) is the only tech category in the top 5 — this is primarily a lifestyle and home goods marketplace.")
add_bullet(doc, "Garden Tools holds the highest single-product volume (56 items sold, R$4,043) and Furniture/Decor the highest items-per-product ratio.")
add_bullet(doc, "Average basket size is 1.15 items per order — nearly all transactions are single-item purchases with no visible cross-sell or bundling behaviour.")

add_heading(doc, "5.5 Seller Analysis", level=2)
add_bullet(doc, "1,654 unique sellers operate in the sampled dataset, listing an average of ~4 products each (6,747 distinct products).")
add_bullet(doc, "Top 20% of sellers (~331) drive 72.2% of total item revenue — the bottom 80% compete for less than 28% of the market.")
add_bullet(doc, "The top 10 sellers each generated R$16,000–R$25,000, clustered tightly with no single runaway outlier.")
add_bullet(doc, "São Paulo state dominates seller geography — consistent with it being Brazil's commercial hub and the largest customer base.")

add_heading(doc, "5.6 Payment Behaviour", level=2)
add_bullet(doc, "Credit card accounts for 76.5% of orders; boleto (cash payment slip) holds 19.7%, reflecting Brazil's significant unbanked and underbanked population.")
add_bullet(doc, "In practical terms this is a two-method marketplace — credit card plus boleto together cover ~96% of all transactions.")
add_bullet(doc, "Debit card (1.5%) and voucher (2.3%) are negligible.")
add_bullet(doc, "48.9% of orders are settled in a single installment; when splits are used the average rises to 2.9 installments — short-term convenience, not long financing cycles.")

add_heading(doc, "5.7 Review & Customer Satisfaction", level=2)
add_bullet(doc, "Average review score: 4.09 out of 5 across 9,912 reviewed orders.")
add_bullet(doc, "76.9% of reviewers gave 4 or 5 stars; 5-star alone accounts for 58% of all reviews.")
add_bullet(doc, "Dissatisfied customers disproportionately choose 1-star (11.4%) rather than 2–3 stars combined (11.8%) — the distribution is bimodal, meaning outcomes are either good or outright failures, rarely middling.")
add_bullet(doc, "The 88 unreviewed orders (0.88%) are too small to shift the average meaningfully even if all were negative.")

add_heading(doc, "5.8 Customer Behaviour & Retention", level=2)
add_bullet(doc, "9,956 unique customers placed the 10,000 orders — 9,913 of them (99.6%) bought exactly once.")
add_bullet(doc, "Only 43 customers (0.4%) placed more than one order. Growth on this platform is entirely acquisition-dependent.")
add_bullet(doc, "High spend does not predict satisfaction: 4 of the top 10 highest-spending customers gave 1-star reviews.")
add_bullet(doc, "The top customer spent R$7,274 on a single order — almost certainly a single high-value item such as furniture or a large appliance.")

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# 6. KEY BUSINESS METRICS SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "6. Key Business Metrics (Sampled — 10,000 Orders)", level=1)

headers6 = ["Metric", "Value"]
rows6 = [
    ("Total Orders (sample)",          "10,000"),
    ("Total Unique Customers",         "9,956"),
    ("Total Revenue (order-level)",    "R$ 1,602,345"),
    ("Average Order Value",            "R$ 160.23"),
    ("Total Items Sold",               "11,383"),
    ("Total Item Revenue",             "R$ 1,581,151"),
    ("Average Item Price",             "R$ 119.00"),
    ("Average Freight Value",          "R$ 19.91  (16.7% of avg item price)"),
    ("Delivery Rate",                  "97.2%"),
    ("On-Time Delivery Rate",          "92.3%"),
    ("Average Delivery Time",          "12.0 days (median 10.0 days)"),
    ("Average Review Score",           "4.09 / 5.00"),
    ("Customer Satisfaction (4-5★)",   "76.9%"),
    ("Repeat Customer Rate",           "0.4%  (43 of 9,956 customers)"),
    ("Top Payment Method",             "Credit Card (76.5%)"),
    ("Top Revenue State (customer)",   "São Paulo (SP)"),
    ("Top Revenue Category",           "Health & Beauty (R$ 141,728)"),
    ("Unique Sellers",                 "1,654"),
    ("Unique Products",                "6,747"),
    ("Revenue — Top 20% Sellers",      "72.2% of total item revenue"),
]
add_table(doc, headers6, rows6, col_widths=[3.0, 3.5])


# ══════════════════════════════════════════════════════════════════════════════
# 7. DATA LIMITATIONS & KNOWN RISKS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "7. Data Limitations & Known Risks", level=1)

add_bullet(doc, "Sample scope: Analysis is based on 10,000 of 99,441 orders (~10%). Findings are representative but absolute figures will differ from full-dataset results.")
add_bullet(doc, "Temporal window: Orders span 2016–2018. No data beyond August 2018. Seasonal patterns or trends post-2018 are not captured.")
add_bullet(doc, "Geolocation gaps: 35 orders (0.35%) could not be mapped to coordinates. These are excluded from any spatial analysis without impacting other metrics.")
add_bullet(doc, "Category unknowns: 151 item-level rows (1.33%) have no product category even in Portuguese. Category revenue totals for affected items are excluded from category rankings.")
add_bullet(doc, "Single-purchase bias: 99.6% of customers appear only once. It is unclear whether this reflects true one-time buying behaviour or whether repeat customers register new accounts per purchase.")
add_bullet(doc, "Review coverage: ~217 full-dataset orders have no review record. In the sample, 88 orders lack reviews. Review averages may overstate satisfaction if non-reviewers skew negative.")
add_bullet(doc, "Product dimension data: Columns product_name_lenght and product_description_lenght contain a known typo in the source schema. These fields were used as-is without renaming to preserve traceability.")
add_bullet(doc, "Currency: All monetary values are in Brazilian Real (BRL). No currency conversion has been applied.")


# ══════════════════════════════════════════════════════════════════════════════
# 8. DATA PIPELINE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "8. Data Pipeline Summary", level=1)

add_bullet(doc, "Stage 1 — Ingestion (data_ingestion.ipynb): 9 raw CSVs loaded; shape, dtypes, and head verified for all tables; data dictionary of 52 columns built with descriptions.")
add_bullet(doc, "Stage 2 — Sampling (data_ingestion.ipynb): 10,000 orders sampled (random_state=42); dependent tables filtered by order/customer/product/seller IDs; geolocation deduplicated from 650K to 6,486 zip-level rows.")
add_bullet(doc, "Stage 3 — Integration (data_integration.ipynb): order_level assembled via 4 sequential left joins; item_level assembled via 5 sequential left joins; item_revenue feature engineered.")
add_bullet(doc, "Stage 4 — Cleaning (data_cleaning.ipynb): date columns cast to datetime64; missing values treated by type (flags, fill, median, 'unknown'); text standardised to lowercase; zero installments corrected to 1; duplicate check run; clean files saved.")
add_bullet(doc, "Stage 5 — Validation (data_validation.ipynb): schema, data types, missing values, business rules, and referential integrity checked across both analytical tables — all checks passed except a minor dtype note on order_id.")
add_bullet(doc, "Stage 6 — EDA (data_cleaning.ipynb EDA section): order status, geographic distribution, delivery performance, category revenue, seller analysis, payment behaviour, review scores, and customer retention all analysed and charted.")

add_para(doc, "", size=10)

# ── Footer note ──────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("─" * 60)
r.font.size = Pt(9)
r.font.color.rgb = RGBColor(150, 150, 150)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Ragavendra  |  AI / ML Engineer  |  Brazilian E-Commerce Public Dataset  |  June 2026")
r.font.size = Pt(9)
r.italic = True
r.font.color.rgb = RGBColor(120, 120, 120)

# ── Save ──────────────────────────────────────────────────────────────────────
out_path = r"D:\projects\Brazilian E-Commerce Public Dataset\reports\Data_Understanding_Report.docx"
doc.save(out_path)
print("Report saved:", out_path)
