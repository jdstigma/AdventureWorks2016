"""
db_report.py
============
Generates a full statistical report on AdventureWorks in PostgreSQL.

Analyses:
  SALES      - Revenue trend with OLS regression + R2
               Log10 transform of order values
               Customer segmentation via K-Means (3 clusters)
               Territory performance ranking
  PRODUCTION - List price vs standard cost linear regression
               Scrap reason breakdown
  HR         - Pay rate distribution with log10 transform
               Headcount by department
  PURCHASING - Vendor spend clustering (K-Means)
               Purchase order trends over time

OUTPUTS
-------
  adventureworks_report.xlsx  -- Excel workbook, one sheet per analysis
  adventureworks_report.html  -- Self-contained HTML report with charts

REQUIREMENTS
------------
    pip install -r requirements.txt

NOTE: Do not commit real passwords to a public GitHub repo.
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------
import os
import base64
import io
import warnings

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from scipy import stats
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sqlalchemy import create_engine, text

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
DB_HOST     = "localhost"
DB_PORT     = 5432
DB_NAME     = "adventureworks"
DB_USER     = "awuser"
DB_PASSWORD = "Gunner!!24"

EXCEL_OUT = os.path.join(SCRIPT_DIR, "adventureworks_report.xlsx")
HTML_OUT  = os.path.join(SCRIPT_DIR, "adventureworks_report.html")

N_CUSTOMER_CLUSTERS = 3
N_VENDOR_CLUSTERS   = 3

# ---------------------------------------------------------------------------
# CONNECTION
# ---------------------------------------------------------------------------
CONNECTION_STRING = (
    f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
    f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"
)

print("Connecting to database...")
engine = create_engine(CONNECTION_STRING)

try:
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print(f"  Connected to: {DB_NAME} on {DB_HOST}:{DB_PORT}\n")
except Exception as e:
    print(f"  Connection failed: {e}")
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def sql(query_str, params=None):
    """Run SQL and return a DataFrame."""
    with engine.connect() as conn:
        return pd.read_sql(text(query_str), conn, params=params)


def fig_to_base64(fig):
    """Convert a matplotlib figure to a base64 PNG string for HTML embedding."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("utf-8")
    plt.close(fig)
    return encoded


def section(title):
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}")


excel_sheets  = {}
html_sections = []


def add_result(sheet_name, df, chart_b64=None, notes=""):
    """Store a result for both Excel and HTML output."""
    excel_sheets[sheet_name] = df
    table_html = df.to_html(index=False, classes="data-table", border=0)
    img_html   = f'<img src="data:image/png;base64,{chart_b64}" class="chart">' if chart_b64 else ""
    notes_html = f'<p class="notes">{notes}</p>' if notes else ""
    html_sections.append(f"""
        <section>
            <h2>{sheet_name}</h2>
            {notes_html}
            {img_html}
            {table_html}
        </section>
    """)


# ===========================================================================
# SALES: Revenue Trend + OLS Regression
# ===========================================================================
section("SALES: Revenue Trend + OLS Regression")

sales_trend = sql("""
    SELECT
        EXTRACT(YEAR FROM orderdate)::int               AS year,
        TO_CHAR(orderdate, 'YYYY-MM')                   AS year_month,
        COUNT(*)                                        AS order_count,
        SUM(totaldue)                                   AS total_revenue,
        AVG(totaldue)                                   AS avg_order_value
    FROM sales.salesorderheader
    GROUP BY year, year_month
    ORDER BY year_month
""")

sales_trend["total_revenue"]   = sales_trend["total_revenue"].round(2)
sales_trend["avg_order_value"] = sales_trend["avg_order_value"].round(2)

x = np.arange(len(sales_trend))
y = sales_trend["total_revenue"].values
slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
r_squared  = round(r_value ** 2, 4)
trend_line = slope * x + intercept

print(f"  OLS slope:  ${slope:,.0f} per month")
print(f"  R-squared:  {r_squared}")
print(f"  P-value:    {p_value:.4f}")

sales_trend["regression_line"] = trend_line.round(2)

fig, ax = plt.subplots(figsize=(12, 5))
ax.bar(sales_trend["year_month"], sales_trend["total_revenue"],
       color="#4C72B0", alpha=0.7, label="Monthly Revenue")
ax.plot(sales_trend["year_month"], trend_line,
        color="red", linewidth=2, label=f"OLS Trend (R2={r_squared})")
ax.set_title("Monthly Sales Revenue with OLS Trend Line", fontsize=14)
ax.set_xlabel("Month")
ax.set_ylabel("Total Revenue ($)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
step = max(1, len(sales_trend) // 12)
ax.set_xticks(sales_trend["year_month"][::step])
ax.set_xticklabels(sales_trend["year_month"][::step], rotation=45, ha="right")
ax.legend()
plt.tight_layout()
chart1 = fig_to_base64(fig)

add_result(
    "Sales Revenue Trend", sales_trend, chart1,
    f"OLS Regression: slope=${slope:,.0f}/month, R2={r_squared}, p={p_value:.4f}"
)


# ===========================================================================
# SALES: Log10 Transform of Order Values
# ===========================================================================
section("SALES: Log10 Transform of Order Values")

order_vals = sql("""
    SELECT totaldue AS order_value
    FROM sales.salesorderheader
    WHERE totaldue > 0
""")

order_vals["log10_order_value"] = np.log10(order_vals["order_value"])
skew_raw = round(float(order_vals["order_value"].skew()), 3)
skew_log = round(float(order_vals["log10_order_value"].skew()), 3)
print(f"  Skewness raw:   {skew_raw}")
print(f"  Skewness log10: {skew_log}")

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
axes[0].hist(order_vals["order_value"], bins=60, color="#4C72B0", alpha=0.8)
axes[0].set_title(f"Raw Order Values  (skew={skew_raw})")
axes[0].set_xlabel("Order Value ($)")
axes[0].xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
axes[1].hist(order_vals["log10_order_value"], bins=60, color="#DD8452", alpha=0.8)
axes[1].set_title(f"Log10(Order Value)  (skew={skew_log})")
axes[1].set_xlabel("log10(Order Value)")
plt.tight_layout()
chart2 = fig_to_base64(fig)

log_summary = order_vals.describe().round(3).reset_index()
log_summary.columns = ["Statistic", "order_value", "log10_order_value"]

add_result(
    "Sales Log10 Transform", log_summary, chart2,
    f"Raw skewness={skew_raw}. After log10 transform skewness={skew_log}."
)


# ===========================================================================
# SALES: Customer Segmentation (K-Means)
# ===========================================================================
section("SALES: Customer Segmentation (K-Means)")

customer_data = sql("""
    SELECT
        customerid,
        COUNT(*)            AS order_count,
        SUM(totaldue)       AS total_spend,
        AVG(totaldue)       AS avg_order_value
    FROM sales.salesorderheader
    GROUP BY customerid
""")

features        = customer_data[["order_count", "total_spend", "avg_order_value"]].fillna(0)
features_scaled = StandardScaler().fit_transform(features)

kmeans = KMeans(n_clusters=N_CUSTOMER_CLUSTERS, random_state=42, n_init=10)
customer_data["cluster"] = kmeans.fit_predict(features_scaled)

cluster_means = customer_data.groupby("cluster")["total_spend"].mean().sort_values()
label_map     = {c: l for c, l in zip(cluster_means.index, ["Low Value", "Mid Value", "High Value"])}
customer_data["segment"] = customer_data["cluster"].map(label_map)

cluster_summary = (
    customer_data.groupby("segment")
    .agg(customers=("customerid","count"), avg_orders=("order_count","mean"),
         avg_spend=("total_spend","mean"), avg_order_val=("avg_order_value","mean"))
    .round(2).reset_index()
)
print(cluster_summary.to_string(index=False))

colors = {"Low Value": "#4C72B0", "Mid Value": "#DD8452", "High Value": "#55A868"}
fig, ax = plt.subplots(figsize=(9, 6))
for seg, grp in customer_data.groupby("segment"):
    ax.scatter(grp["order_count"], np.log10(grp["total_spend"] + 1),
               label=seg, alpha=0.5, s=20, color=colors[seg])
ax.set_title("Customer Segments (K-Means, 3 clusters)")
ax.set_xlabel("Number of Orders")
ax.set_ylabel("log10(Total Spend)")
ax.legend()
plt.tight_layout()
chart3 = fig_to_base64(fig)

add_result(
    "Customer Segments", cluster_summary, chart3,
    "K-Means on order count, total spend, avg order value. Y-axis log10 scaled."
)


# ===========================================================================
# SALES: Territory Performance
# ===========================================================================
section("SALES: Territory Performance")

territory = sql("""
    SELECT
        st.name                                 AS territory,
        st.countryregioncode                    AS country,
        COUNT(soh.salesorderid)                 AS order_count,
        SUM(soh.totaldue)                       AS total_revenue,
        AVG(soh.totaldue)                       AS avg_order_value
    FROM sales.salesorderheader    soh
    JOIN sales.salesterritory      st ON st.territoryid = soh.territoryid
    GROUP BY st.name, st.countryregioncode
    ORDER BY total_revenue DESC
""")
territory[["total_revenue", "avg_order_value"]] = territory[["total_revenue", "avg_order_value"]].round(2)

fig, ax = plt.subplots(figsize=(10, 5))
ax.barh(territory["territory"], territory["total_revenue"], color="#4C72B0", alpha=0.8)
ax.set_title("Total Revenue by Sales Territory")
ax.set_xlabel("Total Revenue ($)")
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
ax.invert_yaxis()
plt.tight_layout()
chart4 = fig_to_base64(fig)

add_result("Territory Performance", territory, chart4)


# ===========================================================================
# PRODUCTION: List Price vs Standard Cost (OLS Regression)
# ===========================================================================
section("PRODUCTION: List Price vs Standard Cost")

products = sql("""
    SELECT
        p.name          AS product_name,
        p.listprice     AS list_price,
        p.standardcost  AS standard_cost,
        pc.name         AS category
    FROM production.product             p
    LEFT JOIN production.productsubcategory  psc ON psc.productsubcategoryid = p.productsubcategoryid
    LEFT JOIN production.productcategory     pc  ON pc.productcategoryid     = psc.productcategoryid
    WHERE p.listprice > 0
      AND p.standardcost > 0
      AND psc.productsubcategoryid IS NOT NULL
""")

x2 = products["standard_cost"].values
y2 = products["list_price"].values
slope2, intercept2, r2, p2, se2 = stats.linregress(x2, y2)
r_sq2 = round(r2 ** 2, 4)
print(f"  slope={slope2:.3f}, intercept={intercept2:.2f}, R2={r_sq2}, p={p2:.4f}")

x2_line = np.linspace(x2.min(), x2.max(), 200)
y2_line = slope2 * x2_line + intercept2

fig, ax = plt.subplots(figsize=(8, 6))
for cat, grp in products.groupby("category"):
    ax.scatter(grp["standard_cost"], grp["list_price"], label=cat, alpha=0.7, s=40)
ax.plot(x2_line, y2_line, color="red", linewidth=2, label=f"OLS (R2={r_sq2})")
ax.set_title("List Price vs Standard Cost by Category")
ax.set_xlabel("Standard Cost ($)")
ax.set_ylabel("List Price ($)")
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
ax.legend(fontsize=8)
plt.tight_layout()
chart5 = fig_to_base64(fig)

prod_stats = pd.DataFrame([{
    "slope": round(slope2, 4),
    "intercept": round(intercept2, 2),
    "r_squared": r_sq2,
    "p_value": round(p2, 6),
    "std_error": round(se2, 4),
    "interpretation": f"For every $1 increase in cost, price rises by ${slope2:.2f}"
}])

add_result(
    "Price vs Cost Regression", prod_stats, chart5,
    f"OLS: ListPrice = {slope2:.3f} * StandardCost + {intercept2:.2f}. R2={r_sq2}."
)


# ===========================================================================
# PRODUCTION: Scrap Reason Breakdown
# ===========================================================================
section("PRODUCTION: Scrap Reason Breakdown")

scrap = sql("""
    SELECT
        sr.name                 AS scrap_reason,
        COUNT(*)                AS occurrences,
        SUM(wo.scrappedqty)     AS total_scrapped_qty
    FROM production.workorder    wo
    JOIN production.scrapreason  sr ON sr.scrapreasonid = wo.scrapreasonid
    GROUP BY sr.name
    ORDER BY total_scrapped_qty DESC
""")

fig, ax = plt.subplots(figsize=(10, 5))
ax.barh(scrap["scrap_reason"], scrap["total_scrapped_qty"], color="#C44E52", alpha=0.8)
ax.set_title("Total Scrapped Quantity by Reason")
ax.set_xlabel("Total Scrapped Units")
ax.invert_yaxis()
plt.tight_layout()
chart6 = fig_to_base64(fig)

add_result("Scrap Reasons", scrap, chart6)


# ===========================================================================
# HR: Pay Rate Distribution + Log10
# ===========================================================================
section("HR: Pay Rate Distribution + Log10 Transform")

pay = sql("""
    SELECT
        e.businessentityid,
        eph.rate                AS pay_rate,
        d.name                  AS department,
        e.jobtitle
    FROM humanresources.employee                        e
    JOIN humanresources.employeepayhistory              eph ON eph.businessentityid = e.businessentityid
    JOIN humanresources.employeedepartmenthistory       edh ON edh.businessentityid = e.businessentityid
                                                           AND edh.enddate IS NULL
    JOIN humanresources.department                      d   ON d.departmentid = edh.departmentid
""")

pay["log10_pay_rate"] = np.log10(pay["pay_rate"])
skew_pay     = round(float(pay["pay_rate"].skew()), 3)
skew_pay_log = round(float(pay["log10_pay_rate"].skew()), 3)
print(f"  Pay rate skewness raw:   {skew_pay}")
print(f"  Pay rate skewness log10: {skew_pay_log}")

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
axes[0].hist(pay["pay_rate"], bins=30, color="#4C72B0", alpha=0.8)
axes[0].set_title(f"Raw Pay Rate  (skew={skew_pay})")
axes[0].set_xlabel("Hourly Pay Rate ($)")
axes[1].hist(pay["log10_pay_rate"], bins=30, color="#DD8452", alpha=0.8)
axes[1].set_title(f"Log10(Pay Rate)  (skew={skew_pay_log})")
axes[1].set_xlabel("log10(Pay Rate)")
plt.tight_layout()
chart7 = fig_to_base64(fig)

pay_summary = (
    pay.groupby("department")["pay_rate"]
    .agg(employees="count", avg_rate="mean", min_rate="min", max_rate="max")
    .round(2).reset_index()
)

add_result(
    "HR Pay Distribution", pay_summary, chart7,
    f"Raw skewness={skew_pay}. Log10 skewness={skew_pay_log}."
)


# ===========================================================================
# HR: Headcount by Department
# ===========================================================================
section("HR: Headcount by Department")

headcount = sql("""
    SELECT
        d.name          AS department,
        d.groupname     AS group_name,
        COUNT(edh.businessentityid) AS employee_count
    FROM humanresources.department d
    LEFT JOIN humanresources.employeedepartmenthistory edh
           ON edh.departmentid = d.departmentid
          AND edh.enddate IS NULL
    GROUP BY d.name, d.groupname
    ORDER BY employee_count DESC
""")

fig, ax = plt.subplots(figsize=(10, 5))
ax.bar(headcount["department"], headcount["employee_count"], color="#4C72B0", alpha=0.8)
ax.set_title("Employee Headcount by Department")
ax.set_ylabel("Headcount")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
chart8 = fig_to_base64(fig)

add_result("HR Headcount", headcount, chart8)


# ===========================================================================
# PURCHASING: Vendor Spend Clustering (K-Means)
# ===========================================================================
section("PURCHASING: Vendor Spend Clustering")

vendor_data = sql("""
    SELECT
        v.name                          AS vendor_name,
        COUNT(poh.purchaseorderid)                              AS order_count,
        SUM(poh.subtotal + poh.taxamt + poh.freight)            AS total_spend,
        AVG(poh.subtotal + poh.taxamt + poh.freight)            AS avg_order_value,
        AVG(poh.freight)                                        AS avg_freight
    FROM purchasing.purchaseorderheader    poh
    JOIN purchasing.vendor                 v ON v.businessentityid = poh.vendorid
    GROUP BY v.name
""")

v_features = vendor_data[["order_count", "total_spend", "avg_order_value"]].fillna(0)
v_scaled   = StandardScaler().fit_transform(v_features)

vkmeans = KMeans(n_clusters=N_VENDOR_CLUSTERS, random_state=42, n_init=10)
vendor_data["cluster"] = vkmeans.fit_predict(v_scaled)

v_means = vendor_data.groupby("cluster")["total_spend"].mean().sort_values()
v_map   = {c: l for c, l in zip(v_means.index, ["Low Spend", "Mid Spend", "High Spend"])}
vendor_data["segment"] = vendor_data["cluster"].map(v_map)

vendor_summary = (
    vendor_data.groupby("segment")
    .agg(vendors=("vendor_name","count"), avg_orders=("order_count","mean"),
         avg_spend=("total_spend","mean"), avg_order_val=("avg_order_value","mean"))
    .round(2).reset_index()
)
print(vendor_summary.to_string(index=False))

v_colors = {"Low Spend": "#4C72B0", "Mid Spend": "#DD8452", "High Spend": "#55A868"}
fig, ax = plt.subplots(figsize=(9, 6))
for seg, grp in vendor_data.groupby("segment"):
    ax.scatter(grp["order_count"], np.log10(grp["total_spend"] + 1),
               label=seg, alpha=0.7, s=60, color=v_colors[seg])
ax.set_title("Vendor Segments (K-Means, 3 clusters)")
ax.set_xlabel("Number of Purchase Orders")
ax.set_ylabel("log10(Total Spend)")
ax.legend()
plt.tight_layout()
chart9 = fig_to_base64(fig)

add_result(
    "Vendor Segments", vendor_summary, chart9,
    "K-Means on order count, total spend, avg order value."
)


# ===========================================================================
# PURCHASING: Purchase Order Trends Over Time
# ===========================================================================
section("PURCHASING: Purchase Order Trends")

po_trend = sql("""
    SELECT
        TO_CHAR(orderdate, 'YYYY-MM')   AS year_month,
        COUNT(*)                                        AS order_count,
        SUM(subtotal + taxamt + freight)                AS total_spend,
        AVG(subtotal + taxamt + freight)                AS avg_order_value
    FROM purchasing.purchaseorderheader
    GROUP BY year_month
    ORDER BY year_month
""")
po_trend[["total_spend", "avg_order_value"]] = po_trend[["total_spend", "avg_order_value"]].round(2)

fig, ax = plt.subplots(figsize=(12, 4))
ax.bar(po_trend["year_month"], po_trend["total_spend"], color="#55A868", alpha=0.8)
ax.set_title("Monthly Purchasing Spend")
ax.set_xlabel("Month")
ax.set_ylabel("Total Spend ($)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
step = max(1, len(po_trend) // 12)
ax.set_xticks(po_trend["year_month"][::step])
ax.set_xticklabels(po_trend["year_month"][::step], rotation=45, ha="right")
plt.tight_layout()
chart10 = fig_to_base64(fig)

add_result("Purchasing Trends", po_trend, chart10)


# ===========================================================================
# OUTPUT: EXCEL
# ===========================================================================
section("Writing Excel Workbook")

with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as writer:
    for sheet, df in excel_sheets.items():
        df.to_excel(writer, sheet_name=sheet[:31], index=False)

print(f"  Saved: {EXCEL_OUT}")


# ===========================================================================
# OUTPUT: HTML REPORT
# ===========================================================================
section("Writing HTML Report")

nav_links = "".join(
    f'<a href="#{s.replace(" ", "-")}">{s}</a>'
    for s in excel_sheets.keys()
)

html_body = "\n".join(
    s.replace("<section>", f'<section id="{list(excel_sheets.keys())[i].replace(" ", "-")}">')
    for i, s in enumerate(html_sections)
)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>AdventureWorks Statistical Report</title>
<style>
  body        {{ font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f5f5f5; color: #222; }}
  header      {{ background: #1a3a5c; color: white; padding: 24px 40px; }}
  header h1   {{ margin: 0; font-size: 1.8em; }}
  header p    {{ margin: 4px 0 0; opacity: 0.8; }}
  nav         {{ background: #234e7a; padding: 10px 40px; display: flex; flex-wrap: wrap; gap: 12px; }}
  nav a       {{ color: #a8d0f5; text-decoration: none; font-size: 0.9em; }}
  nav a:hover {{ color: white; text-decoration: underline; }}
  main        {{ max-width: 1200px; margin: 30px auto; padding: 0 24px; }}
  section     {{ background: white; border-radius: 8px; padding: 28px; margin-bottom: 30px;
                 box-shadow: 0 2px 6px rgba(0,0,0,0.08); }}
  h2          {{ color: #1a3a5c; margin-top: 0; border-bottom: 2px solid #e0e8f0; padding-bottom: 8px; }}
  .chart      {{ width: 100%; max-width: 900px; display: block; margin: 16px 0; border-radius: 4px; }}
  .notes      {{ background: #eef4fb; border-left: 4px solid #4C72B0; padding: 10px 14px;
                 margin: 12px 0; font-size: 0.92em; color: #444; border-radius: 0 4px 4px 0; }}
  table.data-table {{ border-collapse: collapse; width: 100%; font-size: 0.88em; margin-top: 14px; }}
  .data-table th   {{ background: #1a3a5c; color: white; padding: 8px 12px; text-align: left; }}
  .data-table td   {{ padding: 6px 12px; border-bottom: 1px solid #e8e8e8; }}
  .data-table tr:nth-child(even) td {{ background: #f7f9fc; }}
  footer      {{ text-align: center; padding: 20px; color: #888; font-size: 0.85em; }}
</style>
</head>
<body>
<header>
  <h1>AdventureWorks Statistical Report</h1>
  <p>Sales &bull; Production &bull; HR &bull; Purchasing &mdash; Regression, Clustering, Log10 Analysis</p>
</header>
<nav>{nav_links}</nav>
<main>{html_body}</main>
<footer>Generated by db_report.py &mdash; AdventureWorks PostgreSQL</footer>
</body>
</html>"""

with open(HTML_OUT, "w", encoding="utf-8") as f:
    f.write(html)

print(f"  Saved: {HTML_OUT}")
print("\nDone! Open adventureworks_report.html in a browser for the full report.\n")
