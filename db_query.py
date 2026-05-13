"""
db_query.py
===========
Runs SQL queries against AdventureWorks in PostgreSQL and saves results to Excel.

REQUIREMENTS
------------
    pip install -r requirements.txt

HOW TO USE
----------
1. Update the CONFIG section with your connection details if needed.
2. Run:  python db_query.py
3. Excel files are saved next to this script.

NOTE: Do not commit real passwords to a public GitHub repo.
      Move credentials to environment variables before going public.
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------
import os
import pandas as pd
from sqlalchemy import create_engine, text

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
DB_HOST     = "localhost"
DB_PORT     = 5432
DB_NAME     = "adventureworks"
DB_USER     = "awuser"
DB_PASSWORD = "Gunner!!24"

CONNECTION_STRING = (
    f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
    f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"
)

# ---------------------------------------------------------------------------
# CONNECTION
# ---------------------------------------------------------------------------
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

def run_query(sql, params=None, label="Query"):
    """Run a SQL string and return results as a pandas DataFrame."""
    print(f"Running: {label}...")
    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)
    print(f"  {len(df):,} rows returned\n")
    return df


def save_to_excel(df, filename, sheet_name="Results"):
    """Save a DataFrame to an Excel file next to this script."""
    full_path = os.path.join(SCRIPT_DIR, filename)
    df.to_excel(full_path, index=False, sheet_name=sheet_name)
    print(f"  Saved: {full_path}")


def save_to_csv(df, filename):
    """Save a DataFrame to a CSV file next to this script."""
    full_path = os.path.join(SCRIPT_DIR, filename)
    df.to_csv(full_path, index=False)
    print(f"  Saved: {full_path}")


def show(df, max_rows=20):
    """Print a DataFrame neatly."""
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    pd.set_option("display.float_format", "{:,.2f}".format)
    print(df.head(max_rows).to_string(index=False))
    if len(df) > max_rows:
        print(f"  ... ({len(df) - max_rows:,} more rows not shown)")
    print()


# ---------------------------------------------------------------------------
# QUERIES
# ---------------------------------------------------------------------------

# Query 1: Sales order summary by year
q1 = run_query(
    sql="""
        SELECT
            EXTRACT(YEAR FROM orderdate)::int      AS order_year,
            COUNT(*)                               AS order_count,
            SUM(totaldue)                          AS total_sales,
            AVG(totaldue)                          AS avg_order_value,
            COUNT(DISTINCT customerid)             AS unique_customers
        FROM sales.salesorderheader
        GROUP BY order_year
        ORDER BY order_year
    """,
    label="Sales Summary by Year"
)
show(q1)


# Query 2: Top 10 products by revenue
q2 = run_query(
    sql="""
        SELECT
            p.name                                 AS product_name,
            pc.name                                AS category,
            SUM(sod.unitprice * (1.0 - sod.unitpricediscount) * sod.orderqty) AS total_revenue,
            SUM(sod.orderqty)                      AS units_sold,
            AVG(sod.unitprice)                     AS avg_unit_price
        FROM sales.salesorderdetail            sod
        JOIN production.product                p   ON p.productid              = sod.productid
        JOIN production.productsubcategory     psc ON psc.productsubcategoryid = p.productsubcategoryid
        JOIN production.productcategory        pc  ON pc.productcategoryid     = psc.productcategoryid
        GROUP BY p.name, pc.name
        ORDER BY total_revenue DESC
        LIMIT 10
    """,
    label="Top 10 Products by Revenue"
)
show(q2)


# Query 3: Filter orders by territory -- change :territory_name to explore regions
q3 = run_query(
    sql="""
        SELECT
            soh.salesorderid,
            soh.orderdate,
            soh.totaldue,
            st.name::text                          AS territory,
            soh.accountnumber                      AS customer
        FROM sales.salesorderheader    soh
        JOIN sales.salesterritory      st  ON st.territoryid = soh.territoryid
        WHERE st.name::text = :territory_name
        ORDER BY soh.orderdate DESC
        LIMIT 50
    """,
    params={"territory_name": "Northwest"},
    label="Recent Orders - Northwest Territory"
)
show(q3)


# Query 4: Employee headcount by department
q4 = run_query(
    sql="""
        SELECT
            d.name                                 AS department,
            d.groupname                            AS department_group,
            COUNT(edh.businessentityid)            AS employee_count
        FROM humanresources.department             d
        LEFT JOIN humanresources.employeedepartmenthistory edh
               ON edh.departmentid = d.departmentid
              AND edh.enddate IS NULL
        GROUP BY d.name, d.groupname
        ORDER BY employee_count DESC
    """,
    label="Employee Headcount by Department"
)
show(q4)


# ---------------------------------------------------------------------------
# SAVE TO EXCEL
# ---------------------------------------------------------------------------
save_to_excel(q1, "sales_by_year.xlsx",      sheet_name="Sales by Year")
save_to_excel(q2, "top_products.xlsx",       sheet_name="Top Products")
save_to_excel(q3, "northwest_orders.xlsx",   sheet_name="Northwest Orders")
save_to_excel(q4, "employee_headcount.xlsx", sheet_name="Headcount")


# ---------------------------------------------------------------------------
# ADD YOUR OWN QUERY HERE
# ---------------------------------------------------------------------------
# my_query = run_query(
#     sql="""
#         SELECT *
#         FROM schema_name.table_name
#         LIMIT 100
#     """,
#     label="My Custom Query"
# )
# show(my_query)
# save_to_excel(my_query, "my_results.xlsx")
