"""
db_query.py
===========
A reusable template for running SQL queries against the AdventureWorks
SQLite database and loading results into a pandas DataFrame.

REQUIREMENTS
------------
Install these once in your terminal:
    pip install sqlalchemy pandas openpyxl

(No extra drivers needed — SQLite support is built into Python.)

HOW TO USE
----------
1. Make sure adventureworks.db is in the SAME folder as this script.
   (Or update DB_PATH below to the full path of your .db file.)
2. Edit the SQL in the QUERIES section to ask your own questions.
3. Run:  python db_query.py
4. Uncomment any save_to_* line at the bottom to export results.
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------
import pandas as pd
from sqlalchemy import create_engine, text

# ---------------------------------------------------------------------------
# CONFIG  ← Only change DB_PATH if your .db file is in a different folder
# ---------------------------------------------------------------------------
DB_PATH = "adventureworks.db"

# ---------------------------------------------------------------------------
# CONNECTION
# ---------------------------------------------------------------------------
CONNECTION_STRING = f"sqlite:///{DB_PATH}"

print("Connecting to database...")
engine = create_engine(CONNECTION_STRING)

try:
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print(f"  ✓ Connected to '{DB_PATH}'\n")
except Exception as e:
    print(f"  ✗ Connection failed: {e}")
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------------------------

def run_query(sql, params=None, label="Query"):
    """
    Run a SQL string and return results as a pandas DataFrame.

    Parameters
    ----------
    sql    : str   — your SQL statement (use :param_name for parameters)
    params : dict  — optional dictionary of parameter values
    label  : str   — a friendly name printed in the output

    Returns
    -------
    pd.DataFrame — query results, ready to work with in pandas
    """
    print(f"Running: {label}...")
    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)
    print(f"  → {len(df):,} rows returned\n")
    return df


def save_to_excel(df, filename, sheet_name="Results"):
    """
    Save a DataFrame to an Excel file.
    Requires:  pip install openpyxl
    """
    df.to_excel(filename, index=False, sheet_name=sheet_name)
    print(f"  ✓ Saved to Excel: {filename}")


def save_to_csv(df, filename):
    """Save a DataFrame to a CSV file."""
    df.to_csv(filename, index=False)
    print(f"  ✓ Saved to CSV: {filename}")


def show(df, max_rows=20):
    """Print a DataFrame nicely, limiting to max_rows."""
    pd.set_option("display.max_columns", None)   # show all columns
    pd.set_option("display.width",       200)    # wider terminal output
    pd.set_option("display.float_format", "{:,.2f}".format)  # comma-format numbers
    print(df.head(max_rows).to_string(index=False))
    if len(df) > max_rows:
        print(f"  ... ({len(df) - max_rows:,} more rows not shown)")
    print()


# ---------------------------------------------------------------------------
# QUERIES  ← Replace or add your own SQL here
# ---------------------------------------------------------------------------

# ── Query 1: Sales order summary by year ────────────────────────────────────
# strftime is SQLite's way to extract date parts (no EXTRACT() like in PostgreSQL)
q1 = run_query(
    sql="""
        SELECT
            strftime('%Y', OrderDate)           AS order_year,
            COUNT(*)                            AS order_count,
            SUM(SubTotal)                       AS total_sales,
            AVG(SubTotal)                       AS avg_order_value,
            COUNT(DISTINCT CustomerID)          AS unique_customers
        FROM Sales_SalesOrderHeader
        GROUP BY order_year
        ORDER BY order_year
    """,
    label="Sales Summary by Year"
)
show(q1)


# ── Query 2: Top 10 products by revenue ─────────────────────────────────────
q2 = run_query(
    sql="""
        SELECT
            p.Name                              AS product_name,
            pc.Name                             AS category,
            SUM(sod.LineTotal)                  AS total_revenue,
            SUM(sod.OrderQty)                   AS units_sold,
            AVG(sod.UnitPrice)                  AS avg_unit_price
        FROM Sales_SalesOrderDetail         sod
        JOIN Production_Product             p   ON p.ProductID          = sod.ProductID
        JOIN Production_ProductSubcategory  psc ON psc.ProductSubcategoryID = p.ProductSubcategoryID
        JOIN Production_ProductCategory     pc  ON pc.ProductCategoryID  = psc.ProductCategoryID
        GROUP BY p.Name, pc.Name
        ORDER BY total_revenue DESC
        LIMIT 10
    """,
    label="Top 10 Products by Revenue"
)
show(q2)


# ── Query 3: Parameterized — filter orders by territory name ─────────────────
#   Change the territory_name value below to explore different regions
q3 = run_query(
    sql="""
        SELECT
            soh.SalesOrderID,
            soh.OrderDate,
            soh.TotalDue,
            st.Name             AS territory,
            c.AccountNumber     AS customer
        FROM Sales_SalesOrderHeader  soh
        JOIN Sales_SalesTerritory    st  ON st.TerritoryID = soh.TerritoryID
        JOIN Sales_Customer          c   ON c.CustomerID   = soh.CustomerID
        WHERE st.Name = :territory_name
        ORDER BY soh.OrderDate DESC
        LIMIT 50
    """,
    params={"territory_name": "Northwest"},   # ← change this to explore other territories
    label="Recent Orders — Northwest Territory"
)
show(q3)


# ── Query 4: Employee headcount by department ────────────────────────────────
q4 = run_query(
    sql="""
        SELECT
            d.Name                              AS department,
            d.GroupName                         AS department_group,
            COUNT(edh.BusinessEntityID)         AS employee_count
        FROM HumanResources_Department              d
        LEFT JOIN HumanResources_EmployeeDepartmentHistory edh
               ON edh.DepartmentID = d.DepartmentID
              AND edh.EndDate IS NULL            -- current assignments only
        GROUP BY d.Name, d.GroupName
        ORDER BY employee_count DESC
    """,
    label="Employee Headcount by Department"
)
show(q4)


# ---------------------------------------------------------------------------
# SAVE RESULTS TO FILES
# ---------------------------------------------------------------------------
save_to_excel(q1, "sales_by_year.xlsx",    sheet_name="Sales by Year")
save_to_excel(q2, "top_products.xlsx",     sheet_name="Top Products")
save_to_excel(q3, "northwest_orders.xlsx", sheet_name="Northwest Orders")
save_to_excel(q4, "employee_headcount.xlsx", sheet_name="Headcount")


# ---------------------------------------------------------------------------
# ADD YOUR OWN QUERY HERE
# ---------------------------------------------------------------------------
# Copy this block and paste your SQL between the triple quotes:
#
# my_query = run_query(
#     sql="""
#         SELECT *
#         FROM your_schema.your_table
#         LIMIT 100
#     """,
#     label="My Custom Query"
# )
# show(my_query)
# save_to_excel(my_query, "my_results.xlsx")
