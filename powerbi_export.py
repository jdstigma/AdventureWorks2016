"""
powerbi_export.py
=================
Exports AdventureWorks data from the local SQLite database to CSV files
ready for import into Power BI Desktop (no Azure account, no PostgreSQL needed).

HOW IT WORKS
------------
1. Opens adventureworks.db (SQLite) via SQLAlchemy.
2. Queries the six core tables (raw row-level data).
3. Saves each table as a CSV file inside the powerbi_data/ folder.

POWER BI DESKTOP IMPORT (one-time)
-----------------------------------
1. Open Power BI Desktop (free: microsoft.com/en-us/power-bi).
2. Home → Get Data → Text/CSV → select each file in powerbi_data/.
   Or: Home → Get Data → Folder → point at powerbi_data/ to load all at once.
3. Build relationships in Model view:
     SalesOrderDetail.salesorderid  →  SalesOrderHeader.salesorderid
     CustomerSegments.customerid    →  SalesOrderHeader.customerid
     VendorSegments.vendor_name     →  PurchaseOrder.vendor_name
4. Build visuals, then publish to Power BI Service if you have an account.

REFRESHING DATA
---------------
Re-run this script whenever you want fresh data, then in Power BI Desktop
click Home → Refresh.

CONFIGURATION
-------------
By default the script connects to adventureworks.db in the same directory.
Override by setting the DB_URL environment variable to any SQLAlchemy URL:

  # Use a different SQLite file
  export DB_URL="sqlite:////absolute/path/to/other.db"

  # Use PostgreSQL instead
  export DB_URL="postgresql+psycopg2://user:pass@localhost:5432/adventureworks"

USAGE
-----
  pip install -r requirements.txt
  python powerbi_export.py
"""

import os
import sys

import pandas as pd
from sqlalchemy import create_engine, text

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB = os.path.join(SCRIPT_DIR, "adventureworks.db")
DB_URL     = os.environ.get("DB_URL", f"sqlite:///{DEFAULT_DB}")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "powerbi_data")

# Detect dialect so we can use the right SQL syntax
IS_SQLITE = DB_URL.startswith("sqlite")

# ---------------------------------------------------------------------------
# DATABASE
# ---------------------------------------------------------------------------

def make_engine():
    return create_engine(DB_URL)


def query(engine, sql_str: str, label: str = "") -> pd.DataFrame:
    if label:
        print(f"  Querying: {label} ...")
    with engine.connect() as conn:
        df = pd.read_sql(text(sql_str), conn)
    print(f"    {len(df):,} rows")
    return df


def save_csv(df: pd.DataFrame, filename: str):
    path = os.path.join(OUTPUT_DIR, filename)
    df.to_csv(path, index=False)
    print(f"  Saved: {path}")


# ---------------------------------------------------------------------------
# SQL — two flavours: SQLite and PostgreSQL
# ---------------------------------------------------------------------------

# SQLite uses flat table names (no "schema." prefix) and different date funcs.
# PostgreSQL uses schema-qualified names and EXTRACT / TO_CHAR / LATERAL.

QUERIES_SQLITE = {
    "SalesOrderHeader.csv": """
        SELECT
            soh.salesorderid,
            soh.revisionnumber,
            soh.orderdate,
            soh.duedate,
            soh.shipdate,
            soh.status,
            soh.isonlineorder,
            soh.purchaseordernumber,
            soh.accountnumber,
            soh.customerid,
            soh.salespersonid,
            soh.territoryid,
            st.name                                   AS territory,
            st.countryregioncode                      AS country,
            soh.billtoaddressid,
            soh.shiptoaddressid,
            soh.shipmethodid,
            ROUND(soh.subtotal, 2)                    AS subtotal,
            ROUND(soh.taxamt, 2)                      AS taxamt,
            ROUND(soh.freight, 2)                     AS freight,
            ROUND(soh.totaldue, 2)                    AS totaldue,
            CAST(strftime('%Y', soh.orderdate) AS INTEGER) AS order_year,
            strftime('%Y-%m', soh.orderdate)          AS order_month
        FROM salesorderheader    soh
        JOIN salesterritory      st ON st.territoryid = soh.territoryid
        ORDER BY soh.orderdate
    """,

    "SalesOrderDetail.csv": """
        SELECT
            sod.salesorderid,
            sod.salesorderdetailid,
            sod.orderqty,
            sod.productid,
            p.name                                    AS product_name,
            psc.name                                  AS subcategory,
            pc.name                                   AS category,
            sod.specialofferid,
            ROUND(sod.unitprice, 2)                   AS unitprice,
            ROUND(sod.unitpricediscount, 4)           AS unitpricediscount,
            ROUND(sod.linetotal, 2)                   AS linetotal
        FROM salesorderdetail             sod
        JOIN product                      p   ON p.productid              = sod.productid
        LEFT JOIN productsubcategory      psc ON psc.productsubcategoryid = p.productsubcategoryid
        LEFT JOIN productcategory         pc  ON pc.productcategoryid     = psc.productcategoryid
        ORDER BY sod.salesorderid, sod.salesorderdetailid
    """,

    "Employee.csv": """
        SELECT
            e.businessentityid,
            e.nationalidnumber,
            e.jobtitle,
            e.birthdate,
            e.maritalstatus,
            e.gender,
            e.hiredate,
            e.salariedflag,
            e.vacationhours,
            e.sickleavehours,
            d.name                                    AS department,
            d.groupname                               AS department_group,
            ROUND(eph.rate, 2)                        AS pay_rate,
            eph.payfrequency                          AS pay_frequency
        FROM employee                               e
        JOIN employeedepartmenthistory              edh ON edh.businessentityid = e.businessentityid
                                                      AND edh.enddate IS NULL
        JOIN department                             d   ON d.departmentid = edh.departmentid
        JOIN employeepayhistory                     eph ON eph.businessentityid = e.businessentityid
                                                      AND eph.ratechangedate = (
                                                          SELECT MAX(ratechangedate)
                                                          FROM employeepayhistory
                                                          WHERE businessentityid = e.businessentityid
                                                      )
        ORDER BY e.businessentityid
    """,

    "PurchaseOrder.csv": """
        SELECT
            poh.purchaseorderid,
            poh.revisionnumber,
            poh.status,
            poh.employeeid,
            poh.vendorid,
            v.name                                    AS vendor_name,
            poh.shipmethodid,
            poh.orderdate,
            poh.shipdate,
            ROUND(poh.subtotal, 2)                    AS subtotal,
            ROUND(poh.taxamt, 2)                      AS taxamt,
            ROUND(poh.freight, 2)                     AS freight,
            ROUND(poh.subtotal + poh.taxamt + poh.freight, 2) AS totaldue,
            CAST(strftime('%Y', poh.orderdate) AS INTEGER)    AS order_year,
            strftime('%Y-%m', poh.orderdate)          AS order_month
        FROM purchaseorderheader    poh
        JOIN vendor                 v ON v.businessentityid = poh.vendorid
        ORDER BY poh.orderdate
    """,
}

QUERIES_POSTGRES = {
    "SalesOrderHeader.csv": """
        SELECT
            soh.salesorderid,
            soh.revisionnumber,
            soh.orderdate,
            soh.duedate,
            soh.shipdate,
            soh.status,
            soh.isonlineorder,
            soh.purchaseordernumber::text             AS purchaseordernumber,
            soh.accountnumber::text                   AS accountnumber,
            soh.customerid,
            soh.salespersonid,
            soh.territoryid,
            st.name::text                             AS territory,
            st.countryregioncode::text                AS country,
            soh.billtoaddressid,
            soh.shiptoaddressid,
            soh.shipmethodid,
            ROUND(soh.subtotal::numeric, 2)           AS subtotal,
            ROUND(soh.taxamt::numeric, 2)             AS taxamt,
            ROUND(soh.freight::numeric, 2)            AS freight,
            ROUND(soh.totaldue::numeric, 2)           AS totaldue,
            EXTRACT(YEAR FROM soh.orderdate)::int     AS order_year,
            TO_CHAR(soh.orderdate, 'YYYY-MM')         AS order_month
        FROM sales.salesorderheader    soh
        JOIN sales.salesterritory      st ON st.territoryid = soh.territoryid
        ORDER BY soh.orderdate
    """,

    "SalesOrderDetail.csv": """
        SELECT
            sod.salesorderid,
            sod.salesorderdetailid,
            sod.orderqty::int                         AS orderqty,
            sod.productid,
            p.name::text                              AS product_name,
            psc.name::text                            AS subcategory,
            pc.name::text                             AS category,
            sod.specialofferid,
            ROUND(sod.unitprice::numeric, 2)          AS unitprice,
            ROUND(sod.unitpricediscount::numeric, 4)  AS unitpricediscount,
            ROUND(sod.linetotal::numeric, 2)          AS linetotal
        FROM sales.salesorderdetail             sod
        JOIN production.product                 p   ON p.productid              = sod.productid
        LEFT JOIN production.productsubcategory psc ON psc.productsubcategoryid = p.productsubcategoryid
        LEFT JOIN production.productcategory    pc  ON pc.productcategoryid     = psc.productcategoryid
        ORDER BY sod.salesorderid, sod.salesorderdetailid
    """,

    "Employee.csv": """
        SELECT
            e.businessentityid,
            e.nationalidnumber::text                  AS nationalidnumber,
            e.jobtitle::text                          AS jobtitle,
            e.birthdate,
            e.maritalstatus::text                     AS maritalstatus,
            e.gender::text                            AS gender,
            e.hiredate,
            e.salariedflag,
            e.vacationhours,
            e.sickleavehours,
            d.name::text                              AS department,
            d.groupname::text                         AS department_group,
            ROUND(eph.rate::numeric, 2)               AS pay_rate,
            eph.payfrequency                          AS pay_frequency
        FROM humanresources.employee                        e
        JOIN humanresources.employeedepartmenthistory       edh ON edh.businessentityid = e.businessentityid
                                                               AND edh.enddate IS NULL
        JOIN humanresources.department                      d   ON d.departmentid = edh.departmentid
        JOIN LATERAL (
            SELECT rate, payfrequency
            FROM humanresources.employeepayhistory
            WHERE businessentityid = e.businessentityid
            ORDER BY ratechangedate DESC
            LIMIT 1
        ) eph ON TRUE
        ORDER BY e.businessentityid
    """,

    "PurchaseOrder.csv": """
        SELECT
            poh.purchaseorderid,
            poh.revisionnumber,
            poh.status,
            poh.employeeid,
            poh.vendorid,
            v.name::text                              AS vendor_name,
            poh.shipmethodid,
            poh.orderdate,
            poh.shipdate,
            ROUND(poh.subtotal::numeric, 2)           AS subtotal,
            ROUND(poh.taxamt::numeric, 2)             AS taxamt,
            ROUND(poh.freight::numeric, 2)            AS freight,
            ROUND((poh.subtotal+poh.taxamt+poh.freight)::numeric, 2) AS totaldue,
            EXTRACT(YEAR FROM poh.orderdate)::int     AS order_year,
            TO_CHAR(poh.orderdate, 'YYYY-MM')         AS order_month
        FROM purchasing.purchaseorderheader    poh
        JOIN purchasing.vendor                 v ON v.businessentityid = poh.vendorid
        ORDER BY poh.orderdate
    """,
}

QUERIES = QUERIES_SQLITE if IS_SQLITE else QUERIES_POSTGRES

# ---------------------------------------------------------------------------
# K-MEANS SEGMENTS  (require scikit-learn; skipped gracefully if not installed)
# ---------------------------------------------------------------------------

def _segment_raw_customer(engine) -> pd.DataFrame:
    tbl = "salesorderheader" if IS_SQLITE else "sales.salesorderheader"
    return query(engine, f"""
        SELECT customerid,
               COUNT(*)       AS order_count,
               SUM(totaldue)  AS total_spend,
               AVG(totaldue)  AS avg_order_value
        FROM {tbl}
        GROUP BY customerid
    """, label="CustomerSegments (raw)")


def _segment_raw_vendor(engine) -> pd.DataFrame:
    poh = "purchaseorderheader" if IS_SQLITE else "purchasing.purchaseorderheader"
    v   = "vendor"              if IS_SQLITE else "purchasing.vendor"
    return query(engine, f"""
        SELECT v.name                                  AS vendor_name,
               COUNT(poh.purchaseorderid)              AS order_count,
               SUM(poh.subtotal+poh.taxamt+poh.freight) AS total_spend,
               AVG(poh.subtotal+poh.taxamt+poh.freight) AS avg_order_value
        FROM {poh}    poh
        JOIN {v}      v ON v.businessentityid = poh.vendorid
        GROUP BY v.name
    """, label="VendorSegments (raw)")


def _kmeans_segment(raw: pd.DataFrame, feature_col: str,
                    id_col: str, labels: list[str]) -> pd.DataFrame:
    try:
        from sklearn.cluster import KMeans
        from sklearn.preprocessing import StandardScaler
    except ImportError:
        print("  WARNING: scikit-learn not installed; skipping segment.")
        return pd.DataFrame()

    features = raw[["order_count", "total_spend", "avg_order_value"]].fillna(0)
    scaled   = StandardScaler().fit_transform(features)
    km       = KMeans(n_clusters=3, random_state=42, n_init=10)
    raw      = raw.copy()
    raw["cluster"] = km.fit_predict(scaled)

    order     = raw.groupby("cluster")[feature_col].mean().sort_values().index
    label_map = {c: l for c, l in zip(order, labels)}
    raw["segment"] = raw["cluster"].map(label_map)

    return (
        raw[[id_col, "order_count", "total_spend", "avg_order_value", "segment"]]
        .round({"total_spend": 2, "avg_order_value": 2})
        .reset_index(drop=True)
    )


def build_customer_segments(engine) -> pd.DataFrame:
    raw = _segment_raw_customer(engine)
    return _kmeans_segment(raw, "total_spend", "customerid",
                           ["Low Value", "Mid Value", "High Value"])


def build_vendor_segments(engine) -> pd.DataFrame:
    raw = _segment_raw_vendor(engine)
    return _kmeans_segment(raw, "total_spend", "vendor_name",
                           ["Low Spend", "Mid Spend", "High Spend"])


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("\n=== AdventureWorks → Power BI CSV Export ===\n")
    print(f"  Source:  {DB_URL}")
    print(f"  Output:  {OUTPUT_DIR}\n")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    engine = make_engine()
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("  Database connection OK\n")
    except Exception as exc:
        sys.exit(f"  Connection failed: {exc}")

    print("--- Querying and exporting ---")
    for filename, sql_str in QUERIES.items():
        label = filename.replace(".csv", "")
        df    = query(engine, sql_str, label=label)
        save_csv(df, filename)

    df_customers = build_customer_segments(engine)
    if not df_customers.empty:
        save_csv(df_customers, "CustomerSegments.csv")

    df_vendors = build_vendor_segments(engine)
    if not df_vendors.empty:
        save_csv(df_vendors, "VendorSegments.csv")

    print(f"\nDone! Import the CSV files from:\n  {OUTPUT_DIR}\n")
    print("In Power BI Desktop: Home → Get Data → Text/CSV (or Folder)")
    print("Suggested relationships:")
    print("  SalesOrderDetail.salesorderid → SalesOrderHeader.salesorderid")
    print("  CustomerSegments.customerid   → SalesOrderHeader.customerid")
    print("  VendorSegments.vendor_name    → PurchaseOrder.vendor_name\n")


if __name__ == "__main__":
    main()
