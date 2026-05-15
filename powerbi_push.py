"""
powerbi_push.py
===============
Pushes AdventureWorks data from PostgreSQL to Power BI using the Push Datasets API.

HOW IT WORKS
------------
1. Reads environment variables for credentials (Azure AD + DB).
2. Acquires an OAuth 2.0 token from Azure AD via MSAL.
3. Creates a Power BI Push Dataset (or reuses an existing one by name).
4. Queries PostgreSQL for the six core datasets.
5. Pushes all rows to Power BI in batches.

POWER BI SETUP (one-time)
--------------------------
1. Register an app in Azure Active Directory:
     https://portal.azure.com → Azure Active Directory → App registrations → New registration
2. Under API Permissions add:
     Power BI Service → Delegated → Dataset.ReadWrite.All, Workspace.Read.All
3. Grant admin consent for your organisation.
4. Note the Application (client) ID and Tenant ID.
5. For service-principal flow also create a Client Secret and add the SP to your
   Power BI workspace with at least Member role.

ENVIRONMENT VARIABLES
---------------------
Required:
  POWERBI_TENANT_ID      – Azure AD tenant ID (GUID)
  POWERBI_CLIENT_ID      – App registration client ID
  POWERBI_CLIENT_SECRET  – Client secret (service-principal flow)
  POWERBI_WORKSPACE_ID   – Power BI workspace (group) ID  [leave empty for "My Workspace"]

Optional (override DB defaults):
  DB_HOST      DB_PORT      DB_NAME      DB_USER      DB_PASSWORD

USAGE
-----
  pip install -r requirements.txt
  python powerbi_push.py

  # Dry-run (no network calls to Power BI):
  POWERBI_DRY_RUN=1 python powerbi_push.py
"""

import os
import sys
import json
import time
import math
import logging

import pandas as pd
from sqlalchemy import create_engine, text

# Optional: suppress MSAL's noisy INFO logs
logging.getLogger("msal").setLevel(logging.WARNING)

try:
    import msal
except ImportError:
    sys.exit("ERROR: 'msal' not installed.  Run: pip install msal")

try:
    import requests
except ImportError:
    sys.exit("ERROR: 'requests' not installed.  Run: pip install requests")

# ---------------------------------------------------------------------------
# CONFIGURATION  (all secrets come from environment variables)
# ---------------------------------------------------------------------------

# Azure AD / Power BI
TENANT_ID     = os.environ.get("POWERBI_TENANT_ID", "")
CLIENT_ID     = os.environ.get("POWERBI_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("POWERBI_CLIENT_SECRET", "")
WORKSPACE_ID  = os.environ.get("POWERBI_WORKSPACE_ID", "")   # empty = My Workspace
DATASET_NAME  = os.environ.get("POWERBI_DATASET_NAME", "AdventureWorks2016")
DRY_RUN       = os.environ.get("POWERBI_DRY_RUN", "0").strip() not in ("", "0", "false", "no")

# PostgreSQL (fall back to the same defaults used by db_query.py)
DB_HOST     = os.environ.get("DB_HOST",     "localhost")
DB_PORT     = int(os.environ.get("DB_PORT", "5432"))
DB_NAME     = os.environ.get("DB_NAME",     "adventureworks")
DB_USER     = os.environ.get("DB_USER",     "awuser")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "")   # set via env; no hardcoded default

# Power BI REST API
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
PBI_SCOPE     = ["https://analysis.windows.net/powerbi/api/.default"]
PBI_API_BASE  = "https://api.powerbi.com/v1.0/myorg"

PUSH_BATCH_SIZE = 9_000   # Power BI limit: 10,000 rows per POST; stay under it

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def _check_env():
    """Fail fast if required env vars are missing (unless dry-run)."""
    if DRY_RUN:
        return
    missing = [v for v in ("POWERBI_TENANT_ID", "POWERBI_CLIENT_ID", "POWERBI_CLIENT_SECRET")
               if not os.environ.get(v)]
    if missing:
        sys.exit(
            f"ERROR: Missing environment variable(s): {', '.join(missing)}\n"
            "Set them before running, e.g.:\n"
            "  export POWERBI_TENANT_ID=<guid>\n"
            "  export POWERBI_CLIENT_ID=<guid>\n"
            "  export POWERBI_CLIENT_SECRET=<secret>\n"
        )


def get_access_token() -> str:
    """Acquire a bearer token using the client-credentials (service-principal) flow."""
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=PBI_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(
            f"Failed to acquire Power BI token: {result.get('error_description', result)}"
        )
    return result["access_token"]


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _groups_url() -> str:
    if WORKSPACE_ID:
        return f"{PBI_API_BASE}/groups/{WORKSPACE_ID}"
    return f"{PBI_API_BASE}"   # My Workspace has no /groups prefix


def _datasets_url() -> str:
    return f"{_groups_url()}/datasets"


def list_datasets(token: str) -> list:
    """Return list of datasets in the target workspace."""
    resp = requests.get(_datasets_url(), headers=_headers(token), timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def find_dataset_id(token: str, name: str) -> str | None:
    """Return dataset ID if a dataset with *name* already exists, else None."""
    for ds in list_datasets(token):
        if ds.get("name") == name:
            return ds["id"]
    return None


def delete_dataset(token: str, dataset_id: str):
    """Delete an existing dataset so we can recreate it with a fresh schema."""
    url  = f"{_datasets_url()}/{dataset_id}"
    resp = requests.delete(url, headers=_headers(token), timeout=30)
    resp.raise_for_status()
    print(f"  Deleted existing dataset: {dataset_id}")


def create_dataset(token: str, schema: dict) -> str:
    """Create a push dataset and return its ID."""
    resp = requests.post(
        _datasets_url(),
        headers=_headers(token),
        params={"defaultRetentionPolicy": "basicFIFO"},
        data=json.dumps(schema),
        timeout=60,
    )
    resp.raise_for_status()
    dataset_id = resp.json()["id"]
    print(f"  Created dataset '{schema['name']}' → id={dataset_id}")
    return dataset_id


def push_rows(token: str, dataset_id: str, table_name: str, df: pd.DataFrame):
    """Push a DataFrame to a Power BI push dataset table in batches."""
    url        = f"{_datasets_url()}/{dataset_id}/tables/{table_name}/rows"
    total_rows = len(df)
    n_batches  = math.ceil(total_rows / PUSH_BATCH_SIZE)

    for i in range(n_batches):
        chunk = df.iloc[i * PUSH_BATCH_SIZE : (i + 1) * PUSH_BATCH_SIZE]
        rows  = chunk.to_dict(orient="records")
        body  = json.dumps({"rows": rows}, default=str)   # default=str handles dates/Decimals

        if DRY_RUN:
            print(f"    [DRY RUN] Would push {len(rows)} rows to '{table_name}' (batch {i+1}/{n_batches})")
            continue

        resp = requests.post(url, headers=_headers(token), data=body, timeout=60)
        if resp.status_code == 429:           # rate-limited
            wait = int(resp.headers.get("Retry-After", 10))
            print(f"    Rate-limited; sleeping {wait}s …")
            time.sleep(wait)
            resp = requests.post(url, headers=_headers(token), data=body, timeout=60)
        resp.raise_for_status()
        print(f"    Pushed batch {i+1}/{n_batches}: {len(rows):,} rows → '{table_name}'")


def clear_table_rows(token: str, dataset_id: str, table_name: str):
    """Clear all rows from a push dataset table before re-pushing."""
    url  = f"{_datasets_url()}/{dataset_id}/tables/{table_name}/rows"
    resp = requests.delete(url, headers=_headers(token), timeout=30)
    resp.raise_for_status()


# ---------------------------------------------------------------------------
# DATABASE
# ---------------------------------------------------------------------------

def make_engine():
    if not DB_PASSWORD:
        sys.exit(
            "ERROR: DB_PASSWORD environment variable not set.\n"
            "  export DB_PASSWORD=<your_postgres_password>"
        )
    conn_str = (
        f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
        f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    )
    return create_engine(conn_str)


def query(engine, sql_str: str, label: str = "") -> pd.DataFrame:
    if label:
        print(f"  Querying: {label} …")
    with engine.connect() as conn:
        df = pd.read_sql(text(sql_str), conn)
    print(f"    {len(df):,} rows")
    return df


# ---------------------------------------------------------------------------
# DATASET SCHEMA  (Power BI column types: Int64 / Double / String / DateTime / Boolean)
# ---------------------------------------------------------------------------

DATASET_SCHEMA = {
    "name": DATASET_NAME,
    "tables": [
        {
            "name": "SalesByYear",
            "columns": [
                {"name": "order_year",        "dataType": "Int64"},
                {"name": "order_count",        "dataType": "Int64"},
                {"name": "total_sales",        "dataType": "Double"},
                {"name": "avg_order_value",    "dataType": "Double"},
                {"name": "unique_customers",   "dataType": "Int64"},
            ],
        },
        {
            "name": "Products",
            "columns": [
                {"name": "product_name",    "dataType": "String"},
                {"name": "category",        "dataType": "String"},
                {"name": "total_revenue",   "dataType": "Double"},
                {"name": "units_sold",      "dataType": "Int64"},
                {"name": "avg_unit_price",  "dataType": "Double"},
            ],
        },
        {
            "name": "TerritoryPerformance",
            "columns": [
                {"name": "territory",        "dataType": "String"},
                {"name": "country",          "dataType": "String"},
                {"name": "order_count",      "dataType": "Int64"},
                {"name": "total_revenue",    "dataType": "Double"},
                {"name": "avg_order_value",  "dataType": "Double"},
            ],
        },
        {
            "name": "CustomerSegments",
            "columns": [
                {"name": "segment",          "dataType": "String"},
                {"name": "customers",        "dataType": "Int64"},
                {"name": "avg_orders",       "dataType": "Double"},
                {"name": "avg_spend",        "dataType": "Double"},
                {"name": "avg_order_val",    "dataType": "Double"},
            ],
        },
        {
            "name": "EmployeeHeadcount",
            "columns": [
                {"name": "department",        "dataType": "String"},
                {"name": "department_group",  "dataType": "String"},
                {"name": "employee_count",    "dataType": "Int64"},
            ],
        },
        {
            "name": "VendorSegments",
            "columns": [
                {"name": "segment",          "dataType": "String"},
                {"name": "vendors",          "dataType": "Int64"},
                {"name": "avg_orders",       "dataType": "Double"},
                {"name": "avg_spend",        "dataType": "Double"},
                {"name": "avg_order_val",    "dataType": "Double"},
            ],
        },
        {
            "name": "PurchasingTrends",
            "columns": [
                {"name": "year_month",       "dataType": "String"},
                {"name": "order_count",      "dataType": "Int64"},
                {"name": "total_spend",      "dataType": "Double"},
                {"name": "avg_order_value",  "dataType": "Double"},
            ],
        },
        {
            "name": "SalesOrders",
            "columns": [
                {"name": "salesorderid",     "dataType": "Int64"},
                {"name": "orderdate",        "dataType": "DateTime"},
                {"name": "totaldue",         "dataType": "Double"},
                {"name": "territory",        "dataType": "String"},
                {"name": "order_year",       "dataType": "Int64"},
                {"name": "order_month",      "dataType": "String"},
                {"name": "customerid",       "dataType": "Int64"},
                {"name": "status",           "dataType": "Int64"},
            ],
        },
    ],
}

# ---------------------------------------------------------------------------
# QUERIES
# ---------------------------------------------------------------------------

QUERIES = {
    "SalesByYear": """
        SELECT
            EXTRACT(YEAR FROM orderdate)::int   AS order_year,
            COUNT(*)                            AS order_count,
            ROUND(SUM(totaldue)::numeric, 2)    AS total_sales,
            ROUND(AVG(totaldue)::numeric, 2)    AS avg_order_value,
            COUNT(DISTINCT customerid)          AS unique_customers
        FROM sales.salesorderheader
        GROUP BY order_year
        ORDER BY order_year
    """,

    "Products": """
        SELECT
            p.name                                                              AS product_name,
            pc.name                                                             AS category,
            ROUND(SUM(sod.unitprice*(1.0-sod.unitpricediscount)*sod.orderqty)
                  ::numeric, 2)                                                 AS total_revenue,
            SUM(sod.orderqty)::int                                             AS units_sold,
            ROUND(AVG(sod.unitprice)::numeric, 2)                              AS avg_unit_price
        FROM sales.salesorderdetail            sod
        JOIN production.product                p   ON p.productid              = sod.productid
        JOIN production.productsubcategory     psc ON psc.productsubcategoryid = p.productsubcategoryid
        JOIN production.productcategory        pc  ON pc.productcategoryid     = psc.productcategoryid
        GROUP BY p.name, pc.name
        ORDER BY total_revenue DESC
    """,

    "TerritoryPerformance": """
        SELECT
            st.name                                 AS territory,
            st.countryregioncode                    AS country,
            COUNT(soh.salesorderid)                 AS order_count,
            ROUND(SUM(soh.totaldue)::numeric, 2)    AS total_revenue,
            ROUND(AVG(soh.totaldue)::numeric, 2)    AS avg_order_value
        FROM sales.salesorderheader    soh
        JOIN sales.salesterritory      st ON st.territoryid = soh.territoryid
        GROUP BY st.name, st.countryregioncode
        ORDER BY total_revenue DESC
    """,

    "EmployeeHeadcount": """
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

    "PurchasingTrends": """
        SELECT
            TO_CHAR(orderdate, 'YYYY-MM')                    AS year_month,
            COUNT(*)                                          AS order_count,
            ROUND(SUM(subtotal+taxamt+freight)::numeric, 2)  AS total_spend,
            ROUND(AVG(subtotal+taxamt+freight)::numeric, 2)  AS avg_order_value
        FROM purchasing.purchaseorderheader
        GROUP BY year_month
        ORDER BY year_month
    """,

    "SalesOrders": """
        SELECT
            soh.salesorderid,
            soh.orderdate,
            ROUND(soh.totaldue::numeric, 2)         AS totaldue,
            st.name::text                           AS territory,
            EXTRACT(YEAR FROM soh.orderdate)::int   AS order_year,
            TO_CHAR(soh.orderdate, 'YYYY-MM')       AS order_month,
            soh.customerid,
            soh.status
        FROM sales.salesorderheader    soh
        JOIN sales.salesterritory      st ON st.territoryid = soh.territoryid
        ORDER BY soh.orderdate
    """,
}

# CustomerSegments and VendorSegments need K-Means — we compute them separately.

def build_customer_segments(engine) -> pd.DataFrame:
    try:
        from sklearn.cluster import KMeans
        from sklearn.preprocessing import StandardScaler
        import numpy as np
    except ImportError:
        print("  WARNING: scikit-learn not installed; skipping CustomerSegments.")
        return pd.DataFrame()

    raw = query(engine, """
        SELECT customerid,
               COUNT(*)       AS order_count,
               SUM(totaldue)  AS total_spend,
               AVG(totaldue)  AS avg_order_value
        FROM sales.salesorderheader
        GROUP BY customerid
    """, label="CustomerSegments (raw)")

    features = raw[["order_count", "total_spend", "avg_order_value"]].fillna(0)
    scaled   = StandardScaler().fit_transform(features)
    km       = KMeans(n_clusters=3, random_state=42, n_init=10)
    raw["cluster"] = km.fit_predict(scaled)

    order = raw.groupby("cluster")["total_spend"].mean().sort_values().index
    label_map = {c: l for c, l in zip(order, ["Low Value", "Mid Value", "High Value"])}
    raw["segment"] = raw["cluster"].map(label_map)

    return (
        raw.groupby("segment")
        .agg(
            customers    = ("customerid",     "count"),
            avg_orders   = ("order_count",    "mean"),
            avg_spend    = ("total_spend",    "mean"),
            avg_order_val= ("avg_order_value","mean"),
        )
        .round(2)
        .reset_index()
    )


def build_vendor_segments(engine) -> pd.DataFrame:
    try:
        from sklearn.cluster import KMeans
        from sklearn.preprocessing import StandardScaler
    except ImportError:
        print("  WARNING: scikit-learn not installed; skipping VendorSegments.")
        return pd.DataFrame()

    raw = query(engine, """
        SELECT v.name                                            AS vendor_name,
               COUNT(poh.purchaseorderid)                        AS order_count,
               SUM(poh.subtotal+poh.taxamt+poh.freight)          AS total_spend,
               AVG(poh.subtotal+poh.taxamt+poh.freight)          AS avg_order_value
        FROM purchasing.purchaseorderheader    poh
        JOIN purchasing.vendor                 v ON v.businessentityid = poh.vendorid
        GROUP BY v.name
    """, label="VendorSegments (raw)")

    features = raw[["order_count", "total_spend", "avg_order_value"]].fillna(0)
    scaled   = StandardScaler().fit_transform(features)
    km       = KMeans(n_clusters=3, random_state=42, n_init=10)
    raw["cluster"] = km.fit_predict(scaled)

    order = raw.groupby("cluster")["total_spend"].mean().sort_values().index
    label_map = {c: l for c, l in zip(order, ["Low Spend", "Mid Spend", "High Spend"])}
    raw["segment"] = raw["cluster"].map(label_map)

    return (
        raw.groupby("segment")
        .agg(
            vendors      = ("vendor_name",    "count"),
            avg_orders   = ("order_count",    "mean"),
            avg_spend    = ("total_spend",    "mean"),
            avg_order_val= ("avg_order_value","mean"),
        )
        .round(2)
        .reset_index()
    )


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    _check_env()

    # 1. Connect to PostgreSQL
    print("\n=== AdventureWorks → Power BI Push ===\n")
    if not DRY_RUN:
        print(f"DB: {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")
    engine = make_engine()
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("  PostgreSQL connection OK\n")
    except Exception as exc:
        sys.exit(f"  DB connection failed: {exc}")

    # 2. Query all datasets
    print("--- Querying PostgreSQL ---")
    dataframes: dict[str, pd.DataFrame] = {}
    for table_name, sql_str in QUERIES.items():
        dataframes[table_name] = query(engine, sql_str, label=table_name)

    dataframes["CustomerSegments"] = build_customer_segments(engine)
    dataframes["VendorSegments"]   = build_vendor_segments(engine)
    print()

    # 3. Authenticate with Power BI
    if DRY_RUN:
        print("--- DRY RUN: skipping Power BI auth & push ---\n")
        token = "DRY_RUN_TOKEN"
    else:
        print("--- Authenticating with Power BI ---")
        token = get_access_token()
        print("  Token acquired\n")

    # 4. Create or refresh the push dataset
    if not DRY_RUN:
        print("--- Setting up Power BI dataset ---")
        existing_id = find_dataset_id(token, DATASET_NAME)
        if existing_id:
            print(f"  Dataset '{DATASET_NAME}' already exists (id={existing_id})")
            print("  Clearing existing table rows …")
            for tbl in DATASET_SCHEMA["tables"]:
                try:
                    clear_table_rows(token, existing_id, tbl["name"])
                except Exception:
                    pass   # table may not have rows yet
            dataset_id = existing_id
        else:
            dataset_id = create_dataset(token, DATASET_SCHEMA)
        print()
    else:
        dataset_id = "DRY_RUN_ID"

    # 5. Push each table
    print("--- Pushing data to Power BI ---")
    for tbl in DATASET_SCHEMA["tables"]:
        name = tbl["name"]
        df   = dataframes.get(name)
        if df is None or df.empty:
            print(f"  Skipping '{name}' (no data)")
            continue
        print(f"  Table '{name}': {len(df):,} rows")
        push_rows(token, dataset_id, name, df)

    print("\nDone! Open Power BI Service and build visuals from the dataset.\n")
    if not DRY_RUN:
        workspace_hint = (
            f"https://app.powerbi.com/groups/{WORKSPACE_ID}/list"
            if WORKSPACE_ID else
            "https://app.powerbi.com/groups/me/list"
        )
        print(f"  Workspace: {workspace_hint}")


if __name__ == "__main__":
    main()
