"""
db_explore.py
=============
Connects to AdventureWorks in PostgreSQL and generates a full data profile:
all schemas, tables, row counts, column types, null rates, distinct counts,
and top 5 most common values per column.

REQUIREMENTS
------------
    pip install -r requirements.txt

HOW TO USE
----------
1. Update CONFIG if your connection details differ.
2. Run:  python db_explore.py
3. adventureworks_profile.csv is saved next to this script.

NOTE: Do not commit real passwords to a public GitHub repo.
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

# Leave empty to scan ALL schemas, or list specific ones:
# e.g. ["sales", "production", "humanresources", "purchasing", "person"]
SCHEMAS_TO_SCAN = []

TOP_N = 5  # number of top values to show per column

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


def query(sql, params=None):
    """Run SQL and return a DataFrame."""
    with engine.connect() as conn:
        return pd.read_sql(text(sql), conn, params=params)


# ---------------------------------------------------------------------------
# STEP 1: DISCOVER TABLES
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 1 - Schemas and Tables")
print("=" * 60)

schema_filter = ""
if SCHEMAS_TO_SCAN:
    placeholders = ", ".join(f"'{s}'" for s in SCHEMAS_TO_SCAN)
    schema_filter = f"AND table_schema IN ({placeholders})"

tables_df = query(f"""
    SELECT
        table_schema    AS schema,
        table_name,
        table_type      AS type
    FROM information_schema.tables
    WHERE table_schema NOT IN ('pg_catalog', 'information_schema')
      {schema_filter}
    ORDER BY table_schema, table_name
""")

if tables_df.empty:
    print("No tables found. Check SCHEMAS_TO_SCAN.")
    raise SystemExit(0)

print(f"Found {len(tables_df)} objects:\n")
print(tables_df.to_string(index=False))
print()


# ---------------------------------------------------------------------------
# STEP 2: ROW COUNTS
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 2 - Row Counts")
print("=" * 60)

row_counts = []
for _, row in tables_df.iterrows():
    schema   = row["schema"]
    tbl      = row["table_name"]
    tbl_type = row["type"]
    try:
        cnt = int(query(f'SELECT COUNT(*) AS cnt FROM "{schema}"."{tbl}"')["cnt"].iloc[0])
    except Exception:
        cnt = -1
    row_counts.append({"schema": schema, "table": tbl, "type": tbl_type, "row_count": cnt})

counts_df = pd.DataFrame(row_counts).sort_values(["schema", "row_count"], ascending=[True, False])
print(counts_df.to_string(index=False))
print(f"\nTotal rows: {counts_df['row_count'].clip(lower=0).sum():,}\n")


# ---------------------------------------------------------------------------
# STEP 3: COLUMN METADATA
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 3 - Column Profile")
print("=" * 60)

cols_df = query(f"""
    SELECT
        table_schema    AS schema,
        table_name,
        column_name,
        data_type,
        is_nullable,
        numeric_precision,
        numeric_scale
    FROM information_schema.columns
    WHERE table_schema NOT IN ('pg_catalog', 'information_schema')
      {schema_filter}
    ORDER BY table_schema, table_name, ordinal_position
""")

print(f"Found {len(cols_df)} columns across all tables.\n")


# ---------------------------------------------------------------------------
# STEP 4: DEEP PROFILE
# ---------------------------------------------------------------------------
all_profiles = []

for _, tbl_row in counts_df.iterrows():
    schema    = tbl_row["schema"]
    tbl       = tbl_row["table"]
    row_count = tbl_row["row_count"]

    if row_count <= 0:
        continue

    print(f"  Profiling  {schema}.{tbl}  ({row_count:,} rows)...")

    tbl_cols = cols_df[(cols_df["schema"] == schema) & (cols_df["table_name"] == tbl)]

    for _, col_row in tbl_cols.iterrows():
        col_name  = col_row["column_name"]
        data_type = col_row["data_type"]

        profile = {
            "schema":    schema,
            "table":     tbl,
            "column":    col_name,
            "data_type": data_type,
            "row_count": row_count,
        }

        try:
            # Null count and rate
            null_count = int(query(
                f'SELECT COUNT(*) AS n FROM "{schema}"."{tbl}" WHERE "{col_name}" IS NULL'
            )["n"].iloc[0])
            profile["null_count"] = null_count
            profile["null_pct"]   = round(null_count / row_count * 100, 1)
            profile["null_flag"]  = "HIGH" if profile["null_pct"] > 20 else ("WARN" if profile["null_pct"] > 5 else "")

            # Distinct count
            profile["distinct_count"] = int(query(
                f'SELECT COUNT(DISTINCT "{col_name}") AS n FROM "{schema}"."{tbl}"'
            )["n"].iloc[0])

            # Numeric stats
            numeric_types = (
                "integer", "bigint", "smallint", "numeric", "decimal",
                "real", "double precision", "money"
            )
            if data_type in numeric_types:
                stats = query(f"""
                    SELECT
                        MIN("{col_name}")                                                  AS min_val,
                        MAX("{col_name}")                                                  AS max_val,
                        AVG("{col_name}")                                                  AS avg_val,
                        PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY "{col_name}")         AS median_val
                    FROM "{schema}"."{tbl}"
                    WHERE "{col_name}" IS NOT NULL
                """)
                profile["min"]    = round(float(stats["min_val"].iloc[0] or 0), 2)
                profile["max"]    = round(float(stats["max_val"].iloc[0] or 0), 2)
                profile["avg"]    = round(float(stats["avg_val"].iloc[0] or 0), 2)
                profile["median"] = round(float(stats["median_val"].iloc[0] or 0), 2)

            # Top N values
            top_df = query(f"""
                SELECT "{col_name}" AS val, COUNT(*) AS freq
                FROM "{schema}"."{tbl}"
                WHERE "{col_name}" IS NOT NULL
                GROUP BY "{col_name}"
                ORDER BY freq DESC
                LIMIT {TOP_N}
            """)
            profile["top_values"] = " | ".join(
                f"{str(r['val'])[:30]} ({r['freq']:,})" for _, r in top_df.iterrows()
            )

        except Exception as e:
            profile["error"] = str(e)[:80]

        all_profiles.append(profile)


# ---------------------------------------------------------------------------
# STEP 5: SAVE
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("STEP 5 - Saving Profile")
print("=" * 60)

profile_df = pd.DataFrame(all_profiles)
output_file = os.path.join(SCRIPT_DIR, "adventureworks_profile.csv")
profile_df.to_csv(output_file, index=False)
print(f"\n  Saved: {output_file}")
print("  Open in Excel or Power BI for easy filtering.\n")
