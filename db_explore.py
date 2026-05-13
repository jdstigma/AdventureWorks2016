"""
db_explore.py
=============
Connects to the AdventureWorks SQLite database and generates a full
data profile: all tables, row counts, column types, null rates,
distinct counts, and top 5 most common values per column.

REQUIREMENTS
------------
Install these once in your terminal:
    pip install sqlalchemy pandas

HOW TO USE
----------
1. Make sure adventureworks.db is in the SAME folder as this script.
2. Run:  python db_explore.py
3. adventureworks_profile.csv is saved next to this script when done.
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------
import os
import sqlite3
import pandas as pd
from sqlalchemy import create_engine, text

# Always resolve paths relative to this script's location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# CONFIG -- only change DB_FILENAME if your .db file has a different name
# ---------------------------------------------------------------------------
DB_FILENAME = "adventureworks.db"
DB_PATH     = os.path.join(SCRIPT_DIR, DB_FILENAME)

# How many top-value examples to show per column
TOP_N = 5

# ---------------------------------------------------------------------------
# CONNECTION
# ---------------------------------------------------------------------------
CONNECTION_STRING = f"sqlite:///{DB_PATH}"

print("Connecting to database...")
engine = create_engine(CONNECTION_STRING)

try:
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print(f"  Connected to: {DB_PATH}\n")
except Exception as e:
    print(f"  Connection failed: {e}")
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# HELPER: run a SQL query and return a DataFrame
# ---------------------------------------------------------------------------
def query(sql, params=None):
    """Run a SQL string and return the result as a pandas DataFrame."""
    with engine.connect() as conn:
        return pd.read_sql(text(sql), conn, params=params)


# ---------------------------------------------------------------------------
# STEP 1: DISCOVER TABLES
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 1 - Tables")
print("=" * 60)

tables_sql = """
    SELECT name AS table_name, type
    FROM sqlite_master
    WHERE type IN ('table', 'view')
      AND name NOT LIKE 'sqlite_%'
    ORDER BY type, name
"""

tables_df = query(tables_sql)

if tables_df.empty:
    print("No tables found. Is the DB_PATH correct?")
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
    tbl      = row["table_name"]
    tbl_type = row["type"]

    try:
        count_df = query(f'SELECT COUNT(*) AS cnt FROM "{tbl}"')
        cnt = int(count_df["cnt"].iloc[0])
    except Exception:
        cnt = -1  # -1 means we could not count (e.g. permission issue)

    row_counts.append({"table": tbl, "type": tbl_type, "row_count": cnt})

counts_df = pd.DataFrame(row_counts).sort_values("row_count", ascending=False)

print(counts_df.to_string(index=False))
print(f"\nTotal rows across all tables: {counts_df['row_count'].clip(lower=0).sum():,}\n")


# ---------------------------------------------------------------------------
# STEP 3: COLLECT COLUMN METADATA
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 3 - Column Profile (types, nulls, distinct counts, top values)")
print("=" * 60)

# SQLite stores column info in PRAGMA table_info -- gathered per table
sqlite_cur = sqlite3.connect(DB_PATH).cursor()

col_rows = []
for tbl in counts_df["table"].tolist():
    try:
        sqlite_cur.execute(f'PRAGMA table_info("{tbl}")')
        for col in sqlite_cur.fetchall():
            # PRAGMA returns: cid, name, type, notnull, dflt_value, pk
            col_rows.append({
                "table_name":  tbl,
                "column_name": col[1],
                "data_type":   col[2] if col[2] else "TEXT",
                "is_nullable": "NO" if col[3] else "YES",
            })
    except Exception:
        pass

sqlite_cur.close()
cols_df = pd.DataFrame(col_rows)
print(f"Found {len(cols_df)} columns across all tables.\n")


# ---------------------------------------------------------------------------
# STEP 4: PER-TABLE DEEP PROFILE
# ---------------------------------------------------------------------------
all_profiles = []

for _, tbl_row in counts_df.iterrows():
    tbl       = tbl_row["table"]
    row_count = tbl_row["row_count"]

    if row_count < 0:
        continue

    print(f"  Profiling  {tbl}  ({row_count:,} rows)...")

    tbl_cols = cols_df[cols_df["table_name"] == tbl]

    for _, col_row in tbl_cols.iterrows():
        col_name  = col_row["column_name"]
        data_type = col_row["data_type"].upper()

        profile = {
            "table":     tbl,
            "column":    col_name,
            "data_type": data_type,
            "row_count": row_count,
        }

        try:
            # Null count and rate
            null_df = query(
                f'SELECT COUNT(*) AS null_count '
                f'FROM "{tbl}" '
                f'WHERE "{col_name}" IS NULL'
            )
            null_count = int(null_df["null_count"].iloc[0])
            profile["null_count"] = null_count
            profile["null_pct"]   = round(null_count / row_count * 100, 1) if row_count > 0 else 0

            # Flag high null rates in plain text
            if profile["null_pct"] > 20:
                profile["null_flag"] = "HIGH"
            elif profile["null_pct"] > 5:
                profile["null_flag"] = "WARN"
            else:
                profile["null_flag"] = ""

            # Distinct count
            dist_df = query(
                f'SELECT COUNT(DISTINCT "{col_name}") AS dist_count '
                f'FROM "{tbl}"'
            )
            profile["distinct_count"] = int(dist_df["dist_count"].iloc[0])

            # Numeric stats for number-like columns
            numeric_keywords = ("INT", "REAL", "FLOAT", "NUMERIC", "DECIMAL", "DOUBLE", "MONEY")
            if any(kw in data_type for kw in numeric_keywords):
                stats_df = query(
                    f'SELECT '
                    f'  MIN(CAST("{col_name}" AS REAL)) AS min_val, '
                    f'  MAX(CAST("{col_name}" AS REAL)) AS max_val, '
                    f'  AVG(CAST("{col_name}" AS REAL)) AS avg_val '
                    f'FROM "{tbl}" '
                    f'WHERE "{col_name}" IS NOT NULL'
                )
                profile["min"] = round(float(stats_df["min_val"].iloc[0] or 0), 2)
                profile["max"] = round(float(stats_df["max_val"].iloc[0] or 0), 2)
                profile["avg"] = round(float(stats_df["avg_val"].iloc[0] or 0), 2)

            # Top N most common values
            top_df = query(
                f'SELECT "{col_name}" AS val, COUNT(*) AS freq '
                f'FROM "{tbl}" '
                f'WHERE "{col_name}" IS NOT NULL '
                f'GROUP BY "{col_name}" '
                f'ORDER BY freq DESC '
                f'LIMIT {TOP_N}'
            )
            top_values = [
                f"{str(r['val'])[:30]} ({r['freq']:,})"
                for _, r in top_df.iterrows()
            ]
            profile["top_values"] = " | ".join(top_values)

        except Exception as e:
            profile["error"] = str(e)[:80]

        all_profiles.append(profile)


# ---------------------------------------------------------------------------
# STEP 5: SAVE RESULTS
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("STEP 5 - Saving Profile")
print("=" * 60)

profile_df = pd.DataFrame(all_profiles)

output_file = os.path.join(SCRIPT_DIR, "adventureworks_profile.csv")
profile_df.to_csv(output_file, index=False)
print(f"\n  Saved to: {output_file}")
print("  Open this in Excel or Power BI for easy filtering.\n")