# Template for the local secrets.
# Copy this file to `db_config.py` (which is git-ignored) and fill in the values.
#     cp db_config.example.py db_config.py

# --- BIFarma SQL (MSSQL) ---
DB_HOST = "your-sql-server"        # e.g. "SRV-SQL01" or an IP
DB_PORT = 1433                     # default MSSQL port
DB_NAME = "BifarmaCentral"         # database / schema
DB_USER = "your_user"
DB_PASS = "your_password"

# --- ClickHouse (serving layer, `tancament` DB) ---
CLICKHOUSE_PASSWORD = "your_clickhouse_password"

# --- Metabase admin (for the panel's embedded analytics) ---
MB_USER = "admin@example.com"
MB_PASS = "your_metabase_password"
