import time
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from sqlalchemy.exc import OperationalError, DBAPIError


'''
**************************************************************************************************************************************************************************************
Thin wrapper around a SQLAlchemy engine to run read-only SQL and get DataFrames.

Usage:
    db = SQLConnection(
        db_host=_conn.host, db_port=_conn.port or 1433,
        db_database=_conn.schema, db_username=_conn.login,
        db_password=_conn.password, dialect="mssql", driver="pymssql",
    )
    with db:
        df = db.fech_dataframe("SELECT ...")

Author: Marc Eslava
**************************************************************************************************************************************************************************************
'''


class SQLConnection:
    # Stores the connection settings. Nothing opens here (no side effects in
    # __init__); the engine and connection open in the `with` block.
    # login_timeout: seconds to wait when opening the connection (so a dead
    # server fails fast instead of hanging). timeout: per-query timeout in
    # seconds (leave None for heavy queries that legitimately take minutes).
    # connect_retries/retry_delay: on a connection-level failure (server
    # briefly unreachable -- e.g. login requests queued behind a blocking
    # chain on the server, error 20009/'Adaptive Server is unavailable'),
    # retry instead of failing the whole run outright. Default 0 = no retry
    # (keeps /check's "is it up right now" semantics instant).
    def __init__(self, db_host, db_database, db_username, db_password,
                 db_port=1433, dialect="mssql", driver="pymssql",
                 login_timeout=None, timeout=None,
                 connect_retries=0, retry_delay=10):
        self.db_host = db_host
        self.db_port = int(db_port) if db_port else 1433
        self.db_database = db_database
        self.db_username = db_username
        self.db_password = db_password
        self.dialect = dialect
        self.driver = driver
        self.login_timeout = login_timeout
        self.timeout = timeout
        self.connect_retries = connect_retries
        self.retry_delay = retry_delay
        self._engine = None
        self._conn = None

    # Builds a SQLAlchemy connection object from an Airflow-style connection
    # (attributes: host, port, schema, login, password).
    @classmethod
    def from_connection(cls, conn, dialect="mssql", driver="pymssql"):
        return cls(
            db_host=conn.host, db_port=conn.port or 1433,
            db_database=conn.schema, db_username=conn.login,
            db_password=conn.password, dialect=dialect, driver=driver,
        )

    # Composes the SQLAlchemy URL (credentials are safely escaped by URL.create).
    def _url(self):
        return URL.create(
            drivername=f"{self.dialect}+{self.driver}",
            username=self.db_username,
            password=self.db_password,
            host=self.db_host,
            port=self.db_port,
            database=self.db_database,
        )

    # Opens the engine and a live connection when entering a `with` block.
    # Retries connect_retries times (with retry_delay pause) on a connection-
    # level failure only -- a server that's genuinely down still fails after
    # the last attempt, just not on the first blip.
    def __enter__(self):
        connect_args = {}
        if self.login_timeout is not None:
            connect_args["login_timeout"] = self.login_timeout
        if self.timeout is not None:
            connect_args["timeout"] = self.timeout
        self._engine = create_engine(self._url(), pool_pre_ping=True, connect_args=connect_args)

        attempts = self.connect_retries + 1
        for attempt in range(1, attempts + 1):
            try:
                self._conn = self._engine.connect()
                return self
            except (OperationalError, DBAPIError):
                if attempt >= attempts:
                    raise
                print(f"  Connexió SQL fallida (intent {attempt}/{attempts}); "
                      f"reintentant en {self.retry_delay}s...", flush=True)
                time.sleep(self.retry_delay)

    # Closes the connection and disposes the engine when leaving the block.
    def __exit__(self, exc_type, exc, tb):
        if self._conn is not None:
            self._conn.close()
            self._conn = None
        if self._engine is not None:
            self._engine.dispose()
            self._engine = None
        return False

    # Runs a SQL query and returns the result as a DataFrame.
    # exec_driver_sql sends the string straight to the driver, so pre-built
    # SQL is not re-parsed for bind params. `params` (optional dict) uses the
    # driver's own placeholders -- %(name)s for pymssql -- so user-provided
    # values are bound safely instead of interpolated into the string.
    def fech_dataframe(self, query, params=None):
        if self._conn is None:
            raise RuntimeError(
                "SQLConnection is not open -- use it in a `with SQLConnection(...) as db:` block")
        if params is None:
            result = self._conn.exec_driver_sql(query)
        else:
            result = self._conn.exec_driver_sql(query, params)
        return pd.DataFrame(result.fetchall(), columns=list(result.keys()))

    # Alias for the common (correct) spelling.
    fetch_dataframe = fech_dataframe
