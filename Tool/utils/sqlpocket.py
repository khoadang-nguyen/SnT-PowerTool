from pathlib import Path
import json
import pandas as pd

from sqlalchemy import create_engine, text, bindparam

ENGINE = "mssql+pyodbc://@VNSGNCCCVR01P/COMEX?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"


class DBClient:
    def __init__(self, query_folder = None):
        self.engine = create_engine(ENGINE, fast_executemany=True)
        self.query_folder = Path(query_folder) if query_folder else None

    # Load query from .sql file
    def get_query(self, name):
        path = self.query_folder / f"{name}.sql"
        if not path.exists():
            raise FileNotFoundError(f"Query not found: {path}")
        return path.read_text(encoding="utf-8")

    # ------------------------
    # Read SQL from file with params
    # ------------------------
    def sql_read(self, query_name, params=None):
        sql = self.get_query(query_name)
        return self.sql_read_query(sql, params=params)

    # ------------------------
    # Read SQL from string with params
    # ------------------------
    def sql_read_query(self, query_string, params=None):
        stmt = text(query_string)
        if params:
            for k, v in params.items():
                if isinstance(v, (list, tuple)):
                    stmt = stmt.bindparams(bindparam(k, expanding=True))
        return pd.read_sql(stmt, self.engine, params=params)

    # ------------------------
    # Execute SQL from file with params
    # ------------------------
    def sql_execute(self, query_name, params=None):
        sql = self.get_query(query_name)
        self.sql_execute_query(sql, params=params)

    # ------------------------
    # Execute SQL from string with params
    # ------------------------
    def sql_execute_query(self, query_string, params=None):
        stmt = text(query_string)
        if params:
            for k, v in params.items():
                if isinstance(v, (list, tuple)):
                    stmt = stmt.bindparams(bindparam(k, expanding=True))
        with self.engine.begin() as conn:
            conn.execute(stmt, params or {})

    # ------------------------
    # Push dataframe
    # ------------------------
    def sql_push(self, df, table_name, schema=None, if_exists="append", chunksize=5000):
        df.to_sql(
            table_name,
            self.engine,
            schema=schema,
            if_exists=if_exists,
            index=False,
            chunksize=chunksize
        )

    # Replace today's data
    def push_replace_today(self, df, table_name, date_col="Updated"):
        delete_sql = f"""
        DELETE FROM {table_name}
        WHERE {date_col} = CAST(GETDATE() AS DATE)
        """
        with self.engine.begin() as conn:
            conn.execute(text(delete_sql))
        self.sql_push(df, table_name)

    # Close connection
    def close(self):
        self.engine.dispose()