"""数据库辅助类 - 封装 SQLite 数据库操作"""

import sqlite3
import threading
from typing import Any, Optional


class DatabaseHelper:
    """SQLite 数据库操作辅助类（线程安全）"""

    def __init__(self, connection_string: str):
        self._connection_string = connection_string
        self._connection: Optional[sqlite3.Connection] = None
        self._lock = threading.Lock()

    @property
    def connection(self) -> sqlite3.Connection:
        if self._connection is None:
            # check_same_thread=False 允许跨线程使用，配合 _lock 保证安全
            self._connection = sqlite3.connect(self._connection_string, check_same_thread=False)
            self._connection.execute("PRAGMA foreign_keys = ON")
        return self._connection

    def ensure_connection(self):
        _ = self.connection

    def execute_non_query(self, sql: str, params: tuple = ()) -> int:
        conn = self.connection
        with self._lock:
            # 多条语句（如 "DROP VIEW ...; CREATE VIEW ..."）需要用 executescript
            if not params and ";" in sql.strip().rstrip(";"):
                conn.executescript(sql)
                conn.commit()
                return 0
            cursor = conn.cursor()
            cursor.execute(sql, params)
            conn.commit()
            return cursor.rowcount

    def execute_scalar(self, sql: str, params: tuple = ()) -> Any:
        conn = self.connection
        with self._lock:
            cursor = conn.cursor()
            cursor.execute(sql, params)
            row = cursor.fetchone()
            if row is None:
                return None
            return row[0]

    def fill_data_table(self, sql: str, params: tuple = ()) -> tuple:
        """返回 (columns, rows) 元组，columns 为列名列表，rows 为行数据列表"""
        conn = self.connection
        with self._lock:
            cursor = conn.cursor()
            cursor.execute(sql, params)
            columns = [desc[0] for desc in cursor.description] if cursor.description else []
            rows = cursor.fetchall()
            return columns, rows

    def begin_transaction(self):
        return self.connection

    def create_table(self, sql: str):
        self.execute_non_query(sql)

    def drop_table_if_exists(self, table_name: str):
        self.execute_non_query(f"DROP TABLE IF EXISTS {table_name}")

    def drop_view_if_exists(self, view_name: str):
        self.execute_non_query(f"DROP VIEW IF EXISTS {view_name}")

    def close(self):
        with self._lock:
            if self._connection is not None:
                self._connection.close()
                self._connection = None
