"""
msAccessRW.py

A small wrapper around pyodbc + pandas for reading from / writing to
Microsoft Access (.mdb / .accdb) databases.

Classes:
    cDriverType   - Enum of supported Access ODBC driver names.
    cMsAccessAPI  - Thin wrapper around pyodbc: builds the connection
                    string, validates the DB file, and yields managed
                    connections.
    cDBoperations - Higher level read/write helpers (select -> DataFrame,
                    parameterized insert/update, raw SQL execution).

Usage example:
    api = cMsAccessAPI(r"C:/data/mydb.accdb", cDriverType.msAccess_64bit)
    db = cDBoperations(api)
    df = db.selectFromDB("SELECT * FROM MyTable")
"""

import contextlib
import enum
import logging
import os
import typing
import pandas
import pyodbc

logger = logging.getLogger(__name__)


class cDriverType(enum.Enum):
    """Supported Microsoft Access ODBC drivers."""
    msAccess_32bit = "Microsoft Access Driver (*.mdb)"
    msAccess_64bit = "Microsoft Access Driver (*.mdb, *.accdb)"


class cMsAccessAPI:
    """Wrapper around pyodbc that manages the connection string and
    connection lifecycle for a single Microsoft Access database file."""

    def __init__(self, aFilename: str, aDriverType: cDriverType = cDriverType.msAccess_64bit):
        if not isinstance(aDriverType, cDriverType):
            raise TypeError(
                f"aDriverType must be a cDriverType, got {type(aDriverType).__name__}"
            )
        if not aFilename:
            raise ValueError("aFilename must not be empty")
        if not os.path.isfile(aFilename):
            raise FileNotFoundError(f"Database file not found: {aFilename}")

        self.filename = aFilename
        self.driverType = aDriverType
        self.connStr = (
            f"DRIVER={{{self.driverType.value}}};"
            f"DBQ={self.filename};"
        )

    # =====================================
    @contextlib.contextmanager
    def connect(self) -> typing.Iterator[pyodbc.Connection]:
        """Context manager yielding an open pyodbc connection.

        Commits on normal exit, rolls back on exception, and always
        closes the connection.
        """
        conn: typing.Optional[pyodbc.Connection] = None
        try:
            conn = pyodbc.connect(self.connStr)
        except pyodbc.Error as e:
            raise ConnectionError(
                f"Failed to connect to database '{self.filename}': {e}"
            ) from e

        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # =====================================
    def testConnection(self) -> bool:
        """Returns True if a connection can be opened successfully."""
        try:
            with self.connect():
                return True
        except ConnectionError as e:
            logger.error("testConnection failed: %s", e)
            return False


class cDBoperations:
    """Higher level database read/write helpers built on top of cMsAccessAPI."""

    def __init__(self, aApi: cMsAccessAPI):
        self.api = aApi
        self.preferredDecimals: typing.Dict[str, int] = {}
        """Examples: "openVal": 2, "highVal": 2, "lowVal": 2, "closeVal": 2, "volume": 0"""

        self.preferredTypes: typing.Dict[str, type] = {}
        """Examples: "tradeDate": str"""

    # =====================================
    @staticmethod
    def encodeValue(aValue) -> str:
        """Encodes a Python value as a literal for embedding directly in SQL text.

        Prefer parameterized queries (used internally by insertRows/updateRows)
        whenever possible; this helper is for building ad-hoc SQL fragments
        (e.g. dynamic WHERE clauses) where parameter markers aren't practical.
        """
        if pandas.isna(aValue):
            return "NULL"
        elif isinstance(aValue, (pandas.Timestamp,)) or pandas.api.types.is_datetime64_any_dtype(type(aValue)):
            return f"#{aValue.strftime('%Y-%m-%d')}#"
        elif isinstance(aValue, str):
            return "'" + aValue.replace("'", "''") + "'"
        else:
            return str(aValue)

    # =====================================
    @staticmethod
    def _toParam(aValue):
        """Converts a pandas/numpy scalar into a plain Python value/None suitable for pyodbc."""
        if pandas.isna(aValue):
            return None
        if isinstance(aValue, pandas.Timestamp):
            return aValue.to_pydatetime()
        if hasattr(aValue, "item"):
            return aValue.item()
        return aValue

    # =====================================
    def executeSql(self, aSqlStr: str, aParams: typing.Optional[typing.Sequence] = None) -> None:
        """Executes an arbitrary SQL statement (e.g. DDL) with optional parameters."""
        try:
            with self.api.connect() as conn:
                cursor = conn.cursor()
                logger.info("Executing SQL: %s", aSqlStr)
                if aParams is not None:
                    cursor.execute(aSqlStr, aParams)
                else:
                    cursor.execute(aSqlStr)
        except Exception as e:
            logger.error("Error in executeSql: %s", e)
            raise

    # =====================================
    def selectFromDB(self, aSqlStr: str, aParams: typing.Optional[typing.Sequence] = None) -> pandas.DataFrame:
        """Runs a SELECT statement and returns the results as a pandas DataFrame.

        Example with aParams:
            df = db.selectFromDB(
                "SELECT * FROM MyTable WHERE tradeDate >= ? AND symbol = ?",
                [datetime.date(2024, 1, 1), "AAPL"],
            )
        """
        try:
            with self.api.connect() as conn:
                result = pandas.read_sql_query(aSqlStr, conn, params=aParams)

            # Apply type conversion based on preferredTypes
            for col_name, dtype in self.preferredTypes.items():
                if col_name in result.columns:
                    result[col_name] = result[col_name].astype(dtype)

            # Apply rounding based on preferredDecimals
            for col_name, decimals in self.preferredDecimals.items():
                if col_name in result.columns and pandas.api.types.is_numeric_dtype(result[col_name]):
                    result[col_name] = result[col_name].round(decimals)

            return result
        except Exception as e:
            logger.error("Error in selectFromDB: %s", e)
            return pandas.DataFrame()

    # =====================================
    def updateRows(
        self,
        aTablename: str,
        aRows: typing.List[pandas.Series],
        aKeyColumns: typing.List[str],
    ) -> None:
        """Updates rows in aTablename, matching on aKeyColumns, using parameterized SQL."""
        if not aRows:
            return
        sql = None
        try:
            with self.api.connect() as conn:
                cursor = conn.cursor()
                for row in aRows:
                    setCols = [col for col in row.index if col not in aKeyColumns]
                    if not setCols:
                        continue

                    setClause = ", ".join(f"{col} = ?" for col in setCols)
                    whereClause = " AND ".join(f"{col} = ?" for col in aKeyColumns)
                    sql = f"UPDATE {aTablename} SET {setClause} WHERE {whereClause}"
                    params = [self._toParam(row[col]) for col in setCols] + [
                        self._toParam(row[col]) for col in aKeyColumns
                    ]
                    cursor.execute(sql, params)
        except Exception as e:
            logger.error("Error in updateRows: %s. Last sql was: %s", e, sql)
            raise

    # =====================================
    def insertRows(self, aTablename: str, aRows: typing.List[pandas.Series]) -> None:
        """Inserts rows into aTablename using parameterized SQL."""
        if not aRows:
            return
        sql = None
        try:
            with self.api.connect() as conn:
                cursor = conn.cursor()
                columns = aRows[0].index.tolist()
                placeholders = ", ".join(["?"] * len(columns))
                sql = f"INSERT INTO {aTablename} ({', '.join(columns)}) VALUES ({placeholders})"

                for row in aRows:
                    params = [self._toParam(row[col]) for col in columns]
                    cursor.execute(sql, params)
        except Exception as e:
            logger.error("Error in insertRows: %s. Last sql was: %s", e, sql)
            raise
