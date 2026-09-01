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

import collections.abc
import contextlib
import enum
import logging
import pathlib
import pandas
import pyodbc

logger = logging.getLogger(__name__)


# =====================================
def getAvailableDBdrivers() -> list[str]:
    """Returns the names of all ODBC drivers installed on this machine."""
    return pyodbc.drivers()


class cDriverType(enum.Enum):
    """Supported Microsoft Access ODBC drivers."""
    msAccess_32bit = "Microsoft Access Driver (*.mdb)"
    msAccess_64bit = "Microsoft Access Driver (*.mdb, *.accdb)"


class cMsAccessAPI:
    """Wrapper around pyodbc that manages the connection string and
    connection lifecycle for a single Microsoft Access database file."""

    # in ms access, dates should be wrapped with #, for example, #2017-01-31#
    # Always use the YYYY=MM-DD format
    config_dateQualifier : str = "#"  

    def __init__(self, aFullPathFilename: str, aDriverType: cDriverType = cDriverType.msAccess_64bit):
        # validate inputs before building the connection string, so mistakes
        # surface immediately instead of at the first query
        if not isinstance(aDriverType, cDriverType):
            raise TypeError(f"aDriverType must be a cDriverType, got {type(aDriverType).__name__}")
        if not aFullPathFilename:
            raise ValueError("aFilename must not be empty")

        dbPath = pathlib.Path(aFullPathFilename)
        if not dbPath.is_file():
            raise FileNotFoundError(f"Database file not found: {aFullPathFilename}")

        self.filename = dbPath
        self.driverType = aDriverType
        self.connStr = (
            f"DRIVER={{{self.driverType.value}}};"
            f"DBQ={self.filename};"
            )


    # =====================================
    @contextlib.contextmanager
    # Generator[YieldType, SendType, ReturnType]: this function yields pyodbc.Connection
    # values, never receives values sent into it, and never returns a value itself.
    def connect(self) -> collections.abc.Generator[pyodbc.Connection, None, None]:
        """Yields an open pyodbc connection.

        Commits the transaction on normal exit, rolls it back if an
        exception occurs, and always closes the connection afterwards.

        This must be a generator (use "yield", not "return") because
        @contextlib.contextmanager turns a paused generator into a "with"
        block: code before "yield" becomes __enter__ (open the connection),
        the yielded value becomes the "as conn" variable, and code after
        "yield" becomes __exit__ (commit/rollback/close). A plain function
        would run to completion and return before the "with" block even
        started, leaving no way to resume it afterwards to clean up.
        """
        try:
            conn = pyodbc.connect(self.connStr)
        except pyodbc.Error as e:
            raise ConnectionError(f"Failed to connect to database '{self.filename}': {e}") from e

        try:
            # yield hands the open connection to the caller's "with" block
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()



    # =====================================
    def testConnection(self) -> bool:
        """Opens and immediately closes a connection to verify connectivity."""
        try:
            with self.connect():
                return True
        except ConnectionError as e:
            logger.error("testConnection failed: %s", e)
            return False

# ==================================================================

class cDBoperations:
    """Higher level database read/write helpers built on top of cMsAccessAPI."""

    def __init__(self, aApi: cMsAccessAPI):
        self.api = aApi

        # decimal places to round numeric columns to after a select, keyed by column name
        # example: "openVal": 2, "highVal": 2, "lowVal": 2, "closeVal": 2, "volume": 0
        self.preferredDecimals: dict[str, int] = {}

        # python types to cast columns to after a select, keyed by column name
        # example: "tradeDate": str
        self.preferredTypes: dict[str, type] = {}

    # =====================================
    @staticmethod
    def encodeValue(aValue) -> str:
        """Encodes a Python value as a literal for embedding directly in SQL text.

        Prefer parameterized queries (used internally by insertRows/updateRows)
        whenever possible; this helper is for building ad-hoc SQL fragments
        (e.g. dynamic WHERE clauses) where parameter markers aren't practical.
        """
        # pandas.isna() catches None, NaN and NaT in a single check
        if pandas.isna(aValue):
            return "NULL"
        elif isinstance(aValue, pandas.Timestamp) or pandas.api.types.is_datetime64_any_dtype(type(aValue)):
            return f"#{aValue.strftime('%Y-%m-%d')}#"
        elif isinstance(aValue, str):
            return "'" + aValue.replace("'", "''") + "'"
        else:
            return str(aValue)

    # =====================================
    @staticmethod
    def _toParam(aValue):
        """Converts a pandas/numpy scalar into a plain Python value suitable for pyodbc."""
        if pandas.isna(aValue):
            return None
        if isinstance(aValue, pandas.Timestamp):
            return aValue.to_pydatetime()
        # numpy scalar types (int64, float64, ...) expose .item() to unwrap to a plain Python value
        if hasattr(aValue, "item"):
            return aValue.item()
        return aValue

    # =====================================
    def executeSql(self, aSqlStr: str, aParams: list | tuple = ()) -> None:
        """Executes an arbitrary SQL statement (e.g. DDL), with optional bind parameters.

            Example with aParams:
                df = db.selectFromDB(
                    "SELECT * FROM MyTable WHERE tradeDate >= ? AND symbol = ?",
                    [datetime.date(2024, 1, 1), "AAPL"],
                    )

            In pyodbc's parameterized execution, pyodbc binds each ? placeholder to the corresponding value directly 
            via the ODBC driver (as a typed parameter), rather than substituting text into the SQL string. 
            So NO quoting/escaping of strings happens in Python code; the driver handles it, and this also protects against SQL injection.
        
            
        """
        try:
            with self.api.connect() as conn:
                cursor = conn.cursor()
                logger.info("Executing SQL: %s", aSqlStr)
                if aParams:
                    cursor.execute(aSqlStr, aParams)
                else:
                    cursor.execute(aSqlStr)
        except Exception as e:
            logger.error("Error in executeSql: %s", e)
            raise

    # =====================================
    def selectFromDB(self, aSqlStr: str, aParams: list | tuple = ()) -> pandas.DataFrame:
        """Runs a SELECT statement and returns the results as a pandas DataFrame.

        Example with aParams:
            df = db.selectFromDB(
                "SELECT * FROM MyTable WHERE tradeDate >= ? AND symbol = ?",
                [datetime.date(2024, 1, 1), "AAPL"],
                )
        """
        try:
            with self.api.connect() as conn:
                result = pandas.read_sql_query(aSqlStr, conn, params=aParams or None)

            # cast columns to their preferred Python type
            for columnName, dtype in self.preferredTypes.items():
                if columnName in result.columns:
                    result[columnName] = result[columnName].astype(dtype)

            # round numeric columns to their preferred number of decimals
            for columnName, decimals in self.preferredDecimals.items():
                if columnName in result.columns and pandas.api.types.is_numeric_dtype(result[columnName]):
                    result[columnName] = result[columnName].round(decimals)

            return result
        except Exception as e:
            logger.error("Error in selectFromDB: %s", e)
            return pandas.DataFrame()

    # =====================================
    def updateRows(
        self,
        aTablename: str,
        aDF: pandas.DataFrame,
        aKeyColumns: list[str],
        ) -> None:
        """Updates rows in aTablename, matching on aKeyColumns, using parameterized SQL."""
        if aDF.empty:
            return

        missingKeyColumns = [col for col in aKeyColumns if col not in aDF.columns]
        if missingKeyColumns:
            raise ValueError(f"aKeyColumns not found in aDF: {missingKeyColumns}. The avalailable columns are:{aDF.columns}")

        # columns to update are every column that isn't part of the match key
        setCols = [col for col in aDF.columns if col not in aKeyColumns]
        if not setCols:
            return

        setClause = ", ".join(f"{col} = ?" for col in setCols)
        whereClause = " AND ".join(f"{col} = ?" for col in aKeyColumns)
        sqlTemplate = f"UPDATE {aTablename} SET {setClause} WHERE {whereClause}"

        for _, row in aDF.iterrows():
            params = [self._toParam(row[col]) for col in setCols] + [
                self._toParam(row[col]) for col in aKeyColumns
                ]
            self.executeSql(sqlTemplate, params)

    # =====================================
    def insertRows(self, aTablename: str, aDF: pandas.DataFrame) -> None:
        """Inserts rows into aTablename using parameterized SQL."""
        if aDF.empty:
            return

        columns = aDF.columns.tolist()
        placeholders = ", ".join(["?"] * len(columns))
        sqlTemplate = f"INSERT INTO {aTablename} ({', '.join(columns)}) VALUES ({placeholders})"

        for _, row in aDF.iterrows():
            params = [self._toParam(row[col]) for col in columns]
            self.executeSql(sqlTemplate, params)
