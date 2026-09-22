"""
Azure SQL Database backend. Used by Z and O when DOCUMATE_DB_BACKEND=azuresql
(the default).

The shared reading and writing logic is in sources/database.py. This file
only covers what is specific to Azure SQL:

    finding an ODBC driver   find_driver()
    the connection string    connection_string()
    the query timeout        AzureSqlSource.remove_query_timeout()

Requirements
------------
    pip install pyodbc
    Microsoft ODBC Driver 18 (or 17) for SQL Server, installed separately:
    https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server

Your client IP must also be allowed in the Azure portal, under the SQL
server's Security -> Networking settings.

Run  python main.py --check  to see which driver was found.
"""

from sources import queries_tsql
from sources.database import DatabaseSource


# ODBC drivers that can connect to Azure SQL, in order of preference.
# Older drivers do not support the TLS versions Azure requires.
SUPPORTED_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
]

DRIVER_DOWNLOAD = "https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server"

# Seconds to wait when logging in. A serverless Azure database may be paused
# and take a while to wake up on the first connection.
#
# This is passed to pyodbc.connect(timeout=...), not put in the connection
# string. "Connection Timeout" in the string is not a valid ODBC keyword and
# is ignored.
LOGIN_TIMEOUT = 30


def find_driver(pyodbc):
    """
    Return the name of the best installed ODBC driver from SUPPORTED_DRIVERS.

    pyodbc is only the Python wrapper; the Microsoft driver is a separate
    install. If none is found, raises RuntimeError with the download link
    and the list of drivers that are installed.
    """
    installed = pyodbc.drivers()

    for driver in SUPPORTED_DRIVERS:
        if driver in installed:
            return driver

    raise RuntimeError(
        "No Microsoft ODBC driver for SQL Server found. Install "
        + SUPPORTED_DRIVERS[0] + " from:\n  " + DRIVER_DOWNLOAD
        + "\nDrivers currently installed: "
        + (", ".join(installed) if installed else "none")
    )


def odbc_value(value):
    """
    Wrap a connection-string value in braces so it is read literally.

        odbc_value("p;a=ss")   ->  "{p;a=ss}"
        odbc_value("ab}c")     ->  "{ab}}c}"

    Without the braces, a ; or = inside a password would be read as the
    start of the next setting. A } inside the value is doubled, as the ODBC
    format requires.
    """
    return "{" + str(value).replace("}", "}}") + "}"


def connection_string(settings, driver):
    """
    Build the ODBC connection string from the .env settings.

    settings   dict from config.database_settings(): host, port, database,
               user, password
    driver     driver name from find_driver()

    Returns a string like:

        DRIVER={ODBC Driver 18 for SQL Server};SERVER={tcp:host,1433};
        DATABASE={...};UID={...};PWD={...};Encrypt=yes;TrustServerCertificate=no;

    Encryption is always on, and the server certificate is always checked.
    """
    return ";".join([
        "DRIVER=" + odbc_value(driver),
        "SERVER=" + odbc_value("tcp:%s,%s" % (settings["host"], settings["port"])),
        "DATABASE=" + odbc_value(settings["database"]),
        "UID=" + odbc_value(settings["user"]),
        "PWD=" + odbc_value(settings["password"]),
        "Encrypt=yes",
        "TrustServerCertificate=no",
    ]) + ";"


class AzureSqlSource(DatabaseSource):
    """Reads pending applicants from Azure SQL and writes statuses back."""

    label = "Azure SQL"
    queries = queries_tsql

    def connect(self, settings):
        # Imported here, not at the top, so pyodbc is only needed when the
        # Azure SQL backend is actually used.
        import pyodbc

        return pyodbc.connect(
            connection_string(settings, find_driver(pyodbc)),
            timeout=LOGIN_TIMEOUT,
        )

    def remove_query_timeout(self):
        """Set the pyodbc query timeout to 0, which means no limit."""
        self.connection.timeout = 0
