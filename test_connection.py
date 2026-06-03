import os
import pyodbc
from dotenv import load_dotenv


# ============================================================
# POS Book Report - SQL Connection Test
# Purpose:
#   Test kung nakakakonek ang Python sa POS SQL Server database.
#
# Important:
#   Hindi nito ipi-print ang actual password.
#   Password length lang ang ipapakita for checking.
# ============================================================


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")

load_dotenv(ENV_PATH)


def get_required_env(name: str) -> str:
    """
    Reads a required value from .env.
    Kapag blank/missing, mag-e-error agad para madaling malaman ang kulang.
    """
    value = os.getenv(name)

    if value is None or value.strip() == "":
        raise ValueError(f"Missing required .env value: {name}")

    return value.strip()


def main():
    server = get_required_env("SQL_SERVER")
    database = get_required_env("SQL_DATABASE")
    username = get_required_env("SQL_USERNAME")
    password = get_required_env("SQL_PASSWORD")

    print("========== POS BOOK REPORT CONNECTION TEST ==========")
    print("SQL_SERVER:", server)
    print("SQL_DATABASE:", database)
    print("SQL_USERNAME:", username)
    print("SQL_PASSWORD length:", len(password))
    print("=====================================================")

    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )

    try:
        conn = pyodbc.connect(conn_str, timeout=15)
        cursor = conn.cursor()

        cursor.execute("SELECT 1 AS TestResult;")
        row = cursor.fetchone()

        print("Connection successful.")
        print("Test query result:", row.TestResult)

        cursor.close()
        conn.close()

    except Exception as e:
        print("Connection failed.")
        print("Error:")
        print(e)
        raise


if __name__ == "__main__":
    main()