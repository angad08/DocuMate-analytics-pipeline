import os
import psycopg2
import pandas as pd
import datetime as dt


# =========================================================
# CONFIG
# =========================================================

#convert the path in a way that where the xcel file is the path is auto detected and the code can be run in any system without changing the path



EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "test_data", "Test_Insert_data.xlsx")
SHEET_NAME = "Sheet1"

DB_CONFIG = {
    "host": "YOUR_DB_HOST",
    "database": "YOUR_DB_NAME",
    "user": "YOUR_DB_USER",
    "password": "YOUR_DB_PASSWORD",
    "port": 5432,
    "sslmode": "require"
}


# =========================================================
# AUTHORITY MAP
# =========================================================

authority_map = {
    "MONSHI G KATARI,VICE CONSUL": "MEA/IB/405/003",
    "GIRIRAJ SINGH KULSHEKHARA,VICE CONSUL AND ADMIN": "MEA/IB/405/004",
    "SHEESHACHELLAM SWAMINI, HOC AND CONSUL": "MEA/IB/405/002",
    "VEENA SAI RAJJAN KUMAR,CONSUL GENERAL": "MEA/IB/405/001",
}

state_master = {
    "ACT": "Australian Capital Territory",
    "NSW": "New South Wales",
    "NT":  "Northern Territory",
    "QLD": "Queensland",
    "SA":  "South Australia",
    "TAS": "Tasmania",
    "VIC": "Victoria",
    "WA":  "Western Australia"
}


# =========================================================
# INSERT — reads Excel and inserts into all 4 tables
# =========================================================

def insert_data(cur, conn):

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = [c.strip() for c in df.columns]

    # --- State master ---
    for code, name in state_master.items():
        cur.execute(
            """
            INSERT INTO state (state_code, state_name)
            VALUES (%s, %s)
            ON CONFLICT (state_code) DO NOTHING;
            """,
            (code.upper(), name.upper())
        )

    # --- IB Authority master ---
    for full_name, auth_id in authority_map.items():
        cur.execute(
            """
            INSERT INTO ib_authority (
                ib_staff_authority_id,
                ib_staff_authority_name,
                ib_staff_authority_designation
            )
            VALUES (%s, %s, %s)
            ON CONFLICT (ib_staff_authority_id) DO NOTHING;
            """,
            (
                auth_id,
                full_name.split(",")[0].strip(),
                full_name.split(",")[-1].strip()
            )
        )

    # --- Ministry of Home Affairs (parent) ---
    for _, row in df[["MHA_FILE_NUMBER", "MHA_DATE"]].drop_duplicates().iterrows():
        cur.execute(
            """
            INSERT INTO ministryofhomeaffairs (
                mha_file_number,
                mha_date
            )
            VALUES (%s, %s)
            ON CONFLICT (mha_file_number) DO NOTHING;
            """,
            (row["MHA_FILE_NUMBER"], row["MHA_DATE"])
        )

    # --- Applicant (child) ---
    for _, row in df.iterrows():

        ib_staff_id = authority_map.get(row["Signing_Authority"])

        cur.execute(
            """
            INSERT INTO applicant (
                file_number,
                name,
                sex,
                birth_date,
                place,
                state_code,
                name_of_father,
                name_of_mother,
                address_line_1,
                address_line_2,
                address_line_3,
                registration_date,
                mha_file_number,
                ib_staff_authority_id
            )
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (file_number) DO NOTHING;
            """,
            (
                row["File_Number"],
                row["Name"],
                row["Sex"],
                row["Birth_Date"],
                row["Place"],
                row["State"],
                row["Name_Of_Father"],
                row["Name_Of_Mother"],
                row["Address_line_1"],
                row["Address_line_2"],
                row["Address_line_3"],
                row["Registration_Date"],
                row["MHA_FILE_NUMBER"],
                ib_staff_id
            )
        )

    conn.commit()
    print("✅ DocuMate data insertion completed successfully")


# =========================================================
# SELECT — displays IN PROCESS records from database
# =========================================================

def select_data(cur):

    cur.execute("""
        SELECT 
            a.serial,
            UPPER(a.name),
            UPPER(a.sex),
            a.birth_date,
            UPPER(CONCAT(a.place, ', ', a.state_code)) AS place_of_birth,
            UPPER(a.name_of_father),
            UPPER(a.name_of_mother),
            UPPER(CONCAT(a.address_line_1, ', ',a.address_line_2, ', ',a.address_line_3)) AS address,
            a.registration_date,
            b.mha_file_number,
            b.mha_date,
            UPPER(ib.ib_staff_authority_name) AS signing_authority_name,
            UPPER(CONCAT(ib.ib_staff_authority_name, ', ',ib.ib_staff_authority_designation)) AS authority_name_designation
        FROM applicant a
            JOIN ministryofhomeaffairs b
                ON a.mha_file_number = b.mha_file_number
            LEFT JOIN ib_authority ib
                ON a.ib_staff_authority_id = ib.ib_staff_authority_id
        WHERE UPPER(a.status) = 'IN PROCESS'
        ORDER BY a.serial;
    """)

    rows = cur.fetchall()

    if not rows:
        print("📝 No records found with status 'IN PROCESS'.")
        return

    for row in rows:
        print(row)

    print(f"\n✅ {len(rows)} records found.")


# =========================================================
# MENU
# =========================================================

if __name__ == "__main__":

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    print("\n============================")
    print("  DocuMate — birth.py")
    print("============================")
    print("  1. Insert data from Excel")
    print("  2. Select IN PROCESS records")
    print("============================")

    choice = input("\nChoose (1 or 2): ").strip()

    if choice == "1":
        insert_data(cur, conn)
    elif choice == "2":
        select_data(cur)
    else:
        print("❌ Invalid choice.")

    cur.close()
    conn.close()