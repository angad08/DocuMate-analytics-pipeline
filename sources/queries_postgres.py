"""
SQL for the PostgreSQL backend.

sources/queries_tsql.py has the same queries for Azure SQL. Both files
provide the same names, so sources/database.py can use either one:

    PENDING_APPLICANTS        applicants not yet PRINTED
    DATE_ISSUED_EXISTS        does applicant.date_issued exist?
    ADD_DATE_ISSUED_COLUMN    add it if not
    mark_printed_sql(count)   the UPDATE for a batch of serials
    mark_printed_params(...)  the values for that UPDATE

If you change a query here, make the same change in queries_tsql.py.

Table names start with {schema}. DatabaseSource.sql() replaces it with
"<DOCUMATE_DB_SCHEMA>." or with nothing when no schema is set.
"""

# Every applicant not yet PRINTED, with their MHA file details and signing
# authority. The column order must match COLUMN_ORDER in sources/database.py.
PENDING_APPLICANTS = """
    SELECT
        a.file_number,
        a.serial,
        UPPER(a.name),
        UPPER(a.sex),
        a.birth_date,
        UPPER(CONCAT(a.place, ', ', a.state_code))                AS place_of_birth,
        UPPER(a.name_of_father),
        UPPER(a.name_of_mother),
        UPPER(CONCAT(a.address_line_1, ', ',
                     a.address_line_2, ', ',
                     a.address_line_3))                           AS address,
        a.registration_date,
        b.mha_file_number,
        b.mha_date,
        UPPER(ib.ib_staff_authority_name)                         AS signing_authority_name,
        UPPER(CONCAT(ib.ib_staff_authority_name, ', ',
                     ib.ib_staff_authority_designation))          AS authority_name_designation
    FROM {schema}applicant a
        JOIN {schema}ministryofhomeaffairs b
            ON a.mha_file_number = b.mha_file_number
        LEFT JOIN {schema}ib_authority ib
            ON a.ib_staff_authority_id = ib.ib_staff_authority_id
    WHERE UPPER(COALESCE(a.status, '')) <> 'PRINTED'
    ORDER BY a.serial;
"""

# Returns a row if applicant.date_issued exists. Older databases may not
# have this column.
DATE_ISSUED_EXISTS = """
    SELECT 1
    FROM information_schema.columns
    WHERE table_name = 'applicant'
      AND column_name = 'date_issued'{schema_filter};
"""

ADD_DATE_ISSUED_COLUMN = "ALTER TABLE {schema}applicant ADD COLUMN date_issued DATE;"

# Set the given serials to PRINTED with today's date.
#   first %s   the date
#   second %s  the list of serials, passed as one array
# Rows that are already PRINTED are skipped, so running it twice does not
# change their date_issued.
MARK_PRINTED = """
    UPDATE {schema}applicant
    SET status = 'PRINTED',
        date_issued = %s
    WHERE serial = ANY(%s)
      AND COALESCE(UPPER(status), '') <> 'PRINTED';
"""


def mark_printed_sql(count):
    """
    Return the UPDATE for a batch of `count` serials.

    PostgreSQL takes the whole batch as one array parameter, so the SQL is
    the same for any batch size and `count` is not used. It is accepted so
    this function has the same signature as the one in queries_tsql.py.
    """
    return MARK_PRINTED


def mark_printed_params(today, serials):
    """Return the UPDATE's values: (date, [serial, serial, ...])."""
    return (today, list(serials))
