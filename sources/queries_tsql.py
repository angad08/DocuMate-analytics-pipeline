"""
SQL for the Azure SQL backend (T-SQL).

These are the same queries as sources/queries_postgres.py, and this file
provides the same names, so sources/database.py can use either one. If you
change a query here, make the same change in queries_postgres.py.

Differences from the PostgreSQL version:

    PostgreSQL               T-SQL (here)
    ----------               ------------
    %s placeholders          ?             pyodbc uses ? for parameters
    serial = ANY(%s)         IN (?, ?, ?)  T-SQL has no array parameter
    ADD COLUMN date_issued   ADD date_issued

Table names start with {schema}. DatabaseSource.sql() replaces it with
"<DOCUMATE_DB_SCHEMA>." or with nothing when no schema is set.
"""

# Every applicant not yet PRINTED, with their MHA file details and signing
# authority. The column order must match COLUMN_ORDER in sources/database.py.
#
# Every column has an alias. Python reads the columns by position, so the
# aliases are not needed by the code, but they give readable column names
# when the query is run by hand in the Azure portal query editor.
PENDING_APPLICANTS = """
    SELECT
        a.file_number                                             AS file_number,
        a.serial                                                  AS serial,
        UPPER(a.name)                                             AS name,
        UPPER(a.sex)                                              AS sex,
        a.birth_date                                              AS birth_date,
        UPPER(CONCAT(a.place, ', ', a.state_code))                AS place_of_birth,
        UPPER(a.name_of_father)                                   AS name_of_father,
        UPPER(a.name_of_mother)                                   AS name_of_mother,
        UPPER(CONCAT(a.address_line_1, ', ',
                     a.address_line_2, ', ',
                     a.address_line_3))                           AS address,
        a.registration_date                                       AS registration_date,
        b.mha_file_number                                         AS mha_file_number,
        b.mha_date                                                AS mha_date,
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

# T-SQL uses ADD, not ADD COLUMN.
ADD_DATE_ISSUED_COLUMN = "ALTER TABLE {schema}applicant ADD date_issued DATE;"

# Set the given serials to PRINTED with today's date.
#
# {serials} is replaced by one ? per serial in the batch, for example
# "?, ?, ?" for three serials (see mark_printed_sql). Each serial is still
# sent as a parameter, never pasted into the SQL as text.
#
# Rows that are already PRINTED are skipped, so running it twice does not
# change their date_issued.
MARK_PRINTED = """
    UPDATE {schema}applicant
    SET status = 'PRINTED',
        date_issued = ?
    WHERE serial IN ({serials})
      AND COALESCE(UPPER(status), '') <> 'PRINTED';
"""

# SQL Server allows at most 2100 parameters in one statement, and one is
# used for the date. This cap keeps each UPDATE well under that limit.
# DatabaseSource sends 300 serials per batch by default.
MAX_SERIALS_PER_UPDATE = 2000


def mark_printed_sql(count):
    """
    Return the UPDATE for a batch of `count` serials.

        mark_printed_sql(3)  ->  "... date_issued = ? WHERE serial IN (?, ?, ?) ..."

    The SQL has count + 1 placeholders: the date first, then one per
    serial. mark_printed_params() returns the values in the same order.

    Raises ValueError if count is below 1 or above MAX_SERIALS_PER_UPDATE.
    """
    if count < 1:
        raise ValueError("mark_printed_sql needs at least one serial.")

    if count > MAX_SERIALS_PER_UPDATE:
        raise ValueError(
            "Too many serials for one UPDATE (%d). SQL Server allows 2100 "
            "parameters per statement. Lower update_batch_size in "
            "sources/database.py." % count
        )

    # str.replace rather than str.format, so the {schema} placeholders are
    # left in place for DatabaseSource.sql() to fill in.
    return MARK_PRINTED.replace("{serials}", ", ".join(["?"] * count))


def mark_printed_params(today, serials):
    """Return the UPDATE's values: [date, serial, serial, ...]."""
    return [today] + list(serials)
