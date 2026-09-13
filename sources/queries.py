"""
The SQL used by the database source.

Kept in its own file so you can read the queries without reading the Python
around them, and so a schema change is a diff in one place.
"""

# Pending applicants, joined to MHA and the signing authority.
# The "not yet printed" filter happens here rather than in Python, so the
# database only sends the rows we are actually going to use.
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
    FROM applicant a
        JOIN ministryofhomeaffairs b
            ON a.mha_file_number = b.mha_file_number
        LEFT JOIN ib_authority ib
            ON a.ib_staff_authority_id = ib.ib_staff_authority_id
    WHERE UPPER(COALESCE(a.status, '')) <> 'PRINTED'
    ORDER BY a.serial;
"""

# Does the date_issued column exist? Older databases were set up without it.
DATE_ISSUED_EXISTS = """
    SELECT 1
    FROM information_schema.columns
    WHERE table_name = 'applicant'
      AND column_name = 'date_issued';
"""

ADD_DATE_ISSUED_COLUMN = "ALTER TABLE applicant ADD COLUMN date_issued DATE;"

# Mark the given serial numbers as printed.
# The status check at the end means running this twice does nothing the
# second time, instead of re-stamping the date.
MARK_PRINTED = """
    UPDATE applicant
    SET status = 'PRINTED',
        date_issued = %s
    WHERE serial = ANY(%s)
      AND COALESCE(UPPER(status), '') <> 'PRINTED';
"""
