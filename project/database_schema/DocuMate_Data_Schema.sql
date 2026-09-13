-- =========================================================
-- SCHEMA: Birth Registration System (DocuMate)
-- FULL FINAL SCRIPT (STATE NORMALIZED)
-- =========================================================
-- ✔ Normalized
-- ✔ Production-safe
-- ✔ BI-ready
-- ✔ DocuMate-compatible
-- ✔ Single runnable script
-- =========================================================



-- =========================================================
-- TABLE 1: State Master (Dimension)
-- =========================================================

CREATE TABLE IF NOT EXISTS state (
    state_code CHAR(3) PRIMARY KEY,
    state_name TEXT NOT NULL
);

-- =========================================================
-- TABLE 2: MinistryOfHomeAffairs (Parent)
-- =========================================================

CREATE TABLE IF NOT EXISTS ministryofhomeaffairs (
    mha_file_number VARCHAR(50) PRIMARY KEY,
    mha_date DATE NOT NULL
);



-- =========================================================
-- TABLE 3: IB_Authority (Parent)
-- =========================================================

CREATE TABLE IF NOT EXISTS ib_authority (
    ib_staff_authority_id VARCHAR(50) PRIMARY KEY,
    ib_staff_authority_name TEXT NOT NULL,
    ib_staff_authority_designation TEXT NOT NULL
);



-- =========================================================
-- TABLE 4: Applicant (Child / Fact)
-- =========================================================

CREATE TABLE IF NOT EXISTS applicant (
    file_number VARCHAR(30) PRIMARY KEY,
    serial INT GENERATED ALWAYS AS IDENTITY,

    name TEXT NOT NULL,
    sex CHAR(1) CHECK (sex IN ('M','F')),
    birth_date DATE NOT NULL,

    place TEXT NOT NULL,
    state_code CHAR(3) NOT NULL,

    name_of_father TEXT NOT NULL,
    name_of_mother TEXT NOT NULL,

    address_line_1 TEXT,
    address_line_2 TEXT,
    address_line_3 TEXT,

    registration_date DATE NOT NULL,
    status TEXT NOT NULL DEFAULT 'IN PROCESS',
    date_issued DATE,

    mha_file_number VARCHAR(50),
    ib_staff_authority_id VARCHAR(50),

    CONSTRAINT fk_applicant_state
        FOREIGN KEY (state_code)
        REFERENCES state (state_code)
        ON UPDATE CASCADE
        ON DELETE RESTRICT,

    CONSTRAINT fk_applicant_mha
        FOREIGN KEY (mha_file_number)
        REFERENCES ministryofhomeaffairs (mha_file_number)
        ON UPDATE CASCADE
        ON DELETE SET NULL,

    CONSTRAINT fk_applicant_ib_authority
        FOREIGN KEY (ib_staff_authority_id)
        REFERENCES ib_authority (ib_staff_authority_id)
        ON UPDATE CASCADE
        ON DELETE SET NULL
);



-- =========================================================
-- CONSTRAINTS
-- =========================================================

ALTER TABLE applicant
ADD CONSTRAINT chk_applicant_status
CHECK (UPPER(status) IN ('IN PROCESS', 'PRINTED'));



-- =========================================================
-- INDEXES (PERFORMANCE + BI)
-- =========================================================

CREATE INDEX IF NOT EXISTS idx_applicant_serial
ON applicant (serial);

CREATE INDEX IF NOT EXISTS idx_applicant_status_upper
ON applicant (UPPER(status));

CREATE INDEX IF NOT EXISTS idx_applicant_mha
ON applicant (mha_file_number);

CREATE INDEX IF NOT EXISTS idx_applicant_ib
ON applicant (ib_staff_authority_id);

CREATE INDEX IF NOT EXISTS idx_applicant_state
ON applicant (state_code);



-- =========================================================
-- END OF SCRIPT
-- =========================================================
