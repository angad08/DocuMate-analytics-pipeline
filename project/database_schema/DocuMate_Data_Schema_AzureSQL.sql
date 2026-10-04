-- =========================================================
-- SCHEMA: Birth Registration System (DocuMate)
-- AZURE SQL / SQL SERVER (T-SQL)
-- =========================================================
-- The same four tables as DocuMate_Data_Schema.sql, written in T-SQL.
-- Keep the two in step: if you change a column here, change it there too.
--
-- Run it with sqlcmd, the Azure portal query editor, SSMS or Azure Data
-- Studio - anything that understands GO as a batch separator:
--
--     sqlcmd -S YOUR_SERVER.database.windows.net -d YOUR_DB -U YOUR_USER -P YOUR_PASSWORD ^
--            -i project/database_schema/DocuMate_Data_Schema_AzureSQL.sql
--
-- Then load the sample rows with:  python -m tools.seed_database --insert
--
-- Differences from the PostgreSQL script, and why:
--
--     PostgreSQL                        T-SQL (here)
--     ----------                        ------------
--     CREATE TABLE IF NOT EXISTS        IF OBJECT_ID(...) IS NULL
--     TEXT                              NVARCHAR(MAX), NVARCHAR(n) where
--                                       the column is indexed or checked
--     INT GENERATED ALWAYS AS IDENTITY  INT IDENTITY(1,1)
--     ON DELETE RESTRICT                ON DELETE NO ACTION
--     CREATE INDEX IF NOT EXISTS        IF NOT EXISTS (SELECT ... sys.indexes)
--     CREATE INDEX ON t (UPPER(status)) a plain index on status - see note
--
-- Tables are created in your login's default schema, normally dbo. To put
-- them somewhere else, create that schema first, make it your default, and
-- set DOCUMATE_DB_SCHEMA to the same name in .env.
-- =========================================================


-- =========================================================
-- TABLE 1: State Master (Dimension)
-- =========================================================

IF OBJECT_ID('state', 'U') IS NULL
CREATE TABLE state (
    state_code CHAR(3) NOT NULL PRIMARY KEY,
    state_name NVARCHAR(100) NOT NULL
);
GO


-- =========================================================
-- TABLE 2: MinistryOfHomeAffairs (Parent)
-- =========================================================

IF OBJECT_ID('ministryofhomeaffairs', 'U') IS NULL
CREATE TABLE ministryofhomeaffairs (
    mha_file_number VARCHAR(50) NOT NULL PRIMARY KEY,
    mha_date DATE NOT NULL
);
GO


-- =========================================================
-- TABLE 3: IB_Authority (Parent)
-- =========================================================

IF OBJECT_ID('ib_authority', 'U') IS NULL
CREATE TABLE ib_authority (
    ib_staff_authority_id VARCHAR(50) NOT NULL PRIMARY KEY,
    ib_staff_authority_name NVARCHAR(MAX) NOT NULL,
    ib_staff_authority_designation NVARCHAR(MAX) NOT NULL
);
GO


-- =========================================================
-- TABLE 4: Applicant (Child / Fact)
-- =========================================================
-- status is NVARCHAR(50) rather than NVARCHAR(MAX): it carries a CHECK
-- constraint and an index, and NVARCHAR(MAX) can have neither.

IF OBJECT_ID('applicant', 'U') IS NULL
CREATE TABLE applicant (
    file_number VARCHAR(30) NOT NULL PRIMARY KEY,
    serial INT IDENTITY(1,1) NOT NULL,

    name NVARCHAR(MAX) NOT NULL,
    sex CHAR(1) NULL CHECK (sex IN ('M','F')),
    birth_date DATE NOT NULL,

    place NVARCHAR(MAX) NOT NULL,
    state_code CHAR(3) NOT NULL,

    name_of_father NVARCHAR(MAX) NOT NULL,
    name_of_mother NVARCHAR(MAX) NOT NULL,

    address_line_1 NVARCHAR(MAX) NULL,
    address_line_2 NVARCHAR(MAX) NULL,
    address_line_3 NVARCHAR(MAX) NULL,

    registration_date DATE NOT NULL,
    status NVARCHAR(50) NOT NULL CONSTRAINT df_applicant_status DEFAULT 'IN PROCESS',
    date_issued DATE NULL,

    mha_file_number VARCHAR(50) NULL,
    ib_staff_authority_id VARCHAR(50) NULL,

    CONSTRAINT fk_applicant_state
        FOREIGN KEY (state_code)
        REFERENCES state (state_code)
        ON UPDATE CASCADE
        ON DELETE NO ACTION,

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
GO


-- =========================================================
-- CONSTRAINTS
-- =========================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.check_constraints WHERE name = 'chk_applicant_status'
)
ALTER TABLE applicant
ADD CONSTRAINT chk_applicant_status
CHECK (UPPER(status) IN ('IN PROCESS', 'PRINTED'));
GO


-- =========================================================
-- INDEXES (PERFORMANCE + BI)
-- =========================================================
-- The PostgreSQL script indexes UPPER(status). SQL Server cannot index an
-- expression directly - it would need a persisted computed column - and it
-- does not need to: the default collation is case-insensitive, so a plain
-- index on status serves the same lookups.

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'idx_applicant_serial' AND object_id = OBJECT_ID('applicant'))
CREATE INDEX idx_applicant_serial ON applicant (serial);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'idx_applicant_status' AND object_id = OBJECT_ID('applicant'))
CREATE INDEX idx_applicant_status ON applicant (status);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'idx_applicant_mha' AND object_id = OBJECT_ID('applicant'))
CREATE INDEX idx_applicant_mha ON applicant (mha_file_number);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'idx_applicant_ib' AND object_id = OBJECT_ID('applicant'))
CREATE INDEX idx_applicant_ib ON applicant (ib_staff_authority_id);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'idx_applicant_state' AND object_id = OBJECT_ID('applicant'))
CREATE INDEX idx_applicant_state ON applicant (state_code);
GO


-- =========================================================
-- END OF SCRIPT
-- =========================================================
