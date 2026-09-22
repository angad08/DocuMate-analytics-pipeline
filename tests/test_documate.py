"""
Tests for DocuMate's logic: checks, sorting, file names, batching, the
version list, database settings and SQL, and the PDF settings.

Word and a live database are not needed. The tests cover the code that
decides what Word and the database are asked to do, not Word or the
database themselves.

Run them with:   python -m pytest tests -q
"""

import os
import sys

import pandas as pd
import pytest

# Let the tests import the project without installing anything.
# Walk up from this file to the folder that contains flow/ (the project
# root) and add it to the import path.
_here = os.path.dirname(os.path.abspath(__file__))
while _here != os.path.dirname(_here):
    if os.path.isdir(os.path.join(_here, "flow")):
        break
    _here = os.path.dirname(_here)
sys.path.insert(0, _here)

from flow import checks
from engines.output_paths import build_output_name
from engines.pdf import pdf_path_for
from engines.mailmerge_engine import plan_batches
from sources.postgres import build_record
from flow.checks import CheckFailed
from versions import registry


def sheet(**changes):
    """A small valid sheet. Override one column to break one check."""
    rows = {
        "File_Number": ["F1", "F2"],
        "Serial": ["2/2026", "1/2026"],
        "Name": ["ASHA", "BHAVIK"],
        "Registration_date": ["2026-01-05", "2026-01-06"],
        "STATUS": ["IN PROCESS", "IN PROCESS"],
        "Date_Printed": [None, None],
        "Date_Issued": [None, None],
    }
    rows.update(changes)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# The five checks
# ---------------------------------------------------------------------------

def test_a_good_sheet_passes():
    assert len(checks.check_records(sheet())) == 2


def test_empty_input_stops_the_run():
    with pytest.raises(CheckFailed):
        checks.check_records(pd.DataFrame())

    with pytest.raises(CheckFailed):
        checks.check_records(None)


def test_missing_status_column_stops_the_run():
    with pytest.raises(CheckFailed):
        checks.check_records(sheet().drop(columns=["STATUS"]))


def test_printed_rows_are_skipped():
    left = checks.check_records(sheet(STATUS=["PRINTED", "IN PROCESS"]))
    assert list(left["File_Number"]) == ["F2"]


def test_printed_is_matched_ignoring_case_and_spaces():
    with pytest.raises(CheckFailed):
        checks.check_records(sheet(STATUS=["  printed ", "PRINTED"]))


def test_a_blank_status_still_counts_as_pending():
    """A blank status must never make a record get skipped quietly."""
    assert len(checks.check_records(sheet(STATUS=[None, ""]))) == 2


def test_an_empty_required_field_stops_the_run():
    with pytest.raises(CheckFailed):
        checks.check_records(sheet(Name=["ASHA", "   "]))


def test_optional_columns_are_allowed_to_be_empty():
    """Date_Printed and Date_Issued start out empty and that's fine."""
    assert len(checks.check_records(sheet())) == 2


def test_an_empty_field_in_a_finished_row_does_not_stop_the_run():
    """Check 4 runs after filtering, so finished rows can't block us."""
    data = sheet(STATUS=["PRINTED", "IN PROCESS"], Name=[None, "BHAVIK"])
    assert len(checks.check_records(data)) == 1


def test_the_same_file_number_twice_stops_the_run():
    with pytest.raises(CheckFailed):
        checks.check_records(sheet(File_Number=["F1", "F1"]))


def test_the_missing_field_report_uses_excel_row_numbers():
    """Pandas row 1 is Excel row 3: one for the header, one for counting from 0."""
    with pytest.raises(CheckFailed) as caught:
        checks.check_records(sheet(Name=["ASHA", ""]))

    assert "Row 3" in "\n".join(caught.value.extra_lines)


# ---------------------------------------------------------------------------
# Tidying the records
# ---------------------------------------------------------------------------

def test_dates_come_out_day_first():
    data = checks.format_dates(sheet(), ["Registration_date"])
    assert list(data["Registration_date"]) == ["05/01/2026", "06/01/2026"]


def test_an_unreadable_date_comes_out_empty_not_as_the_word_nat():
    data = checks.format_dates(
        sheet(Registration_date=["nonsense", "2026-01-06"]),
        ["Registration_date"],
    )
    assert list(data["Registration_date"]) == ["", "06/01/2026"]


def test_sorting_uses_the_number_inside_serial():
    data = checks.sort_by_serial(sheet(Serial=["10/2026", "2/2026"]))
    assert list(data["Serial"]) == ["2/2026", "10/2026"]


def test_the_sorting_helper_column_is_cleaned_up():
    assert "sort_key" not in checks.sort_by_serial(sheet()).columns


def test_sorting_survives_a_serial_with_no_number_in_it():
    data = checks.sort_by_serial(sheet(Serial=["abc", "2/2026"]))
    assert list(data["Serial"]) == ["abc", "2/2026"]


def test_the_year_gets_added_to_every_record():
    from datetime import datetime

    records = checks.to_records(sheet(), add_year=True)
    for record in records:
        assert record["year"] == datetime.now().year


def test_adding_the_year_can_be_turned_off():
    assert "year" not in checks.to_records(sheet(), add_year=False)[0]


def test_field_names_match_the_column_names_exactly():
    """to_dict beats itertuples: no quiet renaming of odd column names."""
    data = sheet().rename(columns={"Name": "Full Name"})
    assert "Full Name" in checks.to_records(data, add_year=False)[0]


# ---------------------------------------------------------------------------
# Mail Merge batching
# ---------------------------------------------------------------------------

def test_every_record_ends_up_in_exactly_one_batch():
    batches = plan_batches(1440, 250)

    covered = []
    for first, last in batches:
        covered.extend(range(first, last + 1))

    assert covered == list(range(1, 1441))


def test_the_first_and_last_batch_are_right():
    batches = plan_batches(1440, 250)
    assert batches[0] == (1, 250)
    assert batches[-1] == (1251, 1440)


def test_an_exact_multiple_does_not_make_an_empty_batch():
    assert plan_batches(500, 250) == [(1, 250), (251, 500)]


def test_no_batch_size_means_one_single_run():
    """This is what version Y does."""
    assert plan_batches(900, None) == [(1, 900)]


def test_a_batch_size_bigger_than_the_total_is_one_run():
    assert plan_batches(10, 250) == [(1, 10)]


def test_no_records_means_no_batches():
    assert plan_batches(0, 250) == []


# ---------------------------------------------------------------------------
# Reading a database row
# ---------------------------------------------------------------------------

def database_row():
    import datetime

    return (
        "F/123", 7, "ASHA", "FEMALE",
        datetime.date(2026, 1, 2), "MELBOURNE, VIC",
        "FATHER", "MOTHER", "1 A ST, SUBURB, VIC",
        datetime.date(2026, 1, 5),
        "MHA/9", datetime.date(2026, 1, 6),
        "AUTHORITY", "AUTHORITY, CONSUL",
    )


def test_serial_gets_the_year_added():
    assert build_record(database_row(), 2026)["Serial"] == "7/2026"


def test_database_dates_come_out_day_first():
    record = build_record(database_row(), 2026)

    assert record["Registration_date"] == "05/01/2026"
    assert record["When_and_where_born"] == "02/01/2026, MELBOURNE, VIC"
    assert record["MHA_File_And_date"] == "MHA/9, 06/01/2026"


def test_a_database_row_produces_the_template_field_names():
    expected = {
        "File_Number", "Serial", "Name", "Sex", "When_and_where_born",
        "Name_of_the_Father", "Name_of_the_Mother",
        "Description_and_residence_of_informant", "Registration_date",
        "MHA_File_And_date", "Signing_Authority_Name",
        "Signing_Authority_Name_Designation",
    }
    assert set(build_record(database_row(), 2026)) == expected


# ---------------------------------------------------------------------------
# The version list
# ---------------------------------------------------------------------------

def test_all_five_versions_are_there():
    assert set(registry.VERSIONS) == {"v3", "x", "y", "z", "o"}


def test_each_version_key_matches_its_entry():
    for key in registry.VERSIONS:
        assert registry.VERSIONS[key].key == key


def test_every_version_names_a_real_source_and_engine():
    for version in registry.VERSIONS.values():
        assert version.source in ["excel", "database"]
        assert version.engine in ["docxtpl", "mailmerge"]


def test_the_database_versions_have_checks_turned_off():
    """They filter in SQL, so there is nothing left for the checks to do."""
    for version in registry.VERSIONS.values():
        if version.source == "database":
            assert version.check_records is False


def test_the_excel_versions_have_checks_turned_on():
    for version in registry.VERSIONS.values():
        if version.source == "excel":
            assert version.check_records is True


def test_mail_merge_versions_use_the_mail_merge_template():
    """A docxtpl template in a Mail Merge version would come out blank."""
    for version in registry.VERSIONS.values():
        if version.engine == "mailmerge":
            assert version.template.endswith("_MM.docx")
        else:
            assert not version.template.endswith("_MM.docx")


# ---------------------------------------------------------------------------
# The two database backends
#
# No database connection is made. These tests cover how the backend is
# chosen from .env, the settings each backend receives, and the SQL each
# one builds.
# ---------------------------------------------------------------------------

import datetime

from setup import config
from sources import azuresql
from sources import queries_postgres
from sources import queries_tsql


@pytest.fixture
def no_backend_set(monkeypatch):
    """Clear DOCUMATE_DB_BACKEND and skip loading .env, so defaults apply."""
    monkeypatch.delenv("DOCUMATE_DB_BACKEND", raising=False)
    monkeypatch.setattr(config, "load_environment", lambda: None)


def test_the_default_backend_is_one_we_support(no_backend_set):
    assert config.database_backend() in config.DATABASE_BACKENDS


def test_the_backend_can_be_switched_from_the_environment(no_backend_set, monkeypatch):
    for name in config.DATABASE_BACKENDS:
        monkeypatch.setenv("DOCUMATE_DB_BACKEND", name.upper() + "  ")
        assert config.database_backend() == name


def test_an_unknown_backend_is_refused_by_name(no_backend_set, monkeypatch):
    monkeypatch.setenv("DOCUMATE_DB_BACKEND", "mysql")

    with pytest.raises(ValueError) as failure:
        config.database_backend()

    assert "mysql" in str(failure.value)


def test_each_backend_gets_its_own_default_port(no_backend_set, monkeypatch):
    monkeypatch.delenv("DOCUMATE_DB_PORT", raising=False)
    for key in ["HOST", "NAME", "USER", "PASSWORD"]:
        monkeypatch.setenv("DOCUMATE_DB_" + key, "x")

    for name, backend in config.DATABASE_BACKENDS.items():
        monkeypatch.setenv("DOCUMATE_DB_BACKEND", name)
        assert config.database_settings()["port"] == int(backend["port"])


def test_sslmode_only_goes_to_postgres(no_backend_set, monkeypatch):
    """sslmode is a psycopg2 option, so only the PostgreSQL settings include it."""
    for key in ["HOST", "NAME", "USER", "PASSWORD"]:
        monkeypatch.setenv("DOCUMATE_DB_" + key, "x")

    monkeypatch.setenv("DOCUMATE_DB_BACKEND", "postgres")
    assert "sslmode" in config.database_settings()

    monkeypatch.setenv("DOCUMATE_DB_BACKEND", "azuresql")
    assert "sslmode" not in config.database_settings()


def test_both_dialects_offer_the_same_names():
    """sources/database.py uses these names from whichever queries module is selected."""
    for name in ["PENDING_APPLICANTS", "DATE_ISSUED_EXISTS",
                 "ADD_DATE_ISSUED_COLUMN", "mark_printed_sql",
                 "mark_printed_params"]:
        assert hasattr(queries_postgres, name), "postgres is missing " + name
        assert hasattr(queries_tsql, name), "tsql is missing " + name


def test_each_dialect_builds_an_update_its_own_driver_can_run():
    today = datetime.date(2026, 1, 2)
    serials = [1, 2, 3]

    # PostgreSQL: one array parameter, so two placeholders whatever the size.
    pg_sql = queries_postgres.mark_printed_sql(len(serials))
    assert pg_sql.count("%s") == 2
    assert queries_postgres.mark_printed_params(today, serials) == (today, serials)

    # Azure SQL: one ? per serial, plus one for the date.
    tsql = queries_tsql.mark_printed_sql(len(serials))
    assert tsql.count("?") == len(serials) + 1
    assert "IN (?, ?, ?)" in tsql
    assert queries_tsql.mark_printed_params(today, serials) == [today] + serials


def test_tsql_refuses_a_batch_over_the_parameter_limit():
    """An empty batch, or one over MAX_SERIALS_PER_UPDATE, raises a clear error."""
    with pytest.raises(ValueError):
        queries_tsql.mark_printed_sql(0)

    with pytest.raises(ValueError):
        queries_tsql.mark_printed_sql(queries_tsql.MAX_SERIALS_PER_UPDATE + 1)


def test_an_odbc_password_with_punctuation_survives():
    """A password containing ; = and } is wrapped in braces and passed intact."""
    settings = {
        "host": "srv.database.windows.net", "database": "DocuMate",
        "user": "angadadmin", "password": "p;a=ss}word", "port": 1433,
    }

    built = azuresql.connection_string(settings, "ODBC Driver 18 for SQL Server")

    assert "PWD={p;a=ss}}word};" in built
    assert "Encrypt=yes" in built
    assert "sslmode" not in built


def test_no_odbc_driver_says_how_to_get_one():
    class NoDrivers:
        @staticmethod
        def drivers():
            return ["Microsoft Access Driver (*.mdb, *.accdb)"]

    with pytest.raises(RuntimeError) as failure:
        azuresql.find_driver(NoDrivers)

    assert "ODBC Driver 18 for SQL Server" in str(failure.value)


def test_no_two_versions_share_an_output_name():
    prefixes = [v.output_prefix for v in registry.VERSIONS.values()]
    assert len(prefixes) == len(set(prefixes))


def test_a_wrong_version_name_lists_the_real_ones():
    with pytest.raises(SystemExit):
        registry.get("nope")


def test_version_names_are_not_case_sensitive():
    assert registry.get("V3").key == "v3"


# ---------------------------------------------------------------------------
# Output filenames
# ---------------------------------------------------------------------------

def test_the_output_name_is_built_when_called_not_when_imported():
    """
    The old scripts put the timestamp in a default argument, so a version
    that polls stamped the time it started, forever, and overwrote its own
    output. This has to be worked out fresh each time.
    """
    name = os.path.basename(build_output_name("out", "DocuMateZ", "%d%m%Y_%H%M%S"))

    assert name.startswith("DocuMateZ_BIRTH_REGISTRATION_")
    assert name.endswith(".docx")


def test_each_version_produces_its_own_filename():
    names = set()
    for version in registry.VERSIONS.values():
        names.add(os.path.basename(build_output_name(
            "out", version.output_prefix, version.timestamp_format
        )))

    assert len(names) == len(registry.VERSIONS)


# ---------------------------------------------------------------------------
# PDF copy
# ---------------------------------------------------------------------------

def test_the_pdf_sits_next_to_the_word_file_with_the_same_name():
    docx = os.path.join("out", "DocuMateX_BIRTH_REGISTRATION_13092026.docx")
    pdf = pdf_path_for(docx)

    assert os.path.dirname(pdf) == os.path.abspath("out")
    assert os.path.basename(pdf) == "DocuMateX_BIRTH_REGISTRATION_13092026.pdf"


def test_every_version_has_a_yes_or_no_pdf_setting():
    for version in registry.VERSIONS.values():
        assert version.to_pdf in (True, False)


def test_pdf_is_off_unless_a_version_asks_for_it():
    version = registry.Version("t", "T", "excel", "docxtpl", "T.docx", "T")
    assert version.to_pdf is False


def test_both_engines_are_handed_the_pdf_setting():
    """build_engine passes each version's to_pdf setting to its engine."""
    import versions

    for version in registry.VERSIONS.values():
        assert versions.build_engine(version).to_pdf == version.to_pdf
