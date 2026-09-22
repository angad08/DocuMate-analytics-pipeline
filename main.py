"""
DocuMate - run any version from here.

    python main.py v3    Excel + docxtpl (the production version)
    python main.py x     Excel + Word Mail Merge, 250 at a time
    python main.py y     Excel + Word Mail Merge, all in one go
    python main.py z     database + docxtpl
    python main.py o     database + Word Mail Merge

    (z and o read Azure SQL or PostgreSQL, as set by DOCUMATE_DB_BACKEND
     in .env. --check shows which one is selected.)

    python main.py --list    show every version and what it uses
    python main.py --check   check the setup and stop
    python main.py z --once  run Z once instead of on a loop
    python main.py x --poll  make any version check on a loop

"""

import argparse
import multiprocessing
import sys
import versions


def build_parser():
    """Define the command-line options. The module docstring is the --help text."""
    parser = argparse.ArgumentParser(
        prog="documate",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "version",
        nargs="?",
        default="v3",
        type=str.lower,
        help="which version to run (default: v3)",
    )

    parser.add_argument(
        "--list",
        action="store_true",
        help="list the versions and exit",
    )

    parser.add_argument(
        "--check",
        action="store_true",
        help="check the setup: folders, files, packages, database settings",
    )

    # --poll and --once cannot be used together.
    loop = parser.add_mutually_exclusive_group()
    loop.add_argument("--poll", action="store_true", help="keep checking for new records")
    loop.add_argument("--once", action="store_true", help="run once, even if the version polls")

    parser.add_argument(
        "--interval",
        type=int,
        default=None,
        help="seconds between checks (default: whatever the version says)",
    )

    return parser


def list_versions():
    """
    Print a table of every version: source, engine, batch size, PDF on/off.

    BATCH shows the Mail Merge batch size, "all" for no batching, or "-"
    for docxtpl versions, which do not batch.
    """
    print("KEY   SOURCE       ENGINE      BATCH    PDF   NOTES")

    for version in versions.VERSIONS.values():
        if version.engine == "mailmerge" and version.batch_size is None:
            batch = "all"
        elif version.batch_size:
            batch = str(version.batch_size)
        else:
            batch = "-"

        pdf = "yes" if version.to_pdf else "no"

        print("%-5s %-12s %-11s %-8s %-5s %s" % (
            version.key, version.source, version.engine, batch, pdf, version.notes
        ))


def check_odbc_driver():
    """
    Print which ODBC driver Azure SQL will use, or why none was found.

    pyodbc can be installed while the Microsoft ODBC driver it needs is
    missing, so the driver is checked separately.
    """
    try:
        import pyodbc
        from sources.azuresql import find_driver
    except ImportError:
        print("  driver   pyodbc NOT INSTALLED - run: pip install pyodbc")
        return

    try:
        print("  driver   %s" % find_driver(pyodbc))
    except RuntimeError as error:
        print("  driver   %s" % error)


def check_setup():
    """
    Print what DocuMate can and cannot find (python main.py --check).

    Sections: FOLDERS, FILES (workbook and templates), PACKAGES, and
    DATABASE SETTINGS. Run it first when something is not working. The
    password is never printed, only whether it is set.
    """
    import os
    from setup import config

    print("\nDocuMate setup check")
    print("=" * 60)

    print("\nFOLDERS")
    for label, path in [
        ("project root", config.project_root()),
        ("data",         config.data_folder()),
        ("templates",    config.templates_folder()),
        ("output",       config.output_folder()),
    ]:
        mark = "ok     " if os.path.exists(path) else "MISSING"
        print("  %-8s %s  %s" % (label, mark, path))

    print("\nFILES")
    excel = config.excel_path()
    print("  %-8s %s  %s" % ("workbook",
                             "ok     " if os.path.exists(excel) else "MISSING",
                             excel))

    for version in versions.VERSIONS.values():
        path = version.template_path()
        if os.path.exists(path):
            continue
        print("  template MISSING  %s   (needed by %s)" % (path, version.label))

    print("\nPACKAGES")
    for name, module in [
        ("pandas", "pandas"),
        ("openpyxl", "openpyxl"),
        ("docxtpl", "docxtpl"),
        ("docxcompose", "docxcompose"),
        ("python-dotenv", "dotenv"),
        # Only the selected database backend needs its driver, so it is
        # fine for one of these two to be missing. DATABASE SETTINGS below
        # shows which backend is selected.
        ("psycopg2", "psycopg2"),
        ("pyodbc", "pyodbc"),
        ("pywin32", "win32com"),
    ]:
        try:
            __import__(module)
            print("  %-14s ok" % name)
        except ImportError:
            print("  %-14s NOT INSTALLED" % name)

    print("\nDATABASE SETTINGS  (needed by z and o only)")
    settings_file = config.env_file()

    if not settings_file:
        print("  no settings file found in the project root.")
        print("  expected one named .env, or anything ending in .env")
    else:
        print("  reading  %s" % settings_file)

        # The backend and driver are checked in their own try block, so
        # they are still shown when the connection settings below are
        # incomplete.
        try:
            backend = config.database_backend()
            print("  backend  %s   (DOCUMATE_DB_BACKEND)" % backend)

            if backend == "azuresql":
                check_odbc_driver()

        except Exception as error:
            print("  backend  PROBLEM: %s" % error)
            backend = None

        try:
            settings = config.database_settings()
            print("  host     %s" % settings["host"])
            print("  database %s" % settings["database"])
            print("  user     %s" % settings["user"])
            print("  password %s" % ("set" if settings["password"] else "EMPTY"))
            print("  port     %s" % settings["port"])

            schema = config.database_schema()
            print("  schema   %s" % (schema if schema else "(connection default)"))

            if backend == "postgres":
                print("  sslmode  %s" % settings["sslmode"])

        except Exception as error:
            print("\n  PROBLEM: %s" % error)

    print("\n" + "=" * 60 + "\n")


def main(argv=None):
    """
    Run DocuMate with the given arguments (sys.argv when None).

    --check and --list print and stop. Otherwise the version is built and
    run, once or on a loop depending on its poll setting and the flags.
    """
    args = build_parser().parse_args(argv)

    if args.check:
        check_setup()
        return 0

    if args.list:
        list_versions()
        return 0

    version = versions.get(args.version)
    pipeline = versions.build(args.version)

    # Start from the version's own poll setting; --poll or --once override it.
    poll = version.poll
    if args.poll:
        poll = True
    if args.once:
        poll = False

    interval = args.interval
    if interval is None:
        interval = version.poll_interval

    pipeline.start(poll=poll, interval=interval)
    return 0


if __name__ == "__main__":
    # Required in a frozen exe before starting worker processes. Does
    # nothing when run as a normal script.
    multiprocessing.freeze_support()
    sys.exit(main())
