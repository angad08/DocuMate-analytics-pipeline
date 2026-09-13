"""
DocuMate - run any version from here.

    python main.py v3    Excel + docxtpl (the production one)
    python main.py x     Excel + Word Mail Merge, 250 at a time
    python main.py y     Excel + Word Mail Merge, all in one go
    python main.py z     PostgreSQL + docxtpl
    python main.py o     PostgreSQL + Word Mail Merge

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

    # --poll and --once contradict each other, so only one is allowed.
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
    """Print the version table."""
    print("KEY   SOURCE       ENGINE      BATCH    NOTES")

    for version in versions.VERSIONS.values():
        if version.engine == "mailmerge" and version.batch_size is None:
            batch = "all"
        elif version.batch_size:
            batch = str(version.batch_size)
        else:
            batch = "-"

        print("%-5s %-12s %-11s %-8s %s" % (
            version.key, version.source, version.engine, batch, version.notes
        ))


def check_setup():
    """
    Say what DocuMate can and cannot find, so a failed run is obvious
    instead of a guess. Never prints the password.
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
        ("psycopg2", "psycopg2"),
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
        try:
            settings = config.database_settings()
            print("  host     %s" % settings["host"])
            print("  database %s" % settings["database"])
            print("  user     %s" % settings["user"])
            print("  password %s" % ("set" if settings["password"] else "EMPTY"))
            print("  port     %s   sslmode %s" % (settings["port"], settings["sslmode"]))
        except Exception as error:
            print("\n  PROBLEM: %s" % error)

    print("\n" + "=" * 60 + "\n")


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.check:
        check_setup()
        return 0

    if args.list:
        list_versions()
        return 0

    version = versions.get(args.version)
    pipeline = versions.build(args.version)

    # The version has a default; --poll and --once override it.
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
    # Needed before any parallel work in a built .exe. Harmless otherwise.
    multiprocessing.freeze_support()
    sys.exit(main())
