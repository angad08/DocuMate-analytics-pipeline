"""
Builds a ready-to-run Pipeline from a version's settings.

versions/registry.py describes a version with words, e.g. source="excel"
and engine="mailmerge". This file turns those words into real objects:

    build("x")
      -> get("x")                  the Version entry from registry.py
      -> build_source(version)     e.g. ExcelSource(...)
      -> build_engine(version)     e.g. MailMergeEngine(...)
      -> Pipeline(version, source, engine)

To add a new kind of source or engine, add a branch to build_source() or
build_engine() here. Every version can then use it by name.
"""

from setup import config
from engines.mailmerge_engine import MailMergeEngine
from engines.docxtpl_engine import DocxtplEngine
from flow.pipeline import Pipeline
from sources.excel import ExcelSource
from versions.registry import VERSIONS
from versions.registry import Version
from versions.registry import get


def build_source(version):
    """Create the data source named by version.source."""

    if version.source == "excel":
        return ExcelSource(config.excel_path(), config.sheet_name())

    if version.source == "database":
        # The database backend comes from DOCUMATE_DB_BACKEND in .env, not
        # from the version. The backend module is imported only now, so
        # only the selected backend's driver (pyodbc or psycopg2) needs to
        # be installed, and the Excel versions need neither.
        import importlib

        backend = config.DATABASE_BACKENDS[config.database_backend()]
        module = importlib.import_module(backend["module"])
        source_class = getattr(module, backend["cls"])

        return source_class(config.database_settings(), config.database_schema())

    raise ValueError(
        "Version '" + version.key + "' asks for an unknown source: " + str(version.source)
    )


def build_engine(version):
    """Create the engine named by version.engine, with the version's settings."""

    if version.engine == "docxtpl":
        return DocxtplEngine(
            version.template_path(),
            max_workers=version.max_workers,
            to_pdf=version.to_pdf,
        )

    if version.engine == "mailmerge":
        return MailMergeEngine(
            version.template_path(),
            batch_size=version.batch_size,
            to_pdf=version.to_pdf,
        )

    raise ValueError(
        "Version '" + version.key + "' asks for an unknown engine: " + str(version.engine)
    )


def build(key):
    """Return a Pipeline ready to run the version with this key."""
    version = get(key)

    return Pipeline(
        version=version,
        source=build_source(version),
        engine=build_engine(version),
    )
