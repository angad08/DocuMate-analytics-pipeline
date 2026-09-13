"""
Plugging the parts together.

registry.py says a version uses "excel" and "mailmerge". This file turns
those words into the actual objects and hands them to the Pipeline.

It is the only place that connects a name to a class, so adding a new
source or engine means changing it here once, and then every version can
use it.
"""

from documate.setup import config
from documate.engines.mailmerge_engine import MailMergeEngine
from documate.engines.docxtpl_engine import DocxtplEngine
from documate.flow.pipeline import Pipeline
from documate.sources.excel import ExcelSource
from documate.versions.registry import VERSIONS
from documate.versions.registry import Version
from documate.versions.registry import get


def build_source(version):
    """Create the data source this version asked for."""

    if version.source == "excel":
        return ExcelSource(config.excel_path(), config.sheet_name())

    if version.source == "postgres":
        # Imported here so the Excel versions don't need psycopg2 installed.
        from documate.sources.postgres import DatabaseSource
        return DatabaseSource(config.database_settings())

    raise ValueError(
        "Version '" + version.key + "' asks for an unknown source: " + str(version.source)
    )


def build_engine(version):
    """Create the engine this version asked for."""

    if version.engine == "docxtpl":
        return DocxtplEngine(version.template_path(), max_workers=version.max_workers)

    if version.engine == "mailmerge":
        return MailMergeEngine(version.template_path(), batch_size=version.batch_size)

    raise ValueError(
        "Version '" + version.key + "' asks for an unknown engine: " + str(version.engine)
    )


def build(key):
    """Get a Pipeline ready to run the named version."""
    version = get(key)

    return Pipeline(
        version=version,
        source=build_source(version),
        engine=build_engine(version),
    )
