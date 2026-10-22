"""
Keep the tests independent of whatever config.env and the last run hold.

Every DOCUMATE_ setting is cleared and config.env is not read, so each test starts
from nothing and sets exactly the settings it is about with
monkeypatch.setenv. The last-run memory is kept in the test's own folder.
"""

import os

import pytest

from setup import config
from setup import picker

REAL_LOAD_ENVIRONMENT = config.load_environment


@pytest.fixture(autouse=True)
def no_settings(monkeypatch, tmp_path):
    for name in list(os.environ):
        if name.startswith("DOCUMATE_"):
            monkeypatch.delenv(name)
    # A popup would wait for a click that never comes.
    monkeypatch.setenv("DOCUMATE_POPUPS", "0")
    monkeypatch.setattr(config, "load_environment", lambda: None)
    monkeypatch.setattr(picker, "memory_path", lambda: str(tmp_path / "last_run.json"))


@pytest.fixture
def read_config_env(monkeypatch):
    """Read config.env for real - the loader captured before it was replaced."""
    monkeypatch.setattr(config, "load_environment", REAL_LOAD_ENVIRONMENT)


@pytest.fixture
def birth_register(monkeypatch):
    """The birth register, as the picker sets it up on recognising one."""
    monkeypatch.setenv("DOCUMATE_DATA_FORMAT", "birth_register")
    monkeypatch.setenv("DOCUMATE_OUTPUT_NAME", "BIRTH_REGISTRATION")
