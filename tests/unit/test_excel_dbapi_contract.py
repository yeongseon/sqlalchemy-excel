from __future__ import annotations

from importlib.metadata import version

from packaging.version import Version


def test_excel_dbapi_version_within_supported_major_range() -> None:
    installed = Version(version("excel-dbapi"))

    assert installed >= Version("1.0")
    assert installed < Version("2.0")
