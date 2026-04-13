#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PYPROJECT = ROOT / "pyproject.toml"
README = ROOT / "README.md"
COMPAT = ROOT / "docs" / "COMPATIBILITY.md"


def extract_excel_dbapi_spec(dependencies: list[str]) -> str | None:
    for dep in dependencies:
        if dep.startswith("excel-dbapi"):
            spec = dep.removeprefix("excel-dbapi")
            return spec if spec else None
    return None


def extract_project_section(pyproject_text: str) -> str:
    match = re.search(
        r"^\[project\]\n(?P<section>.*?)(?:\n\[|\Z)",
        pyproject_text,
        flags=re.MULTILINE | re.DOTALL,
    )
    if not match:
        return ""
    return match.group("section")


def extract_string_field(section: str, key: str) -> str | None:
    match = re.search(rf'^\s*{re.escape(key)}\s*=\s*"([^"]+)"\s*$', section, re.MULTILINE)
    if not match:
        return None
    return match.group(1)


def extract_dependencies(section: str) -> list[str]:
    dep_block = re.search(
        r"^\s*dependencies\s*=\s*\[(?P<deps>.*?)\]\s*$",
        section,
        flags=re.MULTILINE | re.DOTALL,
    )
    if not dep_block:
        return []
    deps_body = dep_block.group("deps")
    return re.findall(r'"([^"]+)"', deps_body)


def main() -> int:
    errors: list[str] = []

    pyproject_text = PYPROJECT.read_text(encoding="utf-8")
    project_section = extract_project_section(pyproject_text)
    current_version = extract_string_field(project_section, "version")
    requires_python = extract_string_field(project_section, "requires-python")
    dependencies = extract_dependencies(project_section)

    if not current_version:
        errors.append("Could not parse project version from pyproject.toml")
        current_version = ""
    if not requires_python:
        errors.append("Could not parse requires-python from pyproject.toml")
        requires_python = ""
    if not dependencies:
        errors.append("Could not parse project.dependencies from pyproject.toml")

    sa_spec = None
    for dep in dependencies:
        if dep.startswith("sqlalchemy"):
            sa_spec = dep.removeprefix("sqlalchemy")
            break

    if not sa_spec:
        errors.append("Could not parse sqlalchemy dependency spec from pyproject.toml")
        sa_spec = ""

    excel_dbapi_spec = extract_excel_dbapi_spec(dependencies)
    if not excel_dbapi_spec:
        errors.append("Could not parse excel-dbapi dependency spec from pyproject.toml")
        excel_dbapi_spec = ""

    readme_text = README.read_text(encoding="utf-8")
    compat_text = COMPAT.read_text(encoding="utf-8")

    if current_version and f"Current release: `{current_version}`" not in readme_text:
        errors.append(
            f"README.md must contain exact marker `Current release: `{current_version}``"
        )

    if current_version and f"`{current_version}` (current)" not in compat_text:
        errors.append(
            f"docs/COMPATIBILITY.md matrix must include current row marker `{current_version}` (current)"
        )

    if current_version and f"`sqlalchemy-excel` version: `{current_version}`" not in compat_text:
        errors.append(
            "docs/COMPATIBILITY.md baseline must include current sqlalchemy-excel version"
        )

    if excel_dbapi_spec and f"`excel-dbapi` requirement: `{excel_dbapi_spec}`" not in compat_text:
        errors.append(
            "docs/COMPATIBILITY.md baseline must include pyproject excel-dbapi requirement"
        )

    if sa_spec and f"SQLAlchemy requirement: `{sa_spec}`" not in compat_text:
        errors.append(
            "docs/COMPATIBILITY.md baseline must include pyproject SQLAlchemy requirement"
        )

    if requires_python and f"Python requirement: `{requires_python}`" not in compat_text:
        errors.append(
            "docs/COMPATIBILITY.md baseline must include pyproject Python requirement"
        )

    if current_version:
        version_mentions = re.findall(r"\b0\.\d+\.\d+\b", compat_text)
        if current_version not in version_mentions:
            errors.append(
                "docs/COMPATIBILITY.md does not mention current pyproject version"
            )

    if errors:
        print("Version/doc sync check failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print("Version/doc sync check passed.")
    print(f"- sqlalchemy-excel version: {current_version}")
    print(f"- excel-dbapi requirement: {excel_dbapi_spec}")
    print(f"- SQLAlchemy requirement: {sa_spec}")
    print(f"- Python requirement: {requires_python}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
