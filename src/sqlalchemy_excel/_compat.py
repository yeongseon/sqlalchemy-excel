"""Version compatibility and optional dependency helpers."""

from __future__ import annotations

import importlib
from typing import Any


def ensure_defusedxml() -> None:
    """Ensure defusedxml is installed for safe XML processing.

    Raises:
        ImportError: If defusedxml is not installed.
    """
    try:
        import defusedxml  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "defusedxml is required for processing Excel files safely. "
            "Install it with: pip install defusedxml"
        ) from e


def import_optional(
    module_name: str,
    extra_name: str,
) -> Any:
    """Import an optional dependency, raising a helpful error if missing.

    Args:
        module_name: The module to import (e.g., "pandas").
        extra_name: The pip extra name (e.g., "pandas").

    Returns:
        The imported module.

    Raises:
        ImportError: With installation instructions if module is not found.
    """
    try:
        return importlib.import_module(module_name)
    except ImportError as e:
        raise ImportError(
            f"{module_name} is required for this feature. "
            f"Install it with: pip install sqlalchemy-excel[{extra_name}]"
        ) from e
