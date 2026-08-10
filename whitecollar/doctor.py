"""Side-effect-free environment and permission diagnostics."""

from __future__ import annotations

import importlib.util
import platform
import sys
from collections.abc import Callable
from typing import Any

from .authority import Authority


def _module_available(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ModuleNotFoundError):
        # Some optional packages (notably pywin32) are only importable on
        # Windows.  Diagnostics must report that absence rather than fail.
        return False


def diagnose(
    authority: Authority,
    *,
    module_available: Callable[[str], bool] = _module_available,
) -> dict[str, Any]:
    """Return diagnostics without starting Office or reading user documents."""

    pywin32 = module_available("win32com.client")
    pillow = module_available("PIL")
    is_windows = platform.system() == "Windows"
    owner_grants = authority.owner_grants
    built_in_grants = [grant for grant in authority.grants if grant not in owner_grants]

    def com_status(app: str) -> dict[str, str]:
        if not is_windows:
            return {"status": "unavailable", "reason": "requires Windows"}
        if not pywin32:
            return {"status": "unavailable", "reason": "install the office extra for pywin32"}
        return {
            "status": "dependency-ready",
            "permission": "enabled" if any(grant.app == app and grant.backend == "com" for grant in authority.grants) else "disabled",
            "probe": "not-run",
        }

    return {
        "environment": {
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "windows": is_windows,
            "pywin32": pywin32,
            "pillow": pillow,
        },
        "backends": {
            "word": {
                "com": com_status("word"),
            },
            "slides": {
                "com": com_status("slides"),
            },
            "mail": {
                "com": com_status("mail"),
            },
        },
        "permissions": {
            "store": authority.source,
            "built_in_grants": len(built_in_grants),
            "owner_grants": len(owner_grants),
            "owner_apps": sorted({grant.app for grant in owner_grants}),
            "owner_policies": sorted({grant.policy for grant in owner_grants}),
            "targets": "redacted; use permissions show for explicit owner-grant details",
        },
        "safety": {
            "office_probe": "not-run",
            "mail_com_default": "disabled unless an owner grant enables it",
        },
    }
