from __future__ import annotations

from dataclasses import dataclass

from .errors import PolicyError
from .models import POLICY_NAMES, Plan


@dataclass(frozen=True)
class PolicyProfile:
    name: str
    allow_read: bool
    allow_dry_run: bool
    allow_save_as: bool
    allow_in_place: bool


PROFILES = {
    "read-only": PolicyProfile("read-only", True, False, False, False),
    "review": PolicyProfile("review", True, True, True, False),
    "edit": PolicyProfile("edit", True, True, True, True),
}


def require_read(policy: str) -> PolicyProfile:
    if policy not in POLICY_NAMES:
        raise PolicyError(f"unknown policy profile: {policy}")
    profile = PROFILES[policy]
    if not profile.allow_read:
        raise PolicyError(f"policy {policy!r} does not allow reads")
    return profile


def authorize_plan(plan: Plan, *, dry_run: bool) -> PolicyProfile:
    profile = PROFILES[plan.policy]
    if dry_run:
        if not profile.allow_dry_run:
            raise PolicyError(f"policy {plan.policy!r} does not allow mutation plans, including dry-runs")
        return profile
    if plan.write.mode == "save-as" and not profile.allow_save_as:
        raise PolicyError(f"policy {plan.policy!r} does not allow save-as writes")
    if plan.write.mode == "in-place" and not profile.allow_in_place:
        raise PolicyError(f"policy {plan.policy!r} does not allow in-place writes")
    return profile
