from __future__ import annotations

from dataclasses import dataclass

from .errors import PolicyError
from .models import POLICY_NAMES, Plan
from .word_ops import WORD_COM_MUTATING_OPERATIONS, is_read_operation


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
    operations = {operation["op"] for operation in plan.operations}
    is_mutation = any(operation in WORD_COM_MUTATING_OPERATIONS or operation == "replace_text" for operation in operations)
    if not is_mutation and plan.write.mode != "none":
        raise PolicyError("read-only Word operations must use write.mode 'none'")
    if not is_mutation and not all(is_read_operation(operation) or operation == "replace_text" for operation in operations):
        raise PolicyError("plan contains an unknown operation capability")
    if not is_mutation:
        if dry_run:
            raise PolicyError("read operations do not need --dry-run")
        return profile
    if dry_run:
        if not profile.allow_dry_run:
            raise PolicyError(f"policy {plan.policy!r} does not allow mutation plans, including dry-runs")
        return profile
    if plan.write.mode == "none":
        raise PolicyError("mutating Word operations require save-as or in-place write intent")
    if plan.write.mode == "save-as" and not profile.allow_save_as:
        raise PolicyError(f"policy {plan.policy!r} does not allow save-as writes")
    if plan.write.mode == "in-place" and not profile.allow_in_place:
        raise PolicyError(f"policy {plan.policy!r} does not allow in-place writes")
    return profile
