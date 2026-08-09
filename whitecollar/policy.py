from __future__ import annotations

from dataclasses import dataclass

from .authority import Authority
from .errors import PolicyError
from .mail_ops import MAIL_COM_MUTATING_OPERATIONS
from .models import POLICY_NAMES, Plan
from .permissions import (
    PROFILE_CAPABILITIES,
    capability_for_operation,
    require_capability,
)
from .slides_ops import SLIDES_COM_MUTATING_OPERATIONS, is_read_operation as is_slides_read_operation
from .word_ops import WORD_COM_MUTATING_OPERATIONS, is_read_operation as is_word_read_operation


@dataclass(frozen=True)
class PolicyProfile:
    name: str
    allow_read: bool
    allow_dry_run: bool
    allow_save_as: bool
    allow_in_place: bool
    capabilities: frozenset[str]


PROFILES = {
    "read-only": PolicyProfile("read-only", True, False, False, False, PROFILE_CAPABILITIES["read-only"]),
    "review": PolicyProfile("review", True, True, True, False, PROFILE_CAPABILITIES["review"]),
    "edit": PolicyProfile("edit", True, True, True, True, PROFILE_CAPABILITIES["edit"]),
}


def require_read(policy: str) -> PolicyProfile:
    if policy not in POLICY_NAMES:
        raise PolicyError(f"unknown policy profile: {policy}")
    profile = PROFILES[policy]
    if not profile.allow_read:
        raise PolicyError(f"policy {policy!r} does not allow reads")
    return profile


def authorize_plan(
    plan: Plan,
    *,
    dry_run: bool,
    authority: Authority | None = None,
    backend: str = "local",
) -> PolicyProfile:
    profile = PROFILES[plan.policy]
    operations = {operation["op"] for operation in plan.operations}
    if plan.app == "mail":
        if plan.write.mode != "none":
            raise PolicyError("mail mutation plans must use write.mode 'none'")
        if not operations or not operations.issubset(MAIL_COM_MUTATING_OPERATIONS):
            raise PolicyError("mail plan contains an unsupported operation")
        target = plan.target.id
        required_capabilities = tuple(
            capability_for_operation(plan.app, operation)
            for operation in operations
        )
        if dry_run and not profile.allow_dry_run:
            raise PolicyError(f"policy {plan.policy!r} does not allow mutation plans, including dry-runs")
        for capability in required_capabilities:
            require_capability(plan.policy, capability, target=target)
        if authority is not None:
            authority.require_access(plan.app, backend, plan.policy, target, required_capabilities)
        return profile
    if plan.app == "word":
        mutating_operations = WORD_COM_MUTATING_OPERATIONS
        read_operation = is_word_read_operation
        app_name = "Word"
    else:
        mutating_operations = SLIDES_COM_MUTATING_OPERATIONS
        read_operation = is_slides_read_operation
        app_name = "PowerPoint"
    is_mutation = any(operation in mutating_operations or operation == "replace_text" for operation in operations)
    if not is_mutation and plan.write.mode != "none":
        raise PolicyError(f"read-only {app_name} operations must use write.mode 'none'")
    if not is_mutation and not all(read_operation(operation) or operation == "replace_text" for operation in operations):
        raise PolicyError("plan contains an unknown operation capability")
    required_capabilities = {
        capability_for_operation(plan.app, operation)
        for operation in operations
    }
    write_capability = (
        f"{plan.app}.write.{plan.write.mode.replace('-', '_')}"
        if plan.write.mode != "none"
        else None
    )
    if not is_mutation:
        if dry_run:
            raise PolicyError("read operations do not need --dry-run")
        for operation in operations:
            require_capability(plan.policy, capability_for_operation(plan.app, operation), target=plan.target.path)
        if authority is not None:
            authority.require_access(
                plan.app,
                backend,
                plan.policy,
                plan.target.path,
                tuple(required_capabilities),
            )
        return profile
    if dry_run:
        if not profile.allow_dry_run:
            raise PolicyError(f"policy {plan.policy!r} does not allow mutation plans, including dry-runs")
        for operation in operations:
            require_capability(plan.policy, capability_for_operation(plan.app, operation), target=plan.target.path)
        if write_capability is not None:
            require_capability(plan.policy, write_capability, target=plan.target.path)
            required_capabilities.add(write_capability)
        if authority is not None:
            authority.require_access(
                plan.app,
                backend,
                plan.policy,
                plan.target.path,
                tuple(required_capabilities),
            )
            output_target = plan.write.path or plan.write.snapshot
            if write_capability is not None and output_target is not None:
                authority.require_access(
                    plan.app,
                    backend,
                    plan.policy,
                    output_target,
                    (write_capability,),
                )
        return profile
    if plan.write.mode == "none":
        raise PolicyError(f"mutating {app_name} operations require save-as or in-place write intent")
    if plan.write.mode == "save-as" and not profile.allow_save_as:
        raise PolicyError(f"policy {plan.policy!r} does not allow save-as writes")
    if plan.write.mode == "in-place" and not profile.allow_in_place:
        raise PolicyError(f"policy {plan.policy!r} does not allow in-place writes")
    for operation in operations:
        require_capability(plan.policy, capability_for_operation(plan.app, operation), target=plan.target.path)
    assert write_capability is not None
    require_capability(plan.policy, write_capability, target=plan.target.path)
    required_capabilities.add(write_capability)
    if authority is not None:
        authority.require_access(
            plan.app,
            backend,
            plan.policy,
            plan.target.path,
            tuple(required_capabilities),
        )
        output_target = plan.write.path or plan.write.snapshot
        if output_target is not None:
            authority.require_access(
                plan.app,
                backend,
                plan.policy,
                output_target,
                (write_capability,),
            )
    return profile
