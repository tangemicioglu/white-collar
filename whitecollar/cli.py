from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from . import __version__
from .authority import (
    HUMAN_PERMISSION_NOTICE,
    Authority,
    load_authority,
    make_grant,
    replace_app_grants,
    revoke_all_grants,
    revoke_grant,
    save_grant,
)
from .engine import RuntimeAdapters, apply_plan, inspect_document, read_mail, search_mail
from .doctor import diagnose
from .completion import completion_script
from .errors import ValidationError, WhiteCollarError
from .models import Plan, result
from .permissions import (
    CAPABILITIES,
    PROFILE_NAMES,
    SETUP_APP_POLICIES,
    SETUP_POLICY_NAMES,
    SETUP_PRESETS,
    catalog,
    require_capability,
    setup_capabilities,
)


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValidationError(message)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        prog="white-collar",
        description="Narrow, safety-aware Windows Office COM control plane",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  white-collar doctor\n"
            "  white-collar word replace --target C:\\work\\brief.docx --find Draft --replace Final --output C:\\work\\reviewed.docx\n"
            "  white-collar mail draft --to person@example.com --subject Review --body \"Draft body\"\n"
            "  white-collar setup --preset safe\n"
            "  white-collar completions powershell | Out-String | Invoke-Expression"
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    apps = parser.add_subparsers(dest="app", required=True, parser_class=JsonArgumentParser)

    setup = apps.add_parser("setup", help="human-owner-only application permission setup")
    setup.add_argument("--app", dest="setup_app", choices=("word", "slides", "mail"))
    setup.add_argument("--policy", choices=SETUP_POLICY_NAMES)
    setup.add_argument("--preset", choices=tuple(SETUP_PRESETS), help="apply a named application permission bundle")
    setup.add_argument("--json", action="store_true", help="emit the machine-readable result instead of human setup output")
    apps.add_parser("doctor", help="diagnose Windows Office COM dependencies and permission readiness")

    for app in ("word", "slides"):
        app_parser = apps.add_parser(app, help=f"inspect and edit {app} documents")
        commands = app_parser.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
        create = commands.add_parser("create", help=f"create a blank {app} file through {app.title()} COM")
        create.add_argument("--output", required=True, help=f"absolute .{'docx' if app == 'word' else 'pptx'} output path")
        create.add_argument("--policy", choices=("review", "edit"), default="review")
        create.add_argument("--dry-run", action="store_true")
        create.add_argument("--backend", choices=("com",), default="com")
        inspect = commands.add_parser("inspect")
        inspect.add_argument("target")
        inspect.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
        inspect.add_argument("--backend", choices=("com",), default="com")
        inspect.add_argument("--render-dir", help="write one native Office PNG per page or slide")
        apply = commands.add_parser("apply")
        apply.add_argument("--plan", required=True)
        apply.add_argument("--dry-run", action="store_true")
        apply.add_argument("--backend", choices=("com",), default="com")
        replace = commands.add_parser("replace", help="bounded text replacement compiled into a validated plan")
        replace.add_argument("--target", required=True, help="absolute Word document path")
        replace.add_argument("--find", required=True)
        replace.add_argument("--replace", required=True, dest="replacement")
        replace.add_argument("--occurrence", choices=("all", "first"), default="all")
        replace.add_argument("--output", help="distinct save-as output path")
        replace.add_argument("--in-place", action="store_true", help="replace the target after creating --snapshot")
        replace.add_argument("--snapshot", help="required backup path with --in-place")
        replace.add_argument("--policy", choices=("review", "edit"))
        replace.add_argument("--dry-run", action="store_true")
        replace.add_argument("--backend", choices=("com",), default="com")

    mail = apps.add_parser("mail", help="search, read, draft, and send Outlook mail")
    mail_commands = mail.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
    search = mail_commands.add_parser("search")
    search.add_argument("--query", required=True)
    search.add_argument("--limit", type=int, default=20)
    search.add_argument("--folder", default="Inbox", help="Inbox, Sent Items, Drafts, or another supported default folder")
    search.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
    search.add_argument("--backend", choices=("com",), default="com")
    read = mail_commands.add_parser("read")
    read.add_argument("--id", required=True, dest="message_id")
    read.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
    read.add_argument("--include-body", action="store_true", help="request sensitive message body access")
    read.add_argument("--backend", choices=("com",), default="com")
    mail_apply = mail_commands.add_parser("apply")
    mail_apply.add_argument("--plan", required=True)
    mail_apply.add_argument("--dry-run", action="store_true")
    mail_apply.add_argument("--backend", choices=("com",), default="com")
    draft = mail_commands.add_parser("draft", help="create a bounded Outlook draft")
    draft.add_argument("--account", default="mailbox", help="Outlook account address or mailbox")
    draft.add_argument("--to", required=True)
    draft.add_argument("--cc")
    draft.add_argument("--bcc")
    draft.add_argument("--subject", required=True)
    draft.add_argument("--body", required=True)
    draft.add_argument("--dry-run", action="store_true")
    draft.add_argument("--backend", choices=("com",), default="com")
    send = mail_commands.add_parser("send", help="send one existing Outlook draft")
    send.add_argument("--draft-id", required=True)
    send.add_argument("--dry-run", action="store_true")
    send.add_argument("--backend", choices=("com",), default="com")

    permissions = apps.add_parser("permissions", help="inspect and check capability grants")
    permission_commands = permissions.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
    show = permission_commands.add_parser("show")
    show.add_argument("--policy", choices=PROFILE_NAMES, default="read-only")
    show.add_argument("--redacted", action="store_true", help="hide exact owner-grant targets in the status output")
    check = permission_commands.add_parser("check")
    check.add_argument("--capability", required=True)
    check.add_argument("--target")
    check.add_argument("--policy", choices=PROFILE_NAMES, default="read-only")
    check.add_argument("--backend", choices=("com",), default="com")
    grant = permission_commands.add_parser("grant", help="human-owner-only; store a narrowly scoped grant")
    grant.add_argument("--app", dest="grant_app", choices=("word", "slides", "mail"), required=True)
    grant.add_argument("--backend", choices=("com",), required=True)
    grant.add_argument("--policy", choices=PROFILE_NAMES, required=True)
    grant.add_argument("--target", action="append", required=True, help="exact file, message id, or 'mailbox'; repeat for multiple targets")
    grant.add_argument("--capability", action="append", help="narrow the grant; repeat for multiple capabilities")
    grant.add_argument("--json", action="store_true", help="emit the machine-readable result instead of the human confirmation output")
    revoke = permission_commands.add_parser("revoke", help="human-owner-only; revoke a narrowly scoped grant")
    revoke.add_argument("--app", dest="grant_app", choices=("word", "slides", "mail"))
    revoke.add_argument("--backend", choices=("com",))
    revoke.add_argument("--policy", choices=PROFILE_NAMES)
    revoke.add_argument("--target", action="append", help="exact target; repeat for multiple targets")
    revoke.add_argument("--capability", action="append", help="identify the grant; repeat for multiple capabilities")
    revoke.add_argument("--all", action="store_true", help="revoke all owner grants; human confirmation is still required")
    revoke.add_argument("--json", action="store_true", help="emit the machine-readable result instead of the human confirmation output")
    completions = apps.add_parser("completions", help="print shell completion code")
    completions.add_argument("shell", choices=("powershell",), help="shell to generate completion code for")
    return parser


def _load_plan(path: str) -> Plan:
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except OSError as exc:
        raise ValidationError("cannot read plan", details={"path": path, "reason": str(exc)}) from exc
    except json.JSONDecodeError as exc:
        raise ValidationError("plan is not valid JSON", details={"line": exc.lineno, "column": exc.colno}) from exc
    return Plan.from_dict(raw)


def _command_name(args: argparse.Namespace | None) -> str:
    if args is None:
        return "cli"
    return ".".join(part for part in (getattr(args, "app", None), getattr(args, "action", None)) if part) or "cli"


def _grant_summary(grant: Any) -> str:
    value = grant.to_dict()
    return "\n".join(
        (
            "Permission change:",
            f"  application: {value['app']}",
            f"  backend: {value['backend']}",
            f"  policy: {value['policy']}",
            f"  capabilities: {', '.join(value['capabilities'])}",
            f"  targets: {', '.join(value['targets'])}",
        )
    )


def _human_confirmation(*, action: str, summary: str) -> None:
    """Require a real interactive terminal before changing owner grants."""

    if not sys.stdin.isatty() or not sys.stderr.isatty():
        raise ValidationError(
            "owner permission changes require an interactive human terminal",
            details={
                "human_action_required": True,
                "agent_instruction": HUMAN_PERMISSION_NOTICE,
                "requested_action": action,
                "requested_grant": summary,
            },
        )
    print("WHITE-COLLAR OWNER PERMISSION CHANGE", file=sys.stderr)
    print(HUMAN_PERMISSION_NOTICE, file=sys.stderr)
    print(summary, file=sys.stderr)
    print(f"{action.capitalize()} this permission? [y/N] ", end="", file=sys.stderr)
    sys.stderr.flush()
    answer = sys.stdin.readline().strip().casefold()
    if answer not in {"y", "yes"}:
        raise ValidationError(
            "human confirmation was not received; no permission was changed",
            details={
                "human_action_required": True,
                "agent_instruction": HUMAN_PERMISSION_NOTICE,
                "requested_action": action,
            },
        )


def _human_permission_output(args: argparse.Namespace | None) -> bool:
    return bool(
        args
        and ((args.app == "permissions" and args.action in {"grant", "revoke"}) or args.app == "setup")
        and not getattr(args, "json", False)
        and sys.stdin.isatty()
        and sys.stderr.isatty()
    )


def _print_human_permission_result(response: dict[str, Any], *, action: str) -> None:
    if response.get("ok"):
        if action == "setup":
            print("Application permissions updated.")
            print("The settings are stored in protected OS credential storage.")
        else:
            verb = {"grant": "granted", "revoke": "revoked"}[action]
            print(f"Permission {verb}.")
            if action == "grant":
                print("The grant is stored in protected OS credential storage.")
            else:
                print("The grant was removed from protected OS credential storage.")
        return
    error = response.get("error", {})
    print(f"white-collar: {error.get('message', 'permission change failed')}", file=sys.stderr)


def _collect_setup_selections(args: argparse.Namespace) -> dict[str, str]:
    if not sys.stdin.isatty() or not sys.stderr.isatty():
        raise ValidationError(
            "application permission setup requires an interactive human terminal",
            details={
                "human_action_required": True,
                "agent_instruction": HUMAN_PERMISSION_NOTICE,
                "hint": "ask the human owner to run 'white-collar setup' in their own terminal",
            },
        )
    if args.preset is not None:
        if args.setup_app is not None or args.policy is not None:
            raise ValidationError("--preset cannot be combined with --app or --policy")
        return dict(SETUP_PRESETS[args.preset])
    if args.setup_app is None and args.policy is not None:
        raise ValidationError("--policy requires --app")
    applications = (args.setup_app,) if args.setup_app else tuple(SETUP_APP_POLICIES)
    selections: dict[str, str] = {}
    if args.setup_app is None:
        print(
            "White-collar setup: choose the highest permission each application may use.\n"
            "Word and PowerPoint review are safe save-as defaults; Outlook is disabled by default.\n",
            file=sys.stderr,
        )
    for app in applications:
        if args.setup_app == app and args.policy is not None:
            policy = args.policy
        else:
            allowed = SETUP_APP_POLICIES[app]
            default = "disabled" if app == "mail" else "review"
            print(
                f"{app} permission [{default}] ({', '.join(allowed)}): ",
                end="",
                file=sys.stderr,
            )
            sys.stderr.flush()
            policy = sys.stdin.readline().strip().casefold() or default
            if policy not in allowed:
                raise ValidationError(
                    "unsupported setup policy for application",
                    details={"app": app, "policy": policy, "allowed": list(allowed)},
                )
        if policy not in SETUP_APP_POLICIES[app]:
            raise ValidationError(
                "unsupported setup policy for application",
                details={"app": app, "policy": policy, "allowed": list(SETUP_APP_POLICIES[app])},
            )
        selections[app] = policy
    return selections


def _setup_grants(selections: dict[str, str]) -> tuple[dict[str, tuple[Any, ...]], str]:
    labels = {"word": "Word", "slides": "PowerPoint", "mail": "Outlook"}
    grants_by_app: dict[str, tuple[Any, ...]] = {}
    summary_lines = ["Application permission setup:"]
    for app, policy in selections.items():
        capabilities = setup_capabilities(app, policy)
        if policy == "disabled":
            grants_by_app[app] = ()
            suffix = "; built-in Word/PowerPoint review remains" if app in {"word", "slides"} else ""
            summary_lines.append(f"  {labels[app]}: disabled (owner grants removed{suffix})")
            continue
        backends = ("com",)
        target = "*" if app in {"word", "slides"} else "mailbox"
        grants_by_app[app] = tuple(
            make_grant(
                app=app,
                backend=backend,
                policy=policy,
                targets=[target],
                capabilities=list(capabilities),
            )
            for backend in backends
        )
        scope = "app-wide" if app in {"word", "slides"} else "mailbox-wide"
        warning = "; includes sending any existing draft" if app == "mail" and policy == "send" else ""
        summary_lines.append(f"  {labels[app]}: {policy} ({scope}{warning})")
    return grants_by_app, "\n".join(summary_lines)


def _shortcut_result(
    plan: Plan,
    *,
    command: str,
    dry_run: bool,
    adapters: RuntimeAdapters,
    authority: Authority,
    backend: str,
) -> dict[str, Any]:
    data = apply_plan(plan, dry_run=dry_run, adapters=adapters, authority=authority, backend=backend)
    changes = data.pop("changes", [])
    target = plan.target.path if hasattr(plan.target, "path") else plan.target.id
    return result(
        ok=True,
        command=command,
        policy=plan.policy,
        dry_run=dry_run,
        target=target,
        data=data,
        changes=changes,
    )


def _run(
    args: argparse.Namespace,
    adapters: RuntimeAdapters,
    authority: Authority,
    grant_store: Any | None = None,
) -> dict[str, Any]:
    command = _command_name(args)
    if args.app == "doctor":
        return result(ok=True, command=command, policy="read-only", dry_run=False, data=diagnose(authority))
    if args.app == "setup":
        selections = _collect_setup_selections(args)
        grants_by_app, summary = _setup_grants(selections)
        _human_confirmation(action="setup", summary=summary)
        updated = replace_app_grants(authority, grants_by_app, store=grant_store)
        return result(
            ok=True,
            command=command,
            policy="setup",
            dry_run=False,
            data={"selections": selections, "authority": updated.to_dict()},
        )
    if args.app in {"word", "slides"} and args.action == "create":
        target = Path(args.output).resolve()
        operation = "word_live_create_document" if args.app == "word" else "slides_live_create_presentation"
        plan = Plan.from_dict(
            {
                "schema": "white-collar.plan/v1",
                "app": args.app,
                "target": {"path": str(target)},
                "policy": args.policy,
                "operations": [{"op": operation}],
                "write": {"mode": "create"},
            }
        )
        return _shortcut_result(
            plan,
            command=command,
            dry_run=args.dry_run,
            adapters=adapters,
            authority=authority,
            backend=args.backend,
        )
    if args.app == "word" and args.action == "replace":
        target = Path(args.target).resolve()
        if args.in_place:
            if args.output:
                raise ValidationError("--in-place cannot be combined with --output")
            if not args.snapshot:
                raise ValidationError("--in-place requires --snapshot")
            policy = args.policy or "edit"
            write = {"mode": "in-place", "snapshot": str(Path(args.snapshot).resolve())}
        else:
            if not args.output:
                raise ValidationError("word replace requires --output or --in-place")
            if args.snapshot:
                raise ValidationError("--snapshot requires --in-place")
            policy = args.policy or "review"
            write = {"mode": "save-as", "path": str(Path(args.output).resolve())}
        args.policy = policy
        plan = Plan.from_dict(
            {
                "schema": "white-collar.plan/v1",
                "app": "word",
                "target": {"path": str(target)},
                "policy": policy,
                "operations": [{
                    "op": "replace_text",
                    "find": args.find,
                    "replace": args.replacement,
                    "occurrence": args.occurrence,
                }],
                "write": write,
            }
        )
        return _shortcut_result(
            plan,
            command=command,
            dry_run=args.dry_run,
            adapters=adapters,
            authority=authority,
            backend=args.backend,
        )
    if args.app == "mail" and args.action == "draft":
        args.policy = "edit"
        draft_args = {"to": args.to, "subject": args.subject, "body": args.body}
        if args.cc is not None:
            draft_args["cc"] = args.cc
        if args.bcc is not None:
            draft_args["bcc"] = args.bcc
        plan = Plan.from_dict(
            {
                "schema": "white-collar.plan/v1",
                "app": "mail",
                "target": {"id": args.account},
                "policy": "edit",
                "operations": [{"op": "mail_live_create_draft", "args": draft_args}],
                "write": {"mode": "none"},
            }
        )
        return _shortcut_result(
            plan,
            command=command,
            dry_run=args.dry_run,
            adapters=adapters,
            authority=authority,
            backend=args.backend,
        )
    if args.app == "mail" and args.action == "send":
        args.policy = "send"
        plan = Plan.from_dict(
            {
                "schema": "white-collar.plan/v1",
                "app": "mail",
                "target": {"id": args.draft_id},
                "policy": "send",
                "operations": [{"op": "mail_live_send"}],
                "write": {"mode": "none"},
            }
        )
        return _shortcut_result(
            plan,
            command=command,
            dry_run=args.dry_run,
            adapters=adapters,
            authority=authority,
            backend=args.backend,
        )
    if args.app == "permissions" and args.action == "show":
        data = catalog(policy=args.policy, authority=authority)
        data["authority"] = authority.to_dict(redact_targets=args.redacted)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=data)
    if args.app == "permissions" and args.action == "check":
        decision = require_capability(args.policy, args.capability, target=args.target)
        capability = CAPABILITIES[args.capability]
        authority.require_access(
            capability.app,
            args.backend,
            args.policy,
            args.target if args.target is not None else "mailbox",
            (args.capability,),
        )
        return result(ok=True, command=command, policy=args.policy, dry_run=False, target=args.target, data=decision)
    if args.app == "permissions" and args.action == "grant":
        targets = list(args.target)
        grant = make_grant(
            app=args.grant_app,
            backend=args.backend,
            policy=args.policy,
            targets=targets,
            capabilities=args.capability,
        )
        _human_confirmation(action="grant", summary=_grant_summary(grant))
        updated = save_grant(authority, grant, store=grant_store)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=updated.to_dict())
    if args.app == "permissions" and args.action == "revoke":
        if args.all:
            summary = "Revoke every owner grant currently stored by white-collar."
            _human_confirmation(action="revoke all grants", summary=summary)
            updated = revoke_all_grants(authority, store=grant_store)
        else:
            if not args.grant_app or not args.backend or not args.policy or not args.target:
                raise ValidationError("revoke requires --app, --backend, --policy, and --target unless --all is used")
            grant = make_grant(
                app=args.grant_app,
                backend=args.backend,
                policy=args.policy,
                targets=list(args.target),
                capabilities=args.capability,
            )
            _human_confirmation(action="revoke", summary=_grant_summary(grant))
            updated = revoke_grant(authority, grant, store=grant_store)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=updated.to_dict())
    if args.app in {"word", "slides"} and args.action == "inspect":
        target = Path(args.target).resolve()
        render_dir = Path(args.render_dir).resolve() if getattr(args, "render_dir", None) else None
        data = inspect_document(
            args.app,
            target,
            args.policy,
            adapters,
            render_dir=render_dir,
            backend=args.backend,
            authority=authority,
        )
        return result(ok=True, command=command, policy=args.policy, dry_run=False, target=str(target), data=data)
    if args.app in {"word", "slides"} and args.action == "apply":
        plan = _load_plan(args.plan)
        args.policy = plan.policy
        if plan.app != args.app:
            raise ValidationError(
                "plan app does not match command",
                details={"plan_app": plan.app, "command_app": args.app},
            )
        data = apply_plan(plan, dry_run=args.dry_run, adapters=adapters, authority=authority, backend=args.backend)
        changes = data.pop("changes", [])
        return result(
            ok=True,
            command=command,
            policy=plan.policy,
            dry_run=args.dry_run,
            target=plan.target.path,
            data=data,
            changes=changes,
        )
    if args.app == "mail" and args.action == "apply":
        plan = _load_plan(args.plan)
        args.policy = plan.policy
        if plan.app != "mail":
            raise ValidationError(
                "plan app does not match command",
                details={"plan_app": plan.app, "command_app": args.app},
            )
        data = apply_plan(plan, dry_run=args.dry_run, adapters=adapters, authority=authority, backend=args.backend)
        changes = data.pop("changes", [])
        return result(
            ok=True,
            command=command,
            policy=plan.policy,
            dry_run=args.dry_run,
            target=plan.target.id,
            data=data,
            changes=changes,
        )
    if args.app == "mail" and args.action == "search":
        data = search_mail(
            args.query,
            limit=args.limit,
            policy=args.policy,
            folder=args.folder,
            backend=args.backend,
            adapters=adapters,
            authority=authority,
        )
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=data)
    if args.app == "mail" and args.action == "read":
        data = read_mail(
            args.message_id,
            policy=args.policy,
            adapters=adapters,
            include_body=args.include_body,
            backend=args.backend,
            authority=authority,
        )
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=data)
    raise ValidationError("unsupported command")


def main(
    argv: Sequence[str] | None = None,
    *,
    adapters: RuntimeAdapters | None = None,
    authority: Authority | None = None,
    grant_store: Any | None = None,
) -> int:
    parsed: argparse.Namespace | None = None
    try:
        parsed = build_parser().parse_args(argv)
        if parsed.app == "completions":
            print(completion_script(parsed.shell), end="")
            return 0
        active_authority = authority or load_authority(store=grant_store)
        response = _run(
            parsed,
            adapters
            or RuntimeAdapters.live(),
            active_authority,
            grant_store,
        )
        exit_code = 0
    except WhiteCollarError as exc:
        policy = getattr(parsed, "policy", "read-only") if parsed else "read-only"
        if parsed is not None and parsed.app == "setup":
            policy = "setup"
        response = result(
            ok=False,
            command=_command_name(parsed),
            policy=policy,
            dry_run=bool(getattr(parsed, "dry_run", False)) if parsed else False,
            error={"code": exc.code, "message": exc.message, "details": exc.details},
        )
        exit_code = 2
    except OSError as exc:
        policy = getattr(parsed, "policy", "read-only") if parsed else "read-only"
        if parsed is not None and parsed.app == "setup":
            policy = "setup"
        response = result(
            ok=False,
            command=_command_name(parsed),
            policy=policy,
            dry_run=bool(getattr(parsed, "dry_run", False)) if parsed else False,
            error={"code": "io_error", "message": str(exc), "details": {}},
        )
        exit_code = 2
    if _human_permission_output(parsed):
        action = "setup" if parsed.app == "setup" else ("grant" if parsed.action == "grant" else "revoke")
        _print_human_permission_result(response, action=action)
    else:
        print(json.dumps(response, sort_keys=True, separators=(",", ":")))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
