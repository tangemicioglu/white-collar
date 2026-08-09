from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from . import __version__
from .authority import (
    HUMAN_CONFIRMATION_PHRASE,
    HUMAN_PERMISSION_NOTICE,
    Authority,
    load_authority,
    make_grant,
    revoke_all_grants,
    revoke_grant,
    save_grant,
)
from .engine import RuntimeAdapters, apply_plan, inspect_document, read_mail, search_mail
from .errors import ValidationError, WhiteCollarError
from .models import Plan, result
from .permissions import CAPABILITIES, PROFILE_NAMES, catalog, require_capability


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValidationError(message)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(prog="white-collar", description="Narrow local Office control plane")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    apps = parser.add_subparsers(dest="app", required=True, parser_class=JsonArgumentParser)

    for app in ("word", "slides"):
        app_parser = apps.add_parser(app)
        commands = app_parser.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
        inspect = commands.add_parser("inspect")
        inspect.add_argument("target")
        inspect.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
        inspect.add_argument("--backend", choices=("local", "com"), default="local")
        inspect.add_argument("--render-dir", help="write one native Office PNG per page or slide")
        apply = commands.add_parser("apply")
        apply.add_argument("--plan", required=True)
        apply.add_argument("--dry-run", action="store_true")
        apply.add_argument("--backend", choices=("local", "com"), default="local")

    mail = apps.add_parser("mail")
    mail_commands = mail.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
    search = mail_commands.add_parser("search")
    search.add_argument("--query", required=True)
    search.add_argument("--limit", type=int, default=20)
    search.add_argument("--folder", default="Inbox", help="Inbox, Sent Items, Drafts, or another supported default folder")
    search.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
    search.add_argument("--backend", choices=("local", "com"), default="local")
    read = mail_commands.add_parser("read")
    read.add_argument("--id", required=True, dest="message_id")
    read.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
    read.add_argument("--include-body", action="store_true", help="request sensitive message body access")
    read.add_argument("--backend", choices=("local", "com"), default="local")
    mail_apply = mail_commands.add_parser("apply")
    mail_apply.add_argument("--plan", required=True)
    mail_apply.add_argument("--dry-run", action="store_true")
    mail_apply.add_argument("--backend", choices=("local", "com"), default="local")

    permissions = apps.add_parser("permissions", help="inspect and check local capability grants")
    permission_commands = permissions.add_subparsers(dest="action", required=True, parser_class=JsonArgumentParser)
    show = permission_commands.add_parser("show")
    show.add_argument("--policy", choices=PROFILE_NAMES, default="read-only")
    check = permission_commands.add_parser("check")
    check.add_argument("--capability", required=True)
    check.add_argument("--target")
    check.add_argument("--policy", choices=PROFILE_NAMES, default="read-only")
    check.add_argument("--backend", choices=("local", "com"), default="local")
    grant = permission_commands.add_parser("grant", help="human-owner-only; store a narrowly scoped grant")
    grant.add_argument("--app", dest="grant_app", choices=("word", "slides", "mail"), required=True)
    grant.add_argument("--backend", choices=("local", "com"), required=True)
    grant.add_argument("--policy", choices=PROFILE_NAMES, required=True)
    grant.add_argument("--target", action="append", required=True, help="exact file, message id, or 'mailbox'; repeat for multiple targets")
    grant.add_argument("--capability", action="append", help="narrow the grant; repeat for multiple capabilities")
    revoke = permission_commands.add_parser("revoke", help="human-owner-only; revoke a narrowly scoped grant")
    revoke.add_argument("--app", dest="grant_app", choices=("word", "slides", "mail"))
    revoke.add_argument("--backend", choices=("local", "com"))
    revoke.add_argument("--policy", choices=PROFILE_NAMES)
    revoke.add_argument("--target", action="append", help="exact target; repeat for multiple targets")
    revoke.add_argument("--capability", action="append", help="identify the grant; repeat for multiple capabilities")
    revoke.add_argument("--all", action="store_true", help="revoke all owner grants; human confirmation is still required")
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
    print(
        f"If you are the human owner, type {HUMAN_CONFIRMATION_PHRASE!r} to {action}.",
        file=sys.stderr,
    )
    sys.stderr.flush()
    answer = sys.stdin.readline().strip()
    if answer != HUMAN_CONFIRMATION_PHRASE:
        raise ValidationError(
            "human confirmation was not received; no permission was changed",
            details={
                "human_action_required": True,
                "agent_instruction": HUMAN_PERMISSION_NOTICE,
                "requested_action": action,
            },
        )


def _run(
    args: argparse.Namespace,
    adapters: RuntimeAdapters,
    authority: Authority,
    grant_store: Any | None = None,
) -> dict[str, Any]:
    command = _command_name(args)
    if args.app == "permissions" and args.action == "show":
        data = catalog(policy=args.policy, authority=authority)
        data["authority"] = authority.to_dict()
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
        _human_confirmation(action="grant", summary=json.dumps(grant.to_dict(), sort_keys=True))
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
            _human_confirmation(action="revoke", summary=json.dumps(grant.to_dict(), sort_keys=True))
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
        active_authority = authority or load_authority(store=grant_store)
        response = _run(
            parsed,
            adapters
            or RuntimeAdapters.local(
                word_backend=getattr(parsed, "backend", "local") if getattr(parsed, "app", None) == "word" else "local",
                slides_backend=getattr(parsed, "backend", "local") if getattr(parsed, "app", None) == "slides" else "local",
                mail_backend=getattr(parsed, "backend", "local") if getattr(parsed, "app", None) == "mail" else "local",
            ),
            active_authority,
            grant_store,
        )
        exit_code = 0
    except WhiteCollarError as exc:
        policy = getattr(parsed, "policy", "read-only") if parsed else "read-only"
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
        response = result(
            ok=False,
            command=_command_name(parsed),
            policy=policy,
            dry_run=bool(getattr(parsed, "dry_run", False)) if parsed else False,
            error={"code": "io_error", "message": str(exc), "details": {}},
        )
        exit_code = 2
    print(json.dumps(response, sort_keys=True, separators=(",", ":")))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
