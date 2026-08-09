from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from . import __version__
from .engine import RuntimeAdapters, apply_plan, inspect_document, read_mail, search_mail
from .errors import ValidationError, WhiteCollarError
from .models import PLAN_SCHEMA, Plan, result


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
    search.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
    read = mail_commands.add_parser("read")
    read.add_argument("--id", required=True, dest="message_id")
    read.add_argument("--policy", choices=("read-only", "review", "edit"), default="read-only")
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


def _run(args: argparse.Namespace, adapters: RuntimeAdapters) -> dict[str, Any]:
    command = _command_name(args)
    if args.app in {"word", "slides"} and args.action == "inspect":
        target = Path(args.target).resolve()
        render_dir = Path(args.render_dir).resolve() if getattr(args, "render_dir", None) else None
        data = inspect_document(args.app, target, args.policy, adapters, render_dir=render_dir)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, target=str(target), data=data)
    if args.app in {"word", "slides"} and args.action == "apply":
        plan = _load_plan(args.plan)
        args.policy = plan.policy
        if plan.app != args.app:
            raise ValidationError(
                "plan app does not match command",
                details={"plan_app": plan.app, "command_app": args.app},
            )
        data = apply_plan(plan, dry_run=args.dry_run, adapters=adapters)
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
    if args.app == "mail" and args.action == "search":
        data = search_mail(args.query, limit=args.limit, policy=args.policy, adapters=adapters)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=data)
    if args.app == "mail" and args.action == "read":
        data = read_mail(args.message_id, policy=args.policy, adapters=adapters)
        return result(ok=True, command=command, policy=args.policy, dry_run=False, data=data)
    raise ValidationError("unsupported command")


def main(argv: Sequence[str] | None = None, *, adapters: RuntimeAdapters | None = None) -> int:
    parsed: argparse.Namespace | None = None
    try:
        parsed = build_parser().parse_args(argv)
        response = _run(
            parsed,
            adapters
            or RuntimeAdapters.local(
                word_backend=getattr(parsed, "backend", "local") if getattr(parsed, "app", None) == "word" else "local",
                slides_backend=getattr(parsed, "backend", "local") if getattr(parsed, "app", None) == "slides" else "local",
            ),
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
