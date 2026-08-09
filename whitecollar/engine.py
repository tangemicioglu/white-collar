from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .adapters import MailAdapter, OoxmlWordAdapter, OutlookComAdapter, PowerPointComAdapter, SlidesAdapter, UnavailableMailAdapter, UnavailableSlidesAdapter, Win32WordComAdapter, WordAdapter
from .authority import Authority
from .errors import ValidationError
from .models import Plan
from .permissions import require_capability
from .policy import authorize_plan


@dataclass
class RuntimeAdapters:
    word: WordAdapter
    slides: SlidesAdapter
    mail: MailAdapter

    @classmethod
    def local(
        cls,
        *,
        word_backend: str = "local",
        slides_backend: str = "local",
        mail_backend: str = "local",
    ) -> "RuntimeAdapters":
        word = Win32WordComAdapter() if word_backend == "com" else OoxmlWordAdapter()
        slides = PowerPointComAdapter() if slides_backend == "com" else UnavailableSlidesAdapter()
        mail = OutlookComAdapter() if mail_backend == "com" else UnavailableMailAdapter()
        return cls(word, slides, mail)


def inspect_document(
    app: str,
    target: Path,
    policy: str,
    adapters: RuntimeAdapters,
    *,
    render_dir: Path | None = None,
    backend: str = "local",
    authority: Authority | None = None,
) -> dict[str, Any]:
    if not target.is_absolute():
        raise ValidationError("target must be an absolute path")
    require_capability(policy, f"{app}.read", target=str(target))
    if app == "word":
        if render_dir is not None and not render_dir.is_absolute():
            raise ValidationError("render_dir must be an absolute path")
        if render_dir is not None:
            require_capability(policy, "word.render", target=str(target))
        if authority is not None:
            capabilities = ["word.read"] + (["word.render"] if render_dir is not None else [])
            authority.require_access(app, backend, policy, str(target), capabilities)
        return adapters.word.inspect(target, render_dir=render_dir)
    if app == "slides":
        if render_dir is not None and not render_dir.is_absolute():
            raise ValidationError("render_dir must be an absolute path")
        if render_dir is not None:
            require_capability(policy, "slides.render", target=str(target))
        if authority is not None:
            capabilities = ["slides.read"] + (["slides.render"] if render_dir is not None else [])
            authority.require_access(app, backend, policy, str(target), capabilities)
        return adapters.slides.inspect(target, render_dir=render_dir)
    raise ValidationError(f"unsupported app: {app}")


def apply_plan(
    plan: Plan,
    *,
    dry_run: bool,
    adapters: RuntimeAdapters,
    authority: Authority | None = None,
    backend: str = "local",
) -> dict[str, Any]:
    authorize_plan(plan, dry_run=dry_run, authority=authority, backend=backend)
    if plan.app == "word":
        return adapters.word.apply(plan, dry_run=dry_run)
    if plan.app == "slides":
        return adapters.slides.apply(plan, dry_run=dry_run)
    return adapters.mail.apply(plan, dry_run=dry_run)


def search_mail(
    query: str,
    *,
    limit: int,
    policy: str,
    adapters: RuntimeAdapters,
    folder: str = "Inbox",
    backend: str = "local",
    authority: Authority | None = None,
) -> list[dict[str, Any]]:
    require_capability(policy, "mail.metadata.read")
    if not query.strip():
        raise ValidationError("mail query must not be empty")
    if limit < 1 or limit > 100:
        raise ValidationError("mail limit must be between 1 and 100")
    if authority is not None:
        authority.require_access("mail", backend, policy, "mailbox", ("mail.metadata.read",))
    return adapters.mail.search(query, limit=limit, folder=folder)


def read_mail(
    message_id: str,
    *,
    policy: str,
    adapters: RuntimeAdapters,
    include_body: bool = False,
    backend: str = "local",
    authority: Authority | None = None,
) -> dict[str, Any]:
    if not message_id.strip():
        raise ValidationError("message id must not be empty")
    capability = "mail.body.read" if include_body else "mail.metadata.read"
    require_capability(policy, capability, target=message_id)
    if authority is not None:
        authority.require_access("mail", backend, policy, message_id, (capability,))
    return adapters.mail.read(message_id, include_body=include_body)
