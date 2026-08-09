from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .adapters import MailAdapter, OoxmlWordAdapter, PowerPointComAdapter, SlidesAdapter, UnavailableMailAdapter, UnavailableSlidesAdapter, Win32WordComAdapter, WordAdapter
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
    def local(cls, *, word_backend: str = "local", slides_backend: str = "local") -> "RuntimeAdapters":
        word = Win32WordComAdapter() if word_backend == "com" else OoxmlWordAdapter()
        slides = PowerPointComAdapter() if slides_backend == "com" else UnavailableSlidesAdapter()
        return cls(word, slides, UnavailableMailAdapter())


def inspect_document(
    app: str,
    target: Path,
    policy: str,
    adapters: RuntimeAdapters,
    *,
    render_dir: Path | None = None,
) -> dict[str, Any]:
    if not target.is_absolute():
        raise ValidationError("target must be an absolute path")
    require_capability(policy, f"{app}.read", target=str(target))
    if app == "word":
        if render_dir is not None and not render_dir.is_absolute():
            raise ValidationError("render_dir must be an absolute path")
        if render_dir is not None:
            require_capability(policy, "word.render", target=str(target))
        return adapters.word.inspect(target, render_dir=render_dir)
    if app == "slides":
        if render_dir is not None and not render_dir.is_absolute():
            raise ValidationError("render_dir must be an absolute path")
        if render_dir is not None:
            require_capability(policy, "slides.render", target=str(target))
        return adapters.slides.inspect(target, render_dir=render_dir)
    raise ValidationError(f"unsupported app: {app}")


def apply_plan(plan: Plan, *, dry_run: bool, adapters: RuntimeAdapters) -> dict[str, Any]:
    authorize_plan(plan, dry_run=dry_run)
    if plan.app == "word":
        return adapters.word.apply(plan, dry_run=dry_run)
    return adapters.slides.apply(plan, dry_run=dry_run)


def search_mail(query: str, *, limit: int, policy: str, adapters: RuntimeAdapters) -> list[dict[str, Any]]:
    require_capability(policy, "mail.metadata.read")
    if not query.strip():
        raise ValidationError("mail query must not be empty")
    if limit < 1 or limit > 100:
        raise ValidationError("mail limit must be between 1 and 100")
    return adapters.mail.search(query, limit=limit)


def read_mail(
    message_id: str,
    *,
    policy: str,
    adapters: RuntimeAdapters,
    include_body: bool = False,
) -> dict[str, Any]:
    if not message_id.strip():
        raise ValidationError("message id must not be empty")
    capability = "mail.body.read" if include_body else "mail.metadata.read"
    require_capability(policy, capability, target=message_id)
    return adapters.mail.read(message_id, include_body=include_body)
