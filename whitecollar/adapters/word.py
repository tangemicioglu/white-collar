from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Protocol
from xml.etree import ElementTree

from ..errors import ValidationError
from ..models import Plan

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
TEXT_TAG = f"{{{WORD_NS}}}t"
PARAGRAPH_TAG = f"{{{WORD_NS}}}p"
TABLE_TAG = f"{{{WORD_NS}}}tbl"


class WordAdapter(Protocol):
    def inspect(self, target: Path) -> dict[str, Any]: ...

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]: ...


class OoxmlWordAdapter:
    """A constrained DOCX adapter; it never exposes arbitrary OOXML or COM calls."""

    def inspect(self, target: Path) -> dict[str, Any]:
        _validate_docx(target)
        parts = _read_word_parts(target)
        document = parts["word/document.xml"]
        root = ElementTree.fromstring(document)
        text = "".join(node.text or "" for node in root.iter(TEXT_TAG))
        return {
            "format": "docx",
            "sha256": _sha256(target),
            "bytes": target.stat().st_size,
            "paragraphs": sum(1 for _ in root.iter(PARAGRAPH_TAG)),
            "tables": sum(1 for _ in root.iter(TABLE_TAG)),
            "characters": len(text),
            "text": text,
        }

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        source = Path(plan.target.path)
        _validate_docx(source)
        before_sha = _sha256(source)
        if plan.target.expected_sha256 and plan.target.expected_sha256.lower() != before_sha:
            raise ValidationError(
                "target SHA-256 does not match plan",
                details={"expected": plan.target.expected_sha256, "actual": before_sha},
            )

        replacements, changed_parts = _preview_replacements(source, plan)
        changes = [
            {
                "operation": index,
                "op": operation["op"],
                "matches": replacements[index],
            }
            for index, operation in enumerate(plan.operations)
        ]
        if dry_run:
            return {"changes": changes, "changed_parts": sorted(changed_parts), "written": False}
        unmatched = [index for index, count in enumerate(replacements) if count == 0]
        if unmatched:
            raise ValidationError(
                "one or more operations matched no text; no file was written",
                details={"operation_indexes": unmatched},
            )

        destination = Path(plan.write.path) if plan.write.mode == "save-as" else source
        snapshot = Path(plan.write.snapshot) if plan.write.snapshot else None
        for candidate, label in ((destination, "destination"), (snapshot, "snapshot")):
            if candidate is not None and candidate != source and candidate.exists():
                raise ValidationError(f"{label} already exists", details={"path": str(candidate)})
            if candidate is not None:
                candidate.parent.mkdir(parents=True, exist_ok=True)
        if snapshot is not None:
            shutil.copy2(source, snapshot)

        _write_transformed(source, destination, plan)
        return {
            "changes": changes,
            "changed_parts": sorted(changed_parts),
            "written": True,
            "output": str(destination),
            "snapshot": str(snapshot) if snapshot else None,
            "before_sha256": before_sha,
            "after_sha256": _sha256(destination),
        }


def _validate_docx(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise ValidationError("Word targets must have a .docx extension", details={"path": str(path)})
    if not path.is_file():
        raise ValidationError("Word target does not exist", details={"path": str(path)})
    try:
        with zipfile.ZipFile(path) as archive:
            if "word/document.xml" not in archive.namelist():
                raise ValidationError("target is not a Word OOXML document", details={"path": str(path)})
    except zipfile.BadZipFile as exc:
        raise ValidationError("target is not a valid DOCX archive", details={"path": str(path)}) from exc


def _read_word_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if name == "word/document.xml" or name.startswith("word/header") or name.startswith("word/footer")
            if name.endswith(".xml")
        }


def _transform_part(
    content: bytes,
    operations: tuple[dict[str, Any], ...],
    first_done: list[bool],
) -> tuple[bytes, list[int]]:
    root = ElementTree.fromstring(content)
    counts = [0] * len(operations)
    for node in root.iter(TEXT_TAG):
        text = node.text or ""
        for index, operation in enumerate(operations):
            if operation["occurrence"] == "first" and first_done[index]:
                continue
            matches = text.count(operation["find"])
            if not matches:
                continue
            limit = 1 if operation["occurrence"] == "first" else -1
            text = text.replace(operation["find"], operation["replace"], limit)
            applied = 1 if limit == 1 else matches
            counts[index] += applied
            first_done[index] = first_done[index] or applied > 0
        node.text = text
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True), counts


def _preview_replacements(path: Path, plan: Plan) -> tuple[list[int], set[str]]:
    totals = [0] * len(plan.operations)
    changed_parts: set[str] = set()
    first_done = [False] * len(plan.operations)
    for name, content in _read_word_parts(path).items():
        transformed, counts = _transform_part(content, plan.operations, first_done)
        if transformed != content and any(counts):
            changed_parts.add(name)
        totals = [left + right for left, right in zip(totals, counts)]
    return totals, changed_parts


def _write_transformed(source: Path, destination: Path, plan: Plan) -> None:
    temporary = tempfile.NamedTemporaryFile(prefix=".white-collar-", suffix=".docx", dir=destination.parent, delete=False)
    temporary_path = Path(temporary.name)
    temporary.close()
    try:
        transformable = set(_read_word_parts(source))
        first_done = [False] * len(plan.operations)
        with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(temporary_path, "w") as outgoing:
            for item in incoming.infolist():
                content = incoming.read(item.filename)
                if item.filename in transformable:
                    content, _ = _transform_part(content, plan.operations, first_done)
                outgoing.writestr(item, content)
        os.replace(temporary_path, destination)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
