from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest


@pytest.fixture
def make_docx():
    def factory(path: Path, *paragraphs: str) -> Path:
        body = "".join(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>" for text in paragraphs)
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"<w:body>{body}<w:sectPr/></w:body></w:document>"
        )
        content_types = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>"
        )
        relationships = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>'
        )
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", content_types)
            archive.writestr("_rels/.rels", relationships)
            archive.writestr("word/document.xml", document)
        return path

    return factory


def write_plan(
    path: Path,
    target: Path,
    output: Path,
    *,
    app: str = "word",
    policy: str = "review",
    mode: str = "save-as",
    find: str = "Draft",
    replace: str = "Final",
    expected_sha256: str | None = None,
) -> Path:
    target_value = {"path": str(target)}
    if expected_sha256:
        target_value["expected_sha256"] = expected_sha256
    write = {"mode": mode, "path" if mode == "save-as" else "snapshot": str(output)}
    value = {
        "schema": "white-collar.plan/v1",
        "app": app,
        "target": target_value,
        "policy": policy,
        "operations": [{"op": "replace_text", "find": find, "replace": replace, "occurrence": "all"}],
        "write": write,
    }
    path.write_text(json.dumps(value), encoding="utf-8")
    return path
