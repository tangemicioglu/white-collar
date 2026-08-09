from __future__ import annotations

import json
import os
import zipfile
from pathlib import Path
import subprocess

import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--run-real-word",
        action="store_true",
        default=False,
        help="run integration tests against an installed Microsoft Word instance",
    )
    parser.addoption(
        "--run-real-powerpoint",
        action="store_true",
        default=False,
        help="run integration tests against an installed Microsoft PowerPoint instance",
    )
    parser.addoption(
        "--run-real-outlook",
        action="store_true",
        default=False,
        help="run integration tests against an installed Outlook Classic instance",
    )
    parser.addoption(
        "--run-real-office",
        action="store_true",
        default=False,
        help="run integration tests against installed Microsoft Office applications",
    )


def pytest_collection_modifyitems(config, items):
    run_word = config.getoption("--run-real-word") or config.getoption("--run-real-office")
    run_powerpoint = config.getoption("--run-real-powerpoint") or config.getoption("--run-real-office")
    run_outlook = config.getoption("--run-real-outlook") or config.getoption("--run-real-office")
    for item in items:
        if "real_word" in item.keywords and not run_word:
            item.add_marker(pytest.mark.skip(reason="pass --run-real-word to exercise Microsoft Word COM"))
        if "real_powerpoint" in item.keywords and not run_powerpoint:
            item.add_marker(pytest.mark.skip(reason="pass --run-real-powerpoint to exercise Microsoft PowerPoint COM"))
        if "real_outlook" in item.keywords and not run_outlook:
            item.add_marker(pytest.mark.skip(reason="pass --run-real-outlook to exercise Outlook Classic COM"))


def pytest_sessionfinish(session, exitstatus):
    process_id = os.environ.pop("WHITE_COLLAR_REAL_POWERPOINT_PID", "")
    if not process_id.isdigit():
        return
    command = (
        "$p = Get-Process -Id "
        + process_id
        + " -ErrorAction SilentlyContinue; "
        "if ($null -ne $p -and $p.ProcessName -eq 'POWERPNT') { Stop-Process -Id "
        + process_id
        + " -Force }"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", command], check=False, capture_output=True)


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
