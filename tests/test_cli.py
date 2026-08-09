from __future__ import annotations

import json
import io
import subprocess
import sys

from conftest import write_plan
from whitecollar.adapters.word import OoxmlWordAdapter
from whitecollar.authority import Authority, MemoryCredentialStore, load_authority, make_grant, save_grant
from whitecollar import cli
from whitecollar.cli import main
from whitecollar.engine import RuntimeAdapters


class FakeSlides:
    def inspect(self, target, *, render_dir=None):
        return {"slides": 7, "title": "Review", "render_dir": str(render_dir) if render_dir else None}

    def apply(self, plan, *, dry_run):
        return {"changes": [{"operation": 0, "matches": 3}], "written": not dry_run}


class FakeWordRender:
    def inspect(self, target, *, render_dir=None):
        return {"render_dir": str(render_dir) if render_dir else None}

    def apply(self, plan, *, dry_run):
        return {"changes": [], "written": not dry_run}


class FakeMail:
    def __init__(self):
        self.last_plan = None

    def search(self, query, *, limit, folder="Inbox"):
        return [{"id": "m-1", "subject": "Roadmap", "query": query, "folder": folder}][:limit]

    def read(self, message_id, *, include_body=False):
        value = {"id": message_id, "subject": "Roadmap"}
        if include_body:
            value["body"] = "Ship it"
        return value

    def apply(self, plan, *, dry_run):
        self.last_plan = plan
        return {
            "changes": [{"op": plan.operations[0]["op"], "id": plan.target.id}],
            "written": not dry_run,
        }


def adapters() -> RuntimeAdapters:
    return RuntimeAdapters(OoxmlWordAdapter(), FakeSlides(), FakeMail())


def output(capsys):
    return json.loads(capsys.readouterr().out)


class TtyStream(io.StringIO):
    def isatty(self):
        return True


def test_permission_confirmation_uses_normal_yes_no_prompt(monkeypatch):
    stdin = TtyStream("yes\n")
    stderr = TtyStream()
    monkeypatch.setattr(cli.sys, "stdin", stdin)
    monkeypatch.setattr(cli.sys, "stderr", stderr)
    cli._human_confirmation(action="grant", summary="Permission change:\n  target: mailbox")
    prompt = stderr.getvalue()
    assert "Grant this permission? [y/N]" in prompt
    assert "I AM THE HUMAN OWNER" not in prompt


def invoke(argv, *, adapters):
    return main(argv, adapters=adapters, authority=Authority.for_testing())


def test_word_inspect_smoke(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    assert invoke(["word", "inspect", str(target)], adapters=adapters()) == 0
    value = output(capsys)
    assert value["schema"] == "white-collar.result/v1"
    assert value["command"] == "word.inspect"
    assert value["policy"] == "read-only"
    assert "dry_run" not in value


def test_doctor_is_machine_readable_and_does_not_need_office(capsys):
    assert main(["doctor"], adapters=adapters(), authority=Authority.default()) == 0
    value = output(capsys)
    assert value["command"] == "doctor"
    assert value["policy"] == "read-only"
    assert value["data"]["backends"]["word"]["local"]["status"] == "ready"
    assert value["data"]["permissions"]["owner_grants"] == 0
    assert value["data"]["permissions"]["targets"].startswith("redacted")


def test_word_apply_dry_run_smoke(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx")
    assert invoke(["word", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 0
    value = output(capsys)
    assert value["dry_run"] is True
    assert value["changes"][0]["matches"] == 1


def test_word_replace_shortcut_compiles_to_validated_plan(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    output_path = tmp_path / "reviewed.docx"
    assert main(
        [
            "word",
            "replace",
            "--target",
            str(target),
            "--find",
            "Draft",
            "--replace",
            "Final",
            "--output",
            str(output_path),
        ],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 0
    value = output(capsys)
    assert value["command"] == "word.replace"
    assert value["policy"] == "review"
    assert value["data"]["written"] is True
    assert output_path.exists()


def test_word_replace_shortcut_requires_safe_write_shape(tmp_path, capsys):
    assert main(
        ["word", "replace", "--target", str(tmp_path / "brief.docx"), "--find", "Draft", "--replace", "Final"],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 2
    value = output(capsys)
    assert "--output or --in-place" in value["error"]["message"]


def test_default_word_and_slides_review_writes_are_enabled(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    word_plan = write_plan(tmp_path / "word-plan.json", target, tmp_path / "final.docx")
    assert main(["word", "apply", "--plan", str(word_plan)], adapters=adapters(), authority=Authority.default()) == 0
    assert output(capsys)["data"]["written"] is True
    assert (tmp_path / "final.docx").exists()

    deck = tmp_path / "deck.pptx"
    slides_plan = write_plan(tmp_path / "slides-plan.json", deck, tmp_path / "final.pptx", app="slides")
    assert main(["slides", "apply", "--plan", str(slides_plan), "--dry-run"], adapters=adapters(), authority=Authority.default()) == 0
    assert output(capsys)["dry_run"] is True


def test_slides_commands_use_only_slides_adapter(tmp_path, capsys):
    deck = tmp_path / "deck.pptx"
    assert invoke(["slides", "inspect", str(deck)], adapters=adapters()) == 0
    assert output(capsys)["data"]["slides"] == 7
    plan = write_plan(tmp_path / "plan.json", deck, tmp_path / "new.pptx", app="slides")
    assert invoke(["slides", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 0
    assert output(capsys)["changes"][0]["matches"] == 3


def test_slides_inspect_render_dir_smoke(tmp_path, capsys):
    deck = tmp_path / "deck.pptx"
    render_dir = tmp_path / "rendered"
    assert invoke(["slides", "inspect", str(deck), "--render-dir", str(render_dir)], adapters=adapters()) == 0
    value = output(capsys)
    assert value["data"]["render_dir"] == str(render_dir.resolve())


def test_word_inspect_render_dir_smoke(tmp_path, capsys):
    target = tmp_path / "brief.docx"
    render_dir = tmp_path / "rendered"
    runtime = RuntimeAdapters(FakeWordRender(), FakeSlides(), FakeMail())
    assert invoke(["word", "inspect", str(target), "--backend", "com", "--render-dir", str(render_dir)], adapters=runtime) == 0
    value = output(capsys)
    assert value["data"]["render_dir"] == str(render_dir.resolve())


def test_mail_search_and_read_default_to_read_only(capsys):
    assert invoke(["mail", "search", "--query", "roadmap"], adapters=adapters()) == 0
    searched = output(capsys)
    assert searched["policy"] == "read-only"
    assert searched["data"][0]["id"] == "m-1"
    assert invoke(["mail", "read", "--id", "m-1"], adapters=adapters()) == 0
    assert "body" not in output(capsys)["data"]


def test_mail_shortcuts_compile_to_semantic_plans(capsys):
    fake_mail = FakeMail()
    runtime = RuntimeAdapters(OoxmlWordAdapter(), FakeSlides(), fake_mail)
    assert main(
        [
            "mail",
            "draft",
            "--account",
            "mailbox",
            "--to",
            "person@example.com",
            "--subject",
            "Test",
            "--body",
            "Hello",
            "--dry-run",
        ],
        adapters=runtime,
        authority=Authority.for_testing(),
    ) == 0
    value = output(capsys)
    assert value["command"] == "mail.draft"
    assert value["changes"][0]["op"] == "mail_live_create_draft"
    assert fake_mail.last_plan.operations[0]["args"]["to"] == "person@example.com"

    assert main(
        ["mail", "send", "--draft-id", "draft-1", "--dry-run"],
        adapters=runtime,
        authority=Authority.for_testing(),
    ) == 0
    value = output(capsys)
    assert value["command"] == "mail.send"
    assert value["policy"] == "send"
    assert fake_mail.last_plan.operations[0]["op"] == "mail_live_send"


def test_mail_backend_and_folder_flags_reach_the_adapter(capsys):
    assert invoke(
        ["mail", "search", "--backend", "com", "--folder", "Sent Items", "--query", "roadmap"],
        adapters=adapters(),
    ) == 0
    assert output(capsys)["data"][0]["folder"] == "Sent Items"


def test_outlook_backend_is_disabled_by_default(capsys):
    assert main(
        ["mail", "search", "--backend", "com", "--query", "roadmap"],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 2
    value = output(capsys)
    assert value["error"]["details"]["backend"] == "mail:com"


def test_agent_cannot_self_escalate_to_edit_policy(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx", policy="edit")
    assert main(
        ["word", "apply", "--plan", str(plan), "--dry-run"],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 2
    value = output(capsys)
    assert value["error"]["details"]["requested_policy"] == "edit"
    assert value["error"]["details"]["human_action_required"] is True
    assert "permissions grant" in value["error"]["details"]["human_only_command"]


def test_mail_body_requires_explicit_sensitive_policy(capsys):
    assert invoke(["mail", "read", "--id", "m-1", "--include-body"], adapters=adapters()) == 2
    denied = output(capsys)
    assert denied["error"]["code"] == "policy_denied"
    assert denied["error"]["details"]["capability"] == "mail.body.read"
    assert invoke(
        ["mail", "read", "--id", "m-1", "--include-body", "--policy", "review"],
        adapters=adapters(),
    ) == 0
    assert output(capsys)["data"]["body"] == "Ship it"


def test_mail_write_plan_requires_review_or_edit_and_owner_grant(tmp_path, capsys):
    plan_path = tmp_path / "mail.plan.json"
    plan_path.write_text(
        json.dumps(
            {
                "schema": "white-collar.plan/v1",
                "app": "mail",
                "target": {"id": "m-1"},
                "policy": "review",
                "operations": [{"op": "mail_live_mark_read"}],
                "write": {"mode": "none"},
            }
        ),
        encoding="utf-8",
    )
    assert main(
        ["mail", "apply", "--plan", str(plan_path), "--backend", "com", "--dry-run"],
        adapters=adapters(),
        authority=Authority.for_testing(),
    ) == 0
    assert output(capsys)["changes"][0]["op"] == "mail_live_mark_read"

    assert main(
        ["mail", "apply", "--plan", str(plan_path), "--backend", "com", "--dry-run"],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 2
    denied = output(capsys)
    assert denied["error"]["details"]["capabilities"] == ["mail.write.state"]
    assert denied["error"]["details"]["human_action_required"] is True

    store = MemoryCredentialStore()
    granted = save_grant(
        Authority.default(),
        make_grant(
            app="mail",
            backend="com",
            policy="review",
            capabilities=["mail.write.state"],
            targets=["m-1"],
        ),
        store,
    )
    assert main(
        ["mail", "apply", "--plan", str(plan_path), "--backend", "com", "--dry-run"],
        adapters=adapters(),
        authority=granted,
    ) == 0
    assert output(capsys)["changes"][0]["id"] == "m-1"

    plan_path.write_text(
        json.dumps(
            {
                "schema": "white-collar.plan/v1",
                "app": "mail",
                "target": {"id": "mailbox"},
                "policy": "edit",
                "operations": [{
                    "op": "mail_live_create_draft",
                    "args": {"to": "person@example.com", "subject": "Test", "body": "Hello"},
                }],
                "write": {"mode": "none"},
            }
        ),
        encoding="utf-8",
    )
    assert main(
        ["mail", "apply", "--plan", str(plan_path), "--backend", "com", "--dry-run"],
        adapters=adapters(),
        authority=Authority.for_testing(),
    ) == 0
    assert output(capsys)["changes"][0]["op"] == "mail_live_create_draft"

    plan_path.write_text(
        json.dumps(
            {
                "schema": "white-collar.plan/v1",
                "app": "mail",
                "target": {"id": "draft-1"},
                "policy": "send",
                "operations": [{"op": "mail_live_send"}],
                "write": {"mode": "none"},
            }
        ),
        encoding="utf-8",
    )
    assert main(
        ["mail", "apply", "--plan", str(plan_path), "--backend", "com", "--dry-run"],
        adapters=adapters(),
        authority=Authority.for_testing(),
    ) == 0
    assert output(capsys)["changes"][0]["op"] == "mail_live_send"


def test_agent_permission_change_requires_human_terminal(tmp_path, capsys):
    store = MemoryCredentialStore()
    assert main(
        [
            "permissions",
            "grant",
            "--app",
            "word",
            "--backend",
            "local",
            "--policy",
            "review",
            "--target",
            str((tmp_path / "brief.docx").resolve()),
        ],
        authority=Authority.default(),
        grant_store=store,
    ) == 2
    value = output(capsys)
    assert value["error"]["details"]["human_action_required"] is True
    assert "must stop and ask a human" in value["error"]["details"]["agent_instruction"]
    assert store.value is None


def test_setup_requires_human_terminal(tmp_path, capsys):
    store = MemoryCredentialStore()
    assert main(
        ["setup", "--app", "mail", "--policy", "edit"],
        adapters=adapters(),
        authority=Authority.default(),
        grant_store=store,
    ) == 2
    value = output(capsys)
    assert value["command"] == "setup"
    assert value["error"]["details"]["human_action_required"] is True
    assert store.value is None


def test_interactive_setup_writes_mailbox_scoped_owner_grant(monkeypatch, capsys):
    stdin = TtyStream("y\n")
    stderr = TtyStream()
    monkeypatch.setattr(cli.sys, "stdin", stdin)
    monkeypatch.setattr(cli.sys, "stderr", stderr)
    store = MemoryCredentialStore()
    assert main(
        ["setup", "--app", "mail", "--policy", "edit"],
        adapters=adapters(),
        authority=Authority.default(),
        grant_store=store,
    ) == 0
    assert capsys.readouterr().out == (
        "Application permissions updated.\n"
        "The settings are stored in protected OS credential storage.\n"
    )
    loaded = load_authority(store=store)
    loaded.require_access("mail", "com", "edit", "draft-1", ("mail.write.compose",))


def test_permissions_commands_are_machine_readable(tmp_path, capsys):
    assert main(
        ["permissions", "show", "--policy", "review"],
        adapters=adapters(),
        authority=Authority.default(),
    ) == 0
    shown = output(capsys)
    assert shown["data"]["schema"] == "white-collar.permissions/v1"
    assert "mail.body.read" in shown["data"]["profiles"]["review"]
    mail_body = next(item for item in shown["data"]["capabilities"] if item["name"] == "mail.body.read")
    assert mail_body["profile_granted"] is True
    assert mail_body["authority_granted"] is False
    assert mail_body["granted"] is False

    target = str((tmp_path / "brief.docx").resolve())
    assert invoke(
        ["permissions", "check", "--policy", "review", "--capability", "word.write.save_as", "--target", target],
        adapters=adapters(),
    ) == 0
    allowed = output(capsys)
    assert allowed["data"]["decision"] == "allow"

    assert invoke(
        ["permissions", "check", "--capability", "mail.body.read", "--target", "m-1"],
        adapters=adapters(),
    ) == 2
    denied = output(capsys)
    assert denied["error"]["details"]["capability"] == "mail.body.read"


def test_cli_validation_errors_are_json(capsys):
    assert invoke(["mail", "search", "--query", "x", "--limit", "101"], adapters=adapters()) == 2
    value = output(capsys)
    assert value["ok"] is False
    assert value["error"]["code"] == "validation_error"


def test_plan_app_must_match_command(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx")
    assert invoke(["slides", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 2
    assert output(capsys)["error"]["details"]["plan_app"] == "word"


def test_module_entrypoint_is_machine_readable():
    completed = subprocess.run(
        [sys.executable, "-m", "whitecollar.cli", "mail", "search", "--query", "roadmap"],
        check=False,
        capture_output=True,
        text=True,
    )
    value = json.loads(completed.stdout)
    assert completed.returncode == 2
    assert value["schema"] == "white-collar.result/v1"
    assert value["error"]["code"] == "backend_unavailable"
    assert completed.stderr == ""
