# Plans and results

The wire contracts are versioned JSON documents:

- plan: `white-collar.plan/v1`
- result: `white-collar.result/v1`

The repository's authoritative schemas are `schemas/plan-v1.schema.json` and
`schemas/result-v1.schema.json`. Reject unknown fields and preserve the version
string when generating or validating data.

## File mutation plan

```json
{
  "schema": "white-collar.plan/v1",
  "app": "word",
  "target": {
    "path": "C:/work/brief.docx",
    "expected_sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"
  },
  "policy": "review",
  "operations": [{
    "op": "replace_text",
    "find": "Quarterly Draft",
    "replace": "Quarterly Review",
    "occurrence": "all"
  }],
  "write": {"mode": "save-as", "path": "C:/work/brief-reviewed.docx"}
}
```

For in-place Office writes, use `policy: "edit"` and
`write: {"mode": "in-place", "snapshot": "..."}`. For live mail use an
`id` target and `write.mode: "none"`.

## Result handling

The minimal successful shape is:

```json
{
  "schema": "white-collar.result/v1",
  "ok": true,
  "command": "word.inspect",
  "policy": "read-only",
  "data": {"path": "C:/work/brief.docx"}
}
```

Expected failures remain machine-readable:

```json
{
  "schema": "white-collar.result/v1",
  "ok": false,
  "command": "mail.read",
  "policy": "review",
  "error": {
    "code": "authority_denied",
    "message": "The requested capability is not authorized.",
    "details": {"capability": "mail.body.read", "target": "MESSAGE_ENTRY_ID"}
  }
}
```

Do not treat a missing `dry_run` field as false: the CLI omits it for ordinary
non-simulated responses to save tokens. It is present only when a mutation was
explicitly simulated.
