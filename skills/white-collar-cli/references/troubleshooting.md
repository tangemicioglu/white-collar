# Troubleshooting

## Start with doctor

```powershell
white-collar doctor
```

This checks Python, optional dependencies, backend readiness, protected
permission storage, and whether Outlook COM has been enabled. It does not start
Office, open documents, or read mail.

## Common results

- `backend_unavailable`: the selected backend is not available. Use the local
  Word path, install the `office` extra, or run on Windows with the required
  Office application installed.
- `authority_denied`: the requested capability or exact target is not granted.
  Do not modify the plan to evade it; ask the human owner if they want to run
  the appropriate interactive setup/grant.
- `plan_invalid` or validation failure: inspect the schema, app, target, policy,
  operation name, and write mode. Do not add arbitrary COM method fields.
- `target_not_found`, `target_changed`, or `output_exists`: verify the exact
  path/ID, refresh the inspection, or choose a new explicit output/snapshot.
- `no_match`: the requested Word replacement did not match. Word text can be
  split across OOXML nodes; inspect the document and avoid guessing a new
  phrase.

## Backend installation

From a checkout:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pip install -e ".[office]"
python -m pytest -q
```

The normal test suite does not require Microsoft Office. The live COM backends
need Windows, the corresponding installed Office application, and the optional
dependencies. Never install or enable Outlook access merely to work around a
readiness failure without the user's explicit decision.

## PowerShell completion

The bounded command vocabulary can be loaded into the current session:

```powershell
white-collar completions powershell | Out-String | Invoke-Expression
```

The command prints completion code only and does not modify the profile.
