# Doctor

`white-collar doctor` is a side-effect-free readiness check for the Windows
Office COM installation:

```powershell
white-collar doctor
```

It reports Windows and optional dependency availability, COM backend readiness,
protected permission-store status, and whether Outlook COM has been enabled by
an owner grant. It does not start Word, PowerPoint, or Outlook; it does not
open documents or mail; and it redacts owner-grant targets.

The normal output is the machine-readable `white-collar.result/v1` envelope.
Use `white-collar permissions show` when you intentionally need explicit grant
details.
