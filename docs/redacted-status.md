# Redacted permission status

Permission status normally includes exact owner-grant targets because agents
may need to understand why a plan is allowed or denied. When sharing terminal
output or diagnostics, use:

```powershell
white-collar permissions show --redacted
```

This preserves application, backend, policy, capability, and grant counts but
replaces every owner-grant target with `<redacted>`. It does not change the
active authority or the behavior of `permissions check` and plan execution.
`white-collar doctor` is redacted by default.
