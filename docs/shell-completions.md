# Shell completion and help

The CLI ships a PowerShell completer for its bounded command and option
vocabulary:

```powershell
white-collar completions powershell | Out-String | Invoke-Expression
```

To load it automatically for future PowerShell sessions, add the same command
to `$PROFILE`. The command prints only completion code and does not modify the
profile itself.

The root help includes copyable examples:

```powershell
white-collar --help
white-collar word --help
white-collar mail --help
white-collar setup --help
```
