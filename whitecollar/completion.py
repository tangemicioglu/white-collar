"""Shell completion scripts shipped by the CLI."""

from __future__ import annotations

import textwrap


POWERSHELL_COMPLETION = textwrap.dedent(
    r'''
    Register-ArgumentCompleter -Native -CommandName white-collar -ScriptBlock {
        param($wordToComplete, $commandAst, $cursorPosition)

        $tokens = @($commandAst.CommandElements | Select-Object -Skip 1 | ForEach-Object {
            $_.Extent.Text.Trim("'`"")
        })
        $options = @()
        $root = @('word', 'slides', 'mail', 'setup', 'doctor', 'permissions', 'completions', '--help', '--version')
        $policy = @('read-only', 'review', 'edit', 'send')
        $setupPolicy = @('disabled', 'read-only', 'review', 'edit', 'send')

        if ($tokens.Count -eq 0) {
            $options = $root
        } elseif ($tokens[-1] -eq '--policy') {
            if ($tokens[0] -eq 'setup') { $options = $setupPolicy } else { $options = $policy }
        } elseif ($tokens[-1] -eq '--backend') {
            $options = @('com')
        } elseif ($tokens[-1] -eq '--app') {
            $options = @('word', 'slides', 'mail')
        } elseif ($tokens[-1] -eq '--preset') {
            $options = @('safe', 'office-authoring', 'outlook-review', 'outlook-send')
        } elseif ($tokens[-1] -eq '--shell') {
            $options = @('powershell')
        } elseif ($tokens[0] -eq 'word') {
            if ($tokens.Count -eq 1) { $options = @('inspect', 'apply', 'replace') }
            elseif ($tokens[1] -eq 'inspect') { $options = @('--policy', '--backend', '--render-dir', '--help') }
            elseif ($tokens[1] -eq 'apply') { $options = @('--plan', '--dry-run', '--backend', '--help') }
            elseif ($tokens[1] -eq 'replace') { $options = @('--target', '--find', '--replace', '--occurrence', '--output', '--in-place', '--snapshot', '--policy', '--dry-run', '--backend', '--help') }
        } elseif ($tokens[0] -eq 'slides') {
            if ($tokens.Count -eq 1) { $options = @('inspect', 'apply') }
            else { $options = @('--plan', '--dry-run', '--backend', '--render-dir', '--policy', '--help') }
        } elseif ($tokens[0] -eq 'mail') {
            if ($tokens.Count -eq 1) { $options = @('search', 'read', 'apply', 'draft', 'send') }
            elseif ($tokens[1] -eq 'search') { $options = @('--query', '--limit', '--folder', '--policy', '--backend', '--help') }
            elseif ($tokens[1] -eq 'read') { $options = @('--id', '--include-body', '--policy', '--backend', '--help') }
            elseif ($tokens[1] -eq 'apply') { $options = @('--plan', '--dry-run', '--backend', '--help') }
            elseif ($tokens[1] -eq 'draft') { $options = @('--account', '--to', '--cc', '--bcc', '--subject', '--body', '--dry-run', '--backend', '--help') }
            elseif ($tokens[1] -eq 'send') { $options = @('--draft-id', '--dry-run', '--backend', '--help') }
        } elseif ($tokens[0] -eq 'setup') {
            $options = @('--app', '--policy', '--preset', '--json', '--help')
        } elseif ($tokens[0] -eq 'doctor') {
            $options = @('--help')
        } elseif ($tokens[0] -eq 'permissions') {
            if ($tokens.Count -eq 1) { $options = @('show', 'check', 'grant', 'revoke') }
            elseif ($tokens[1] -eq 'show') { $options = @('--policy', '--redacted', '--help') }
            elseif ($tokens[1] -eq 'check') { $options = @('--capability', '--target', '--policy', '--backend', '--help') }
            elseif ($tokens[1] -eq 'grant') { $options = @('--app', '--backend', '--policy', '--target', '--capability', '--json', '--help') }
            elseif ($tokens[1] -eq 'revoke') { $options = @('--app', '--backend', '--policy', '--target', '--capability', '--all', '--json', '--help') }
        } elseif ($tokens[0] -eq 'completions') {
            $options = @('powershell', '--help')
        }

        $options | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object {
            [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterName', $_)
        }
    }
    ''').strip() + "\n"


def completion_script(shell: str) -> str:
    if shell != "powershell":
        raise ValueError(f"unsupported shell: {shell}")
    return POWERSHELL_COMPLETION
