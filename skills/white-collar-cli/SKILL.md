---
name: white-collar-cli
description: Use the white-collar Windows Office COM CLI for bounded Word, PowerPoint, and Outlook workflows. Apply this skill when an agent needs to inspect, render, edit, search, read, draft, or send Office content through white-collar, including choosing a command, plan, policy, backend, target, dry-run, or permission-safe workflow.
---

# White-Collar CLI

Use this skill when the task is to operate on a Windows `.docx`, `.pptx`, or
Outlook Classic mailbox through the `white-collar` command. It teaches the
small semantic COM command surface and the safety boundary; it is not a guide
to arbitrary COM automation or to upstream MCP servers.

## First pass

1. Confirm the command exists with `white-collar --help` or use the repository's editable install.
2. Run `white-collar doctor` when backend or permission readiness is unclear. It is side-effect-free and emits a compact `white-collar.result/v1` JSON envelope.
3. Choose a bounded shortcut for a common task, otherwise write a versioned plan. Never invent a target path, message ID, or account.
4. Inspect before changing anything. For a mutation, run the exact command with `--dry-run`, review its JSON, then repeat without `--dry-run` only when the requested policy and authority allow it.
5. Parse the result rather than scraping prose. Check `ok`; on failure preserve `error.code`, `error.message`, and `error.details`.

## Command routing

- Word document text, rendering, or metadata: `word inspect`.
- New blank Word document: `word create --output FILE.docx`.
- Simple Word replacement: `word replace`; complex or live Word semantics: `word apply --plan`.
- PowerPoint inspection or slide images: `slides inspect`; semantic live edits: `slides apply --plan`.
- PowerPoint title/body edits: use the native `Title`, `Body`, `Content`, or `Subtitle` selectors; use `slides_live_add_textbox` only for deliberate freeform text.
- New blank PowerPoint presentation: `slides create --output FILE.pptx`.
- Outlook metadata search or read: `mail search` and `mail read`.
- Outlook draft composition: `mail draft`; sending an existing draft: `mail send`.
- Permission readiness: `doctor`, `permissions show --redacted`, or `permissions check`.
- Human onboarding: `setup`. Agents must not run or confirm it on a human's behalf.

## Default policy and authority

- Word/PowerPoint: `read-only` reads; `review` creation and save-as writes; `edit` in-place writes with a snapshot. All operations use the live COM adapters; screen capture and in-place edits can still need authority.
- Outlook: `read-only` is the default. Metadata search/read is read-only; body reads are sensitive and need `review` or `edit`; mark read/unread is `review`; move/delete and draft composition are `edit`; sending is `send` only.
- A plan's policy is never authority. Do not bypass a denial by changing the plan, backend, target, or a local file.
- Use exact file targets for Office work and exact message IDs for mail work. `mailbox` is a deliberately broad human-configured scope, not an agent shortcut.
- Normal commands return one compact JSON object. Permission grants, revokes, and setup are intentionally human-facing when interactive; an agent must stop and ask the human to perform them.

## Progressive disclosure

Load only the reference needed for the task:

- Command selection and shortcut-versus-plan examples: `references/command-routing.md`.
- Policies, grants, setup presets, and the human boundary: `references/permissions-and-safety.md`.
- Word inspection, replacement, plan writes, and live operation names: `references/word.md`.
- PowerPoint inspection, native slide rendering, and live operation names: `references/powerpoint.md`.
- Outlook search/read, drafts, organization, and send: `references/outlook.md`.
- Plan/result schemas and JSON examples: `references/schemas.md`.
- Readiness failures and recovery checks: `references/troubleshooting.md`.

## Hard boundaries

- Do not expose or request raw COM methods, object paths, macros, arbitrary dispatch, or a high-cardinality MCP tool list. Use the finite semantic operations documented in the app references.
- Do not create an authority file, edit the protected permission store, fabricate a human confirmation, or retry a denied Outlook action with a broader target.
- Treat message bodies, recipients, drafts, and sends as potentially sensitive. Preview and validate before any mailbox write; never send unless the user explicitly requested that exact send and the separate `send` authority is already present.
- Keep machine output machine-readable. Ask for human-only permission changes in ordinary language rather than directing the agent to silently navigate a permission path.
