# Case 04 — FTP model exchange alerts (ntfy / Telegram)

[← Portfolio hub](../../README.md) · [RU](README.ru.md)

## Problem

Teams exchange Revit (and related) models over a shared FTP/FTPS folder. BIM managers need a **phone push** when a folder or file changes — without sitting in FileZilla all day. Mobile FTP clients do not offer reliable remote folder watch + background alerts (FTP has no push events; iOS limits background polling).

Corporate workstations often **block outbound HTTPS** to notification services (`ntfy.sh`, `api.telegram.org`), so a watcher cannot live on the locked-down office PC.

## Solution

PowerShell toolkit that:

1. Polls one or more FTP folders (plain FTP or optional TLS)
2. Diffs names / sizes / FTP timestamps against a local `state.json`
3. Sends push via **ntfy** (default) or **Telegram**
4. Runs every N minutes via Windows Task Scheduler on a machine that **can** reach both FTP and the notify API (e.g. a home always-on PC)

Includes: credential store (`Export-Clixml`), connectivity tests, multi-path watch, NEW/UPD/DEL lines with **FTP time** + local check time.

## Impact

| Item | Result |
|------|--------|
| Status | **Production-ready** (field-tested) |
| Writes models? | **No** — FTP list only |
| Scale | Multiple `RemotePaths`, recursive depth configurable |

## Safety

- Never uploads / deletes / renames on the FTP server
- Do not commit `config.json`, `credentials.xml`, or `state.json`
- Use a **secret** ntfy topic (treat like a password); rotate Telegram bot tokens if leaked
- Prefer a dedicated FTP account with **read-only** rights for the watcher

## How to run

1. Copy [`src/`](src/) to a folder on a PC with internet + FTP access
2. `config.example.json` → `config.json` (host, paths, `NotifyProvider`)
3. Phone: install **ntfy**, subscribe to your topic (or set up a Telegram bot)
4. `Save-FtpCredentials.ps1` → store FTP password
5. `Test-Ntfy.ps1` (or `Test-Telegram.ps1`) → confirm push
6. `Watch-FtpModels.ps1` once → baseline (no spam)
7. Change something on FTP → run again → notification
8. `Register-ScheduledTask.ps1` → every 5 minutes

Code: [`src/`](src/)
