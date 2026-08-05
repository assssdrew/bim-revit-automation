# FTP model exchange → phone alerts

PowerShell watcher for shared FTP folders used in BIM/Revit model exchange.

## Files

| File | Role |
|------|------|
| `Watch-FtpModels.ps1` | Poll FTP, diff state, send ntfy/Telegram |
| `Save-FtpCredentials.ps1` | Encrypt FTP password for current Windows user |
| `Register-ScheduledTask.ps1` | Task Scheduler every 5 minutes |
| `Test-Ntfy.ps1` | DNS/TCP + test push to ntfy |
| `Test-Telegram.ps1` | DNS/TCP + test message to Telegram |
| `config.example.json` | Template — copy to `config.json` |

## Config notes

- `UseSsl: false` — many internal FTP hosts reject AUTH TLS (534)
- `NotifyProvider`: `ntfy` or `telegram`
- `RemotePaths` — array of folders to watch
- `RecursiveDepth` — how deep to list under each path
- Notification body includes FTP listing timestamp + local check time

## Secrets

Never commit: `config.json`, `credentials.xml`, `state.json` (listed in `.gitignore`).
