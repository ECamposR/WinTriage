# WinTriage

`WinTriage.ps1` is a read-only Windows 11 triage CLI for helpdesk and support workflows.

It is designed for Windows PowerShell 5.1 and avoids external modules.

## Current Scope

- Read-only diagnostic skeleton.
- JSON report generation.
- Markdown report generation.
- Safe execution without administrator privileges.
- Fallback output handling for RMM and remote runs.

## Usage

```powershell
.\WinTriage.ps1
.\WinTriage.ps1 -Quick
.\WinTriage.ps1 -Full -Days 14
.\WinTriage.ps1 -UseExitCode
.\WinTriage.ps1 -JsonOnly
.\WinTriage.ps1 -OpenReport
```

## Output

Default report root:

```text
C:\ProgramData\WinTriage\Reports
```

Fallback report root:

```text
%TEMP%\WinTriage\Reports
```

Each run creates a per-host, per-timestamp folder containing:

- `WinTriage.json`
- `WinTriage.md` unless `-JsonOnly` is used

## Notes

- The script is read-only by design.
- It does not repair, delete, restart, or change system settings.
- Future work will add real collectors for system, events, Defender, Windows Update, services, disk, and performance.

