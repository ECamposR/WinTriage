# WinTriage Worklog

## 2026-04-29

- Created the initial `WinTriage.ps1` skeleton.
- Added parameter validation and read-only reporting flow.
- Added JSON and Markdown report generation.
- Added fallback output-path handling for constrained environments.
- Added optional exit-code behavior for manual use versus RMM use.
- Added report-opening support for Markdown output.
- Added repository metadata files and project documentation.
- Corrected final report export order so JSON and Markdown reflect final metadata values.
- Hardened disk model to flag unknown-size volumes without treating them as low-space conditions.
