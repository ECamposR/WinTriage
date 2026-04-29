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
- Added first real diagnostic module for system, disk, and performance inventory plus initial rules and overview sections.
- Fixed initialization crashes caused by null-binding in helpers and added safe-step handling plus partial report recovery on fatal errors.
- Tightened report naming to use millisecond timestamps plus incremental suffixes, added `RunId`, and corrected build rendering in console and Markdown output.
- Updated report naming consistency and metadata/output formatting so repeated runs do not collide and build text is always populated or marked unknown.
