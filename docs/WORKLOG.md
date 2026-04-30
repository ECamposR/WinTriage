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
- Added first real System/Application event module with CSV export, event overview Markdown, and initial shutdown/bugcheck/service/app-crash findings plus basic correlations.
- Refined event classification to reduce WER false positives, added process extraction and crash summaries, and made Markdown counts more robust.
- Hotfixed a PowerShell parser error in event sorting and bumped the script version to `0.3.2`.
- Replaced invalid `Sort-Object` syntax with PowerShell 5.1-compatible property hashtables to keep the script parseable.
- Refined event process extraction for Spanish Application Error 1000 messages, stabilized Markdown `Exit code`, and reprioritized recent important events away from non-critical WER noise.
- Tightened Spanish `Application Error 1000` parsing, added an internal warning when process extraction fails, and reduced WER noise in `Recent Important Events`.
- Restored APPCRASH/Application Error enrichment for `AffectedProcess`, `AffectedPath`, and `ProcessSourcePattern`, then rebalanced crash summaries and event prioritization.
- Repaired regression in APPCRASH/Application Error parsing so enriched process fields are populated again and crash summaries no longer collapse to `Unknown`.
- Added a dedicated `Application Error 1000` parser path for Spanish faulting application names and paths so `mxdhcp.exe` style crashes enrich correctly again.
- Bumped script version to `0.3.6` after restoring the Spanish `Application Error 1000` parser path.
- Refined the Spanish `Application Error 1000` parser to use line-oriented matching and improved the internal warning path for future regressions.
- Bumped script version to `0.3.7` after correcting the last pending kernel-uptime and boot-state semantics work.
- Added `-SelfTestEventParser` to validate `Get-WTEventProcessName` against Spanish Application Error and WER crash samples before running the full triage.
- Bumped script version to `0.3.8` after hardening the Spanish `Application Error 1000` parser path and adding parser self-tests.
- Added a read-only Windows Update, CBS, servicing, Store and Edge/WebView diagnostics module with pending reboot checks, hotfix inventory, servicing WER classification and update-focused report sections.
- Bumped script version to `0.4.0` after integrating the Windows Update / servicing diagnostics module.
- Corrected Windows Update pending reboot semantics so `PendingFileRenameOperations` is treated as a generic restart signal instead of a Windows Update reboot requirement.
- Added explicit `PendingRebootSource`, `GenericRestartSignal`, and pending file rename sampling fields to keep update metadata clean and more actionable.
- Suppressed accidental pipeline output from the Updates collector and top-level framework findings so `Raw.Updates` and `Normalized.Updates` remain clean objects.
- Bumped script version to `0.4.1` after fixing pending reboot false positives and report-object contamination.
