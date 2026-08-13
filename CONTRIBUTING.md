# Contributing

Thanks for considering a contribution to Intune Patch Monitor.

## Getting set up
This is a Windows-only PowerShell WPF tool (it needs `PresentationFramework`), so development and
manual testing of the GUI require Windows PowerShell 5.1 or PowerShell 7+ on Windows. The pure
logic in `IntunePatchMonitor/Private/` (device/status matching, summary counts) has no Windows or
Graph dependency and can be edited/tested cross-platform.

## Project layout
See the "Project Structure" section of [README.md](README.md). In short:
- `Public/Show-IntunePatchMonitor.ps1` owns the window, controls, and event wiring.
- `Private/` holds everything that doesn't need a live window: Graph calls, data shaping, logging,
  background job orchestration.
- Keep Graph/UI-coupled code out of `Private/` functions that are meant to be pure and testable
  (`Get-RingDeviceStatus`, `Get-PatchMonitorSummary`) - if a change needs live data, add a new
  function rather than mixing concerns into an existing pure one.

## Making a change
1. Create a branch off `main`.
2. Make your change. If you touch logic in `Private/`, prefer adding/extending Pester tests in
   `Tests/IntunePatchMonitor.Tests.ps1` over relying on manual GUI testing alone.
3. Run the linter and tests locally:
   ```powershell
   Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
   Invoke-Pester -Path .\Tests
   ```
4. If your change affects the GUI, launch `Start-IntunePatchMonitor.ps1` on Windows and exercise the
   golden path (load rings, select a ring, filter, export, auto-refresh) plus any edge case your
   change touches.
5. Open a pull request describing what changed and why. CI runs the same lint/test steps
   automatically.

## Reporting issues
Open a GitHub issue with:
- What you expected vs. what happened
- PowerShell version (`$PSVersionTable`) and Graph module versions
  (`Get-Module Microsoft.Graph.* -ListAvailable`)
- The relevant lines from `%TEMP%\IntunePatchMonitor\IntunePatchMonitor_yyyyMMdd.log`, if available
