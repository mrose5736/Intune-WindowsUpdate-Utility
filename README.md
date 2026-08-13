# Intune Patch Monitor

A lightweight PowerShell GUI tool to monitor Windows Update patch status across an Intune Tenant.

## Features
- **View Update Rings**: Lists all Windows Update for Business configurations, including those managed by Autopatch.
- **Device Patch Status**: Select a ring/group to see detailed status of assigned devices.
- **Compliance Visibility**: Quickly see if devices are `Compliant`, `Pending`, or in `Error` state, with a live summary dashboard.
- **Filter/Search**: Narrow the device grid by name, user, model, or status as you type.
- **Auto-Refresh**: Optionally re-poll the selected ring on an interval (1/5/10/15 minutes).
- **Device Details**: Double-click a row to see the full raw status record for that device.
- **CSV Export**: Export the currently loaded device grid to CSV.
- **No Installation Required**: Runs as a standalone PowerShell script.

## Prerequisites
- Windows 10/11
- PowerShell 5.1 or 7+
- Microsoft Graph PowerShell Modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.DeviceManagement`. (The script will attempt to install them if missing).
- An Intune Administrator account (or appropriate delegated permissions).

## Usage
1. Open PowerShell.
2. Navigate to the directory containing the script.
3. Run:
   ```powershell
   .\Start-IntunePatchMonitor.ps1
   ```
4. Sign in with your Intune credentials when prompted.
5. Select an Update Ring from the left panel to load device statuses.

## Permissions
The script requires the following Graph API scopes:
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `Group.Read.All`

## Project Structure
```
Start-IntunePatchMonitor.ps1          Thin launcher - imports the module and starts the GUI
IntunePatchMonitor/
  IntunePatchMonitor.psd1             Module manifest
  IntunePatchMonitor.psm1             Loads Public/Private functions
  Public/Show-IntunePatchMonitor.ps1  GUI wiring (window, event handlers)
  Private/                            Graph calls, data shaping, logging, background jobs
  Resources/MainWindow.xaml           WPF layout
Tests/IntunePatchMonitor.Tests.ps1    Pester tests for the pure logic (no live Graph calls)
```

Splitting the tool this way keeps the parts that need a live Graph connection and a WPF window
separate from the pure data-shaping logic (device/status matching, summary counts), so the latter
can be unit tested in CI without a Windows GUI environment.

## Logging
Each run writes a timestamped log to `%TEMP%\IntunePatchMonitor\IntunePatchMonitor_yyyyMMdd.log`,
useful for diagnosing issues without needing to capture console output from the GUI process.

## Development

### Running tests
```powershell
Install-Module Pester -MinimumVersion 5.0.0 -Scope CurrentUser
Invoke-Pester -Path .\Tests
```

### Linting
```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
```

Both run automatically in CI (`.github/workflows/ci.yml`) on every push and pull request.

See [CONTRIBUTING.md](CONTRIBUTING.md) for more on contributing changes.

## Known Follow-ups
A couple of ideas from the original review are intentionally **not** implemented here because they
need things this repo/session doesn't have access to:
- **Explicit Autopatch Groups support**: the current ring list covers Windows Update for Business
  configurations (which includes Autopatch-managed rings), but querying Autopatch Groups directly
  would require the `Microsoft.Graph.Beta` module and a tenant with Autopatch enabled to verify
  against - worth a follow-up once that's available.
- **PowerShell Gallery packaging**: the module manifest (`IntunePatchMonitor.psd1`) is publish-ready,
  but actually publishing to PSGallery requires an API key/account decision that belongs to whoever
  owns this project, not something to do unilaterally.
