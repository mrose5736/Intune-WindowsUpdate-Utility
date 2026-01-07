# Intune Patch Monitor

A lightweight PowerShell GUI tool to monitor Windows Update patch status across an Intune Tenant.

## Features
- **View Update Rings**: Lists all Windows Update for Business configurations, including those managed by Autopatch.
- **Device Patch Status**: Select a ring/group to see detailed status of assigned devices.
- **Compliance Visibility**: Quickly see if devices are `Compliant`, `Pending`, or in `Error` state.
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
