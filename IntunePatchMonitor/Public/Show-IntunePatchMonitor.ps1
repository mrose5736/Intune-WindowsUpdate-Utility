function Show-IntunePatchMonitor {
    <#
    .SYNOPSIS
        Launches the Intune Patch Monitor GUI.
    .DESCRIPTION
        Installs prerequisites, authenticates to Microsoft Graph, and shows a WPF window that
        lists Windows Update for Business update rings and per-device patch status.
    #>
    [CmdletBinding()]
    param()

    Install-PatchMonitorPrerequisite
    Connect-PatchMonitorGraph

    Add-Type -AssemblyName PresentationFramework

    $ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    $XamlPath = Join-Path -Path $ModuleRoot -ChildPath 'Resources/MainWindow.xaml'
    [xml]$Xaml = Get-Content -Path $XamlPath -Raw

    $Reader = New-Object System.Xml.XmlNodeReader $Xaml
    try {
        $Window = [Windows.Markup.XamlReader]::Load($Reader)
    }
    catch {
        Write-PatchMonitorLog -Level ERROR -Message "Failed to parse XAML: $_"
        throw
    }

    function Get-Ctrl { param($Name) return $Window.FindName($Name) }

    $BtnRefresh         = Get-Ctrl 'BtnRefresh'
    $BtnExport          = Get-Ctrl 'BtnExport'
    $ListRings          = Get-Ctrl 'ListRings'
    $GridDevices        = Get-Ctrl 'GridDevices'
    $TxtStatus          = Get-Ctrl 'TxtStatus'
    $LblLastUpdated     = Get-Ctrl 'LblLastUpdated'
    $PbLoading          = Get-Ctrl 'PbLoading'
    $TxtCompliant       = Get-Ctrl 'TxtCompliant'
    $TxtError           = Get-Ctrl 'TxtError'
    $TxtPending         = Get-Ctrl 'TxtPending'
    $TxtTotal           = Get-Ctrl 'TxtTotal'
    $TxtFilter          = Get-Ctrl 'TxtFilter'
    $ChkAutoRefresh     = Get-Ctrl 'ChkAutoRefresh'
    $CmbRefreshInterval = Get-Ctrl 'CmbRefreshInterval'

    # --- State (scoped to this function/closure, not the module) ---
    $State = [PSCustomObject]@{
        RingLookup      = @{}
        DeviceCache     = @{}
        CurrentJob      = $null
        CurrentRingName = $null
        ModulePath      = (Join-Path -Path $ModuleRoot -ChildPath 'IntunePatchMonitor.psd1')
    }

    $JobTimer = New-Object System.Windows.Threading.DispatcherTimer
    $JobTimer.Interval = [TimeSpan]::FromMilliseconds(200)

    $AutoRefreshTimer = New-Object System.Windows.Threading.DispatcherTimer

    function Update-Status {
        param($Message)
        $TxtStatus.Dispatcher.Invoke([action]{ $TxtStatus.Text = $Message })
        Write-PatchMonitorLog -Message $Message -NoConsole
    }

    function Show-Loading {
        param($Show)
        $PbLoading.Dispatcher.Invoke([action]{
            $PbLoading.Visibility = if ($Show) { 'Visible' } else { 'Collapsed' }
        })
    }

    function Update-VisualSummary {
        param($Devices)
        $Summary = Get-PatchMonitorSummary -Devices $Devices
        $TxtTotal.Dispatcher.Invoke([action]{
            $TxtTotal.Text = $Summary.Total
            $TxtCompliant.Text = $Summary.Compliant
            $TxtError.Text = $Summary.Error
            $TxtPending.Text = $Summary.Pending
        })
    }

    function Sync-DeviceCache {
        Update-Status 'Caching Windows Device Details...'
        $State.DeviceCache = Get-PatchMonitorDeviceCache
        Update-Status "Cached $($State.DeviceCache.Count) Windows devices."
    }

    function Sync-Rings {
        Update-Status 'Loading Update Rings...'
        $ListRings.Items.Clear()
        $State.RingLookup.Clear()

        try {
            $Rings = Get-PatchMonitorRingList
            foreach ($Ring in $Rings) {
                $State.RingLookup[$Ring.DisplayName] = $Ring.Id
                $ListRings.Items.Add($Ring.DisplayName) | Out-Null
            }
            Update-Status "Ready. Found $($Rings.Count) rings."
            $LblLastUpdated.Content = (Get-Date).ToString('HH:mm:ss')
        }
        catch {
            Update-Status "Error loading rings: $_"
        }
    }

    # Applies the current filter text to whatever is bound to the grid right now.
    function Update-DeviceFilter {
        $View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($GridDevices.ItemsSource)
        if (-not $View) { return }

        $FilterText = $TxtFilter.Text
        if ([string]::IsNullOrWhiteSpace($FilterText)) {
            $View.Filter = $null
        }
        else {
            $View.Filter = {
                param($Item)
                $Needle = $FilterText
                @($Item.DeviceName, $Item.UserPrincipalName, $Item.DeviceModel, $Item.Status) -join ' ' -match [regex]::Escape($Needle)
            }.GetNewClosure()
        }
        $View.Refresh()
    }

    function Complete-RingDeviceJob {
        if (-not $State.CurrentJob) { return }

        $JobTimer.Stop()
        Show-Loading $false

        $DeviceList = Stop-PatchMonitorRingDeviceJob -Job $State.CurrentJob
        $State.CurrentJob = $null

        $ObservableDevices = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        if ($DeviceList) {
            $DeviceList | ForEach-Object { $ObservableDevices.Add($_) }
        }

        $GridDevices.ItemsSource = $ObservableDevices
        Update-DeviceFilter
        Update-Status "Loaded $($ObservableDevices.Count) devices for '$($State.CurrentRingName)'."
        Update-VisualSummary $ObservableDevices
    }

    # Single Tick handler registered once - fixes the original bug where a new handler was
    # added to the shared timer on every ring click, so old callbacks kept firing forever.
    $JobTimer.Add_Tick({
        if ($State.CurrentJob -and $State.CurrentJob.PowerShell.InvocationStateInfo.State -in @('Completed', 'Failed', 'Stopped')) {
            Complete-RingDeviceJob
        }
    })

    function Sync-RingDevices {
        param($RingName)
        $RingId = $State.RingLookup[$RingName]
        if (-not $RingId) { return }

        # If a previous job is still in flight (e.g. the user switched rings quickly), stop and
        # discard it rather than letting it race with the new one.
        if ($State.CurrentJob) {
            $JobTimer.Stop()
            Stop-PatchMonitorRingDeviceJob -Job $State.CurrentJob | Out-Null
            $State.CurrentJob = $null
        }

        Update-Status "Loading devices for '$RingName'..."
        Show-Loading $true
        $GridDevices.ItemsSource = $null

        if ($State.DeviceCache.Count -eq 0) {
            Sync-DeviceCache
        }

        $State.CurrentRingName = $RingName
        $State.CurrentJob = Start-PatchMonitorRingDeviceJob -RingId $RingId -DeviceCache $State.DeviceCache -ModulePath $State.ModulePath
        $JobTimer.Start()
    }

    # --- Filter wiring ---
    $TxtFilter.Add_TextChanged({ Update-DeviceFilter })

    # --- Auto-refresh wiring ---
    $AutoRefreshTimer.Add_Tick({
        if ($ListRings.SelectedItem) {
            Sync-RingDevices -RingName $ListRings.SelectedItem
        }
    })

    function Update-AutoRefreshTimer {
        $AutoRefreshTimer.Stop()
        if ($ChkAutoRefresh.IsChecked -eq $true) {
            $SelectedText = $CmbRefreshInterval.Text
            $Minutes = 5
            if ($SelectedText -match '(\d+)\s*min') { $Minutes = [int]$Matches[1] }
            $AutoRefreshTimer.Interval = [TimeSpan]::FromMinutes($Minutes)
            $AutoRefreshTimer.Start()
            Update-Status "Auto-refresh enabled every $Minutes minute(s)."
        }
        else {
            Update-Status 'Auto-refresh disabled.'
        }
    }
    $ChkAutoRefresh.Add_Click({ Update-AutoRefreshTimer })
    $CmbRefreshInterval.Add_SelectionChanged({ Update-AutoRefreshTimer })

    # --- Detail popup on double-click ---
    $GridDevices.Add_MouseDoubleClick({
        param($sender, $e)
        $Selected = $GridDevices.SelectedItem
        if (-not $Selected) { return }

        $Details = $Selected.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }
        [System.Windows.MessageBox]::Show(
            ($Details -join "`n"),
            "Device Details - $($Selected.DeviceName)",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    })

    # --- Export handler ---
    $BtnExport.Add_Click({
        if ($GridDevices.ItemsSource -and $GridDevices.ItemsSource.Count -gt 0) {
            $SaveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $SaveDialog.Filter = 'CSV File (*.csv)|*.csv'
            $SaveDialog.FileName = "IntunePatchReport_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

            if ($SaveDialog.ShowDialog() -eq $true) {
                try {
                    $GridDevices.ItemsSource | Select-Object DeviceName, UserPrincipalName, DeviceModel, OSVersion, Status, LastCheckin |
                        Export-Csv -Path $SaveDialog.FileName -NoTypeInformation
                    Update-Status "Exported to: $($SaveDialog.FileName)"
                    [System.Windows.MessageBox]::Show('Export Successful!', 'Export', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
                }
                catch {
                    Write-PatchMonitorLog -Level ERROR -Message "Failed to export CSV: $_"
                    [System.Windows.MessageBox]::Show("Failed to export: $_", 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                }
            }
        }
        else {
            [System.Windows.MessageBox]::Show('No data to export.', 'Warning', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    })

    # --- Remaining event handlers ---
    $BtnRefresh.Add_Click({ Sync-Rings })

    $ListRings.Add_SelectionChanged({
        if ($ListRings.SelectedItem) {
            Sync-RingDevices -RingName $ListRings.SelectedItem
        }
    })

    $Window.Add_Loaded({
        Sync-DeviceCache
        Sync-Rings
    })

    $Window.Add_Closed({
        $JobTimer.Stop()
        $AutoRefreshTimer.Stop()
        if ($State.CurrentJob) {
            Stop-PatchMonitorRingDeviceJob -Job $State.CurrentJob | Out-Null
        }
    })

    Write-PatchMonitorLog -Message "Log file for this session: $(Get-PatchMonitorLogPath)" -NoConsole
    $Window.ShowDialog() | Out-Null
}
