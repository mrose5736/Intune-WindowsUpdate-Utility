function Get-RingDeviceStatus {
    <#
    .SYNOPSIS
        Joins raw Graph device-configuration status records with the cached device list.
    .DESCRIPTION
        Pure data-shaping function with no Graph/network calls, so it can be unit tested with
        mocked input. Takes the raw status objects returned by
        Get-MgDeviceManagementDeviceConfigurationDeviceStatus and the device cache built by
        Get-PatchMonitorDeviceCache, and returns the flattened rows the UI grid binds to.
    .PARAMETER Statuses
        Raw status objects. Each must expose DeviceDisplayName, UserName/UserPrincipalName,
        DeviceModel, Status and LastReportedDateTime.
    .PARAMETER DeviceCache
        Hashtable keyed by device name, values expose OSVersion.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Statuses,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$DeviceCache
    )

    $Results = [System.Collections.Generic.List[object]]::new()

    foreach ($Status in $Statuses) {
        $DevName = $Status.DeviceDisplayName
        if (-not $DevName) { continue }
        if (-not $DeviceCache.ContainsKey($DevName)) { continue }

        $CachedDev = $DeviceCache[$DevName]

        $Results.Add([PSCustomObject]@{
            DeviceName        = $DevName
            UserPrincipalName = if ($Status.UserName) { $Status.UserName } else { $Status.UserPrincipalName }
            DeviceModel       = if ($Status.DeviceModel) { $Status.DeviceModel } else { 'N/A' }
            OSVersion         = $CachedDev.OSVersion
            Status            = $Status.Status
            LastCheckin       = $Status.LastReportedDateTime
        })
    }

    return $Results.ToArray()
}
