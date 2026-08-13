function Get-PatchMonitorDeviceCache {
    <#
    .SYNOPSIS
        Fetches all Windows managed devices from Graph and returns them keyed by device name.
    #>
    [CmdletBinding()]
    param()

    $Cache = @{}
    try {
        $Devices = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'Windows'" -Property DeviceName, OSVersion, UserId -All
        foreach ($Dev in $Devices) {
            if ($Dev.DeviceName) {
                $Cache[$Dev.DeviceName] = $Dev
            }
        }
        Write-PatchMonitorLog -Message "Cached $($Cache.Count) Windows devices."
    }
    catch {
        Write-PatchMonitorLog -Level ERROR -Message "Failed to cache devices: $_"
    }

    return $Cache
}
