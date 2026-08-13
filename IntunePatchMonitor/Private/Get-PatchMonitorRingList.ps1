function Get-PatchMonitorRingList {
    <#
    .SYNOPSIS
        Fetches Windows Update for Business update ring configurations from Graph.
    .OUTPUTS
        Array of objects with DisplayName and Id.
    #>
    [CmdletBinding()]
    param()

    try {
        $Rings = Get-MgDeviceManagementDeviceConfiguration -Filter "isof('microsoft.graph.windowsUpdateForBusinessConfiguration')" -Property Id, DisplayName -All -ErrorAction Stop
        Write-PatchMonitorLog -Message "Found $($Rings.Count) update rings."
        return $Rings
    }
    catch {
        Write-PatchMonitorLog -Level ERROR -Message "Failed to load update rings: $_"
        throw
    }
}
