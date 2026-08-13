function Get-PatchMonitorSummary {
    <#
    .SYNOPSIS
        Computes the Compliant/Error/Pending/Total tile counts for a set of device rows.
    .DESCRIPTION
        Pure function extracted from the UI update path so the categorization rules can be
        unit tested without a live window.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Devices
    )

    [PSCustomObject]@{
        Total     = $Devices.Count
        Compliant = ($Devices | Where-Object { $_.Status -match 'Compliant|Succeeded' } | Measure-Object).Count
        Error     = ($Devices | Where-Object { $_.Status -match 'Error|Failed' } | Measure-Object).Count
        Pending   = ($Devices | Where-Object { $_.Status -match 'Pending' } | Measure-Object).Count
    }
}
