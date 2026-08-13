function Start-PatchMonitorRingDeviceJob {
    <#
    .SYNOPSIS
        Starts a background runspace that fetches device configuration status for one update ring.
    .DESCRIPTION
        Runs the Graph call off the UI thread so the WPF window stays responsive. The caller is
        responsible for polling the returned job object's PowerShell.InvocationStateInfo.State and
        calling Stop-PatchMonitorRingDeviceJob to collect results and dispose resources exactly once.
    .PARAMETER RingId
        The device configuration (update ring) id to query.
    .PARAMETER DeviceCache
        Hashtable keyed by device name, as built by Get-PatchMonitorDeviceCache.
    .PARAMETER ModulePath
        Path to the IntunePatchMonitor module's root .psd1/.psm1, so the background runspace can
        import it and reuse Get-RingDeviceStatus rather than duplicating the join logic.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RingId,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$DeviceCache,

        [Parameter(Mandatory)]
        [string]$ModulePath
    )

    # Grab the token on the caller's (UI) thread - GraphSession is tied to the runspace that
    # authenticated, and reading it here keeps the access-token handling in one place.
    $Token = Get-PatchMonitorAccessToken

    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.Open()
    $PowerShell = [powershell]::Create()
    $PowerShell.Runspace = $Runspace

    $ScriptBlock = {
        param($RingId, $DeviceCache, $AccessToken, $ModulePath)

        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
        Import-Module $ModulePath -Force -ErrorAction Stop

        if (-not $AccessToken) {
            throw 'No access token was provided to the background job.'
        }

        # Get-PatchMonitorAccessToken on the UI thread already normalized this to a plain
        # string; Connect-MgGraph -AccessToken accepts a plain string in all supported
        # module versions, so no further type juggling is needed here.
        Connect-MgGraph -AccessToken $AccessToken -ErrorAction Stop | Out-Null

        $Statuses = Get-MgDeviceManagementDeviceConfigurationDeviceStatus -DeviceConfigurationId $RingId -All -ErrorAction Stop

        return Get-RingDeviceStatus -Statuses $Statuses -DeviceCache $DeviceCache
    }

    $PowerShell.AddScript($ScriptBlock).
        AddArgument($RingId).
        AddArgument($DeviceCache).
        AddArgument($Token).
        AddArgument($ModulePath) | Out-Null

    $AsyncResult = $PowerShell.BeginInvoke()

    [PSCustomObject]@{
        PowerShell  = $PowerShell
        Runspace    = $Runspace
        AsyncResult = $AsyncResult
        RingId      = $RingId
    }
}

function Stop-PatchMonitorRingDeviceJob {
    <#
    .SYNOPSIS
        Collects the result of a job started by Start-PatchMonitorRingDeviceJob and disposes it.
    .DESCRIPTION
        Safe to call on a still-running job (it will be stopped) or a completed one (its result is
        returned). Always disposes the PowerShell instance and runspace, so this is the single
        cleanup path callers should use instead of calling EndInvoke/Dispose themselves.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Job
    )

    $Result = $null
    try {
        $State = $Job.PowerShell.InvocationStateInfo.State
        if ($State -eq 'Completed') {
            $Result = $Job.PowerShell.EndInvoke($Job.AsyncResult)
        }
        elseif ($State -notin @('Failed', 'Stopped')) {
            $Job.PowerShell.Stop()
        }

        if ($Job.PowerShell.Streams.Error.Count -gt 0) {
            foreach ($Err in $Job.PowerShell.Streams.Error) {
                Write-PatchMonitorLog -Level ERROR -Message "Background job error: $($Err.Exception.Message)"
            }
        }
    }
    catch {
        Write-PatchMonitorLog -Level ERROR -Message "Failed to collect background job result: $_"
    }
    finally {
        $Job.PowerShell.Dispose()
        $Job.Runspace.Dispose()
    }

    return $Result
}
