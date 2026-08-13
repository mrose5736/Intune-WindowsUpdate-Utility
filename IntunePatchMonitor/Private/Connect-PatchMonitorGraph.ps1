function Connect-PatchMonitorGraph {
    <#
    .SYNOPSIS
        Connects the current session to Microsoft Graph with the scopes Intune Patch Monitor needs.
    #>
    [CmdletBinding()]
    param(
        [string[]]$Scopes = @(
            'DeviceManagementConfiguration.Read.All',
            'DeviceManagementManagedDevices.Read.All',
            'Group.Read.All'
        )
    )

    Write-PatchMonitorLog -Message 'Authenticating to Microsoft Graph...'
    try {
        Connect-MgGraph -Scopes $Scopes -ErrorAction Stop | Out-Null
        Write-PatchMonitorLog -Message 'Successfully connected to Microsoft Graph.'
    }
    catch {
        Write-PatchMonitorLog -Level ERROR -Message "Failed to connect to Microsoft Graph: $_"
        throw
    }
}

function Get-PatchMonitorAccessToken {
    <#
    .SYNOPSIS
        Retrieves the current Graph session's access token as a plain string, for use only when
        an authenticated context must be handed to a background runspace.
    .DESCRIPTION
        Microsoft.Graph module versions differ on whether CurrentAccessToken is exposed as a
        SecureString or a plain String. This helper checks the actual runtime type instead of
        assuming one or the other, and always returns a plain string (or $null) so callers have
        a single, predictable contract.
    #>
    [CmdletBinding()]
    param()

    try {
        $RawToken = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.CurrentAccessToken
    }
    catch {
        Write-PatchMonitorLog -Level WARN -Message "Unable to read current Graph access token: $_"
        return $null
    }

    if (-not $RawToken) {
        return $null
    }

    if ($RawToken -is [System.Security.SecureString]) {
        # ConvertFrom-SecureString -AsPlainText requires PS 7+; use the Marshal-based
        # conversion instead so this works on Windows PowerShell 5.1 too.
        $Bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RawToken)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($Bstr)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Bstr)
        }
    }

    if ($RawToken -is [string]) {
        return $RawToken
    }

    Write-PatchMonitorLog -Level WARN -Message "Unexpected access token type: $($RawToken.GetType().FullName)"
    return $null
}
