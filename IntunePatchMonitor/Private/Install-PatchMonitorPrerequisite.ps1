function Install-PatchMonitorPrerequisite {
    <#
    .SYNOPSIS
        Ensures the Microsoft Graph modules required by Intune Patch Monitor are available.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DeviceManagement')
    )

    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-PatchMonitorLog -Level WARN -Message "Module '$Module' is missing. Attempting to install..."
            try {
                Install-Module -Name $Module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-PatchMonitorLog -Message "Installed module '$Module'."
            }
            catch {
                Write-PatchMonitorLog -Level ERROR -Message "Failed to install module '$Module': $_"
                throw "Failed to install required module '$Module'. Please install it manually with: Install-Module $Module -Scope CurrentUser"
            }
        }
    }
}
