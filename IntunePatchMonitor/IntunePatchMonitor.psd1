@{
    RootModule        = 'IntunePatchMonitor.psm1'
    ModuleVersion     = '0.2.0'
    GUID              = '74cb432f-29f5-48e9-ac6e-02f7557cc8e4'
    Author            = 'Antigravity'
    Description       = 'WPF GUI for monitoring Windows Update for Business patch status across an Intune tenant via Microsoft Graph.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Show-IntunePatchMonitor')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('Intune', 'WindowsUpdate', 'MicrosoftGraph', 'WPF')
            ProjectUri = ''
        }
    }
}
