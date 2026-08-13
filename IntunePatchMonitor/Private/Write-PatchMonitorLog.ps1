function Write-PatchMonitorLog {
    <#
    .SYNOPSIS
        Writes a timestamped line to the Intune Patch Monitor log file and (optionally) the console.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$NoConsole
    )

    if (-not $Script:PatchMonitorLogPath) {
        # [System.IO.Path]::GetTempPath() (rather than $env:TEMP) works on both Windows and
        # non-Windows PowerShell hosts, e.g. when tests run outside a Windows GUI environment.
        $LogDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'IntunePatchMonitor'
        if (-not (Test-Path -Path $LogDir)) {
            New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        }
        $Script:PatchMonitorLogPath = Join-Path -Path $LogDir -ChildPath "IntunePatchMonitor_$(Get-Date -Format 'yyyyMMdd').log"
    }

    $Line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    try {
        Add-Content -Path $Script:PatchMonitorLogPath -Value $Line -ErrorAction Stop
    }
    catch {
        # If we can't write to the log file, fall back to console only - never let logging break the app.
        Write-Warning "Unable to write to log file '$Script:PatchMonitorLogPath': $_"
    }

    if (-not $NoConsole) {
        switch ($Level) {
            'WARN'  { Write-Warning $Message }
            'ERROR' { Write-Warning $Message }
            default { Write-Verbose $Message }
        }
    }
}

function Get-PatchMonitorLogPath {
    <#
    .SYNOPSIS
        Returns the path of the current session's log file (creating it if needed).
    #>
    [CmdletBinding()]
    param()

    if (-not $Script:PatchMonitorLogPath) {
        Write-PatchMonitorLog -Message 'Log initialized.' -NoConsole
    }
    return $Script:PatchMonitorLogPath
}
