#Requires -Modules Pester

BeforeAll {
    $ModulePath = Join-Path -Path $PSScriptRoot -ChildPath '../IntunePatchMonitor/IntunePatchMonitor.psd1'
    Import-Module -Name $ModulePath -Force
}

Describe 'Get-RingDeviceStatus' {
    InModuleScope IntunePatchMonitor {

        It 'joins a status record with its cached device by device name' {
            $DeviceCache = @{
                'PC-001' = [PSCustomObject]@{ DeviceName = 'PC-001'; OSVersion = '10.0.22631' }
            }
            $Statuses = @(
                [PSCustomObject]@{
                    DeviceDisplayName = 'PC-001'
                    UserName          = 'alice@contoso.com'
                    UserPrincipalName = $null
                    DeviceModel       = 'Surface Laptop 5'
                    Status            = 'compliant'
                    LastReportedDateTime = '2026-08-01T12:00:00Z'
                }
            )

            $Result = Get-RingDeviceStatus -Statuses $Statuses -DeviceCache $DeviceCache

            $Result.Count | Should -Be 1
            $Result[0].DeviceName | Should -Be 'PC-001'
            $Result[0].UserPrincipalName | Should -Be 'alice@contoso.com'
            $Result[0].OSVersion | Should -Be '10.0.22631'
            $Result[0].Status | Should -Be 'compliant'
        }

        It 'falls back to UserPrincipalName when UserName is absent' {
            $DeviceCache = @{ 'PC-002' = [PSCustomObject]@{ DeviceName = 'PC-002'; OSVersion = '10.0.19045' } }
            $Statuses = @(
                [PSCustomObject]@{
                    DeviceDisplayName = 'PC-002'
                    UserName          = $null
                    UserPrincipalName = 'bob@contoso.com'
                    DeviceModel       = $null
                    Status            = 'error'
                    LastReportedDateTime = '2026-08-01T12:00:00Z'
                }
            )

            $Result = Get-RingDeviceStatus -Statuses $Statuses -DeviceCache $DeviceCache

            $Result[0].UserPrincipalName | Should -Be 'bob@contoso.com'
            $Result[0].DeviceModel | Should -Be 'N/A'
        }

        It 'skips statuses for devices that are not in the cache' {
            $DeviceCache = @{ 'PC-001' = [PSCustomObject]@{ DeviceName = 'PC-001'; OSVersion = '10.0.22631' } }
            $Statuses = @(
                [PSCustomObject]@{ DeviceDisplayName = 'PC-999'; UserName = 'x'; DeviceModel = 'y'; Status = 'compliant'; LastReportedDateTime = $null }
            )

            $Result = Get-RingDeviceStatus -Statuses $Statuses -DeviceCache $DeviceCache

            $Result.Count | Should -Be 0
        }

        It 'skips statuses with no device display name' {
            $DeviceCache = @{ 'PC-001' = [PSCustomObject]@{ DeviceName = 'PC-001'; OSVersion = '10.0.22631' } }
            $Statuses = @(
                [PSCustomObject]@{ DeviceDisplayName = $null; UserName = 'x'; DeviceModel = 'y'; Status = 'compliant'; LastReportedDateTime = $null }
            )

            $Result = Get-RingDeviceStatus -Statuses $Statuses -DeviceCache $DeviceCache

            $Result.Count | Should -Be 0
        }

        It 'returns an empty array when there are no statuses' {
            $Result = Get-RingDeviceStatus -Statuses @() -DeviceCache @{}
            $Result.Count | Should -Be 0
        }
    }
}

Describe 'Get-PatchMonitorSummary' {
    InModuleScope IntunePatchMonitor {

        It 'categorizes devices into compliant, error and pending buckets' {
            $Devices = @(
                [PSCustomObject]@{ Status = 'compliant' }
                [PSCustomObject]@{ Status = 'Succeeded' }
                [PSCustomObject]@{ Status = 'error' }
                [PSCustomObject]@{ Status = 'Failed' }
                [PSCustomObject]@{ Status = 'pending' }
                [PSCustomObject]@{ Status = 'unknown' }
            )

            $Summary = Get-PatchMonitorSummary -Devices $Devices

            $Summary.Total | Should -Be 6
            $Summary.Compliant | Should -Be 2
            $Summary.Error | Should -Be 2
            $Summary.Pending | Should -Be 1
        }

        It 'returns all zeros for an empty device list' {
            $Summary = Get-PatchMonitorSummary -Devices @()

            $Summary.Total | Should -Be 0
            $Summary.Compliant | Should -Be 0
            $Summary.Error | Should -Be 0
            $Summary.Pending | Should -Be 0
        }
    }
}

Describe 'Get-PatchMonitorAccessToken' {
    InModuleScope IntunePatchMonitor {

        It 'returns null when the Graph session has no current access token' {
            Mock -CommandName Write-PatchMonitorLog -MockWith {}

            # In a real host without an authenticated Graph session, the type reference either
            # throws (module not loaded) or CurrentAccessToken is $null - both should yield $null.
            $Result = Get-PatchMonitorAccessToken
            $Result | Should -BeNullOrEmpty
        }
    }
}
