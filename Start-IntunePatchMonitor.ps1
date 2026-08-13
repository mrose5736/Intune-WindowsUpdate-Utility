<#
.SYNOPSIS
    Intune Patch Management Monitor GUI
.DESCRIPTION
    A GUI tool to view patch levels of Windows devices across Intune Update Rings via Microsoft Graph.
    This is a thin launcher - all logic lives in the IntunePatchMonitor module alongside this script.
.NOTES
    Author: Antigravity
    Version: 0.2.0
#>

#Requires -Version 5.1

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'IntunePatchMonitor/IntunePatchMonitor.psd1'
Import-Module -Name $ModulePath -Force -ErrorAction Stop

Show-IntunePatchMonitor
