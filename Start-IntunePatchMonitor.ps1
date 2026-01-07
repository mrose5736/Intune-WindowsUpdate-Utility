<#
.SYNOPSIS
    Intune Patch Management Monitor GUI
.DESCRIPTION
    A GUI tool to view patch levels of devices across Intune Update Rings and Autopatch Groups.
.NOTES
    Author: Antigravity
    Version: 0.1
#>

# --- Dependencies Check ---
$RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement")
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Warning "Module '$Module' is missing. Attempting to install..."
        try {
            Install-Module $Module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to install module '$Module'. Please install it manually."
            exit
        }
    }
}

# --- Authentication ---
Write-Host "Authenticating to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All", "Group.Read.All" -ErrorAction Stop
    Write-Host "Successfully connected." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

# --- Load WPF Assemblies ---
Add-Type -AssemblyName PresentationFramework

# --- XAML Definition ---
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Patch Monitor" Height="600" Width="1000" Background="#F0F0F0">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header / Controls -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <Button Name="BtnRefresh" Content="Refresh Data" Width="120"/>
            <Label Content="Last Updated: " VerticalAlignment="Center" Margin="20,0,0,0"/>
            <Label Name="LblLastUpdated" Content="Never" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Main Content -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Update Rings List -->
            <GroupBox Header="Update Rings &amp; Groups" Grid.Column="0" Margin="0,0,10,0">
                <ListBox Name="ListRings" Margin="5"/>
            </GroupBox>

            <!-- Device Status Grid -->
            <GroupBox Header="Device Status" Grid.Column="1">
                <DataGrid Name="GridDevices" Margin="5" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#E6E6E6">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="*"/>
                        <DataGridTextColumn Header="User" Binding="{Binding UserPrincipalName}" Width="150"/>
                        <DataGridTextColumn Header="Model" Binding="{Binding DeviceModel}" Width="150"/>
                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                        <DataGridTextColumn Header="Last Check-in" Binding="{Binding LastCheckin}" Width="150"/>
                    </DataGrid.Columns>
                </DataGrid>
            </GroupBox>
        </Grid>

        <!-- Status Bar -->
        <StatusBar Grid.Row="2">
            <StatusBarItem>
                <TextBlock Name="TxtStatus" Text="Ready"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@

# --- XAML Loading Helper ---
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
try {
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
}
catch {
    Write-Error "Failed to parse XAML: $_"
    exit
}

# --- Connect Controls ---
$Props = $Window.GetType().GetProperties("Instance, Public") | Where-Object { $_.Name -eq "FindName" }
# Helper to get control by name
function Get-Ctrl { param($Name) return $Window.FindName($Name) }

$BtnRefresh = Get-Ctrl "BtnRefresh"
$ListRings = Get-Ctrl "ListRings"
$GridDevices = Get-Ctrl "GridDevices"
$TxtStatus = Get-Ctrl "TxtStatus"
$LblLastUpdated = Get-Ctrl "LblLastUpdated"

# --- Logic Functions ---

$Script:RingLookup = @{}

function Update-Status {
    param($Message)
    $TxtStatus.Dispatcher.Invoke([action]{ $TxtStatus.Text = $Message })
}

function Load-Rings {
    Update-Status "Loading Update Rings..."
    $ListRings.Items.Clear()
    $Script:RingLookup.Clear()
    
    try {
        # Fetch Windows Update for Business configurations
        $Rings = Get-MgDeviceManagementDeviceConfiguration -Filter "isof('microsoft.graph.windowsUpdateForBusinessConfiguration')" -Property Id, DisplayName
        
        foreach ($Ring in $Rings) {
            $Script:RingLookup[$Ring.DisplayName] = $Ring.Id
            $ListRings.Items.Add($Ring.DisplayName) | Out-Null
        }
        
        Update-Status "Ready. Found $($Rings.Count) rings."
        $LblLastUpdated.Content = (Get-Date).ToString("HH:mm:ss")
    }
    catch {
        Update-Status "Error loading rings: $_"
        Write-Error $_
    }
}

function Load-RingDevices {
    param($RingName)
    $RingId = $Script:RingLookup[$RingName]
    if (-not $RingId) { return }

    Update-Status "Loading devices for '$RingName'..."
    $GridDevices.ItemsSource = $null
    
    try {
        # Fetch device statuses for this config
        # Note: This checks the assignment status (Succeeded/Error/Pending) for the policy
        $Statuses = Get-MgDeviceManagementDeviceConfigurationDeviceStatus -DeviceConfigurationId $RingId -All
        
        $DeviceList = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        foreach ($Status in $Statuses) {
            # We assume UserPrincipalName/DeviceName might be populated, depending on the expanding
            # If not, we might need to fetch device details separately if names are missing
            
            $DeviceList.Add([PSCustomObject]@{
                DeviceName      = $Status.DeviceName
                UserPrincipalName = $Status.UserPrincipalName
                DeviceModel     = $Status.DeviceModel
                Status          = $Status.ComplianceStatus
                LastCheckin     = $Status.LastReportedDateTime
            })
        }
        
        $GridDevices.ItemsSource = $DeviceList
        Update-Status "Loaded $($DeviceList.Count) devices for '$RingName'."
    }
    catch {
        Update-Status "Error loading devices: $_"
        Write-Error $_
    }
}

# --- Event Handlers ---

$BtnRefresh.Add_Click({
    Load-Rings
})

$ListRings.Add_SelectionChanged({
    if ($ListRings.SelectedItem) {
        Load-RingDevices -RingName $ListRings.SelectedItem
    }
})

# --- Initial Load ---
# Delay slightly to let window render
$Window.Add_Loaded({
    Load-Rings
})

# --- Show Window ---
$Window.ShowDialog() | Out-Null
