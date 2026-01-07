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
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal">
                <Button Name="BtnRefresh" Content="Refresh Data" Width="120"/>
                <Button Name="BtnExport" Content="Export CSV" Width="120" Background="#107C10"/>
                <Label Content="Last Updated: " VerticalAlignment="Center" Margin="20,0,0,0"/>
                <Label Name="LblLastUpdated" Content="Never" VerticalAlignment="Center"/>
            </StackPanel>
            
            <!-- Progress Bar -->
            <ProgressBar Name="PbLoading" Height="4" IsIndeterminate="True" Visibility="Collapsed" Margin="0,5,0,0" Foreground="#0078D7"/>
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
            <Grid Grid.Column="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Visual Summary Dashboard -->
                <Border Grid.Row="0" Margin="0,0,0,10" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="4" Padding="10">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                        <Border Background="#107C10" CornerRadius="4" Padding="10,5" Margin="5">
                             <TextBlock Foreground="White" FontWeight="Bold"><Run Text="Compliant: "/><Run Name="TxtCompliant" Text="0"/></TextBlock>
                        </Border>
                        <Border Background="#D83B01" CornerRadius="4" Padding="10,5" Margin="5">
                             <TextBlock Foreground="White" FontWeight="Bold"><Run Text="Error: "/><Run Name="TxtError" Text="0"/></TextBlock>
                        </Border>
                        <Border Background="#FFB900" CornerRadius="4" Padding="10,5" Margin="5">
                             <TextBlock Foreground="Black" FontWeight="Bold"><Run Text="Pending: "/><Run Name="TxtPending" Text="0"/></TextBlock>
                        </Border>
                         <Border Background="#666666" CornerRadius="4" Padding="10,5" Margin="5">
                             <TextBlock Foreground="White" FontWeight="Bold"><Run Text="Total: "/><Run Name="TxtTotal" Text="0"/></TextBlock>
                        </Border>
                    </StackPanel>
                </Border>

                <!-- Grid -->
                <GroupBox Header="Device Status" Grid.Row="1">
                    <DataGrid Name="GridDevices" Margin="5" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#E6E6E6">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="*"/>
                            <DataGridTextColumn Header="User" Binding="{Binding UserPrincipalName}" Width="150"/>
                            <DataGridTextColumn Header="Model" Binding="{Binding DeviceModel}" Width="150"/>
                            <DataGridTextColumn Header="OS Version" Binding="{Binding OSVersion}" Width="120"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                            <DataGridTextColumn Header="Last Check-in" Binding="{Binding LastCheckin}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </GroupBox>
            </Grid>
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
$BtnExport = Get-Ctrl "BtnExport"
$ListRings = Get-Ctrl "ListRings"
$GridDevices = Get-Ctrl "GridDevices"
$TxtStatus = Get-Ctrl "TxtStatus"
$LblLastUpdated = Get-Ctrl "LblLastUpdated"
$PbLoading = Get-Ctrl "PbLoading"
$TxtCompliant = Get-Ctrl "TxtCompliant"
$TxtError = Get-Ctrl "TxtError"
$TxtPending = Get-Ctrl "TxtPending"
$TxtTotal = Get-Ctrl "TxtTotal"

# --- Logic Functions ---

$Script:RingLookup = @{}
$Script:DeviceCache = @{}

# Timer for Async Jobs
$Script:JobTimer = New-Object System.Windows.Threading.DispatcherTimer
$Script:JobTimer.Interval = [TimeSpan]::FromMilliseconds(200)

function Update-Status {
    param($Message)
    $TxtStatus.Dispatcher.Invoke([action]{ $TxtStatus.Text = $Message })
}

function Show-Loading {
    param($Show)
    $PbLoading.Dispatcher.Invoke([action]{ 
        if ($Show) { $PbLoading.Visibility = "Visible" } else { $PbLoading.Visibility = "Collapsed" }
    })
}

function Update-VisualSummary {
    param($Devices)
    
    $TxtTotal.Dispatcher.Invoke([action]{
        $TxtTotal.Text = $Devices.Count
        $TxtCompliant.Text = ($Devices | Where-Object { $_.Status -match "Compliant|Succeeded" }).Count
        $TxtError.Text = ($Devices | Where-Object { $_.Status -match "Error|Failed" }).Count
        $TxtPending.Text = ($Devices | Where-Object { $_.Status -match "Pending" }).Count
    })
}

function Load-DeviceCache {
    Update-Status "Caching Windows Device Details..."
    $Script:DeviceCache.Clear()
    try {
        # Fetch only Windows devices, get OS version
        $Devices = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'Windows'" -Property DeviceName, OSVersion, UserId
        foreach ($Dev in $Devices) {
            if ($Dev.DeviceName) {
                $Script:DeviceCache[$Dev.DeviceName] = $Dev
            }
        }
        Update-Status "Cached $($Devices.Count) Windows devices."
    }
    catch {
        Write-Warning "Failed to cache devices: $_"
    }
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
    Show-Loading $true
    $GridDevices.ItemsSource = $null
    
    # Ensure cache is populated (lazy load if needed, but we do it on start usually)
    if ($Script:DeviceCache.Count -eq 0) {
       Load-DeviceCache
    }

    # Get Access Token from main session to pass to background job
    # Try different methods to retrieve the token safely
    $Token = $null
    try {
        if ([Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.CurrentAccessToken) {
             # SecureString to PlainText if needed, or mostly it's a bearer string
             $Token = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.CurrentAccessToken
        }
    } catch {
        # Fallback or older versions
    }

    # Pass data to Runspace
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.Open()
    $PowerShell = [powershell]::Create()
    $PowerShell.Runspace = $Runspace
    
    $ScriptBlock = {
        param($RingId, $DeviceCache, $AccessToken)
        
        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
        
        # Re-authenticate in this runspace using the passed token
        if ($AccessToken) {
            # Convert SecureString to string if necessary, but usually Connect-MgGraph takes SecureString or String depending on version
            # We assume it accepts the object we passed.
            try {
                Connect-MgGraph -AccessToken $AccessToken -ErrorAction Stop
            } catch {
                Write-Error "Failed to authenticate in background job: $_"
                return
            }
        } else {
             Write-Error "No access token provided to background job."
             return
        }

        # Fetch device statuses
        try {
            $Statuses = Get-MgDeviceManagementDeviceConfigurationDeviceStatus -DeviceConfigurationId $RingId -All -ErrorAction Stop
        } catch {
            Write-Error "Failed to fetch statuses: $_"
            return
        }
        
        $Results = @()
        foreach ($Status in $Statuses) {
            $DevName = $Status.DeviceDisplayName
            
            # Check Filter
            if ($DeviceCache.ContainsKey($DevName)) {
                $CachedDev = $DeviceCache[$DevName]
                
                $Results += [PSCustomObject]@{
                    DeviceName      = $DevName
                    UserPrincipalName = if ($Status.UserName) { $Status.UserName } else { $Status.UserPrincipalName }
                    DeviceModel     = if ($Status.DeviceModel) { $Status.DeviceModel } else { "N/A" }
                    OSVersion       = $CachedDev.OSVersion
                    Status          = $Status.Status
                    LastCheckin     = $Status.LastReportedDateTime
                }
            }
        }
        return $Results
    }

    $PowerShell.AddScript($ScriptBlock).AddArgument($RingId).AddArgument($Script:DeviceCache).AddArgument($Token) | Out-Null
    
    $AsyncResult = $PowerShell.BeginInvoke()
    
    # Handle Completion with Timer
    $Script:CurrentPowerShell = $PowerShell
    $Script:JobTimer.Add_Tick({
        param($sender, $e)
        
        if ($Script:CurrentPowerShell.InvocationStateInfo.State -eq 'Completed' -or 
            $Script:CurrentPowerShell.InvocationStateInfo.State -eq 'Failed') {
            
            $Script:JobTimer.Stop()
            $Script:JobTimer.Remove_Tick($Script:JobTimer.Tick) 
            
            Show-Loading $false
            
            try {
                $DeviceList = $Script:CurrentPowerShell.EndInvoke($AsyncResult)
                
                # Check for errors in the stream
                if ($Script:CurrentPowerShell.Streams.Error.Count -gt 0) {
                    foreach ($Err in $Script:CurrentPowerShell.Streams.Error) {
                        Write-Warning "Background Job Error: $($Err.Exception.Message)"
                    }
                     Update-Status "Error loading data. See console for details."
                }
                
                $Script:CurrentPowerShell.Dispose()
                
                $ObservableDevices = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                if ($DeviceList) {
                    $DeviceList | ForEach-Object { $ObservableDevices.Add($_) }
                }

                $GridDevices.ItemsSource = $ObservableDevices
                Update-Status "Loaded $($ObservableDevices.Count) devices for '$RingName'."
                Update-VisualSummary $ObservableDevices
            }
            catch {
                Update-Status "Critical Error in background job: $_"
            }
        }
    })
    $Script:JobTimer.Start()
}

# Export Handler
$BtnExport.Add_Click({
    if ($GridDevices.ItemsSource -and $GridDevices.ItemsSource.Count -gt 0) {
        $SaveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $SaveDialog.Filter = "CSV File (*.csv)|*.csv"
        $SaveDialog.FileName = "IntunePatchReport_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        
        if ($SaveDialog.ShowDialog() -eq $true) {
            try {
                $GridDevices.ItemsSource | Select-Object DeviceName, UserPrincipalName, Model, OSVersion, Status, LastCheckin | Export-Csv -Path $SaveDialog.FileName -NoTypeInformation
                Update-Status "Exported to: $($SaveDialog.FileName)"
                [System.Windows.MessageBox]::Show("Export Successful!", "Export", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    } else {
         [System.Windows.MessageBox]::Show("No data to export.", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
})

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
    Load-DeviceCache
    Load-Rings
})

# --- Show Window ---
$Window.ShowDialog() | Out-Null
