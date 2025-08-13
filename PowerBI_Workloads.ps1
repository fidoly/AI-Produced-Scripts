<#
.SYNOPSIS
    A PowerShell script with a GUI to retrieve details about a Power BI tenant.
.DESCRIPTION
    This script provides a user-friendly interface to connect to a Power BI service
    and fetch various details like workspaces, capacities, datasets, reports, and users.
    It uses a non-blocking login flow and can generate a high-level executive summary.
.NOTES
    Author: Gemini
    Version: 2.0
    Requires: PowerShell 7+, Windows PowerShell
.EXAMPLE
    .\Get-PowerBIDetailsGUI.ps1
    This will launch the GUI application.
#>

#region WPF GUI Definition
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Power BI Tenant Details Retriever" Height="700" Width="900"
        WindowStartupLocation="CenterScreen" Icon="https://www.microsoft.com/favicon.ico?v2">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Power BI Tenant Details Retriever" FontSize="24" FontWeight="Bold" Foreground="#0078D4"/>
            <TextBlock Text="Connect to your Power BI tenant and select the details you want to retrieve." Foreground="Gray" TextWrapping="Wrap"/>
        </StackPanel>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,15,0">
                <StackPanel>
                    <Button x:Name="InstallButton" Content="Install Required Modules" Margin="0,5" Padding="10,5" Background="#0078D4" Foreground="White" FontWeight="Bold" ToolTip="Install missing PowerShell modules." Visibility="Collapsed"/>
                    <Button x:Name="ConnectButton" Content="Connect to Power BI" Margin="0,5" Padding="10,5" Background="#0078D4" Foreground="White" FontWeight="Bold" ToolTip="Connect to your Power BI account." IsEnabled="False"/>
                    <TextBlock Text="Select details to retrieve:" FontWeight="Bold" Margin="0,15,0,5"/>
                    <CheckBox x:Name="WorkspacesCheck" Content="Workspaces" IsChecked="True"/>
                    <CheckBox x:Name="CapacitiesCheck" Content="Capacities"/>
                    <CheckBox x:Name="DatasetsCheck" Content="Datasets"/>
                    <CheckBox x:Name="ReportsCheck" Content="Reports"/>
                    <CheckBox x:Name="UsersCheck" Content="Users in Workspaces"/>
                    <CheckBox x:Name="ExternalUsersCheck" Content="Highlight External Users"/>
                    
                    <Button x:Name="GetDetailsButton" Content="Get Detailed Report" Margin="0,20,0,0" Padding="10,5" IsEnabled="False" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                    <Button x:Name="ExecutiveReportButton" Content="Get Executive Summary" Margin="0,10,0,0" Padding="10,5" IsEnabled="False" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                    <Button x:Name="ExportButton" Content="Export Detailed Report" Margin="0,10,0,0" Padding="10,5" IsEnabled="False" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="1" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="5" Padding="10">
                <DockPanel>
                    <TextBlock x:Name="StatusLabel" DockPanel.Dock="Top" Text="Status: Initializing..." FontStyle="Italic" Foreground="Gray" Margin="0,0,0,5"/>
                    <TextBox x:Name="OutputTextBox" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas" FontSize="12" BorderThickness="0"/>
                </DockPanel>
            </Border>
        </Grid>

        <DockPanel Grid.Row="2" Margin="0,10,0,0">
            <Button x:Name="CloseButton" Content="Close" HorizontalAlignment="Right" Padding="10,5" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
            <ProgressBar x:Name="ProgressBar" Height="8" Margin="5,5,5,5" Visibility="Hidden"/>
        </DockPanel>
    </Grid>
</Window>
"@
#endregion

#region Script Logic
try {
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Error loading XAML: $_"
    exit
}

# Get references to the GUI elements
$installButton = $window.FindName("InstallButton")
$connectButton = $window.FindName("ConnectButton")
$getDetailsButton = $window.FindName("GetDetailsButton")
$executiveReportButton = $window.FindName("ExecutiveReportButton")
$exportButton = $window.FindName("ExportButton")
$closeButton = $window.FindName("CloseButton")
$outputTextBox = $window.FindName("OutputTextBox")
$statusLabel = $window.FindName("StatusLabel")
$progressBar = $window.FindName("ProgressBar")

$workspacesCheck = $window.FindName("WorkspacesCheck")
$capacitiesCheck = $window.FindName("CapacitiesCheck")
$datasetsCheck = $window.FindName("DatasetsCheck")
$reportsCheck = $window.FindName("ReportsCheck")
$usersCheck = $window.FindName("UsersCheck")
$externalUsersCheck = $window.FindName("ExternalUsersCheck")

# Store results for export
$script:Results = @()
$script:ModulesToInstall = @()
$script:isJobCleaningUp = $false

# --- Functions ---

function Update-Status ($message, $color = "Gray") {
    $statusLabel.Dispatcher.Invoke({
        $statusLabel.Text = "Status: $message"
        $statusLabel.Foreground = $color
    })
}

function Show-Progress($isVisible) {
    $progressBar.Dispatcher.Invoke({
        $progressBar.Visibility = if ($isVisible) { "Visible" } else { "Hidden" }
        $progressBar.IsIndeterminate = $isVisible
    })
}

function Get-MissingModules {
    $requiredModules = @("MicrosoftPowerBIMgmt", "MSAL.PS")
    $installedModules = Get-Module -ListAvailable
    $modulesToInstall = $requiredModules | Where-Object { $installedModules.Name -notcontains $_ }
    return $modulesToInstall
}

function Set-ActionButtonsState($isEnabled) {
    $getDetailsButton.IsEnabled = $isEnabled
    $executiveReportButton.IsEnabled = $isEnabled
    $exportButton.IsEnabled = $false # Always disable export initially
}

function Check-Prerequisites {
    Show-Progress $true
    Update-Status "Checking for required modules..." "Blue"
    $requiredModulesList = @("MicrosoftPowerBIMgmt", "MSAL.PS")
    $outputTextBox.Text = "Welcome! Initializing...`n`nChecking for the following required modules:`n - " + ($requiredModulesList -join "`n - ")
    
    $window.Dispatcher.Invoke([System.Action]{}, "Background")

    $script:ModulesToInstall = Get-MissingModules
    
    if ($script:ModulesToInstall.Count -gt 0) {
        $moduleList = $script:ModulesToInstall -join ", "
        Update-Status "Required modules missing. Please install." "Orange"
        $outputTextBox.Text += "`n`nAction Required: The following modules are missing: $moduleList`nPlease click the 'Install' button to proceed."
        $installButton.Visibility = "Visible"
        $connectButton.Visibility = "Collapsed"
    } else {
        Update-Status "Ready to connect." "Green"
        $outputTextBox.Text += "`n`nSuccess! All required modules are installed. Please click 'Connect to Power BI'."
        $connectButton.IsEnabled = $true
    }
    Show-Progress $false
}

# --- Event Handlers ---

$window.Add_Loaded({
    Check-Prerequisites
})

$installButton.Add_Click({
    $moduleList = $script:ModulesToInstall -join ", "
    $message = "The following required modules are not installed: $moduleList.`n`nThis script will now attempt to install them from the PSGallery. This may take several minutes and might require administrator privileges.`n`nDo you want to continue with the installation?"
    $result = [System.Windows.MessageBox]::Show($message, "Required Modules Missing", "YesNo", "Warning")

    if ($result -eq "Yes") {
        Show-Progress $true
        Update-Status "Installing modules: $moduleList. Please wait..." "Orange"
        $installButton.IsEnabled = $false
        $script:isJobCleaningUp = $false # Reset the cleanup flag
        
        $installJob = Start-Job -ScriptBlock {
            param($modules)
            Install-Module -Name $modules -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        } -ArgumentList (,$script:ModulesToInstall)

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(1)
        $timer.Tag = $installJob

        $tickHandler = $null
        $tickHandler = {
            param($sender, $e)
            $timerObject = $sender
            $job = $timerObject.Tag
            
            if ($null -eq $job -or $job.State -ne 'Running') {
                if (-not $script:isJobCleaningUp) {
                    $script:isJobCleaningUp = $true # Set the flag to prevent re-entry
                    $timerObject.Stop()
                    $timerObject.remove_Tick($tickHandler)

                    try {
                        if ($null -ne $job) { Receive-Job -Job $job -ErrorAction Stop }
                        Update-Status "Modules installed. Ready to connect." "Green"
                        $outputTextBox.Text = "Module installation successful. Please click 'Connect to Power BI'."
                        $installButton.Visibility = "Collapsed"
                        $connectButton.Visibility = "Visible"
                        $connectButton.IsEnabled = $true
                    } catch {
                        Update-Status "Module installation failed. See output for details." "Red"
                        $outputTextBox.Text += "`n`nERROR: Module installation failed.`n$($_.Exception.Message)"
                        [System.Windows.MessageBox]::Show("Failed to install modules. Please try running PowerShell as an Administrator and try again.", "Installation Error", "OK", "Error")
                        $installButton.IsEnabled = $true
                    } finally {
                        if ($null -ne $job) { Remove-Job -Job $job -Force }
                        Show-Progress $false
                    }
                }
            }
        }
        $timer.Add_Tick($tickHandler)
        $timer.Start()
    }
})

$connectButton.Add_Click({
    Show-Progress $true
    Set-ActionButtonsState $false
    $connectButton.IsEnabled = $false
    Update-Status "Waiting for user login..." "Blue"
    $outputTextBox.Text = "A login window should appear. Please sign in to continue."
    $script:isJobCleaningUp = $false # Reset the cleanup flag

    $loginJob = Start-Job -ScriptBlock {
        Import-Module MicrosoftPowerBIMgmt -Force
        # This will pop a standard interactive login window in the background job's process space
        Connect-PowerBIServiceAccount
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(2)
    $timer.Tag = $loginJob

    $tickHandler = $null
    $tickHandler = {
        param($sender, $e)
        $timerObject = $sender
        $job = $timerObject.Tag

        if ($null -eq $job) {
            if (-not $script:isJobCleaningUp) {
                $script:isJobCleaningUp = $true
                $timerObject.Stop()
                $timerObject.remove_Tick($tickHandler)
            }
            return
        }
        
        if ($job.State -ne 'Running') {
            if (-not $script:isJobCleaningUp) {
                $script:isJobCleaningUp = $true # Set the flag to prevent re-entry
                $timerObject.Stop()
                $timerObject.remove_Tick($tickHandler)

                try {
                    # This will throw an error if the login in the job failed
                    Receive-Job -Job $job -ErrorAction Stop
                    
                    # *** FIX: Import the module in the main session to make cmdlets available ***
                    Update-Status "Finalizing connection..." "Blue"
                    Import-Module MicrosoftPowerBIMgmt -Force
                    
                    # Now, get the context in the main thread
                    $context = Get-PowerBIServiceAccountContext
                    Update-Status "Connected as $($context.User.Identifier)" "Green"
                    $outputTextBox.Text = "Successfully connected as $($context.User.Identifier).`nReady to get details."
                    Set-ActionButtonsState $true
                } catch {
                    Update-Status "Connection failed or was cancelled. Please try again." "Red"
                    $outputTextBox.Text = "The login process failed or was cancelled.`nError: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Failed to connect to Power BI.", "Connection Error", "OK", "Error")
                    $connectButton.IsEnabled = $true
                } finally {
                    Remove-Job -Job $job -Force
                    Show-Progress $false
                }
            }
        }
    }
    $timer.Add_Tick($tickHandler)
    $timer.Start()
})


$getDetailsButton.Add_Click({
    $outputTextBox.Clear()
    $script:Results = @()
    Show-Progress $true
    Set-ActionButtonsState $false
    Update-Status "Getting detailed report..." "Blue"

    $job = Start-Job -ScriptBlock {
        param($options)
        $output = ""
        $results = @()
        Import-Module MicrosoftPowerBIMgmt -Force

        if ($options.Workspaces) {
            $output += "`n" + ("-"*20) + " WORKSPACES " + ("-"*20) + "`n"
            $workspaces = Get-PowerBIWorkspace -Scope Organization -Include All
            foreach ($ws in $workspaces) {
                $output += "Name: $($ws.Name)`n  ID: $($ws.Id)`n  State: $($ws.State)`n  IsOnDedicatedCapacity: $($ws.IsOnDedicatedCapacity)`n`n"
                $results += [pscustomobject]@{ Type = "Workspace"; Name = $ws.Name; ID = $ws.Id; State = $ws.State; IsOnDedicatedCapacity = $ws.IsOnDedicatedCapacity }
                if ($options.Users) {
                    $output += "  -- Users --`n"
                    $users = Get-PowerBIWorkspaceUser -Workspace $ws
                    foreach ($user in $users) {
                        $isExternal = $user.UserPrincipalName -like "*#ext#*"
                        $userLine = "    - $($user.UserPrincipalName) (Access: $($user.AccessRight))"
                        if ($isExternal -and $options.ExternalUsers) { $userLine += " <-- EXTERNAL USER" }
                        $output += "$userLine`n"
                        $results += [pscustomobject]@{ Type = "User"; Workspace = $ws.Name; UserPrincipalName = $user.UserPrincipalName; AccessRight = $user.AccessRight; IsExternal = $isExternal }
                    }
                    $output += "`n"
                }
            }
        }
        if ($options.Capacities) {
            $output += "`n" + ("-"*20) + " CAPACITIES " + ("-"*20) + "`n"
            $capacities = Get-PowerBICapacity
            foreach ($cap in $capacities) {
                $output += "DisplayName: $($cap.DisplayName)`n  ID: $($cap.Id)`n  Sku: $($cap.Sku)`n  State: $($cap.State)`n`n"
                 $results += [pscustomobject]@{ Type = "Capacity"; DisplayName = $cap.DisplayName; ID = $cap.Id; Sku = $cap.Sku; State = $cap.State }
            }
        }
        if ($options.Datasets) {
            $output += "`n" + ("-"*20) + " DATASETS " + ("-"*20) + "`n"
            $datasets = Get-PowerBIDataset -Scope Organization
            foreach ($ds in $datasets) {
                $output += "Name: $($ds.Name)`n  ID: $($ds.Id)`n  Workspace ID: $($ds.WorkspaceId)`n`n"
                $results += [pscustomobject]@{ Type = "Dataset"; Name = $ds.Name; ID = $ds.Id; WorkspaceId = $ds.WorkspaceId }
            }
        }
        if ($options.Reports) {
            $output += "`n" + ("-"*20) + " REPORTS " + ("-"*20) + "`n"
            $reports = Get-PowerBIReport -Scope Organization
            foreach ($rpt in $reports) {
                $output += "Name: $($rpt.Name)`n  ID: $($rpt.Id)`n  Dataset ID: $($rpt.DatasetId)`n  Workspace ID: $($rpt.WorkspaceId)`n`n"
                $results += [pscustomobject]@{ Type = "Report"; Name = $rpt.Name; ID = $rpt.Id; DatasetId = $rpt.DatasetId; WorkspaceId = $rpt.WorkspaceId }
            }
        }
        return @{Output = $output; Results = $results}
    } -ArgumentList @{ Workspaces = $workspacesCheck.IsChecked; Capacities = $capacitiesCheck.IsChecked; Datasets = $datasetsCheck.IsChecked; Reports = $reportsCheck.IsChecked; Users = $usersCheck.IsChecked; ExternalUsers = $externalUsersCheck.IsChecked }

    while ($job.State -eq 'Running') { Start-Sleep -Seconds 1 }
    $jobResult = Receive-Job $job
    $outputTextBox.Text = $jobResult.Output
    $script:Results = $jobResult.Results
    Remove-Job $job -Force
    Show-Progress $false
    Update-Status "Detailed report retrieved successfully." "Green"
    Set-ActionButtonsState $true
    $exportButton.IsEnabled = ($script:Results.Count -gt 0)
})

$executiveReportButton.Add_Click({
    $outputTextBox.Clear()
    $script:Results = @()
    Show-Progress $true
    Set-ActionButtonsState $false
    Update-Status "Generating Executive Summary..." "Blue"

    $job = Start-Job -ScriptBlock {
        Import-Module MicrosoftPowerBIMgmt -Force
        $workspaces = Get-PowerBIWorkspace -Scope Organization -All -Include Users
        $reports = Get-PowerBIReport -Scope Organization -All
        $datasets = Get-PowerBIDataset -Scope Organization -All

        $totalWorkspaces = $workspaces.Count
        $premiumWorkspaces = ($workspaces | Where-Object { $_.IsOnDedicatedCapacity -eq $true }).Count
        $allUsers = $workspaces.Users.UserPrincipalName | Select-Object -Unique
        $totalUsers = $allUsers.Count
        $externalUsers = ($allUsers | Where-Object { $_ -like "*#ext#*" }).Count
        
        $output = @"
Power BI Tenant - Executive Summary
Generated on: $(Get-Date)
=================================================

Workspace Metrics:
------------------
- Total Workspaces: $totalWorkspaces
- Workspaces on Premium/Fabric Capacity: $premiumWorkspaces

User Metrics:
-------------
- Total Unique Users with Access: $totalUsers
- Total External (Guest) Users: $externalUsers

Content Metrics:
----------------
- Total Reports: $($reports.Count)
- Total Datasets: $($datasets.Count)

"@
        return $output
    }

    while ($job.State -eq 'Running') { Start-Sleep -Seconds 1 }
    $outputTextBox.Text = Receive-Job $job
    Remove-Job $job -Force
    Show-Progress $false
    Update-Status "Executive Summary generated successfully." "Green"
    Set-ActionButtonsState $true
})


$exportButton.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.MessageBox]::Show("There is no data to export. Please run a detailed report first.", "No Data", "OK", "Information")
        return
    }

    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Power BI Details"
    $saveFileDialog.FileName = "PowerBI_Details_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"

    if ($saveFileDialog.ShowDialog() -eq $true) {
        try {
            $script:Results | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)", "Export Successful", "OK", "Information")
            Update-Status "Data exported to $($saveFileDialog.FileName)" "Green"
        } catch {
            [System.Windows.MessageBox]::Show("Failed to export data. Error: $($_.Exception.Message)", "Export Error", "OK", "Error")
            Update-Status "Export failed." "Red"
        }
    }
})

$closeButton.Add_Click({
    try {
        if (Get-PowerBIServiceAccountContext -ErrorAction SilentlyContinue) {
            Disconnect-PowerBIServiceAccount
        }
    } catch {}
    $window.Close()
})


# Show the window
$window.ShowDialog() | Out-Null
#endregion