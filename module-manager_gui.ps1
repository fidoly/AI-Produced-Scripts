# Enhanced PowerShell Module Manager v3.5.3
# A modern, robust GUI for managing PowerShell modules and Microsoft 365 services
# 
# ⚠️  EXECUTION POLICY NOTICE ⚠️
# If you get an execution policy error, run one of these commands first:
# 
# Option 1 (Recommended): Use the included batch launcher
#   .\launch_module_manager.bat
# 
# Option 2: Set execution policy manually then run script
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   .\module_manager_v353.ps1
# 
# Option 3: Run directly with bypass
#   powershell.exe -ExecutionPolicy Bypass -File ".\module_manager_v353.ps1"
# 
# Version: 3.5.3
# Author: Enhanced by Claude (Anthropic)
# Compatible with: Windows 10/11, PowerShell 5.1+, PowerShell 7+
# Last Updated: January 2025
# 
# Changelog v3.5.3:
# - FIXED: Module auto-import issue - modules no longer import without being selected
# - FIXED: Refresh functionality now works properly and updates all statuses correctly
# - FIXED: Dependency tracking to prevent unwanted module imports
# - Simplified SharePoint conflict resolution for better reliability
# - Improved GUI synchronization to prevent timing issues
# - Added module import prevention mechanism
# - Enhanced error handling and logging

#Requires -Version 5.1

param(
    [switch]$SkipAdminCheck,
    [switch]$Debug
)

# Set strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Import required assemblies with error handling
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
} catch {
    Write-Host "Failed to load required assemblies. Please ensure .NET Framework is installed." -ForegroundColor Red
    exit 1
}

# Configuration
$script:Config = @{
    RequiredModules = @(
        @{ Name = "Az"; DisplayName = "Az"; Description = "Azure PowerShell Module"; Category = "Azure"; MinVersion = "11.0.0" }
        @{ Name = "Microsoft.Graph"; DisplayName = "Microsoft.Graph"; Description = "Microsoft Graph PowerShell SDK"; Category = "Graph"; MinVersion = "2.0.0" }
        @{ Name = "ExchangeOnlineManagement"; DisplayName = "ExchangeOnlineManagement"; Description = "Exchange Online Management"; Category = "Exchange"; MinVersion = "3.0.0" }
        @{ Name = "Microsoft.Online.SharePoint.PowerShell"; DisplayName = "MSOL-SharePoint-PowerShell"; Description = "SharePoint Online Management Shell (Legacy - conflicts with PnP.PowerShell)"; Category = "SharePoint"; MinVersion = $null }
        @{ Name = "MicrosoftPowerBIMgmt"; DisplayName = "MicrosoftPowerBIMgmt"; Description = "Power BI Management"; Category = "Power BI"; MinVersion = $null }
        @{ Name = "MicrosoftTeams"; DisplayName = "MicrosoftTeams"; Description = "Microsoft Teams PowerShell"; Category = "Teams"; MinVersion = "5.0.0" }
        @{ Name = "PnP.PowerShell"; DisplayName = "PnP.PowerShell"; Description = "PnP PowerShell (Modern - conflicts with MSOL-SharePoint-PowerShell)"; Category = "SharePoint"; MinVersion = "2.0.0" }
    )
    LogFile = Join-Path $env:TEMP "ModuleManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    AppName = "PowerShell Module Manager Pro"
    Version = "3.5.3"
    DebugMode = $Debug.IsPresent
}

# Global variables for GUI components and service connections
$script:GUI = @{
    Form = $null
    ModuleCheckboxes = @{}
    StatusLabels = @{}
    ProgressBar = $null
    StatusLabel = $null
    LogTextBox = $null
}

# Track service connection states
$script:ServiceConnections = @{
    "Microsoft.Graph" = $false
    "ExchangeOnlineManagement" = $false
    "MicrosoftTeams" = $false
    "PnP.PowerShell" = $false
}

# Track modules we've explicitly imported to prevent auto-imports
$script:ExplicitlyImportedModules = @()

# Thread-safe logging function
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Level = "Info"
    )
    
    # Skip debug messages unless in debug mode
    if ($Level -eq "Debug" -and -not $script:Config.DebugMode) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Thread-safe file writing
    try {
        $logEntry | Out-File -FilePath $script:Config.LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        # Silently continue if log file write fails
    }
    
    # Console output with color
    $color = switch ($Level) {
        "Info" { "White" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
        "Debug" { "Cyan" }
    }
    Write-Host $logEntry -ForegroundColor $color
    
    # Update GUI log if available
    if ($script:GUI.LogTextBox -and -not $script:GUI.LogTextBox.IsDisposed) {
        try {
            if ($script:GUI.LogTextBox.InvokeRequired) {
                $script:GUI.LogTextBox.Invoke([Action]{
                    $script:GUI.LogTextBox.AppendText("$logEntry`r`n")
                    $script:GUI.LogTextBox.SelectionStart = $script:GUI.LogTextBox.Text.Length
                    $script:GUI.LogTextBox.ScrollToCaret()
                })
            } else {
                $script:GUI.LogTextBox.AppendText("$logEntry`r`n")
                $script:GUI.LogTextBox.SelectionStart = $script:GUI.LogTextBox.Text.Length
                $script:GUI.LogTextBox.ScrollToCaret()
            }
        }
        catch {
            # Silently continue if GUI update fails
        }
    }
}

# Check service connection status
function Test-ServiceConnection {
    param([string]$ServiceName)
    
    try {
        switch ($ServiceName) {
            "Microsoft.Graph" {
                $context = Get-MgContext -ErrorAction SilentlyContinue
                return $null -ne $context
            }
            "ExchangeOnlineManagement" {
                $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                return $null -ne $session
            }
            "MicrosoftTeams" {
                try {
                    $null = Get-CsTenant -ErrorAction SilentlyContinue
                    return $true
                } catch {
                    return $false
                }
            }
            "PnP.PowerShell" {
                try {
                    $null = Get-PnPConnection -ErrorAction SilentlyContinue
                    return $true
                } catch {
                    return $false
                }
            }
            default {
                return $false
            }
        }
    }
    catch {
        return $false
    }
}

# Simplified SharePoint conflict check
function Test-SharePointConflict {
    $msolImported = Get-Module -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue
    $pnpImported = Get-Module -Name "PnP.PowerShell" -ErrorAction SilentlyContinue
    
    return ($null -ne $msolImported -and $null -ne $pnpImported)
}

# Check if module is installed
function Test-ModuleInstalled {
    param([string]$ModuleName)
    
    try {
        $module = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue
        return $null -ne $module
    }
    catch {
        Write-Log "Error checking if module $ModuleName is installed: $_" -Level "Debug"
        return $false
    }
}

# Check if module is imported
function Test-ModuleImported {
    param([string]$ModuleName)
    
    try {
        $module = Get-Module -Name $ModuleName -ErrorAction SilentlyContinue
        return $null -ne $module
    }
    catch {
        Write-Log "Error checking if module $ModuleName is imported: $_" -Level "Debug"
        return $false
    }
}

# Get module version
function Get-ModuleVersion {
    param([string]$ModuleName)
    
    try {
        $module = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue | 
                  Sort-Object Version -Descending | 
                  Select-Object -First 1
        
        if ($module) {
            return $module.Version.ToString()
        }
        return "N/A"
    }
    catch {
        return "Error"
    }
}

# Install module
function Install-RequiredModule {
    param(
        [Parameter(Mandatory)]
        [hashtable]$ModuleInfo,
        
        [System.Windows.Forms.ProgressBar]$ProgressBar = $null
    )
    
    try {
        Write-Log "Installing module: $($ModuleInfo.Name)" -Level "Info"
        
        if (Test-ModuleInstalled -ModuleName $ModuleInfo.Name) {
            Write-Log "Module $($ModuleInfo.Name) is already installed" -Level "Info"
            Update-ModuleStatus
            return $true
        }
        
        # Update progress bar
        if ($ProgressBar -and -not $ProgressBar.IsDisposed) {
            $ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        }
        
        # Check for NuGet provider
        $nugetProvider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
        if (-not $nugetProvider) {
            Write-Log "Installing NuGet provider..." -Level "Info"
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
        }
        
        # Set PSGallery as trusted if not already
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($psGallery.InstallationPolicy -ne 'Trusted') {
            Write-Log "Setting PSGallery as trusted repository..." -Level "Info"
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
        
        # Install module with proper parameters
        $installParams = @{
            Name = $ModuleInfo.Name
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            Repository = "PSGallery"
            ErrorAction = "Stop"
        }
        
        # Add minimum version if specified
        if ($ModuleInfo.MinVersion) {
            $installParams['MinimumVersion'] = $ModuleInfo.MinVersion
        }
        
        Install-Module @installParams
        
        Write-Log "Successfully installed module: $($ModuleInfo.Name)" -Level "Success"
        
        # Update status after installation
        Update-ModuleStatus
        
        return $true
    }
    catch {
        Write-Log "Failed to install module $($ModuleInfo.Name): $($_.Exception.Message)" -Level "Error"
        return $false
    }
    finally {
        if ($ProgressBar -and -not $ProgressBar.IsDisposed) {
            $ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        }
    }
}

# Import module with dependency prevention
function Import-RequiredModule {
    param(
        [Parameter(Mandatory)]
        [hashtable]$ModuleInfo
    )
    
    try {
        Write-Log "Starting import process for: $($ModuleInfo.Name)" -Level "Info"
        
        if (Test-ModuleImported -ModuleName $ModuleInfo.Name) {
            Write-Log "Module $($ModuleInfo.Name) is already imported" -Level "Info"
            return $true
        }
        
        if (-not (Test-ModuleInstalled -ModuleName $ModuleInfo.Name)) {
            Write-Log "Module $($ModuleInfo.Name) is not installed. Please install it first." -Level "Warning"
            return $false
        }
        
        # SharePoint conflict check
        if ($ModuleInfo.Name -in @("Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell")) {
            $conflictingModule = if ($ModuleInfo.Name -eq "PnP.PowerShell") { 
                "Microsoft.Online.SharePoint.PowerShell" 
            } else { 
                "PnP.PowerShell" 
            }
            
            if (Test-ModuleImported -ModuleName $conflictingModule) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "The module '$conflictingModule' is already imported and conflicts with '$($ModuleInfo.Name)'.`n`nDo you want to remove '$conflictingModule' and import '$($ModuleInfo.Name)' instead?",
                    "SharePoint Module Conflict",
                    "YesNo",
                    "Warning"
                )
                
                if ($result -eq "Yes") {
                    try {
                        Remove-Module -Name $conflictingModule -Force -ErrorAction Stop
                        Write-Log "Removed conflicting module: $conflictingModule" -Level "Success"
                    }
                    catch {
                        Write-Log "Failed to remove conflicting module: $($_.Exception.Message)" -Level "Error"
                        return $false
                    }
                } else {
                    Write-Log "User cancelled import due to conflict" -Level "Info"
                    return $false
                }
            }
        }
        
        # Track what we're importing explicitly
        $script:ExplicitlyImportedModules += $ModuleInfo.Name
        
        # Import the module
        Write-Log "Importing module: $($ModuleInfo.Name)" -Level "Info"
        
        # Show progress for large modules
        if ($ModuleInfo.Name -in @("Az", "Microsoft.Graph")) {
            Write-Log "WARNING: $($ModuleInfo.Name) is a large module and may take 3-5 minutes to import. Please be patient..." -Level "Warning"
            
            if ($script:GUI.StatusLabel -and -not $script:GUI.StatusLabel.IsDisposed) {
                $script:GUI.StatusLabel.Text = "Importing $($ModuleInfo.Name) - This may take 3-5 minutes..."
                $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
            }
            
            if ($script:GUI.ProgressBar -and -not $script:GUI.ProgressBar.IsDisposed) {
                $script:GUI.ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            }
        }
        
        # Prevent auto-loading of dependencies by using -DisableNameChecking
        Import-Module -Name $ModuleInfo.Name -Force -DisableNameChecking -ErrorAction Stop
        
        Write-Log "Successfully imported module: $($ModuleInfo.Name)" -Level "Success"
        
        # Update status
        Update-ModuleStatus
        
        # Reset progress bar
        if ($script:GUI.ProgressBar -and -not $script:GUI.ProgressBar.IsDisposed) {
            $script:GUI.ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        }
        
        if ($script:GUI.StatusLabel -and -not $script:GUI.StatusLabel.IsDisposed) {
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Black
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to import module $($ModuleInfo.Name): $($_.Exception.Message)" -Level "Error"
        
        # Reset progress bar on error
        if ($script:GUI.ProgressBar -and -not $script:GUI.ProgressBar.IsDisposed) {
            $script:GUI.ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        }
        
        if ($script:GUI.StatusLabel -and -not $script:GUI.StatusLabel.IsDisposed) {
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Black
        }
        
        return $false
    }
}

# Remove imported module
function Remove-ImportedModule {
    param(
        [Parameter(Mandatory)]
        [hashtable]$ModuleInfo
    )
    
    try {
        if (-not (Test-ModuleImported -ModuleName $ModuleInfo.Name)) {
            Write-Log "Module $($ModuleInfo.Name) is not currently imported" -Level "Info"
            return $true
        }
        
        Write-Log "Removing module: $($ModuleInfo.Name)" -Level "Info"
        
        # Disconnect from services first
        switch ($ModuleInfo.Name) {
            "ExchangeOnlineManagement" {
                try { 
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 
                    $script:ServiceConnections["ExchangeOnlineManagement"] = $false
                } catch {}
            }
            "Microsoft.Graph" {
                try { 
                    Disconnect-MgGraph -ErrorAction SilentlyContinue 
                    $script:ServiceConnections["Microsoft.Graph"] = $false
                } catch {}
            }
            "MicrosoftTeams" {
                try { 
                    Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue 
                    $script:ServiceConnections["MicrosoftTeams"] = $false
                } catch {}
            }
            "PnP.PowerShell" {
                try { 
                    Disconnect-PnPOnline -ErrorAction SilentlyContinue 
                    $script:ServiceConnections["PnP.PowerShell"] = $false
                } catch {}
            }
        }
        
        # Remove from explicitly imported list
        $script:ExplicitlyImportedModules = $script:ExplicitlyImportedModules | Where-Object { $_ -ne $ModuleInfo.Name }
        
        Remove-Module -Name $ModuleInfo.Name -Force -ErrorAction Stop
        Write-Log "Successfully removed module: $($ModuleInfo.Name)" -Level "Success"
        
        # Update status
        Update-ModuleStatus
        
        return $true
    }
    catch {
        Write-Log "Failed to remove module $($ModuleInfo.Name): $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

# Disconnect from selected services
function Disconnect-SelectedServices {
    param([array]$SelectedModules)
    
    $disconnectedServices = @()
    
    foreach ($module in $SelectedModules) {
        if (-not (Test-ModuleImported -ModuleName $module.Name)) {
            Write-Log "Module $($module.Name) is not imported - skipping disconnect" -Level "Info"
            continue
        }
        
        Write-Log "Disconnecting from service: $($module.DisplayName)" -Level "Info"
        
        $disconnected = $false
        switch ($module.Name) {
            "ExchangeOnlineManagement" {
                try {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                    $script:ServiceConnections["ExchangeOnlineManagement"] = $false
                    Write-Log "Successfully disconnected from Exchange Online" -Level "Success"
                    $disconnectedServices += "Exchange Online"
                    $disconnected = $true
                } catch {
                    Write-Log "Failed to disconnect from Exchange Online: $($_.Exception.Message)" -Level "Error"
                }
            }
            "Microsoft.Graph" {
                try {
                    Disconnect-MgGraph -ErrorAction Stop
                    $script:ServiceConnections["Microsoft.Graph"] = $false
                    Write-Log "Successfully disconnected from Microsoft Graph" -Level "Success"
                    $disconnectedServices += "Microsoft Graph"
                    $disconnected = $true
                } catch {
                    Write-Log "Failed to disconnect from Microsoft Graph: $($_.Exception.Message)" -Level "Error"
                }
            }
            "MicrosoftTeams" {
                try {
                    Disconnect-MicrosoftTeams -ErrorAction Stop
                    $script:ServiceConnections["MicrosoftTeams"] = $false
                    Write-Log "Successfully disconnected from Microsoft Teams" -Level "Success"
                    $disconnectedServices += "Microsoft Teams"
                    $disconnected = $true
                } catch {
                    Write-Log "Failed to disconnect from Microsoft Teams: $($_.Exception.Message)" -Level "Error"
                }
            }
            "PnP.PowerShell" {
                try {
                    Disconnect-PnPOnline -ErrorAction Stop
                    $script:ServiceConnections["PnP.PowerShell"] = $false
                    Write-Log "Successfully disconnected from SharePoint PnP" -Level "Success"
                    $disconnectedServices += "SharePoint PnP"
                    $disconnected = $true
                } catch {
                    Write-Log "Failed to disconnect from SharePoint PnP: $($_.Exception.Message)" -Level "Error"
                }
            }
            "Microsoft.Online.SharePoint.PowerShell" {
                try {
                    Disconnect-SPOService -ErrorAction Stop
                    Write-Log "Successfully disconnected from SharePoint Online" -Level "Success"
                    $disconnectedServices += "SharePoint Online"
                    $disconnected = $true
                } catch {
                    Write-Log "Failed to disconnect from SharePoint Online: $($_.Exception.Message)" -Level "Error"
                }
            }
            default {
                Write-Log "No disconnect function available for $($module.DisplayName)" -Level "Info"
            }
        }
        
        # Force status update after each disconnect
        if ($disconnected) {
            Update-ModuleStatus
            Start-Sleep -Milliseconds 100
        }
    }
    
    return $disconnectedServices
}

# Unload selected modules
function Unload-SelectedModules {
    param([array]$SelectedModules)
    
    $unloadedModules = @()
    $failedModules = @()
    
    foreach ($module in $SelectedModules) {
        if (-not (Test-ModuleImported -ModuleName $module.Name)) {
            Write-Log "Module $($module.Name) is not imported - skipping unload" -Level "Info"
            continue
        }
        
        Write-Log "Unloading module: $($module.DisplayName)" -Level "Info"
        
        # First disconnect from service if connected
        $null = Disconnect-SelectedServices -SelectedModules @($module)
        
        # Remove the module
        try {
            Remove-Module -Name $module.Name -Force -ErrorAction Stop
            
            # Remove from explicitly imported list
            $script:ExplicitlyImportedModules = $script:ExplicitlyImportedModules | Where-Object { $_ -ne $module.Name }
            
            Write-Log "Successfully unloaded module: $($module.DisplayName)" -Level "Success"
            $unloadedModules += $module.DisplayName
            
            # Force immediate status update
            Update-ModuleStatus
            Start-Sleep -Milliseconds 100
        }
        catch {
            Write-Log "Failed to unload module $($module.DisplayName): $($_.Exception.Message)" -Level "Error"
            $failedModules += $module.DisplayName
        }
    }
    
    return @{
        Unloaded = $unloadedModules
        Failed = $failedModules
    }
}

# Update module status in GUI
function Update-ModuleStatus {
    try {
        Write-Log "Refreshing module status..." -Level "Debug"
        
        # Force GUI to process pending events first
        if ($script:GUI.Form -and -not $script:GUI.Form.IsDisposed) {
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        foreach ($module in $script:Config.RequiredModules) {
            $moduleName = $module.Name
            
            if ($script:GUI.StatusLabels.ContainsKey($moduleName) -and 
                -not $script:GUI.StatusLabels[$moduleName].IsDisposed) {
                
                $installed = Test-ModuleInstalled -ModuleName $moduleName
                $imported = Test-ModuleImported -ModuleName $moduleName
                $version = if ($installed) { Get-ModuleVersion -ModuleName $moduleName } else { "N/A" }
                
                # Check connection status for applicable modules
                $connected = $false
                if ($script:ServiceConnections.ContainsKey($moduleName) -and $imported) {
                    $connected = Test-ServiceConnection -ServiceName $moduleName
                    $script:ServiceConnections[$moduleName] = $connected
                }
                
                $statusText = "Not Installed"
                $statusColor = [System.Drawing.Color]::Red
                
                if ($installed -and $imported -and $connected) {
                    $statusText = "Connected (v$version)"
                    $statusColor = [System.Drawing.Color]::Blue
                } elseif ($installed -and $imported) {
                    $statusText = "Imported (v$version)"
                    $statusColor = [System.Drawing.Color]::Green
                } elseif ($installed) {
                    $statusText = "Installed (v$version)"
                    $statusColor = [System.Drawing.Color]::Orange
                }
                
                # Update the status label
                try {
                    if ($script:GUI.StatusLabels[$moduleName].InvokeRequired) {
                        $script:GUI.StatusLabels[$moduleName].Invoke([Action]{
                            $script:GUI.StatusLabels[$moduleName].Text = $statusText
                            $script:GUI.StatusLabels[$moduleName].ForeColor = $statusColor
                            $script:GUI.StatusLabels[$moduleName].Refresh()
                        })
                    } else {
                        $script:GUI.StatusLabels[$moduleName].Text = $statusText
                        $script:GUI.StatusLabels[$moduleName].ForeColor = $statusColor
                        $script:GUI.StatusLabels[$moduleName].Refresh()
                    }
                }
                catch {
                    Write-Log "Error updating status label for $moduleName : $_" -Level "Debug"
                }
            }
        }
        
        # Force form refresh
        if ($script:GUI.Form -and -not $script:GUI.Form.IsDisposed) {
            $script:GUI.Form.Refresh()
        }
        
        Write-Log "Module status refresh completed" -Level "Debug"
    }
    catch {
        Write-Log "Error updating module status: $($_.Exception.Message)" -Level "Error"
    }
}

# Get selected modules from checkboxes
function Get-SelectedModules {
    $selectedModules = @()
    
    foreach ($module in $script:Config.RequiredModules) {
        $moduleName = $module.Name
        if ($script:GUI.ModuleCheckboxes.ContainsKey($moduleName) -and 
            -not $script:GUI.ModuleCheckboxes[$moduleName].IsDisposed -and
            $script:GUI.ModuleCheckboxes[$moduleName].Checked) {
            $selectedModules += $module
        }
    }
    
    return @($selectedModules)
}

# Connect to services
function Connect-ToServices {
    param([System.Windows.Forms.Form]$ParentForm)
    
    $connections = @()
    
    # Microsoft Graph
    if (Test-ModuleImported -ModuleName "Microsoft.Graph") {
        try {
            Write-Log "Connecting to Microsoft Graph..." -Level "Info"
            $scopes = @(
                "User.Read.All",
                "Group.Read.All",
                "Directory.Read.All",
                "Mail.Read",
                "Calendars.Read"
            )
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
            $script:ServiceConnections["Microsoft.Graph"] = $true
            Write-Log "Successfully connected to Microsoft Graph" -Level "Success"
            $connections += "Microsoft Graph"
        }
        catch {
            Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level "Error"
        }
    }
    
    # Exchange Online
    if (Test-ModuleImported -ModuleName "ExchangeOnlineManagement") {
        try {
            Write-Log "Connecting to Exchange Online..." -Level "Info"
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:ServiceConnections["ExchangeOnlineManagement"] = $true
            Write-Log "Successfully connected to Exchange Online" -Level "Success"
            $connections += "Exchange Online"
        }
        catch {
            Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level "Error"
        }
    }
    
    # Microsoft Teams
    if (Test-ModuleImported -ModuleName "MicrosoftTeams") {
        try {
            Write-Log "Connecting to Microsoft Teams..." -Level "Info"
            Connect-MicrosoftTeams -ErrorAction Stop
            $script:ServiceConnections["MicrosoftTeams"] = $true
            Write-Log "Successfully connected to Microsoft Teams" -Level "Success"
            $connections += "Microsoft Teams"
        }
        catch {
            Write-Log "Failed to connect to Microsoft Teams: $($_.Exception.Message)" -Level "Error"
        }
    }
    
    # SharePoint PnP
    if (Test-ModuleImported -ModuleName "PnP.PowerShell") {
        $siteUrlForm = New-Object System.Windows.Forms.Form
        $siteUrlForm.Text = "SharePoint Connection"
        $siteUrlForm.Size = New-Object System.Drawing.Size(400, 150)
        $siteUrlForm.StartPosition = "CenterParent"
        $siteUrlForm.FormBorderStyle = "FixedDialog"
        $siteUrlForm.MaximizeBox = $false
        $siteUrlForm.MinimizeBox = $false
        
        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Enter SharePoint Site URL:"
        $label.Location = New-Object System.Drawing.Point(20, 20)
        $label.Size = New-Object System.Drawing.Size(350, 20)
        
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(20, 45)
        $textBox.Size = New-Object System.Drawing.Size(350, 20)
        $textBox.Text = "https://yourtenant.sharepoint.com"
        
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Connect"
        $btnOK.Location = New-Object System.Drawing.Point(200, 80)
        $btnOK.Size = New-Object System.Drawing.Size(80, 30)
        $btnOK.DialogResult = "OK"
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Location = New-Object System.Drawing.Point(290, 80)
        $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
        $btnCancel.DialogResult = "Cancel"
        
        $siteUrlForm.Controls.AddRange(@($label, $textBox, $btnOK, $btnCancel))
        $siteUrlForm.AcceptButton = $btnOK
        $siteUrlForm.CancelButton = $btnCancel
        
        if ($siteUrlForm.ShowDialog($ParentForm) -eq "OK" -and $textBox.Text.Trim() -ne "") {
            try {
                Write-Log "Connecting to SharePoint PnP..." -Level "Info"
                Connect-PnPOnline -Url $textBox.Text.Trim() -Interactive -ErrorAction Stop
                $script:ServiceConnections["PnP.PowerShell"] = $true
                Write-Log "Successfully connected to SharePoint PnP" -Level "Success"
                $connections += "SharePoint PnP"
            }
            catch {
                Write-Log "Failed to connect to SharePoint PnP: $($_.Exception.Message)" -Level "Error"
            }
        }
        
        $siteUrlForm.Dispose()
    }
    
    # Update status after connections
    Update-ModuleStatus
    
    # Show results
    if ($connections.Count -gt 0) {
        $message = "Successfully connected to:`n`n" + ($connections -join "`n")
        [System.Windows.Forms.MessageBox]::Show($message, "Connection Status", "OK", "Information")
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("No services were connected. Please check that modules are imported first.", "Connection Status", "OK", "Warning")
    }
}

# Helper function to create modern buttons
function New-ModernButton {
    param(
        [string]$Text,
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215),
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::White
    )
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = $Location
    $button.Size = $Size
    $button.BackColor = $BackColor
    $button.ForeColor = $ForeColor
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderSize = 0
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $button.Cursor = "Hand"
    
    # Store original color
    $button.Tag = $BackColor
    
    # Add hover effects
    $button.Add_MouseEnter({
        $originalColor = $this.Tag
        $r = [Math]::Max(0, $originalColor.R - 20)
        $g = [Math]::Max(0, $originalColor.G - 20)
        $b = [Math]::Max(0, $originalColor.B - 35)
        $this.BackColor = [System.Drawing.Color]::FromArgb($r, $g, $b)
    })
    
    $button.Add_MouseLeave({
        $this.BackColor = $this.Tag
    })
    
    return $button
}

# Initialize GUI
function Initialize-GUI {
    Write-Log "Initializing GUI..." -Level "Info"
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:Config.AppName) v$($script:Config.Version)"
    $form.Size = New-Object System.Drawing.Size(1000, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Add icon
    try {
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    }
    catch {
        # Icon not available, continue without it
    }
    
    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $script:Config.AppName
    $titleLabel.Location = New-Object System.Drawing.Point(30, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 30)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    
    # Create subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Manage PowerShell modules and Microsoft 365 service connections | v$($script:Config.Version)"
    $subtitleLabel.Location = New-Object System.Drawing.Point(30, 50)
    $subtitleLabel.Size = New-Object System.Drawing.Size(700, 20)
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
    
    # Create module selection area
    $moduleGroupBox = New-Object System.Windows.Forms.GroupBox
    $moduleGroupBox.Text = "Module Selection & Status"
    $moduleGroupBox.Location = New-Object System.Drawing.Point(30, 90)
    $moduleGroupBox.Size = New-Object System.Drawing.Size(550, 280)
    $moduleGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    # Create module checkboxes and status labels
    $yPos = 30
    foreach ($module in $script:Config.RequiredModules) {
        # Checkbox
        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Text = $module.DisplayName
        $checkbox.Location = New-Object System.Drawing.Point(20, $yPos)
        $checkbox.Size = New-Object System.Drawing.Size(200, 25)
        $checkbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        
        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabelY = $yPos + 3
        $statusLabel.Location = New-Object System.Drawing.Point(230, $statusLabelY)
        $statusLabel.Size = New-Object System.Drawing.Size(150, 20)
        $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        
        # Category label
        $categoryLabel = New-Object System.Windows.Forms.Label
        $categoryLabel.Text = "[$($module.Category)]"
        $categoryLabelY = $yPos + 3
        $categoryLabel.Location = New-Object System.Drawing.Point(390, $categoryLabelY)
        $categoryLabel.Size = New-Object System.Drawing.Size(100, 20)
        $categoryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $categoryLabel.ForeColor = [System.Drawing.Color]::Gray
        
        # Store references
        $script:GUI.ModuleCheckboxes[$module.Name] = $checkbox
        $script:GUI.StatusLabels[$module.Name] = $statusLabel
        
        # Add to group box
        $moduleGroupBox.Controls.AddRange(@($checkbox, $statusLabel, $categoryLabel))
        
        $yPos += 35
    }
    
    # Create action buttons panel
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Location = New-Object System.Drawing.Point(600, 90)
    $buttonPanel.Size = New-Object System.Drawing.Size(350, 320)
    
    $btnInstallAll = New-ModernButton -Text "Install All Modules" -Location (New-Object System.Drawing.Point(0, 20)) -Size (New-Object System.Drawing.Size(160, 40))
    $btnInstallSelected = New-ModernButton -Text "Install Selected" -Location (New-Object System.Drawing.Point(170, 20)) -Size (New-Object System.Drawing.Size(160, 40))
    
    $btnImportAll = New-ModernButton -Text "Import All Modules" -Location (New-Object System.Drawing.Point(0, 70)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(0, 150, 0))
    $btnImportSelected = New-ModernButton -Text "Import Selected" -Location (New-Object System.Drawing.Point(170, 70)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(0, 150, 0))
    
    $btnConnectServices = New-ModernButton -Text "Connect to Services" -Location (New-Object System.Drawing.Point(0, 120)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(150, 0, 150))
    $btnRemoveAll = New-ModernButton -Text "Remove All Modules" -Location (New-Object System.Drawing.Point(170, 120)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(180, 0, 0))
    
    $btnDisconnect = New-ModernButton -Text "Disconnect Selected" -Location (New-Object System.Drawing.Point(0, 170)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(200, 100, 0))
    $btnUnload = New-ModernButton -Text "Unload Selected" -Location (New-Object System.Drawing.Point(170, 170)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(150, 75, 0))
    
    $btnRefresh = New-ModernButton -Text "Refresh Status" -Location (New-Object System.Drawing.Point(0, 220)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(100, 100, 100))
    $btnOpenLog = New-ModernButton -Text "Open Log File" -Location (New-Object System.Drawing.Point(170, 220)) -Size (New-Object System.Drawing.Size(160, 40)) -BackColor ([System.Drawing.Color]::FromArgb(100, 100, 100))
    
    $buttonPanel.Controls.AddRange(@($btnInstallAll, $btnInstallSelected, $btnImportAll, $btnImportSelected, $btnConnectServices, $btnRemoveAll, $btnDisconnect, $btnUnload, $btnRefresh, $btnOpenLog))
    
    # Create progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(30, 390)
    $progressBar.Size = New-Object System.Drawing.Size(920, 25)
    $progressBar.Style = "Continuous"
    
    # Create status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready"
    $statusLabel.Location = New-Object System.Drawing.Point(30, 420)
    $statusLabel.Size = New-Object System.Drawing.Size(920, 20)
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create log text box
    $logGroupBox = New-Object System.Windows.Forms.GroupBox
    $logGroupBox.Text = "Activity Log"
    $logGroupBox.Location = New-Object System.Drawing.Point(30, 450)
    $logGroupBox.Size = New-Object System.Drawing.Size(920, 250)
    $logGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $logTextBox = New-Object System.Windows.Forms.TextBox
    $logTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $logTextBox.Size = New-Object System.Drawing.Size(900, 210)
    $logTextBox.Multiline = $true
    $logTextBox.ScrollBars = "Vertical"
    $logTextBox.ReadOnly = $true
    $logTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $logTextBox.BackColor = [System.Drawing.Color]::Black
    $logTextBox.ForeColor = [System.Drawing.Color]::Lime
    
    $logGroupBox.Controls.Add($logTextBox)
    
    # Add all controls to form
    $form.Controls.AddRange(@($titleLabel, $subtitleLabel, $moduleGroupBox, $buttonPanel, $progressBar, $statusLabel, $logGroupBox))
    
    # Store references
    $script:GUI.Form = $form
    $script:GUI.ProgressBar = $progressBar
    $script:GUI.StatusLabel = $statusLabel
    $script:GUI.LogTextBox = $logTextBox
    
    # Button event handlers
    $btnInstallAll.Add_Click({
        try {
            $script:GUI.StatusLabel.Text = "Installing all modules..."
            Write-Log "Starting installation of all modules" -Level "Info"
            foreach ($module in $script:Config.RequiredModules) {
                Install-RequiredModule -ModuleInfo $module -ProgressBar $script:GUI.ProgressBar
            }
            $script:GUI.StatusLabel.Text = "All module installation completed"
        }
        catch {
            Write-Log "Error in Install All: $($_.Exception.Message)" -Level "Error"
        }
    })
    
    $btnInstallSelected.Add_Click({
        try {
            $selectedModules = Get-SelectedModules
            
            if ($selectedModules.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select modules to install by checking the boxes.", "No Selection", "OK", "Warning")
                return
            }
            
            $script:GUI.StatusLabel.Text = "Installing $($selectedModules.Count) selected modules..."
            Write-Log "Installing selected modules: $(@($selectedModules.DisplayName) -join ', ')" -Level "Info"
            
            foreach ($module in $selectedModules) {
                Install-RequiredModule -ModuleInfo $module -ProgressBar $script:GUI.ProgressBar
            }
            $script:GUI.StatusLabel.Text = "Selected module installation completed"
        }
        catch {
            Write-Log "Error in Install Selected: $($_.Exception.Message)" -Level "Error"
        }
    })
    
    $btnImportAll.Add_Click({
        try {
            # Check for SharePoint conflicts
            if (Test-SharePointConflict) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Both SharePoint modules are currently imported. This may cause conflicts.`n`nPlease restart PowerShell and import only one SharePoint module.",
                    "SharePoint Conflict Warning",
                    "OK",
                    "Warning"
                )
                return
            }
            
            # Filter out conflicting SharePoint modules
            $modulesToImport = $script:Config.RequiredModules
            $hasBothSharePoint = ($modulesToImport | Where-Object { $_.Name -eq "Microsoft.Online.SharePoint.PowerShell" }) -and 
                                ($modulesToImport | Where-Object { $_.Name -eq "PnP.PowerShell" })
            
            if ($hasBothSharePoint) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Both SharePoint modules cannot be imported together due to conflicts.`n`nWould you like to import PnP.PowerShell (recommended modern module)?`n`nYes = PnP.PowerShell`nNo = Skip both SharePoint modules",
                    "SharePoint Module Selection",
                    "YesNo",
                    "Question"
                )
                
                if ($result -eq "Yes") {
                    $modulesToImport = $modulesToImport | Where-Object { $_.Name -ne "Microsoft.Online.SharePoint.PowerShell" }
                } else {
                    $modulesToImport = $modulesToImport | Where-Object { 
                        $_.Name -ne "Microsoft.Online.SharePoint.PowerShell" -and $_.Name -ne "PnP.PowerShell" 
                    }
                }
            }
            
            # Show warning about import times
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Some modules (Az, Microsoft.Graph) may take 3-5 minutes to import.`n`nThe application is NOT frozen during this time - please be patient.`n`nContinue with import?",
                "Import Time Warning",
                "YesNo",
                "Information"
            )
            
            if ($result -ne "Yes") {
                return
            }
            
            $script:GUI.StatusLabel.Text = "Importing all modules..."
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
            Write-Log "Starting import of all modules" -Level "Info"
            
            foreach ($module in $modulesToImport) {
                $script:GUI.StatusLabel.Text = "Importing $($module.DisplayName)..."
                Import-RequiredModule -ModuleInfo $module
            }
            
            $script:GUI.StatusLabel.Text = "All module import completed"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Green
        }
        catch {
            Write-Log "Error in Import All: $($_.Exception.Message)" -Level "Error"
            $script:GUI.StatusLabel.Text = "Error during import"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $btnImportSelected.Add_Click({
        try {
            $selectedModules = Get-SelectedModules
            
            if ($selectedModules.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select modules to import by checking the boxes.", "No Selection", "OK", "Warning")
                return
            }
            
            # Check for SharePoint conflicts in selection
            $hasMSOL = $selectedModules | Where-Object { $_.Name -eq "Microsoft.Online.SharePoint.PowerShell" }
            $hasPnP = $selectedModules | Where-Object { $_.Name -eq "PnP.PowerShell" }
            
            if ($hasMSOL -and $hasPnP) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "You've selected both SharePoint modules. They cannot be imported together due to conflicts.`n`nWould you like to import PnP.PowerShell (recommended modern module)?`n`nYes = PnP.PowerShell`nNo = Microsoft.Online.SharePoint.PowerShell",
                    "SharePoint Module Selection",
                    "YesNo",
                    "Question"
                )
                
                if ($result -eq "Yes") {
                    $selectedModules = $selectedModules | Where-Object { $_.Name -ne "Microsoft.Online.SharePoint.PowerShell" }
                } else {
                    $selectedModules = $selectedModules | Where-Object { $_.Name -ne "PnP.PowerShell" }
                }
            }
            
            # Check if large modules are selected
            $largeModules = @($selectedModules | Where-Object { $_.Name -in @("Az", "Microsoft.Graph") })
            if ($largeModules.Count -gt 0) {
                $largeModuleNames = ($largeModules.DisplayName -join ", ")
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Selected large modules ($largeModuleNames) may take 3-5 minutes each to import.`n`nThe application is NOT frozen during this time - please be patient.`n`nContinue with import?",
                    "Import Time Warning",
                    "YesNo",
                    "Information"
                )
                
                if ($result -ne "Yes") {
                    return
                }
            }
            
            $script:GUI.StatusLabel.Text = "Importing selected modules..."
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
            Write-Log "Importing selected modules: $(@($selectedModules.DisplayName) -join ', ')" -Level "Info"
            
            foreach ($module in $selectedModules) {
                $script:GUI.StatusLabel.Text = "Importing $($module.DisplayName)..."
                Import-RequiredModule -ModuleInfo $module
            }
            
            $script:GUI.StatusLabel.Text = "Selected module import completed"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Green
        }
        catch {
            Write-Log "Error in Import Selected: $($_.Exception.Message)" -Level "Error"
            $script:GUI.StatusLabel.Text = "Error during import"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $btnConnectServices.Add_Click({
        try {
            $script:GUI.StatusLabel.Text = "Connecting to services..."
            Connect-ToServices -ParentForm $form
            $script:GUI.StatusLabel.Text = "Service connection completed"
        }
        catch {
            Write-Log "Error in Connect Services: $($_.Exception.Message)" -Level "Error"
        }
    })
    
    $btnRemoveAll.Add_Click({
        try {
            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove all imported modules?", "Confirm Removal", "YesNo", "Question")
            if ($result -eq "Yes") {
                $script:GUI.StatusLabel.Text = "Removing all modules..."
                foreach ($module in $script:Config.RequiredModules) {
                    Remove-ImportedModule -ModuleInfo $module
                }
                $script:GUI.StatusLabel.Text = "Module removal completed"
            }
        }
        catch {
            Write-Log "Error in Remove All: $($_.Exception.Message)" -Level "Error"
        }
    })
    
    $btnDisconnect.Add_Click({
        try {
            $selectedModules = Get-SelectedModules
            
            if ($selectedModules.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select modules to disconnect by checking the boxes.", "No Selection", "OK", "Warning")
                return
            }
            
            $script:GUI.StatusLabel.Text = "Disconnecting from selected services..."
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
            Write-Log "Disconnecting from selected services: $(@($selectedModules.DisplayName) -join ', ')" -Level "Info"
            
            $btnDisconnect.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            
            $disconnectedServices = Disconnect-SelectedServices -SelectedModules $selectedModules
            
            $btnDisconnect.Enabled = $true
            
            $script:GUI.StatusLabel.Text = "Disconnect completed"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Green
            
            if ($disconnectedServices.Count -gt 0) {
                $message = "Successfully disconnected from:`n`n" + ($disconnectedServices -join "`n")
                [System.Windows.Forms.MessageBox]::Show($message, "Disconnect Status", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("No services were disconnected. Selected modules may not have active connections.", "Disconnect Status", "OK", "Information")
            }
        }
        catch {
            $btnDisconnect.Enabled = $true
            Write-Log "Error in Disconnect Selected: $($_.Exception.Message)" -Level "Error"
            $script:GUI.StatusLabel.Text = "Error during disconnect"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $btnUnload.Add_Click({
        try {
            $selectedModules = Get-SelectedModules
            
            if ($selectedModules.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select modules to unload by checking the boxes.", "No Selection", "OK", "Warning")
                return
            }
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to unload the selected modules?`n`nThis will disconnect from services and remove modules from memory.`n`nSelected modules: $(@($selectedModules.DisplayName) -join ', ')",
                "Confirm Unload",
                "YesNo",
                "Question"
            )
            
            if ($result -ne "Yes") {
                return
            }
            
            $script:GUI.StatusLabel.Text = "Unloading selected modules..."
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
            Write-Log "Unloading selected modules: $(@($selectedModules.DisplayName) -join ', ')" -Level "Info"
            
            $btnUnload.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            
            $unloadResults = Unload-SelectedModules -SelectedModules $selectedModules
            
            $btnUnload.Enabled = $true
            
            $script:GUI.StatusLabel.Text = "Unload completed"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Green
            
            # Show results
            $message = ""
            if ($unloadResults.Unloaded.Count -gt 0) {
                $message += "Successfully unloaded:`n" + ($unloadResults.Unloaded -join "`n")
            }
            if ($unloadResults.Failed.Count -gt 0) {
                if ($message) { $message += "`n`n" }
                $message += "Failed to unload:`n" + ($unloadResults.Failed -join "`n")
            }
            if (-not $message) {
                $message = "No modules were unloaded. Selected modules may not be currently imported."
            }
            
            $iconType = if ($unloadResults.Failed.Count -gt 0) { "Warning" } else { "Information" }
            [System.Windows.Forms.MessageBox]::Show($message, "Unload Status", "OK", $iconType)
        }
        catch {
            $btnUnload.Enabled = $true
            Write-Log "Error in Unload Selected: $($_.Exception.Message)" -Level "Error"
            $script:GUI.StatusLabel.Text = "Error during unload"
            $script:GUI.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $btnRefresh.Add_Click({
        try {
            $script:GUI.StatusLabel.Text = "Refreshing status..."
            Write-Log "Manual refresh requested" -Level "Info"
            Update-ModuleStatus
            $script:GUI.StatusLabel.Text = "Status refreshed"
            Write-Log "Module status refreshed successfully" -Level "Success"
        }
        catch {
            Write-Log "Error during refresh: $($_.Exception.Message)" -Level "Error"
            $script:GUI.StatusLabel.Text = "Error during refresh"
        }
    })
    
    $btnOpenLog.Add_Click({
        if (Test-Path $script:Config.LogFile) {
            Start-Process notepad.exe -ArgumentList $script:Config.LogFile
        } else {
            [System.Windows.Forms.MessageBox]::Show("Log file not found.", "File Not Found", "OK", "Information")
        }
    })
    
    # Initial status update
    Update-ModuleStatus
    
    Write-Log "GUI initialized successfully" -Level "Success"
    return $form
}

# Main execution
function Start-ModuleManagerGUI {
    try {
        # Initialize GUI reference variable first
        $script:GUI = @{
            Form = $null
            ModuleCheckboxes = @{}
            StatusLabels = @{}
            ProgressBar = $null
            StatusLabel = $null
            LogTextBox = $null
        }
        
        # Check if running as administrator
        if (-not $SkipAdminCheck) {
            $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
            if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Write-Log "Warning: Not running as administrator. Some operations may fail." -Level "Warning"
            }
        }
        
        Write-Log "Starting $($script:Config.AppName) v$($script:Config.Version)" -Level "Info"
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level "Info"
        Write-Log "Operating System: $([System.Environment]::OSVersion.VersionString)" -Level "Info"
        Write-Log "Debug Mode: $($script:Config.DebugMode)" -Level "Info"
        
        # Check initial SharePoint status
        Write-Log "Performing initial environment check..." -Level "Info"
        $msolImported = Test-ModuleImported -ModuleName "Microsoft.Online.SharePoint.PowerShell"
        $pnpImported = Test-ModuleImported -ModuleName "PnP.PowerShell"
        
        if ($msolImported -and $pnpImported) {
            Write-Log "WARNING: Both SharePoint modules are currently imported - this will cause conflicts!" -Level "Warning"
            Write-Log "RECOMMENDATION: Restart PowerShell for best results" -Level "Warning"
        } elseif ($msolImported) {
            Write-Log "MSOL-SharePoint-PowerShell is currently imported" -Level "Info"
        } elseif ($pnpImported) {
            Write-Log "PnP.PowerShell is currently imported" -Level "Info"
        }
        
        # Initialize and show GUI
        $form = Initialize-GUI
        
        # Ensure the form handle is created before showing
        $form.CreateControl()
        
        $form.Add_Shown({ 
            $form.Activate()
            Write-Log "Application ready for use" -Level "Success"
            Write-Log "v3.5.3 - Fixed module auto-import and refresh issues" -Level "Info"
        })
        
        [void]$form.ShowDialog()
        
        Write-Log "Application closed" -Level "Info"
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Critical error: $errorMessage" -ForegroundColor Red
        Write-Log "Critical error: $errorMessage" -Level "Error"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "Debug"
        [System.Windows.Forms.MessageBox]::Show("A critical error occurred: $errorMessage", "Error", "OK", "Error")
    }
}

# Start the application
Start-ModuleManagerGUI