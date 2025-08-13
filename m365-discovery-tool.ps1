# Microsoft 365 Discovery Tool with GUI
# Version: 1.1.0
# Requires PowerShell 7+ and pre-installed modules:
# - ExchangeOnlineManagement (V3.8.0)
# - PnP.PowerShell (V3.1.0)
# - MicrosoftPowerBIMgmt (V1.2.1111)
# - Az.Accounts, Az.Resources

$script:ToolVersion = "1.1.0"

# Set console encoding for emoji support
if ($PSVersionTable.PSVersion.Major -ge 7) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

# Load required assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Test emoji support
$script:EmojiSupported = $true
try {
    $testLabel = New-Object System.Windows.Forms.Label
    $testLabel.Text = "🚀"
    $testLabel.Dispose()
} catch {
    $script:EmojiSupported = $false
}

# Helper function to get display text with or without emoji
function Get-DisplayText {
    param(
        [string]$Emoji,
        [string]$Text
    )
    
    if ($script:EmojiSupported) {
        return "$Emoji $Text"
    } else {
        return $Text
    }
}

<#
.SYNOPSIS
    Microsoft 365 Discovery Tool with GUI - Version 1.1.0
.DESCRIPTION
    Comprehensive M365 environment discovery tool with streaming support for large tenants
.NOTES
    Author: Enhanced Discovery Tool
    Version: 1.1.0
    Requires pre-installed PowerShell modules
#>

# Global variables for the script
$Global:ScriptConfig = @{
    Version = $script:ToolVersion
    WorkingFolder = ""
    RunFolder = ""
    OpCo = ""
    GlobalAdminAccount = ""
    UseMFA = $false
    SkipPowerBI = $false
    SkipTeams = $false
    SkipSharePoint = $false
    SkipExchange = $false
    SkipAzureAD = $false
    LogFile = ""
    SharePointAdminSite = ""
    Form = $null
    ProgressForm = $null
    ErrorSummary = @()
    DryRun = $false
    TenantSize = "Small" # Small, Medium, Large
    StreamingEnabled = $false
    ConfigFile = ""
}

#region Helper Functions

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to console
    Write-Host $LogEntry -ForegroundColor $(switch($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    })
    
    # Write to log file if configured
    if (![string]::IsNullOrEmpty($Global:ScriptConfig.LogFile)) {
        try {
            Add-Content -Path $Global:ScriptConfig.LogFile -Value $LogEntry -ErrorAction SilentlyContinue
        } catch {
            # Ignore log file errors
        }
    }
    
    # Update GUI if available
    Update-ProgressDisplay $Message
}

function Update-ProgressDisplay {
    param([string]$Message)
    
    if ($Global:ScriptConfig.ProgressForm -and !$Global:ScriptConfig.ProgressForm.IsDisposed) {
        try {
            $statusLabel = $Global:ScriptConfig.ProgressForm.Controls["statusLabel"]
            $progressBar = $Global:ScriptConfig.ProgressForm.Controls["progressBar"]
            
            if ($statusLabel) {
                $statusLabel.Text = $Message
                $Global:ScriptConfig.ProgressForm.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            # Ignore GUI update errors
        }
    }
}

function Get-SharePointAdminUrl {
    param([string]$AdminEmail)
    
    if ([string]::IsNullOrEmpty($AdminEmail)) {
        return ""
    }
    
    try {
        $domain = $AdminEmail.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com') {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, extract the main part
            $tenantName = $domain.Split('.')[0]
        }
        
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        Write-LogMessage "SharePoint Admin URL determined: $adminUrl" "INFO"
        return $adminUrl
    }
    catch {
        Write-LogMessage "Error determining SharePoint Admin URL: $($_.Exception.Message)" "WARNING"
        return ""
    }
}

function Test-ModuleAvailability {
    param([string]$ModuleName, [string]$RequiredVersion = "")
    
    $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $module) {
        Write-LogMessage "Module $ModuleName not found" "WARNING"
        return $false
    }
    
    if (![string]::IsNullOrEmpty($RequiredVersion)) {
        if ($module.Version -lt [Version]$RequiredVersion) {
            Write-LogMessage "$ModuleName version $($module.Version) is below required version $RequiredVersion" "WARNING"
            return $false
        }
    }
    
    Write-LogMessage "Module $ModuleName version $($module.Version) is available" "SUCCESS"
    return $true
}

function Save-Configuration {
    param([string]$FilePath)
    
    try {
        $config = @{
            Version = $Global:ScriptConfig.Version
            WorkingFolder = $Global:ScriptConfig.WorkingFolder
            OpCo = $Global:ScriptConfig.OpCo
            GlobalAdminAccount = $Global:ScriptConfig.GlobalAdminAccount
            UseMFA = $Global:ScriptConfig.UseMFA
            TenantSize = $Global:ScriptConfig.TenantSize
            StreamingEnabled = $Global:ScriptConfig.StreamingEnabled
        }
        
        $config | ConvertTo-Json | Out-File -FilePath $FilePath -Encoding UTF8
        Write-LogMessage "Configuration saved to: $FilePath" "SUCCESS"
        return $true
    }
    catch {
        Write-LogMessage "Failed to save configuration: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Load-Configuration {
    param([string]$FilePath)
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-LogMessage "Configuration file not found: $FilePath" "WARNING"
            return $false
        }
        
        $config = Get-Content -Path $FilePath -Raw | ConvertFrom-Json
        
        # Update global config with loaded values
        $Global:ScriptConfig.WorkingFolder = $config.WorkingFolder
        $Global:ScriptConfig.OpCo = $config.OpCo
        $Global:ScriptConfig.GlobalAdminAccount = $config.GlobalAdminAccount
        $Global:ScriptConfig.UseMFA = $config.UseMFA
        
        if ($config.PSObject.Properties.Name -contains 'TenantSize') {
            $Global:ScriptConfig.TenantSize = $config.TenantSize
        }
        if ($config.PSObject.Properties.Name -contains 'StreamingEnabled') {
            $Global:ScriptConfig.StreamingEnabled = $config.StreamingEnabled
        }
        
        Write-LogMessage "Configuration loaded from: $FilePath" "SUCCESS"
        return $true
    }
    catch {
        Write-LogMessage "Failed to load configuration: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Connect-ServicesWithValidation {
    param(
        [System.Management.Automation.PSCredential]$Credentials,
        [bool]$MFAEnabled,
        [bool]$DryRun = $false
    )
    
    if ($DryRun) {
        Write-LogMessage "DRY RUN MODE: Testing service connections..." "INFO"
    }
    
    $maxRetries = 3
    $retryCount = 0
    $connectedServices = @()
    
    while ($retryCount -lt $maxRetries) {
        Write-LogMessage "Connecting to services (Attempt $($retryCount + 1) of $maxRetries)..." "INFO"
        $authenticationFailed = $false
        
        # Exchange Online
        if (-not $Global:ScriptConfig.SkipExchange) {
            try {
                Write-LogMessage "Connecting to Exchange Online..." "INFO"
                
                if (-not $DryRun) {
                    if ($MFAEnabled) {
                        Connect-ExchangeOnline -UserPrincipalName $Global:ScriptConfig.GlobalAdminAccount -ShowBanner:$false
                    } else {
                        Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false
                    }
                    
                    # Test connection
                    $null = Get-OrganizationConfig -ErrorAction Stop
                }
                
                Write-LogMessage "Exchange Online connected successfully" "SUCCESS"
                $connectedServices += "Exchange Online"
            }
            catch {
                Write-LogMessage "Exchange connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled") {
                    $authenticationFailed = $true
                }
                if (-not $DryRun) {
                    $Global:ScriptConfig.SkipExchange = $true
                }
                $Global:ScriptConfig.ErrorSummary += "Exchange connection failed: $($_.Exception.Message)"
            }
        }
        
        # SharePoint Online
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipSharePoint) {
            try {
                Write-LogMessage "Connecting to SharePoint Online..." "INFO"
                
                if (-not $DryRun) {
                    if ($MFAEnabled) {
                        Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Interactive
                    } else {
                        Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Credential $Credentials
                    }
                    
                    # Test connection
                    $null = Get-PnPContext -ErrorAction Stop
                }
                
                Write-LogMessage "SharePoint Online connected successfully" "SUCCESS"
                $connectedServices += "SharePoint Online"
            }
            catch {
                Write-LogMessage "SharePoint connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled") {
                    $authenticationFailed = $true
                }
                if (-not $DryRun) {
                    $Global:ScriptConfig.SkipSharePoint = $true
                }
                $Global:ScriptConfig.ErrorSummary += "SharePoint connection failed: $($_.Exception.Message)"
            }
        }
        
        # PowerBI
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipPowerBI) {
            try {
                Write-LogMessage "Connecting to PowerBI..." "INFO"
                
                if (-not $DryRun) {
                    if ($MFAEnabled) {
                        Connect-PowerBIServiceAccount
                    } else {
                        Connect-PowerBIServiceAccount -Credential $Credentials
                    }
                }
                
                Write-LogMessage "PowerBI connected successfully" "SUCCESS"
                $connectedServices += "PowerBI"
            }
            catch {
                Write-LogMessage "PowerBI connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled") {
                    $authenticationFailed = $true
                }
                if (-not $DryRun) {
                    $Global:ScriptConfig.SkipPowerBI = $true
                }
                $Global:ScriptConfig.ErrorSummary += "PowerBI connection failed: $($_.Exception.Message)"
            }
        }
        
        # Azure AD
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipAzureAD) {
            try {
                Write-LogMessage "Connecting to Azure..." "INFO"
                
                if (-not $DryRun) {
                    if ($MFAEnabled) {
                        Connect-AzAccount
                    } else {
                        Connect-AzAccount -Credential $Credentials
                    }
                    
                    # Test connection
                    $null = Get-AzContext -ErrorAction Stop
                }
                
                Write-LogMessage "Azure connected successfully" "SUCCESS"
                $connectedServices += "Azure"
            }
            catch {
                Write-LogMessage "Azure connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled") {
                    $authenticationFailed = $true
                }
                if (-not $DryRun) {
                    $Global:ScriptConfig.SkipAzureAD = $true
                }
                $Global:ScriptConfig.ErrorSummary += "Azure connection failed: $($_.Exception.Message)"
            }
        }
        
        # Check results
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Successfully connected to: $($connectedServices -join ', ')" "SUCCESS"
            return $true
        }
        
        # Handle retry
        if ($authenticationFailed -and -not $DryRun) {
            $retryCount++
            
            if ($retryCount -lt $maxRetries) {
                $retryMessage = "Authentication failed. Attempt $retryCount of $maxRetries.`n`nDo you want to try again?"
                $retryResult = [System.Windows.Forms.MessageBox]::Show($retryMessage, "Authentication Failed", "YesNo", "Question")
                
                if ($retryResult -eq "No") {
                    return $false
                }
                
                # Disconnect and retry
                Disconnect-AllServices
                Start-Sleep -Seconds 2
                Continue
            }
        } else {
            break
        }
    }
    
    if ($connectedServices.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No services could be connected.", "Connection Failed", "OK", "Error")
    }
    
    return $false
}

function Disconnect-AllServices {
    Write-LogMessage "Disconnecting from all services..." "INFO"
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        }
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
        }
        if (Get-Command Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue } catch { }
        }
        if (Get-Command Disconnect-AzAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-AzAccount -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        }
        Write-LogMessage "Disconnected from all services" "SUCCESS"
    }
    catch {
        Write-LogMessage "Warning: Some connections may not have been properly closed" "WARNING"
    }
}

function Export-DataSafely {
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string]$Description,
        [bool]$StreamingEnabled = $false
    )
    
    try {
        Write-LogMessage "Exporting $Description to $FilePath..." "INFO"
        
        if ($Data -and $Data.Count -gt 0) {
            if ($StreamingEnabled -and $Data.Count -gt 1000) {
                Write-LogMessage "Using streaming export for large dataset ($($Data.Count) records)" "INFO"
                
                # Export in chunks
                $chunkSize = 1000
                $chunks = [Math]::Ceiling($Data.Count / $chunkSize)
                
                for ($i = 0; $i -lt $chunks; $i++) {
                    $start = $i * $chunkSize
                    $end = [Math]::Min(($i + 1) * $chunkSize - 1, $Data.Count - 1)
                    $chunk = $Data[$start..$end]
                    
                    if ($i -eq 0) {
                        $chunk | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
                    } else {
                        $chunk | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Append
                    }
                    
                    Write-LogMessage "Exported chunk $($i + 1) of $chunks" "INFO"
                }
            } else {
                $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            }
            
            Write-LogMessage "Successfully exported $($Data.Count) records for $Description" "SUCCESS"
        } else {
            Write-LogMessage "No data found for $Description" "WARNING"
            "No data found" | Out-File $FilePath
        }
    }
    catch {
        Write-LogMessage "Failed to export $Description: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "Export failed for $Description: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Discovery Functions

function Invoke-PowerBIDiscovery {
    if ($Global:ScriptConfig.SkipPowerBI -or $Global:ScriptConfig.DryRun) {
        Write-LogMessage "Skipping PowerBI discovery..." "INFO"
        return
    }
    
    Write-LogMessage "Starting PowerBI discovery..." "INFO"
    
    try {
        # Test PowerBI connection
        try {
            $testWorkspaces = Get-PowerBIWorkspace -Scope Individual -First 1
        } catch {
            Write-LogMessage "PowerBI connection test failed: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "PowerBI connection test failed"
            return
        }
        
        # PowerBI Activity Events (last 30 days)
        Write-LogMessage "Getting PowerBI activity events..." "INFO"
        try {
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $StartDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddT00:00:00")
            $EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            
            $activities = Get-PowerBIActivityEvent -StartDateTime $StartDate -EndDateTime $EndDate
            if ($activities) {
                $activities | Out-File (Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_UsageActivity.json")
                Write-LogMessage "Exported PowerBI activity events" "SUCCESS"
            }
        } catch {
            Write-LogMessage "Error getting PowerBI activity events: $($_.Exception.Message)" "WARNING"
            $Global:ScriptConfig.ErrorSummary += "PowerBI activity events export failed"
        }
        
        # PowerBI Workspaces
        Write-LogMessage "Getting PowerBI workspaces..." "INFO"
        try {
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
            Export-DataSafely -Data $Workspaces -FilePath (Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_Workspaces.csv") -Description "PowerBI Workspaces" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        } catch {
            Write-LogMessage "Error getting PowerBI workspaces: $($_.Exception.Message)" "WARNING"
            $Global:ScriptConfig.ErrorSummary += "PowerBI workspaces export failed"
        }
        
        Write-LogMessage "PowerBI discovery completed" "SUCCESS"
    }
    catch {
        Write-LogMessage "PowerBI discovery failed: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "PowerBI discovery failed: $($_.Exception.Message)"
    }
}

function Invoke-TeamsDiscovery {
    if ($Global:ScriptConfig.SkipTeams -or $Global:ScriptConfig.DryRun) {
        Write-LogMessage "Skipping Teams discovery..." "INFO"
        return
    }
    
    Write-LogMessage "Starting Teams and M365 Groups discovery..." "INFO"
    
    try {
        # Test Exchange connection
        try {
            $testGroups = Get-UnifiedGroup -ResultSize 1
        } catch {
            Write-LogMessage "Exchange Online connection test failed: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "Teams discovery requires Exchange Online connection"
            return
        }
        
        Write-LogMessage "Getting all Teams and M365 Groups..." "INFO"
        $AllTeamsAndGroups = Get-UnifiedGroup -ResultSize Unlimited
        
        if ($Global:ScriptConfig.StreamingEnabled -and $AllTeamsAndGroups.Count -gt 100) {
            Write-LogMessage "Processing $($AllTeamsAndGroups.Count) groups in streaming mode..." "INFO"
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $outputFile = Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_TeamsGroups.csv"
            
            $processed = 0
            $batch = @()
            $batchSize = 50
            
            foreach ($Team in $AllTeamsAndGroups) {
                $processed++
                Write-LogMessage "Processing group $processed of $($AllTeamsAndGroups.Count): $($Team.DisplayName)" "INFO"
                
                try {
                    $GroupType = if ($Team.ResourceProvisioningOptions -contains "Team") { "Team" } else { "M365 Group" }
                    
                    # Get mailbox statistics
                    $MailboxStats = $null
                    try {
                        $MailboxStats = Get-EXOMailboxStatistics $Team.ExchangeGuid -ErrorAction Stop
                    } catch {
                        Write-LogMessage "Could not get mailbox stats for $($Team.PrimarySmtpAddress)" "WARNING"
                    }
                    
                    # Get SharePoint site info
                    $SharePointSite = $null
                    if ($Team.SharePointSiteUrl -and !$Global:ScriptConfig.SkipSharePoint) {
                        try {
                            $SharePointSite = Get-PnPTenantSite -Identity $Team.SharePointSiteUrl -ErrorAction Stop
                        } catch {
                            Write-LogMessage "Could not get SharePoint site for $($Team.PrimarySmtpAddress)" "WARNING"
                        }
                    }
                    
                    $TeamDetails = [PSCustomObject]@{
                        GroupName = $Team.Alias
                        DisplayName = $Team.DisplayName
                        Notes = $Team.Notes
                        ObjectID = $Team.ExternalDirectoryObjectId
                        GroupType = $GroupType
                        GroupMemberCount = $Team.GroupMemberCount
                        GroupExternalMemberCount = $Team.GroupExternalMemberCount
                        EmailAddress = $Team.PrimarySmtpAddress
                        Mailbox_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                        Mailbox_Size = if ($MailboxStats) { $MailboxStats.TotalItemSize } else { "N/A" }
                        SharePointSiteUrl = $Team.SharePointSiteUrl
                        SP_Status = if ($SharePointSite) { $SharePointSite.Status } else { "N/A" }
                        SP_StorageUsageCurrent = if ($SharePointSite) { $SharePointSite.StorageUsageCurrent } else { "N/A" }
                    }
                    
                    $batch += $TeamDetails
                    
                    # Export batch when it reaches the size limit
                    if ($batch.Count -ge $batchSize) {
                        if ($processed -eq $batchSize) {
                            $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                        } else {
                            $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8 -Append
                        }
                        $batch = @()
                    }
                }
                catch {
                    Write-LogMessage "Error processing team $($Team.PrimarySmtpAddress): $($_.Exception.Message)" "WARNING"
                    $Global:ScriptConfig.ErrorSummary += "Failed to process team: $($Team.PrimarySmtpAddress)"
                }
            }
            
            # Export remaining items
            if ($batch.Count -gt 0) {
                if ($processed -eq $batch.Count) {
                    $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                } else {
                    $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8 -Append
                }
            }
            
            Write-LogMessage "Teams and Groups discovery completed - processed $processed groups" "SUCCESS"
        } else {
            # Non-streaming mode for smaller datasets
            $ArrayToExport = @()
            $Counter = 0
            $Total = $AllTeamsAndGroups.Count
            
            foreach ($Team in $AllTeamsAndGroups) {
                $Counter++
                Write-LogMessage "Processing $($Team.PrimarySmtpAddress) ($Counter of $Total)" "INFO"
                
                try {
                    $GroupType = if ($Team.ResourceProvisioningOptions -contains "Team") { "Team" } else { "M365 Group" }
                    
                    $MailboxStats = $null
                    try {
                        $MailboxStats = Get-EXOMailboxStatistics $Team.ExchangeGuid -ErrorAction Stop
                    } catch { }
                    
                    $SharePointSite = $null
                    if ($Team.SharePointSiteUrl -and !$Global:ScriptConfig.SkipSharePoint) {
                        try {
                            $SharePointSite = Get-PnPTenantSite -Identity $Team.SharePointSiteUrl -ErrorAction Stop
                        } catch { }
                    }
                    
                    $TeamDetails = [PSCustomObject]@{
                        GroupName = $Team.Alias
                        DisplayName = $Team.DisplayName
                        Notes = $Team.Notes
                        ObjectID = $Team.ExternalDirectoryObjectId
                        GroupType = $GroupType
                        GroupMemberCount = $Team.GroupMemberCount
                        GroupExternalMemberCount = $Team.GroupExternalMemberCount
                        EmailAddress = $Team.PrimarySmtpAddress
                        Mailbox_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                        Mailbox_Size = if ($MailboxStats) { $MailboxStats.TotalItemSize } else { "N/A" }
                        SharePointSiteUrl = $Team.SharePointSiteUrl
                        SP_Status = if ($SharePointSite) { $SharePointSite.Status } else { "N/A" }
                        SP_StorageUsageCurrent = if ($SharePointSite) { $SharePointSite.StorageUsageCurrent } else { "N/A" }
                    }
                    $ArrayToExport += $TeamDetails
                }
                catch {
                    Write-LogMessage "Error processing team $($Team.PrimarySmtpAddress): $($_.Exception.Message)" "WARNING"
                }
            }
            
            Export-DataSafely -Data $ArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_TeamsGroups.csv") -Description "Teams and M365 Groups"
        }
        
        Write-LogMessage "Teams and Groups discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Teams discovery failed: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "Teams discovery failed: $($_.Exception.Message)"
    }
}

function Invoke-SharePointDiscovery {
    if ($Global:ScriptConfig.SkipSharePoint -or $Global:ScriptConfig.DryRun) {
        Write-LogMessage "Skipping SharePoint discovery..." "INFO"
        return
    }
    
    Write-LogMessage "Starting SharePoint and OneDrive discovery..." "INFO"
    
    try {
        # Test SharePoint connection
        try {
            $testContext = Get-PnPContext -ErrorAction Stop
        } catch {
            Write-LogMessage "SharePoint Online connection test failed: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "SharePoint connection test failed"
            return
        }
        
        Write-LogMessage "Getting all SharePoint sites..." "INFO"
        $SharePointSites = Get-PnPTenantSite -IncludeOneDriveSites
        
        Write-LogMessage "Processing OneDrive sites..." "INFO"
        $OneDriveSites = $SharePointSites | Where-Object { $_.Template -match "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner
        
        Export-DataSafely -Data $OneDriveSites -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_OneDrive.csv") -Description "OneDrive Sites" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        
        Write-LogMessage "Processing SharePoint sites..." "INFO"
        $SharePointSitesFiltered = $SharePointSites | Where-Object { $_.Template -notmatch "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner
        
        Export-DataSafely -Data $SharePointSitesFiltered -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_SharePoint.csv") -Description "SharePoint Sites" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        
        Write-LogMessage "SharePoint and OneDrive discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "SharePoint discovery failed: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "SharePoint discovery failed: $($_.Exception.Message)"
    }
}

function Invoke-ExchangeDiscovery {
    if ($Global:ScriptConfig.SkipExchange -or $Global:ScriptConfig.DryRun) {
        Write-LogMessage "Skipping Exchange discovery..." "INFO"
        return
    }
    
    Write-LogMessage "Starting Exchange mailbox discovery..." "INFO"
    
    try {
        # Test Exchange connection
        try {
            $testMailboxes = Get-EXOMailbox -ResultSize 1
        } catch {
            Write-LogMessage "Exchange Online connection test failed: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "Exchange connection test failed"
            return
        }
        
        Write-LogMessage "Getting all mailboxes..." "INFO"
        $AllMailboxes = Get-EXOMailbox -IncludeInactiveMailbox -ResultSize Unlimited -PropertySets All
        
        if ($Global:ScriptConfig.StreamingEnabled -and $AllMailboxes.Count -gt 500) {
            Write-LogMessage "Processing $($AllMailboxes.Count) mailboxes in streaming mode..." "INFO"
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $outputFile = Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_Mailboxes.csv"
            
            $processed = 0
            $batch = @()
            $batchSize = 100
            
            foreach ($Mailbox in $AllMailboxes) {
                $processed++
                Write-LogMessage "Processing mailbox $processed of $($AllMailboxes.Count): $($Mailbox.UserPrincipalName)" "INFO"
                
                try {
                    # Get mailbox statistics
                    $MailboxStats = $null
                    try {
                        $MailboxStats = if ($Mailbox.IsInactiveMailbox) {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -IncludeSoftDeletedRecipients
                        } else {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid
                        }
                    } catch {
                        Write-LogMessage "Could not get mailbox stats for $($Mailbox.UserPrincipalName)" "WARNING"
                    }
                    
                    # Handle archive mailbox
                    $ArchiveStats = $null
                    if ($Mailbox.ArchiveName) {
                        try {
                            $ArchiveStats = if ($Mailbox.IsInactiveMailbox) {
                                Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive -IncludeSoftDeletedRecipients
                            } else {
                                Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive
                            }
                        } catch {
                            Write-LogMessage "Could not get archive stats for $($Mailbox.UserPrincipalName)" "WARNING"
                        }
                    }
                    
                    $MailboxDetails = [PSCustomObject]@{
                        UserPrincipalName = $Mailbox.UserPrincipalName
                        DisplayName = $Mailbox.DisplayName
                        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                        HasArchive = if ($Mailbox.ArchiveName) { "Yes" } else { "No" }
                        IsInactiveMailbox = $Mailbox.IsInactiveMailbox
                        LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                        MB_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                        MB_TotalItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalItemSize) {
                            try {
                                [math]::Round(($MailboxStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                            } catch { "N/A" }
                        } else { "N/A" }
                        Archive_ItemCount = if ($ArchiveStats) { $ArchiveStats.ItemCount } else { "N/A" }
                        Archive_TotalItemSizeMB = if ($ArchiveStats -and $ArchiveStats.TotalItemSize) {
                            try {
                                [math]::Round(($ArchiveStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                            } catch { "N/A" }
                        } else { "N/A" }
                    }
                    
                    $batch += $MailboxDetails
                    
                    # Export batch when it reaches the size limit
                    if ($batch.Count -ge $batchSize) {
                        if ($processed -eq $batchSize) {
                            $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                        } else {
                            $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8 -Append
                        }
                        $batch = @()
                    }
                }
                catch {
                    Write-LogMessage "Error processing mailbox $($Mailbox.UserPrincipalName): $($_.Exception.Message)" "WARNING"
                    $Global:ScriptConfig.ErrorSummary += "Failed to process mailbox: $($Mailbox.UserPrincipalName)"
                }
            }
            
            # Export remaining items
            if ($batch.Count -gt 0) {
                if ($processed -eq $batch.Count) {
                    $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                } else {
                    $batch | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8 -Append
                }
            }
            
            Write-LogMessage "Exchange discovery completed - processed $processed mailboxes" "SUCCESS"
        } else {
            # Non-streaming mode for smaller datasets
            $MailboxArrayToExport = @()
            $Counter = 0
            $Total = $AllMailboxes.Count
            
            foreach ($Mailbox in $AllMailboxes) {
                $Counter++
                Write-LogMessage "Processing mailbox: $($Mailbox.UserPrincipalName) ($Counter of $Total)" "INFO"
                
                try {
                    $MailboxStats = $null
                    try {
                        $MailboxStats = if ($Mailbox.IsInactiveMailbox) {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -IncludeSoftDeletedRecipients
                        } else {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid
                        }
                    } catch { }
                    
                    $ArchiveStats = $null
                    if ($Mailbox.ArchiveName) {
                        try {
                            $ArchiveStats = if ($Mailbox.IsInactiveMailbox) {
                                Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive -IncludeSoftDeletedRecipients
                            } else {
                                Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive
                            }
                        } catch { }
                    }
                    
                    $MailboxDetails = [PSCustomObject]@{
                        UserPrincipalName = $Mailbox.UserPrincipalName
                        DisplayName = $Mailbox.DisplayName
                        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                        HasArchive = if ($Mailbox.ArchiveName) { "Yes" } else { "No" }
                        IsInactiveMailbox = $Mailbox.IsInactiveMailbox
                        LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                        MB_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                        MB_TotalItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalItemSize) {
                            try {
                                [math]::Round(($MailboxStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                            } catch { "N/A" }
                        } else { "N/A" }
                        Archive_ItemCount = if ($ArchiveStats) { $ArchiveStats.ItemCount } else { "N/A" }
                        Archive_TotalItemSizeMB = if ($ArchiveStats -and $ArchiveStats.TotalItemSize) {
                            try {
                                [math]::Round(($ArchiveStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                            } catch { "N/A" }
                        } else { "N/A" }
                    }
                    $MailboxArrayToExport += $MailboxDetails
                }
                catch {
                    Write-LogMessage "Error processing mailbox $($Mailbox.UserPrincipalName): $($_.Exception.Message)" "WARNING"
                }
            }
            
            Export-DataSafely -Data $MailboxArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Mailboxes.csv") -Description "Exchange Mailboxes"
        }
        
        Write-LogMessage "Exchange discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Exchange discovery failed: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "Exchange discovery failed: $($_.Exception.Message)"
    }
}

function Invoke-AzureADDiscovery {
    if ($Global:ScriptConfig.SkipAzureAD -or $Global:ScriptConfig.DryRun) {
        Write-LogMessage "Skipping Azure discovery..." "INFO"
        return
    }
    
    Write-LogMessage "Starting Azure discovery..." "INFO"
    
    try {
        # Test Azure connection
        try {
            $testContext = Get-AzContext -ErrorAction Stop
            $TenantId = $testContext.Tenant.Id
        } catch {
            Write-LogMessage "Azure connection test failed: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "Azure connection test failed"
            return
        }
        
        # Enterprise Applications (Service Principals)
        Write-LogMessage "Getting enterprise applications..." "INFO"
        try {
            $EnterpriseApps = Get-AzADServicePrincipal | 
                Where-Object { $_.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp" -or $_.ServicePrincipalType -eq "Application" } | 
                Sort-Object DisplayName | 
                Select-Object Id, AccountEnabled, DisplayName, AppId, ServicePrincipalType
            
            Export-DataSafely -Data $EnterpriseApps -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_EnterpriseApplications.csv") -Description "Enterprise Applications" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        } catch {
            Write-LogMessage "Error getting enterprise applications: $($_.Exception.Message)" "WARNING"
            $Global:ScriptConfig.ErrorSummary += "Enterprise applications export failed"
        }
        
        # Azure AD Users
        Write-LogMessage "Getting Azure AD users..." "INFO"
        try {
            $AzureADUsers = Get-AzADUser | 
                Select-Object UserPrincipalName, DisplayName, Id, UserType, AccountEnabled, Mail
            Export-DataSafely -Data $AzureADUsers -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Users.csv") -Description "Azure AD Users" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        } catch {
            Write-LogMessage "Error getting Azure AD users: $($_.Exception.Message)" "ERROR"
            $Global:ScriptConfig.ErrorSummary += "Azure AD users export failed"
        }
        
        # Azure AD Groups
        Write-LogMessage "Getting Azure AD groups..." "INFO"
        try {
            $AzureADGroups = Get-AzADGroup | 
                Select-Object DisplayName, Id, MailEnabled, SecurityEnabled, Mail, Description
            Export-DataSafely -Data $AzureADGroups -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Groups.csv") -Description "Azure AD Groups" -StreamingEnabled $Global:ScriptConfig.StreamingEnabled
        } catch {
            Write-LogMessage "Error getting Azure AD groups: $($_.Exception.Message)" "WARNING"
            $Global:ScriptConfig.ErrorSummary += "Azure AD groups export failed"
        }
        
        # DevOps Organizations discovery URL
        try {
            $DevOpsUrl = "https://app.vsaex.visualstudio.com/_apis/EnterpriseCatalog/Organizations?tenantId=$TenantId"
            Write-LogMessage "DevOps Organizations discovery URL: $DevOpsUrl" "INFO"
            $DevOpsUrl | Out-File (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_DevOps_Discovery_URL.txt")
        } catch {
            Write-LogMessage "Error generating DevOps URL: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "Azure discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Azure discovery failed: $($_.Exception.Message)" "ERROR"
        $Global:ScriptConfig.ErrorSummary += "Azure discovery failed: $($_.Exception.Message)"
    }
}

#endregion

#region Main Discovery Function

function Start-DiscoveryProcess {
    try {
        Write-LogMessage "=== Microsoft 365 and Azure Discovery Started (v$($Global:ScriptConfig.Version)) ===" "INFO"
        Write-LogMessage "Working Folder: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        Write-LogMessage "Organization: $($Global:ScriptConfig.OpCo)" "INFO"
        Write-LogMessage "Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)" "INFO"
        Write-LogMessage "MFA Enabled: $($Global:ScriptConfig.UseMFA)" "INFO"
        Write-LogMessage "Dry Run Mode: $($Global:ScriptConfig.DryRun)" "INFO"
        Write-LogMessage "Tenant Size: $($Global:ScriptConfig.TenantSize)" "INFO"
        Write-LogMessage "Streaming Enabled: $($Global:ScriptConfig.StreamingEnabled)" "INFO"
        
        # Initialize error summary
        $Global:ScriptConfig.ErrorSummary = @()
        
        # Ensure SharePoint Admin URL is set
        if ([string]::IsNullOrEmpty($Global:ScriptConfig.SharePointAdminSite)) {
            $Global:ScriptConfig.SharePointAdminSite = Get-SharePointAdminUrl -AdminEmail $Global:ScriptConfig.GlobalAdminAccount
        }
        
        # Create working directory if it doesn't exist
        if (-not (Test-Path $Global:ScriptConfig.WorkingFolder)) {
            New-Item -ItemType Directory -Path $Global:ScriptConfig.WorkingFolder -Force | Out-Null
            Write-LogMessage "Created working directory: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        }
        
        # Setup log file with timestamp
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Global:ScriptConfig.LogFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Log_$timestamp.txt"
        
        # Create a run-specific subfolder for this discovery
        $Global:ScriptConfig.RunFolder = Join-Path $Global:ScriptConfig.WorkingFolder "Discovery_$timestamp"
        New-Item -ItemType Directory -Path $Global:ScriptConfig.RunFolder -Force | Out-Null
        Write-LogMessage "Created run folder: $($Global:ScriptConfig.RunFolder)" "INFO"
        
        # Check module availability
        Write-LogMessage "Checking module availability..." "INFO"
        
        $moduleChecks = @(
            @{Name = "ExchangeOnlineManagement"; Version = "3.8.0"; Services = @("Exchange", "Teams")}
            @{Name = "PnP.PowerShell"; Version = "3.1.0"; Services = @("SharePoint", "Teams")}
            @{Name = "MicrosoftPowerBIMgmt"; Version = "1.2.1111"; Services = @("PowerBI")}
            @{Name = "Az.Accounts"; Version = ""; Services = @("Azure")}
            @{Name = "Az.Resources"; Version = ""; Services = @("Azure")}
        )
        
        foreach ($moduleCheck in $moduleChecks) {
            if (-not (Test-ModuleAvailability -ModuleName $moduleCheck.Name -RequiredVersion $moduleCheck.Version)) {
                foreach ($service in $moduleCheck.Services) {
                    switch ($service) {
                        "Exchange" { $Global:ScriptConfig.SkipExchange = $true }
                        "Teams" { $Global:ScriptConfig.SkipTeams = $true }
                        "SharePoint" { $Global:ScriptConfig.SkipSharePoint = $true }
                        "PowerBI" { $Global:ScriptConfig.SkipPowerBI = $true }
                        "Azure" { $Global:ScriptConfig.SkipAzureAD = $true }
                    }
                }
            }
        }
        
        # Check if any modules are still enabled
        if ($Global:ScriptConfig.SkipPowerBI -and $Global:ScriptConfig.SkipTeams -and 
            $Global:ScriptConfig.SkipSharePoint -and $Global:ScriptConfig.SkipExchange -and 
            $Global:ScriptConfig.SkipAzureAD) {
            
            $errorMsg = "No discovery modules can run due to missing PowerShell modules"
            Write-LogMessage $errorMsg "ERROR"
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
            return $false
        }
        
        # Show what will actually run
        $activeModules = @()
        if (-not $Global:ScriptConfig.SkipPowerBI) { $activeModules += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipTeams) { $activeModules += "Teams" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $activeModules += "SharePoint" }
        if (-not $Global:ScriptConfig.SkipExchange) { $activeModules += "Exchange" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $activeModules += "Azure" }
        
        Write-LogMessage "Active discovery modules: $($activeModules -join ', ')" "INFO"
        
        # Get credentials if not dry run
        if (-not $Global:ScriptConfig.DryRun) {
            $Credentials = $null
            if (-not $Global:ScriptConfig.UseMFA) {
                Write-LogMessage "Getting credentials..." "INFO"
                
                $Credentials = Get-Credential -UserName $Global:ScriptConfig.GlobalAdminAccount -Message "Enter password for $($Global:ScriptConfig.GlobalAdminAccount)"
                if (-not $Credentials) {
                    Write-LogMessage "Credentials are required when MFA is not enabled" "ERROR"
                    return $false
                }
                
                # Validate username match
                if ($Credentials.UserName -ne $Global:ScriptConfig.GlobalAdminAccount) {
                    Write-LogMessage "Username mismatch detected" "ERROR"
                    return $false
                }
            }
            
            # Connect to services
            Write-LogMessage "Attempting to connect to services..." "INFO"
            Update-ProgressDisplay "Authenticating with Microsoft 365 services..."
            
            if (-not (Connect-ServicesWithValidation -Credentials $Credentials -MFAEnabled $Global:ScriptConfig.UseMFA -DryRun $Global:ScriptConfig.DryRun)) {
                Write-LogMessage "Service connection failed" "ERROR"
                return $false
            }
        } else {
            Write-LogMessage "DRY RUN MODE: Skipping actual connections" "INFO"
        }
        
        # Run discovery modules
        Write-LogMessage "Starting discovery modules..." "INFO"
        
        try {
            Invoke-PowerBIDiscovery
        } catch {
            $errorMsg = "PowerBI discovery failed: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
        }
        
        try {
            Invoke-TeamsDiscovery
        } catch {
            $errorMsg = "Teams discovery failed: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
        }
        
        try {
            Invoke-SharePointDiscovery
        } catch {
            $errorMsg = "SharePoint discovery failed: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
        }
        
        try {
            Invoke-ExchangeDiscovery
        } catch {
            $errorMsg = "Exchange discovery failed: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
        }
        
        try {
            Invoke-AzureADDiscovery
        } catch {
            $errorMsg = "Azure discovery failed: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
        }
        
        # Generate summary report
        Write-LogMessage "Generating summary report..." "INFO"
        $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
        $SummaryFile = Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_Summary.txt"
        $Summary = @"
Microsoft 365 and Azure Discovery Summary
=========================================
Tool Version: $($Global:ScriptConfig.Version)
Organization: $($Global:ScriptConfig.OpCo)
Discovery Date: $(Get-Date)
Working Folder: $($Global:ScriptConfig.WorkingFolder)
Run Folder: $outputFolder
Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)
Dry Run Mode: $($Global:ScriptConfig.DryRun)
Tenant Size: $($Global:ScriptConfig.TenantSize)
Streaming Enabled: $($Global:ScriptConfig.StreamingEnabled)

Discovery Modules Executed:
- PowerBI Discovery: $(if (-not $Global:ScriptConfig.SkipPowerBI) { "✓ Executed" } else { "✗ Skipped" })
- Teams & Groups Discovery: $(if (-not $Global:ScriptConfig.SkipTeams) { "✓ Executed" } else { "✗ Skipped" })
- SharePoint & OneDrive Discovery: $(if (-not $Global:ScriptConfig.SkipSharePoint) { "✓ Executed" } else { "✗ Skipped" })
- Exchange Discovery: $(if (-not $Global:ScriptConfig.SkipExchange) { "✓ Executed" } else { "✗ Skipped" })
- Azure Discovery: $(if (-not $Global:ScriptConfig.SkipAzureAD) { "✓ Executed" } else { "✗ Skipped" })

Log File: $($Global:ScriptConfig.LogFile)
Summary File: $SummaryFile

Generated Files:
$(Get-ChildItem $outputFolder -Filter "*$($Global:ScriptConfig.OpCo)*" | ForEach-Object { "- $($_.Name)" })

Errors Encountered: $($Global:ScriptConfig.ErrorSummary.Count)
$(if ($Global:ScriptConfig.ErrorSummary.Count -gt 0) {
    "`nError Details:"
    $Global:ScriptConfig.ErrorSummary | ForEach-Object { "- $_" }
})
"@
        $Summary | Out-File $SummaryFile
        
        Write-LogMessage "=== Discovery completed successfully ===" "SUCCESS"
        Write-LogMessage "All reports saved to: $outputFolder" "SUCCESS"
        Write-LogMessage "Summary report: $SummaryFile" "SUCCESS"
        
        return $true
    }
    catch {
        Write-LogMessage "Discovery failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Clean up connections
        if (-not $Global:ScriptConfig.DryRun) {
            Disconnect-AllServices
        }
        Write-LogMessage "=== Discovery process completed ===" "INFO"
    }
}

#endregion

#region GUI Functions

function Show-ProgressForm {
    $Global:ScriptConfig.ProgressForm = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.ProgressForm.Text = "Discovery in Progress"
    $Global:ScriptConfig.ProgressForm.Size = New-Object System.Drawing.Size(500, 300)
    $Global:ScriptConfig.ProgressForm.StartPosition = "CenterParent"
    $Global:ScriptConfig.ProgressForm.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.ProgressForm.MaximizeBox = $false
    $Global:ScriptConfig.ProgressForm.MinimizeBox = $false
    $Global:ScriptConfig.ProgressForm.BackColor = [System.Drawing.Color]::White
    $Global:ScriptConfig.ProgressForm.TopMost = $true
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = Get-DisplayText -Emoji "🚀" -Text "Microsoft 365 Discovery in Progress..."
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $Global:ScriptConfig.ProgressForm.Controls.Add($titleLabel)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Name = "statusLabel"
    $statusLabel.Text = "Initializing discovery process..."
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $statusLabel.Size = New-Object System.Drawing.Size(460, 80)
    $statusLabel.Location = New-Object System.Drawing.Point(20, 60)
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $Global:ScriptConfig.ProgressForm.Controls.Add($statusLabel)
    
    # Add progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Name = "progressBar"
    $progressBar.Size = New-Object System.Drawing.Size(460, 25)
    $progressBar.Location = New-Object System.Drawing.Point(20, 150)
    $progressBar.Style = "Marquee"
    $progressBar.MarqueeAnimationSpeed = 30
    $Global:ScriptConfig.ProgressForm.Controls.Add($progressBar)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = Get-DisplayText -Emoji "❌" -Text "Cancel Discovery"
    $cancelButton.Size = New-Object System.Drawing.Size(130, 35)
    $cancelButton.Location = New-Object System.Drawing.Point(185, 220)
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.FlatStyle = "Flat"
    $cancelButton.Name = "cancelButton"
    
    $cancelButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to cancel the discovery process?", "Cancel Discovery", "YesNo", "Question")
        if ($result -eq "Yes") {
            Write-LogMessage "Discovery cancelled by user" "WARNING"
            $Global:ScriptConfig.ProgressForm.Tag = "CANCELLED"
            Close-ProgressForm
        }
    })
    
    $Global:ScriptConfig.ProgressForm.Controls.Add($cancelButton)
    
    $Global:ScriptConfig.ProgressForm.Show()
    $Global:ScriptConfig.ProgressForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Close-ProgressForm {
    if ($Global:ScriptConfig.ProgressForm -and !$Global:ScriptConfig.ProgressForm.IsDisposed) {
        $Global:ScriptConfig.ProgressForm.Close()
        $Global:ScriptConfig.ProgressForm.Dispose()
        $Global:ScriptConfig.ProgressForm = $null
    }
}

function Start-DiscoveryFromGUI {
    try {
        # Hide main form and show progress
        $Global:ScriptConfig.Form.Hide()
        Show-ProgressForm
        
        # Check if user cancelled during progress form display
        if ($Global:ScriptConfig.ProgressForm.Tag -eq "CANCELLED") {
            $Global:ScriptConfig.Form.Show()
            return
        }
        
        # Start discovery
        $success = Start-DiscoveryProcess
        
        # Check if user cancelled during discovery
        if ($Global:ScriptConfig.ProgressForm.Tag -eq "CANCELLED") {
            Write-LogMessage "Discovery process was cancelled by user" "WARNING"
            [System.Windows.Forms.MessageBox]::Show("Discovery process was cancelled.", "Discovery Cancelled", "OK", "Information")
            $Global:ScriptConfig.Form.Show()
            return
        }
        
        # Close progress form
        Close-ProgressForm
        
        if ($success) {
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $message = "Discovery completed successfully!`n`nReports saved to: $outputFolder"
            
            # Add error summary if there were any errors
            if ($Global:ScriptConfig.ErrorSummary.Count -gt 0) {
                $message += "`n`n⚠️ Some issues were encountered:`n"
                $Global:ScriptConfig.ErrorSummary | ForEach-Object { $message += "• $_`n" }
            }
            
            $message += "`n`nWould you like to open the folder?"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Discovery Complete", "YesNo", "Information")
            
            if ($result -eq "Yes") {
                Start-Process "explorer.exe" $outputFolder
            }
        } else {
            $errorMessage = "Discovery failed. Please check the log file for details."
            
            # Add error summary if available
            if ($Global:ScriptConfig.ErrorSummary.Count -gt 0) {
                $errorMessage += "`n`nErrors encountered:`n"
                $Global:ScriptConfig.ErrorSummary | ForEach-Object { $errorMessage += "• $_`n" }
            }
            
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Discovery Failed", "OK", "Error")
        }
        
        # Show main form again
        $Global:ScriptConfig.Form.Show()
    }
    catch {
        Close-ProgressForm
        [System.Windows.Forms.MessageBox]::Show("Error during discovery: $($_.Exception.Message)", "Error", "OK", "Error")
        $Global:ScriptConfig.Form.Show()
    }
}

function New-DiscoveryGUI {
    # Create the main form
    $Global:ScriptConfig.Form = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.Form.Text = "Microsoft 365 Discovery Tool v$($Global:ScriptConfig.Version)"
    $Global:ScriptConfig.Form.Size = New-Object System.Drawing.Size(850, 800)
    $Global:ScriptConfig.Form.StartPosition = "CenterScreen"
    $Global:ScriptConfig.Form.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.Form.MaximizeBox = $false
    $Global:ScriptConfig.Form.MinimizeBox = $false
    $Global:ScriptConfig.Form.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)

    # Create a header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(830, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Header title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = Get-DisplayText -Emoji "🚀" -Text "Microsoft 365 Discovery Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($titleLabel)
    
    # Version label
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version $($Global:ScriptConfig.Version)"
    $versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $versionLabel.ForeColor = [System.Drawing.Color]::White
    $versionLabel.Size = New-Object System.Drawing.Size(100, 25)
    $versionLabel.Location = New-Object System.Drawing.Point(700, 20)
    $headerPanel.Controls.Add($versionLabel)

    # Header subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Comprehensive M365 environment discovery with streaming support"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
    $subtitleLabel.Size = New-Object System.Drawing.Size(600, 25)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 45)
    $headerPanel.Controls.Add($subtitleLabel)

    $Global:ScriptConfig.Form.Controls.Add($headerPanel)

    # Main content panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Size = New-Object System.Drawing.Size(830, 670)
    $mainPanel.Location = New-Object System.Drawing.Point(10, 100)
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.BorderStyle = "FixedSingle"

    # Configuration Group Box
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = Get-DisplayText -Emoji "🔧" -Text "Basic Configuration"
    $configGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $configGroupBox.Size = New-Object System.Drawing.Size(800, 180)
    $configGroupBox.Location = New-Object System.Drawing.Point(15, 15)
    $configGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Working Folder
    $workingFolderLabel = New-Object System.Windows.Forms.Label
    $workingFolderLabel.Text = "Working Folder *"
    $workingFolderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $workingFolderLabel.Size = New-Object System.Drawing.Size(120, 20)
    $workingFolderLabel.Location = New-Object System.Drawing.Point(20, 30)
    $configGroupBox.Controls.Add($workingFolderLabel)

    $workingFolderTextBox = New-Object System.Windows.Forms.TextBox
    $workingFolderTextBox.Size = New-Object System.Drawing.Size(400, 25)
    $workingFolderTextBox.Location = New-Object System.Drawing.Point(20, 50)
    $workingFolderTextBox.Text = "C:\M365Discovery\"
    $workingFolderTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $configGroupBox.Controls.Add($workingFolderTextBox)

    # Add tooltip for working folder
    $workingFolderTooltip = New-Object System.Windows.Forms.ToolTip
    $workingFolderTooltip.SetToolTip($workingFolderTextBox, "Folder where all discovery reports will be saved")

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse..."
    $browseButton.Size = New-Object System.Drawing.Size(80, 25)
    $browseButton.Location = New-Object System.Drawing.Point(430, 50)
    $browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $browseButton.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $configGroupBox.Controls.Add($browseButton)

    # OpCo
    $opcoLabel = New-Object System.Windows.Forms.Label
    $opcoLabel.Text = "Organization Code *"
    $opcoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $opcoLabel.Size = New-Object System.Drawing.Size(120, 20)
    $opcoLabel.Location = New-Object System.Drawing.Point(530, 30)
    $configGroupBox.Controls.Add($opcoLabel)

    $opcoTextBox = New-Object System.Windows.Forms.TextBox
    $opcoTextBox.Size = New-Object System.Drawing.Size(150, 25)
    $opcoTextBox.Location = New-Object System.Drawing.Point(530, 50)
    $opcoTextBox.Text = "AKQA"
    $opcoTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $configGroupBox.Controls.Add($opcoTextBox)

    # Add tooltip for organization code
    $opcoTooltip = New-Object System.Windows.Forms.ToolTip
    $opcoTooltip.SetToolTip($opcoTextBox, "Short code for your organization (used in file naming)")

    # Save/Load config buttons
    $saveConfigButton = New-Object System.Windows.Forms.Button
    $saveConfigButton.Text = "Save Config"
    $saveConfigButton.Size = New-Object System.Drawing.Size(80, 25)
    $saveConfigButton.Location = New-Object System.Drawing.Point(690, 50)
    $saveConfigButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $saveConfigButton.BackColor = [System.Drawing.Color]::FromArgb(100, 149, 237)
    $saveConfigButton.ForeColor = [System.Drawing.Color]::White
    $configGroupBox.Controls.Add($saveConfigButton)

    $loadConfigButton = New-Object System.Windows.Forms.Button
    $loadConfigButton.Text = "Load Config"
    $loadConfigButton.Size = New-Object System.Drawing.Size(80, 25)
    $loadConfigButton.Location = New-Object System.Drawing.Point(690, 80)
    $loadConfigButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $loadConfigButton.BackColor = [System.Drawing.Color]::FromArgb(100, 149, 237)
    $loadConfigButton.ForeColor = [System.Drawing.Color]::White
    $configGroupBox.Controls.Add($loadConfigButton)

    # Admin Account
    $adminAccountLabel = New-Object System.Windows.Forms.Label
    $adminAccountLabel.Text = "Global Administrator Account *"
    $adminAccountLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $adminAccountLabel.Size = New-Object System.Drawing.Size(200, 20)
    $adminAccountLabel.Location = New-Object System.Drawing.Point(20, 85)
    $configGroupBox.Controls.Add($adminAccountLabel)

    $adminAccountTextBox = New-Object System.Windows.Forms.TextBox
    $adminAccountTextBox.Size = New-Object System.Drawing.Size(400, 25)
    $adminAccountTextBox.Location = New-Object System.Drawing.Point(20, 105)
    $adminAccountTextBox.Text = "admin@yourcompany.com"
    $adminAccountTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $configGroupBox.Controls.Add($adminAccountTextBox)

    # Add tooltip for admin account
    $adminTooltip = New-Object System.Windows.Forms.ToolTip
    $adminTooltip.SetToolTip($adminAccountTextBox, "Global Administrator account with full access to M365 services")

    # MFA Checkbox
    $mfaCheckBox = New-Object System.Windows.Forms.CheckBox
    $mfaCheckBox.Text = "Multi-Factor Authentication (MFA) Enabled"
    $mfaCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mfaCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $mfaCheckBox.Location = New-Object System.Drawing.Point(430, 105)
    $mfaCheckBox.Checked = $false
    $configGroupBox.Controls.Add($mfaCheckBox)

    # Dry Run Checkbox
    $dryRunCheckBox = New-Object System.Windows.Forms.CheckBox
    $dryRunCheckBox.Text = "Dry Run Mode (Test connections only)"
    $dryRunCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dryRunCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $dryRunCheckBox.Location = New-Object System.Drawing.Point(20, 140)
    $dryRunCheckBox.Checked = $false
    $configGroupBox.Controls.Add($dryRunCheckBox)

    # Add tooltip for dry run
    $dryRunTooltip = New-Object System.Windows.Forms.ToolTip
    $dryRunTooltip.SetToolTip($dryRunCheckBox, "Test connections without performing actual discovery")

    # Tenant Size dropdown
    $tenantSizeLabel = New-Object System.Windows.Forms.Label
    $tenantSizeLabel.Text = "Tenant Size:"
    $tenantSizeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tenantSizeLabel.Size = New-Object System.Drawing.Size(80, 20)
    $tenantSizeLabel.Location = New-Object System.Drawing.Point(280, 140)
    $configGroupBox.Controls.Add($tenantSizeLabel)

    $tenantSizeCombo = New-Object System.Windows.Forms.ComboBox
    $tenantSizeCombo.Items.AddRange(@("Small (<100 users)", "Medium (100-1000 users)", "Large (1000+ users)"))
    $tenantSizeCombo.SelectedIndex = 0
    $tenantSizeCombo.Size = New-Object System.Drawing.Size(150, 25)
    $tenantSizeCombo.Location = New-Object System.Drawing.Point(365, 138)
    $tenantSizeCombo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tenantSizeCombo.DropDownStyle = "DropDownList"
    $configGroupBox.Controls.Add($tenantSizeCombo)

    # Streaming checkbox
    $streamingCheckBox = New-Object System.Windows.Forms.CheckBox
    $streamingCheckBox.Text = "Enable streaming for large datasets"
    $streamingCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $streamingCheckBox.Size = New-Object System.Drawing.Size(220, 25)
    $streamingCheckBox.Location = New-Object System.Drawing.Point(530, 140)
    $streamingCheckBox.Checked = $false
    $configGroupBox.Controls.Add($streamingCheckBox)

    # Add tooltip for streaming
    $streamingTooltip = New-Object System.Windows.Forms.ToolTip
    $streamingTooltip.SetToolTip($streamingCheckBox, "Process large datasets in chunks to reduce memory usage")

    $mainPanel.Controls.Add($configGroupBox)

    # Discovery Modules Group Box
    $modulesGroupBox = New-Object System.Windows.Forms.GroupBox
    $modulesGroupBox.Text = Get-DisplayText -Emoji "📊" -Text "Discovery Modules (Uncheck to Skip)"
    $modulesGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $modulesGroupBox.Size = New-Object System.Drawing.Size(800, 200)
    $modulesGroupBox.Location = New-Object System.Drawing.Point(15, 205)
    $modulesGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Module checkboxes
    $powerBICheckBox = New-Object System.Windows.Forms.CheckBox
    $powerBICheckBox.Text = Get-DisplayText -Emoji "📈" -Text "PowerBI Discovery (Workspaces, Reports, Datasets)"
    $powerBICheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $powerBICheckBox.Size = New-Object System.Drawing.Size(380, 25)
    $powerBICheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $powerBICheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($powerBICheckBox)

    $teamsCheckBox = New-Object System.Windows.Forms.CheckBox
    $teamsCheckBox.Text = Get-DisplayText -Emoji "👥" -Text "Teams & Groups Discovery"
    $teamsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $teamsCheckBox.Size = New-Object System.Drawing.Size(380, 25)
    $teamsCheckBox.Location = New-Object System.Drawing.Point(410, 30)
    $teamsCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($teamsCheckBox)

    $sharepointCheckBox = New-Object System.Windows.Forms.CheckBox
    $sharepointCheckBox.Text = Get-DisplayText -Emoji "📁" -Text "SharePoint & OneDrive Discovery"
    $sharepointCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $sharepointCheckBox.Size = New-Object System.Drawing.Size(380, 25)
    $sharepointCheckBox.Location = New-Object System.Drawing.Point(20, 65)
    $sharepointCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($sharepointCheckBox)

    $exchangeCheckBox = New-Object System.Windows.Forms.CheckBox
    $exchangeCheckBox.Text = Get-DisplayText -Emoji "📧" -Text "Exchange Discovery (Mailboxes, Archives)"
    $exchangeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exchangeCheckBox.Size = New-Object System.Drawing.Size(380, 25)
    $exchangeCheckBox.Location = New-Object System.Drawing.Point(410, 65)
    $exchangeCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($exchangeCheckBox)

    $azureADCheckBox = New-Object System.Windows.Forms.CheckBox
    $azureADCheckBox.Text = Get-DisplayText -Emoji "🔑" -Text "Azure Discovery (Users, Groups, Apps)"
    $azureADCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $azureADCheckBox.Size = New-Object System.Drawing.Size(380, 25)
    $azureADCheckBox.Location = New-Object System.Drawing.Point(20, 100)
    $azureADCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($azureADCheckBox)

    # Progress information
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = Get-DisplayText -Emoji "💡" -Text "All modules are enabled by default. Uncheck modules you want to skip."
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $progressLabel.Size = New-Object System.Drawing.Size(750, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(20, 140)
    $progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $modulesGroupBox.Controls.Add($progressLabel)

    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Text = Get-DisplayText -Emoji "⚠️" -Text "Ensure all required PowerShell modules are pre-installed before running."
    $requirementsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $requirementsLabel.Size = New-Object System.Drawing.Size(750, 20)
    $requirementsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $requirementsLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
    $modulesGroupBox.Controls.Add($requirementsLabel)

    $mainPanel.Controls.Add($modulesGroupBox)

    # Prerequisites Group Box
    $prereqGroupBox = New-Object System.Windows.Forms.GroupBox
    $prereqGroupBox.Text = Get-DisplayText -Emoji "📋" -Text "Prerequisites"
    $prereqGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $prereqGroupBox.Size = New-Object System.Drawing.Size(800, 100)
    $prereqGroupBox.Location = New-Object System.Drawing.Point(15, 415)
    $prereqGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    $prereqText = New-Object System.Windows.Forms.Label
    $prereqText.Text = @"
Required PowerShell modules (must be pre-installed):
• ExchangeOnlineManagement (V3.8.0+) • PnP.PowerShell (V3.1.0+) • MicrosoftPowerBIMgmt (V1.2.1111+) • Az.Accounts, Az.Resources

Ensure you have Global Administrator permissions and PowerShell 7+ for best compatibility.
Module installation is handled separately - this tool will check availability before running.
"@
    $prereqText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $prereqText.Size = New-Object System.Drawing.Size(770, 70)
    $prereqText.Location = New-Object System.Drawing.Point(20, 25)
    $prereqText.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $prereqGroupBox.Controls.Add($prereqText)

    $mainPanel.Controls.Add($prereqGroupBox)

    # Action buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = Get-DisplayText -Emoji "🚀" -Text "Start Discovery"
    $runButton.Size = New-Object System.Drawing.Size(150, 40)
    $runButton.Location = New-Object System.Drawing.Point(520, 540)
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($runButton)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = Get-DisplayText -Emoji "✅" -Text "Check Prerequisites"
    $validateButton.Size = New-Object System.Drawing.Size(150, 40)
    $validateButton.Location = New-Object System.Drawing.Point(360, 540)
    $validateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $validateButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $validateButton.ForeColor = [System.Drawing.Color]::White
    $validateButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($validateButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = Get-DisplayText -Emoji "❌" -Text "Exit"
    $exitButton.Size = New-Object System.Drawing.Size(100, 40)
    $exitButton.Location = New-Object System.Drawing.Point(690, 540)
    $exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($exitButton)

    # Status bar
    $statusBar = New-Object System.Windows.Forms.StatusBar
    $statusBar.Text = "Ready - Version $($Global:ScriptConfig.Version)"
    $Global:ScriptConfig.Form.Controls.Add($statusBar)

    $Global:ScriptConfig.Form.Controls.Add($mainPanel)

    # Event handlers
    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select folder for discovery reports"
        $folderBrowser.SelectedPath = $workingFolderTextBox.Text
        
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $workingFolderTextBox.Text = $folderBrowser.SelectedPath + "\"
        }
    })

    $saveConfigButton.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "JSON files (*.json)|*.json"
        $saveDialog.DefaultExt = "json"
        $saveDialog.FileName = "M365Discovery_Config.json"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            # Update global config from form
            $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
            $Global:ScriptConfig.OpCo = $opcoTextBox.Text
            $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
            $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
            $Global:ScriptConfig.TenantSize = @("Small", "Medium", "Large")[$tenantSizeCombo.SelectedIndex]
            $Global:ScriptConfig.StreamingEnabled = $streamingCheckBox.Checked
            
            if (Save-Configuration -FilePath $saveDialog.FileName) {
                [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully!", "Save Configuration", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Failed to save configuration.", "Save Configuration", "OK", "Error")
            }
        }
    })

    $loadConfigButton.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Filter = "JSON files (*.json)|*.json"
        $openDialog.DefaultExt = "json"
        
        if ($openDialog.ShowDialog() -eq "OK") {
            if (Load-Configuration -FilePath $openDialog.FileName) {
                # Update form from global config
                $workingFolderTextBox.Text = $Global:ScriptConfig.WorkingFolder
                $opcoTextBox.Text = $Global:ScriptConfig.OpCo
                $adminAccountTextBox.Text = $Global:ScriptConfig.GlobalAdminAccount
                $mfaCheckBox.Checked = $Global:ScriptConfig.UseMFA
                
                # Update tenant size combo
                $index = switch ($Global:ScriptConfig.TenantSize) {
                    "Small" { 0 }
                    "Medium" { 1 }
                    "Large" { 2 }
                    default { 0 }
                }
                $tenantSizeCombo.SelectedIndex = $index
                
                $streamingCheckBox.Checked = $Global:ScriptConfig.StreamingEnabled
                
                [System.Windows.Forms.MessageBox]::Show("Configuration loaded successfully!", "Load Configuration", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Failed to load configuration.", "Load Configuration", "OK", "Error")
            }
        }
    })

    $validateButton.Add_Click({
        $moduleStatus = @"
Checking module availability...

Required modules and their status:
"@
        
        $modules = @(
            @{Name = "ExchangeOnlineManagement"; Version = "3.8.0"}
            @{Name = "PnP.PowerShell"; Version = "3.1.0"}
            @{Name = "MicrosoftPowerBIMgmt"; Version = "1.2.1111"}
            @{Name = "Az.Accounts"; Version = ""}
            @{Name = "Az.Resources"; Version = ""}
        )
        
        $allAvailable = $true
        foreach ($module in $modules) {
            $available = Test-ModuleAvailability -ModuleName $module.Name -RequiredVersion $module.Version
            $status = if ($available) { "✅ Available" } else { "❌ Not Found/Outdated"; $allAvailable = $false }
            $moduleStatus += "`n$($module.Name): $status"
        }
        
        if ($allAvailable) {
            $moduleStatus += "`n`n🎉 All required modules are available!"
        } else {
            $moduleStatus += "`n`n⚠️ Some modules are missing or outdated. Please install/update them before running discovery."
        }
        
        [System.Windows.Forms.MessageBox]::Show($moduleStatus, "Prerequisites Check", "OK", "Information")
    })

    $runButton.Add_Click({
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($workingFolderTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a working folder.", "Validation Error", "OK", "Warning")
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($opcoTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify an organization code.", "Validation Error", "OK", "Warning")
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($adminAccountTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a global administrator account.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Check for default/placeholder values
        if ($adminAccountTextBox.Text -eq "admin@yourcompany.com" -or $adminAccountTextBox.Text -match "yourcompany|example|test") {
            $placeholderMessage = @"
It looks like you're using a placeholder email address:
$($adminAccountTextBox.Text)

Please enter your actual Global Administrator account email address.

Example: admin@contoso.onmicrosoft.com
"@
            [System.Windows.Forms.MessageBox]::Show($placeholderMessage, "Update Admin Account", "OK", "Warning")
            $adminAccountTextBox.Focus()
            $adminAccountTextBox.SelectAll()
            return
        }
        
        # Validate email format
        if ($adminAccountTextBox.Text -notmatch "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid email address for the administrator account.", "Validation Error", "OK", "Warning")
            $adminAccountTextBox.Focus()
            $adminAccountTextBox.SelectAll()
            return
        }

        # Check if any modules are selected
        if ($powerBICheckBox.Checked -eq $false -and $teamsCheckBox.Checked -eq $false -and 
            $sharepointCheckBox.Checked -eq $false -and $exchangeCheckBox.Checked -eq $false -and 
            $azureADCheckBox.Checked -eq $false) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one discovery module.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Update configuration
        $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
        $Global:ScriptConfig.OpCo = $opcoTextBox.Text
        $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
        $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
        $Global:ScriptConfig.DryRun = $dryRunCheckBox.Checked
        $Global:ScriptConfig.TenantSize = @("Small", "Medium", "Large")[$tenantSizeCombo.SelectedIndex]
        $Global:ScriptConfig.StreamingEnabled = $streamingCheckBox.Checked
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Set SharePoint Admin URL
        $Global:ScriptConfig.SharePointAdminSite = Get-SharePointAdminUrl -AdminEmail $Global:ScriptConfig.GlobalAdminAccount
        
        # Final confirmation with all details
        $selectedModules = @()
        if (-not $Global:ScriptConfig.SkipPowerBI) { $selectedModules += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipTeams) { $selectedModules += "Teams & Groups" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $selectedModules += "SharePoint & OneDrive" }
        if (-not $Global:ScriptConfig.SkipExchange) { $selectedModules += "Exchange" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $selectedModules += "Azure" }
        
        $confirmMessage = @"
Ready to start Microsoft 365 Discovery with these settings:

👤 Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
🔐 MFA Enabled: $($Global:ScriptConfig.UseMFA)
🏢 Organization: $($Global:ScriptConfig.OpCo)
📁 Output Folder: $($Global:ScriptConfig.WorkingFolder)
🔍 Dry Run Mode: $($Global:ScriptConfig.DryRun)
📊 Tenant Size: $($Global:ScriptConfig.TenantSize)
⚡ Streaming: $($Global:ScriptConfig.StreamingEnabled)

📊 Discovery Modules:
$($selectedModules -join "`n")

$(if ($Global:ScriptConfig.DryRun) {
    "⚠️ DRY RUN MODE: Only connections will be tested"
} else {
    "⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed"
})

Do you want to continue?
"@
        
        $result = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Discovery Settings", "YesNo", "Question")
        
        if ($result -eq "Yes") {
            Start-DiscoveryFromGUI
        }
    })

    $exitButton.Add_Click({
        $Global:ScriptConfig.Form.Close()
    })

    # Update status bar when tenant size or streaming changes
    $tenantSizeCombo.Add_SelectedIndexChanged({
        $statusBar.Text = "Tenant size: $($tenantSizeCombo.SelectedItem)"
    })

    $streamingCheckBox.Add_CheckedChanged({
        if ($streamingCheckBox.Checked) {
            $statusBar.Text = "Streaming enabled for large datasets"
        } else {
            $statusBar.Text = "Standard processing mode"
        }
    })

    # Auto-enable streaming for large tenants
    $tenantSizeCombo.Add_SelectedIndexChanged({
        if ($tenantSizeCombo.SelectedIndex -eq 2) { # Large tenant
            $streamingCheckBox.Checked = $true
            $statusBar.Text = "Streaming automatically enabled for large tenant"
        }
    })

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Entry point
Write-Host @"
===============================================
Microsoft 365 Discovery Tool
Version: $($Global:ScriptConfig.Version)
===============================================

This tool requires the following pre-installed PowerShell modules:
- ExchangeOnlineManagement (V3.8.0+)
- PnP.PowerShell (V3.1.0+) 
- MicrosoftPowerBIMgmt (V1.2.1111+)
- Az.Accounts, Az.Resources

Please ensure all modules are installed before proceeding.
Starting GUI...

"@ -ForegroundColor Cyan

# Start the GUI
New-DiscoveryGUI