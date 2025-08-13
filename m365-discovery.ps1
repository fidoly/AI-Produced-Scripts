# Microsoft 365 Discovery Tool with GUI
# Requires PowerShell 7+ for best compatibility with modern modules

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
    Complete Microsoft 365 Discovery Tool with GUI
.DESCRIPTION
    Single-file solution with GUI interface for comprehensive M365 environment discovery
.NOTES
    Author: Enhanced Discovery Tool
    Version: 3.0
    All-in-one solution with GUI and discovery engine
#>

# Global variables for the script
$Global:ScriptConfig = @{
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
}

#region Helper Functions

function Get-SharePointAdminUrl {
    param([string]$AdminEmail)
    
    if ([string]::IsNullOrEmpty($AdminEmail)) {
        return ""
    }
    
    try {
        $domain = $AdminEmail.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com

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
            if ($statusLabel) {
                $statusLabel.Text = $Message
                $Global:ScriptConfig.ProgressForm.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            # Ignore GUI update errors
        }
    }
    
    # Also update console
    if ($Message -match "Connecting to|Successfully connected|Failed to connect") {
        Write-Host $Message -ForegroundColor Cyan
    }
}

function Test-ModuleAndAdjustConfig {
    param([string]$ModuleName, [string]$ModuleDescription)
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-LogMessage "Module $ModuleName not found - $ModuleDescription will be skipped" "WARNING"
        $Global:ScriptConfig.ErrorSummary += "Module $ModuleName not found - $ModuleDescription was skipped"
        return $false
    }
    
    # Module version requirements
    $versionRequirements = @{
        'ExchangeOnlineManagement' = '3.0.0'
        'PnP.PowerShell' = '1.11.0'
        'MicrosoftPowerBIMgmt' = '1.2.0'
        'Az.Accounts' = '2.10.0'
        'Az.Resources' = '6.0.0'
    }
    
    # Check module version if requirements exist
    if ($versionRequirements.ContainsKey($ModuleName)) {
        $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        $requiredVersion = [Version]$versionRequirements[$ModuleName]
        
        if ($module.Version -lt $requiredVersion) {
            $errorMsg = "$ModuleName module version $($module.Version) is outdated. Version $requiredVersion or higher is required."
            Write-LogMessage $errorMsg "WARNING"
            Write-LogMessage "Please update using: Update-Module $ModuleName -Force" "WARNING"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
            return $false
        }
    }
    
    # Try to import the module if it's available but not loaded
    if (-not (Get-Module -Name $ModuleName)) {
        try {
            Write-LogMessage "Importing module $ModuleName..." "INFO"
            Import-Module $ModuleName -Force -ErrorAction Stop
            Write-LogMessage "Successfully imported $ModuleName" "SUCCESS"
        } catch {
            $errorMsg = "Failed to import $ModuleName: $($_.Exception.Message)"
            Write-LogMessage $errorMsg "ERROR"
            $Global:ScriptConfig.ErrorSummary += $errorMsg
            return $false
        }
    }
    
    return $true
}

function Connect-ServicesWithValidation {
    param(
        [System.Management.Automation.PSCredential]$Credentials,
        [bool]$MFAEnabled
    )
    
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        Write-LogMessage "Connecting to services (Attempt $($retryCount + 1) of $maxRetries)..." "INFO"
        
        $connectedServices = @()
        $authenticationFailed = $false
        
        # Try Exchange Online
        if (-not $Global:ScriptConfig.SkipExchange) {
            Write-LogMessage "Connecting to Exchange Online..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, use UserPrincipalName to pre-fill the account
                    Connect-ExchangeOnline -UserPrincipalName $Global:ScriptConfig.GlobalAdminAccount -ShowBanner:$false
                } else {
                    Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false
                }
                
                $testResult = Get-OrganizationConfig -ErrorAction Stop
                Write-LogMessage "Exchange Online connected successfully" "SUCCESS"
                $connectedServices += "Exchange Online"
            }
            catch {
                Write-LogMessage "Exchange connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|0x80070520|cancelled|OAuth") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipExchange = $true
            }
        }
        
        # Try Azure AD (using Az module)
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "Connecting to Azure (Az module)..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, browser will open automatically
                    Connect-AzAccount
                } else {
                    Connect-AzAccount -Credential $Credentials
                }
                
                $testResult = Get-AzContext -ErrorAction Stop
                Write-LogMessage "Azure connected successfully" "SUCCESS"
                $connectedServices += "Azure"
            }
            catch {
                Write-LogMessage "Azure connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled|OAuth") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipAzureAD = $true
            }
        }
        
        # Try SharePoint Online (using PnP.PowerShell)
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipSharePoint) {
            Write-LogMessage "Connecting to SharePoint Online (PnP)..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, use interactive login
                    Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Interactive
                } else {
                    # Create PSCredential object for PnP
                    Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Credential $Credentials
                }
                
                $testResult = Get-PnPContext -ErrorAction Stop
                Write-LogMessage "SharePoint Online connected successfully" "SUCCESS"
                $connectedServices += "SharePoint Online"
            }
            catch {
                Write-LogMessage "SharePoint connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|OAuth|cancelled") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipSharePoint = $true
            }
        }
        
        # Try PowerBI
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipPowerBI) {
            Write-LogMessage "Connecting to PowerBI..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, don't pass any parameters to trigger browser auth
                    Connect-PowerBIServiceAccount
                } else {
                    Connect-PowerBIServiceAccount -Credential $Credentials
                }
                
                Write-LogMessage "PowerBI connected successfully" "SUCCESS"
                $connectedServices += "PowerBI"
            }
            catch {
                Write-LogMessage "PowerBI connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|OAuth|cancelled") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        # Check results
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Successfully connected to: $($connectedServices -join ', ')" "SUCCESS"
            return $true
        }
        
        # Handle retry
        if ($authenticationFailed) {
            $retryCount++
            
            if ($retryCount -lt $maxRetries) {
                $retryMessage = "Authentication failed. Attempt $retryCount of $maxRetries.`n`nDo you want to try again?"
                $retryResult = [System.Windows.Forms.MessageBox]::Show($retryMessage, "Authentication Failed", "YesNo", "Question")
                
                if ($retryResult -eq "No") {
                    return $false
                }
                
                Write-LogMessage "Retrying authentication..." "INFO"
                try {
                    Disconnect-AllServices
                } catch {
                    # Ignore disconnect errors
                }
                Start-Sleep -Seconds 2
                Continue
            } else {
                [System.Windows.Forms.MessageBox]::Show("Maximum authentication attempts reached.", "Authentication Failed", "OK", "Error")
                return $false
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No services could be connected.", "Connection Failed", "OK", "Error")
            return $false
        }
    }
    
    return $false
}

function Disconnect-AllServices {
    Write-LogMessage "Disconnecting from all services..." "INFO"
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
        }
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-PnPOnline } catch { }
        }
        if (Get-Command Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-PowerBIServiceAccount } catch { }
        }
        if (Get-Command Disconnect-AzAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-AzAccount -Confirm:$false } catch { }
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
        [string]$Description
    )
    
    try {
        Write-LogMessage "Exporting $Description to $FilePath..." "INFO"
        if ($Data -and $Data.Count -gt 0) {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            Write-LogMessage "Successfully exported $($Data.Count) records for $Description" "SUCCESS"
        } else {
            Write-LogMessage "No data found for $Description" "WARNING"
            "No data found" | Out-File $FilePath
        }
    }
    catch {
        Write-LogMessage "Failed to export $Description`: $($_.Exception.Message)" "ERROR"
        throw
    }
}

#endregion

#region Discovery Functions

function Invoke-PowerBIDiscovery {
    if ($Global:ScriptConfig.SkipPowerBI) {
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
            return
        }
        
        # PowerBI Activity Events
        Write-LogMessage "Getting PowerBI activity events..." "INFO"
        try {
            # Use RunFolder if available, otherwise WorkingFolder
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $OutputReport = Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI.txt"
            Start-Transcript $OutputReport
            
            $StartDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddT00:00:00")
            $EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            
            Get-PowerBIActivityEvent -StartDateTime $StartDate -EndDateTime $EndDate | 
                Out-File (Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_UsageActivity.txt")
            
            Stop-Transcript
        } catch {
            Write-LogMessage "Error getting PowerBI activity events: $($_.Exception.Message)" "WARNING"
            if (Get-Command Stop-Transcript -ErrorAction SilentlyContinue) { Stop-Transcript }
        }
        
        # PowerBI Workspaces
        Write-LogMessage "Getting PowerBI workspaces..." "INFO"
        try {
            $outputFolder = if ($Global:ScriptConfig.RunFolder) { $Global:ScriptConfig.RunFolder } else { $Global:ScriptConfig.WorkingFolder }
            $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
            Export-DataSafely -Data $Workspaces -FilePath (Join-Path $outputFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_Workspaces.csv") -Description "PowerBI Workspaces"
        } catch {
            Write-LogMessage "Error getting PowerBI workspaces: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "PowerBI discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "PowerBI discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-TeamsDiscovery {
    if ($Global:ScriptConfig.SkipTeams) {
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
            return
        }
        
        Write-LogMessage "Getting all Teams and M365 Groups..." "INFO"
        $AllTeamsAndGroups = Get-UnifiedGroup -ResultSize Unlimited
        $ArrayToExport = @()
        $Counter = 0
        $Total = $AllTeamsAndGroups.Count
        
        Write-LogMessage "Found $Total Teams and M365 Groups to process..." "INFO"
        
        foreach ($Team in $AllTeamsAndGroups) {
            $Counter++
            Write-LogMessage "Processing $($Team.PrimarySmtpAddress) ($Counter of $Total)" "INFO"
            
            try {
                $GroupType = if ($Team.ResourceProvisioningOptions -contains "Team") { "Team" } else { "M365 Group" }
                
                # Get mailbox statistics with error handling
                $MailboxStats = $null
                try {
                    $MailboxStats = Get-EXOMailboxStatistics $Team.ExchangeGuid -ErrorAction Stop
                } catch {
                    Write-LogMessage "Warning: Could not get mailbox stats for $($Team.PrimarySmtpAddress)" "WARNING"
                }
                
                # Get SharePoint site with error handling
                $SharePointSite = $null
                try {
                    if ($Team.SharePointSiteUrl -and !$Global:ScriptConfig.SkipSharePoint) {
                        # Use PnP to get the site
                        $SharePointSite = Get-PnPTenantSite -Identity $Team.SharePointSiteUrl -ErrorAction Stop
                    }
                } catch {
                    Write-LogMessage "Warning: Could not get SharePoint site for $($Team.PrimarySmtpAddress)" "WARNING"
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
                    SP_LockState = if ($SharePointSite) { $SharePointSite.LockState } else { "N/A" }
                    SP_StorageUsageCurrent = if ($SharePointSite) { $SharePointSite.StorageUsageCurrent } else { "N/A" }
                    SP_WebsCount = if ($SharePointSite) { $SharePointSite.WebsCount } else { "N/A" }
                    SP_Template = if ($SharePointSite) { $SharePointSite.Template } else { "N/A" }
                }
                $ArrayToExport += $TeamDetails
            }
            catch {
                Write-LogMessage "Error processing team $($Team.PrimarySmtpAddress): $($_.Exception.Message)" "WARNING"
            }
        }
        
        Export-DataSafely -Data $ArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_TeamsGroups.csv") -Description "Teams and M365 Groups"
        Write-LogMessage "Teams and Groups discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Teams discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-SharePointDiscovery {
    if ($Global:ScriptConfig.SkipSharePoint) {
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
            return
        }
        
        Write-LogMessage "Getting all SharePoint sites..." "INFO"
        # Get all sites using PnP cmdlets
        $SharePointSites = Get-PnPTenantSite -IncludeOneDriveSites
        
        Write-LogMessage "Processing OneDrive sites..." "INFO"
        $OneDriveSites = $SharePointSites | Where-Object { $_.Template -match "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner, ConditionalAccessPolicy
        
        Export-DataSafely -Data $OneDriveSites -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_OneDrive.csv") -Description "OneDrive Sites"
        
        Write-LogMessage "Processing SharePoint sites..." "INFO"
        $SharePointSitesFiltered = $SharePointSites | Where-Object { $_.Template -notmatch "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner, ConditionalAccessPolicy
        
        Export-DataSafely -Data $SharePointSitesFiltered -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_SharePoint.csv") -Description "SharePoint Sites"
        
        Write-LogMessage "SharePoint and OneDrive discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "SharePoint discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-ExchangeDiscovery {
    if ($Global:ScriptConfig.SkipExchange) {
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
            return
        }
        
        Write-LogMessage "Getting all mailboxes..." "INFO"
        $AllMailboxes = Get-EXOMailbox -IncludeInactiveMailbox -ResultSize Unlimited -PropertySets All
        $MailboxArrayToExport = @()
        $Counter = 0
        $Total = $AllMailboxes.Count
        
        Write-LogMessage "Found $Total mailboxes to process..." "INFO"
        
        foreach ($Mailbox in $AllMailboxes) {
            $Counter++
            Write-LogMessage "Processing mailbox: $($Mailbox.UserPrincipalName) ($Counter of $Total)" "INFO"
            
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
                    Write-LogMessage "Warning: Could not get mailbox stats for $($Mailbox.UserPrincipalName)" "WARNING"
                }
                
                # Handle archive mailbox
                $ArchiveDetails = @{
                    HasArchive = "No"
                    DisplayName = ""
                    MailboxGuid = ""
                    ItemCount = ""
                    TotalItemSize = ""
                    DeletedItemCount = ""
                    TotalDeletedItemSize = ""
                }
                
                if ($Mailbox.ArchiveName) {
                    try {
                        $MailboxArchiveStats = if ($Mailbox.IsInactiveMailbox) {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive -IncludeSoftDeletedRecipients
                        } else {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive
                        }
                        
                        $ArchiveDetails = @{
                            HasArchive = "Yes"
                            DisplayName = $MailboxArchiveStats.DisplayName
                            MailboxGuid = $MailboxArchiveStats.MailboxGuid
                            ItemCount = $MailboxArchiveStats.ItemCount
                            TotalItemSize = if ($MailboxArchiveStats.TotalItemSize) {
                                try {
                                    [math]::Round(($MailboxArchiveStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                                } catch { "N/A" }
                            } else { "N/A" }
                            DeletedItemCount = $MailboxArchiveStats.DeletedItemCount
                            TotalDeletedItemSize = if ($MailboxArchiveStats.TotalDeletedItemSize) {
                                try {
                                    [math]::Round(($MailboxArchiveStats.TotalDeletedItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                                } catch { "N/A" }
                            } else { "N/A" }
                        }
                    }
                    catch {
                        Write-LogMessage "Warning: Could not get archive stats for $($Mailbox.UserPrincipalName)" "WARNING"
                    }
                }
                
                $MailboxDetails = [PSCustomObject]@{
                    UserPrincipalName = $Mailbox.UserPrincipalName
                    DisplayName = $Mailbox.DisplayName
                    PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                    ExchangeGuid = $Mailbox.ExchangeGuid
                    RecipientType = $Mailbox.RecipientType
                    RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                    HasArchive = $ArchiveDetails.HasArchive
                    AccountDisabled = $Mailbox.AccountDisabled
                    LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                    RetentionHoldEnabled = $Mailbox.RetentionHoldEnabled
                    IsMailboxEnabled = $Mailbox.IsMailboxEnabled
                    IsInactiveMailbox = $Mailbox.IsInactiveMailbox
                    WhenSoftDeleted = $Mailbox.WhenSoftDeleted
                    ArchiveStatus = $Mailbox.ArchiveStatus
                    WhenMailboxCreated = $Mailbox.WhenMailboxCreated
                    RetentionPolicy = $Mailbox.RetentionPolicy
                    MB_DisplayName = if ($MailboxStats) { $MailboxStats.DisplayName } else { "N/A" }
                    MB_MailboxGuid = if ($MailboxStats) { $MailboxStats.MailboxGuid } else { "N/A" }
                    MB_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                    MB_TotalItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalItemSize) {
                        try {
                            [math]::Round(($MailboxStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                        } catch { "N/A" }
                    } else { "N/A" }
                    Archive_DisplayName = $ArchiveDetails.DisplayName
                    Archive_MailboxGuid = $ArchiveDetails.MailboxGuid
                    Archive_ItemCount = $ArchiveDetails.ItemCount
                    Archive_TotalItemSizeMB = $ArchiveDetails.TotalItemSize
                    MaxSendSize = if ($Mailbox.MaxSendSize) { $Mailbox.MaxSendSize.Split(" (")[0] } else { "N/A" }
                    MaxReceiveSize = if ($Mailbox.MaxReceiveSize) { $Mailbox.MaxReceiveSize.Split(" (")[0] } else { "N/A" }
                    MB_DeletedItemCount = if ($MailboxStats) { $MailboxStats.DeletedItemCount } else { "N/A" }
                    MB_TotalDeletedItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalDeletedItemSize) {
                        try {
                            [math]::Round(($MailboxStats.TotalDeletedItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                        } catch { "N/A" }
                    } else { "N/A" }
                    Archive_DeletedItemCount = $ArchiveDetails.DeletedItemCount
                    Archive_TotalDeletedItemSizeMB = $ArchiveDetails.TotalDeletedItemSize
                }
                $MailboxArrayToExport += $MailboxDetails
            }
            catch {
                Write-LogMessage "Error processing mailbox $($Mailbox.UserPrincipalName): $($_.Exception.Message)" "WARNING"
            }
        }
        
        Export-DataSafely -Data $MailboxArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Mailboxes.csv") -Description "Exchange Mailboxes"
        Write-LogMessage "Exchange discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Exchange discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-AzureADDiscovery {
    if ($Global:ScriptConfig.SkipAzureAD) {
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
            return
        }
        
        # Enterprise Applications (Service Principals)
        Write-LogMessage "Getting enterprise applications..." "INFO"
        try {
            # Get service principals (enterprise apps)
            $EnterpriseApps = Get-AzADServicePrincipal | 
                Where-Object { $_.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp" -or $_.ServicePrincipalType -eq "Application" } | 
                Sort-Object DisplayName
            
            Export-DataSafely -Data ($EnterpriseApps | Select-Object Id, AccountEnabled, DisplayName, AppId, ServicePrincipalType, Tags) -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_EnterpriseApplications.csv") -Description "Enterprise Applications"
        } catch {
            Write-LogMessage "Error getting enterprise applications: $($_.Exception.Message)" "WARNING"
        }
        
        # Azure AD Users
        Write-LogMessage "Getting Azure AD users..." "INFO"
        try {
            $AzureADUsers = Get-AzADUser | 
                Select-Object UserPrincipalName, DisplayName, Id, UserType, AccountEnabled, Mail
            Export-DataSafely -Data $AzureADUsers -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Users.csv") -Description "Azure AD Users"
        } catch {
            Write-LogMessage "Error getting Azure AD users: $($_.Exception.Message)" "ERROR"
        }
        
        # Azure AD Groups
        Write-LogMessage "Getting Azure AD groups..." "INFO"
        try {
            $AzureADGroups = Get-AzADGroup | 
                Select-Object DisplayName, Id, MailEnabled, SecurityEnabled, Mail, Description
            Export-DataSafely -Data $AzureADGroups -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Groups.csv") -Description "Azure AD Groups"
        } catch {
            Write-LogMessage "Error getting Azure AD groups: $($_.Exception.Message)" "WARNING"
        }
        
        # DevOps Organizations discovery URL
        try {
            $DevOpsUrl = "https://app.vsaex.visualstudio.com/_apis/EnterpriseCatalog/Organizations?tenantId=$TenantId"
            Write-LogMessage "DevOps Organizations discovery URL: $DevOpsUrl" "INFO"
            
            # Save DevOps URL to file
            $DevOpsUrl | Out-File (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_DevOps_Discovery_URL.txt")
        } catch {
            Write-LogMessage "Error generating DevOps URL: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "Azure discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Azure discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

#endregion

#region Main Discovery Function

function Start-DiscoveryProcess {
    try {
        Write-LogMessage "=== Microsoft 365 and Azure Discovery Started ===" "INFO"
        Write-LogMessage "Working Folder: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        Write-LogMessage "Organization: $($Global:ScriptConfig.OpCo)" "INFO"
        Write-LogMessage "Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)" "INFO"
        Write-LogMessage "MFA Enabled: $($Global:ScriptConfig.UseMFA)" "INFO"
        
        # Ensure SharePoint Admin URL is set
        if ([string]::IsNullOrEmpty($Global:ScriptConfig.SharePointAdminSite)) {
            $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
            if ($domain -match '\.onmicrosoft\.com
        
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
        
        # Auto-adjust configuration based on available modules
        Write-LogMessage "Checking available modules and adjusting configuration..." "INFO"
        
        # Check and adjust each module
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'ExchangeOnlineManagement' -ModuleDescription 'Exchange and Teams discovery')) {
                Write-LogMessage "Exchange module not available - disabling Exchange and Teams discovery" "WARNING"
                $Global:ScriptConfig.SkipExchange = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires Exchange module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'PnP.PowerShell' -ModuleDescription 'SharePoint and Teams discovery')) {
                Write-LogMessage "PnP.PowerShell module not available - disabling SharePoint discovery" "WARNING"
                $Global:ScriptConfig.SkipSharePoint = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires SharePoint module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'MicrosoftPowerBIMgmt' -ModuleDescription 'PowerBI discovery')) {
                Write-LogMessage "PowerBI module not available - disabling PowerBI discovery" "WARNING"
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Accounts' -ModuleDescription 'Azure discovery')) {
                Write-LogMessage "Az module not available - disabling Azure discovery" "WARNING"
                $Global:ScriptConfig.SkipAzureAD = $true
            } else {
                # Also check for Az.Resources module which is needed for service principal operations
                if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Resources' -ModuleDescription 'Azure service principal discovery')) {
                    Write-LogMessage "Az.Resources module not available - some Azure discovery features may be limited" "WARNING"
                }
            }
        }
        
        # Check if any modules are still enabled
        if ($Global:ScriptConfig.SkipPowerBI -and $Global:ScriptConfig.SkipTeams -and 
            $Global:ScriptConfig.SkipSharePoint -and $Global:ScriptConfig.SkipExchange -and 
            $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "No discovery modules can run due to missing PowerShell modules" "ERROR"
            
            # Provide specific install instructions
            $missingModulesMessage = @"
No discovery modules can run because required PowerShell modules are not installed.

To fix this issue, open PowerShell 7+ as Administrator and run:

For Azure Discovery:
Install-Module Az -Force -AllowClobber
# Or just the required sub-modules:
Install-Module Az.Accounts, Az.Resources -Force -AllowClobber

For Exchange Discovery:  
Install-Module ExchangeOnlineManagement -Force -AllowClobber

For SharePoint Discovery:
Install-Module PnP.PowerShell -Force -AllowClobber

For PowerBI Discovery:
Install-Module MicrosoftPowerBIMgmt -Force -AllowClobber

After installing the required modules, restart PowerShell and run this tool again.
"@
            [System.Windows.Forms.MessageBox]::Show($missingModulesMessage, "Install Required Modules", "OK", "Information")
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
        # Get credentials if MFA is not enabled
        $Credentials = $null
        if (-not $Global:ScriptConfig.UseMFA) {
            Write-LogMessage "Getting credentials..." "INFO"
            
            # Show confirmation dialog with the account being used
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Is this correct? Click No to go back and fix the account details.
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Authentication Details", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled authentication - returning to main form" "INFO"
                return $false
            }
            
            $Credentials = Get-Credential -UserName $Global:ScriptConfig.GlobalAdminAccount -Message "Enter password for $($Global:ScriptConfig.GlobalAdminAccount)"
            if (-not $Credentials) {
                Write-LogMessage "Credentials are required when MFA is not enabled" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Credentials are required when MFA is not enabled", "Authentication Error", "OK", "Error")
                return $false
            }
            
            # Validate that the entered username matches our expected account
            if ($Credentials.UserName -ne $Global:ScriptConfig.GlobalAdminAccount) {
                $mismatchMessage = @"
Username mismatch detected!

Expected: $($Global:ScriptConfig.GlobalAdminAccount)
Entered: $($Credentials.UserName)

Please use the correct administrator account or update the configuration.
"@
                Write-LogMessage "Username mismatch detected" "ERROR"
                [System.Windows.Forms.MessageBox]::Show($mismatchMessage, "Authentication Error", "OK", "Error")
                return $false
            }
        } else {
            # For MFA, still show confirmation but explain browser authentication
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Since MFA is enabled, a browser window will open for EACH service requiring authentication:
• Exchange Online
• Azure 
• SharePoint Online  
• PowerBI

IMPORTANT:
✓ Complete authentication in EACH browser window that opens
✓ Do NOT close browser windows until authentication is complete
✓ Have your authenticator app or phone ready
✓ The script will wait for you to complete each authentication

If you cancel any authentication, that service will be skipped.

Is this correct and are you ready to proceed?
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm MFA Authentication", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled MFA authentication - returning to main form" "INFO"
                return $false
            }
            
            Write-LogMessage "MFA authentication confirmed - will use browser-based authentication" "INFO"
        }
        
        # Connect to services with retry logic
        Write-LogMessage "Attempting to connect to services..." "INFO"
        Update-ProgressDisplay "Authenticating with Microsoft 365 services..."
        
        if (-not (Connect-ServicesWithValidation -Credentials $Credentials -MFAEnabled $Global:ScriptConfig.UseMFA)) {
            Write-LogMessage "Service connection failed or was cancelled by user" "ERROR"
            
            # Show final termination message
            $terminationMessage = @"
Discovery process terminated due to authentication failure.

The script cannot continue without valid service connections.

You can:
• Check your account credentials and permissions
• Verify MFA settings
• Ensure all required PowerShell modules are installed
• Try running the script again

Click OK to return to the main screen.
"@
            [System.Windows.Forms.MessageBox]::Show($terminationMessage, "Discovery Terminated", "OK", "Information")
            return $false
        }
        
        Write-LogMessage "Service connections established successfully" "SUCCESS"
        Update-ProgressDisplay "Authentication successful - starting discovery..."
        
        # Show connection success summary
        $connectedServices = @()
        if (-not $Global:ScriptConfig.SkipExchange) { $connectedServices += "Exchange Online" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $connectedServices += "SharePoint Online" }
        if (-not $Global:ScriptConfig.SkipPowerBI) { $connectedServices += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $connectedServices += "Azure" }
        
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Connected services: $($connectedServices -join ', ')" "SUCCESS"
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
Organization: $($Global:ScriptConfig.OpCo)
Discovery Date: $(Get-Date)
Working Folder: $($Global:ScriptConfig.WorkingFolder)
Run Folder: $outputFolder
Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

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
        Disconnect-AllServices
        Write-LogMessage "=== Discovery process completed ===" "INFO"
    }
}

#endregion

#region GUI Functions

function Show-ProgressForm {
    $Global:ScriptConfig.ProgressForm = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.ProgressForm.Text = "Discovery in Progress"
    $Global:ScriptConfig.ProgressForm.Size = New-Object System.Drawing.Size(500, 250)
    $Global:ScriptConfig.ProgressForm.StartPosition = "CenterParent"
    $Global:ScriptConfig.ProgressForm.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.ProgressForm.MaximizeBox = $false
    $Global:ScriptConfig.ProgressForm.MinimizeBox = $false
    $Global:ScriptConfig.ProgressForm.BackColor = [System.Drawing.Color]::White
    $Global:ScriptConfig.ProgressForm.TopMost = $true
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery in Progress..."
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
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "❌ Cancel Discovery"
    $cancelButton.Size = New-Object System.Drawing.Size(130, 35)
    $cancelButton.Location = New-Object System.Drawing.Point(185, 170)
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
    $Global:ScriptConfig.Form.Text = "Microsoft 365 Discovery Tool"
    $Global:ScriptConfig.Form.Size = New-Object System.Drawing.Size(800, 700)
    $Global:ScriptConfig.Form.StartPosition = "CenterScreen"
    $Global:ScriptConfig.Form.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.Form.MaximizeBox = $false
    $Global:ScriptConfig.Form.MinimizeBox = $false
    $Global:ScriptConfig.Form.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)

    # Create a header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(780, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Header title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($titleLabel)

    # Header subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "All-in-one solution for comprehensive M365 environment discovery"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
    $subtitleLabel.Size = New-Object System.Drawing.Size(600, 25)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 45)
    $headerPanel.Controls.Add($subtitleLabel)

    $Global:ScriptConfig.Form.Controls.Add($headerPanel)

    # Main content panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Size = New-Object System.Drawing.Size(780, 570)
    $mainPanel.Location = New-Object System.Drawing.Point(10, 100)
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.BorderStyle = "FixedSingle"

    # Configuration Group Box
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = "🔧 Basic Configuration"
    $configGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $configGroupBox.Size = New-Object System.Drawing.Size(750, 150)
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

    # MFA Checkbox
    $mfaCheckBox = New-Object System.Windows.Forms.CheckBox
    $mfaCheckBox.Text = "Multi-Factor Authentication (MFA) Enabled"
    $mfaCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mfaCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $mfaCheckBox.Location = New-Object System.Drawing.Point(430, 105)
    $mfaCheckBox.Checked = $false
    $configGroupBox.Controls.Add($mfaCheckBox)

    $mainPanel.Controls.Add($configGroupBox)

    # Discovery Modules Group Box
    $modulesGroupBox = New-Object System.Windows.Forms.GroupBox
    $modulesGroupBox.Text = "📊 Discovery Modules (Uncheck to Skip)"
    $modulesGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $modulesGroupBox.Size = New-Object System.Drawing.Size(750, 200)
    $modulesGroupBox.Location = New-Object System.Drawing.Point(15, 175)
    $modulesGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Module checkboxes
    $powerBICheckBox = New-Object System.Windows.Forms.CheckBox
    $powerBICheckBox.Text = "📈 PowerBI Discovery (Workspaces, Reports, Datasets)"
    $powerBICheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $powerBICheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $powerBICheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $powerBICheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($powerBICheckBox)

    $teamsCheckBox = New-Object System.Windows.Forms.CheckBox
    $teamsCheckBox.Text = "👥 Teams & Groups Discovery"
    $teamsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $teamsCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $teamsCheckBox.Location = New-Object System.Drawing.Point(380, 30)
    $teamsCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($teamsCheckBox)

    $sharepointCheckBox = New-Object System.Windows.Forms.CheckBox
    $sharepointCheckBox.Text = "📁 SharePoint & OneDrive Discovery"
    $sharepointCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $sharepointCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $sharepointCheckBox.Location = New-Object System.Drawing.Point(20, 65)
    $sharepointCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($sharepointCheckBox)

    $exchangeCheckBox = New-Object System.Windows.Forms.CheckBox
    $exchangeCheckBox.Text = "📧 Exchange Discovery (Mailboxes, Archives)"
    $exchangeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exchangeCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $exchangeCheckBox.Location = New-Object System.Drawing.Point(380, 65)
    $exchangeCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($exchangeCheckBox)

    $azureADCheckBox = New-Object System.Windows.Forms.CheckBox
    $azureADCheckBox.Text = "🔑 Azure Discovery (Users, Groups, Apps)"
    $azureADCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $azureADCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $azureADCheckBox.Location = New-Object System.Drawing.Point(20, 100)
    $azureADCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($azureADCheckBox)

    # Progress information
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "💡 All modules are enabled by default. Uncheck modules you want to skip."
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $progressLabel.Size = New-Object System.Drawing.Size(700, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(20, 140)
    $progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $modulesGroupBox.Controls.Add($progressLabel)

    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Text = "⚠️ Required modules will be checked automatically before execution."
    $requirementsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $requirementsLabel.Size = New-Object System.Drawing.Size(700, 20)
    $requirementsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $requirementsLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
    $modulesGroupBox.Controls.Add($requirementsLabel)

    $mainPanel.Controls.Add($modulesGroupBox)

    # Prerequisites Group Box
    $prereqGroupBox = New-Object System.Windows.Forms.GroupBox
    $prereqGroupBox.Text = "📋 Prerequisites"
    $prereqGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $prereqGroupBox.Size = New-Object System.Drawing.Size(750, 100)
    $prereqGroupBox.Location = New-Object System.Drawing.Point(15, 385)
    $prereqGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    $prereqText = New-Object System.Windows.Forms.Label
    $prereqText.Text = "Required PowerShell modules (install with 'Install-Module ModuleName -Force'):`n• ExchangeOnlineManagement  • PnP.PowerShell  • MicrosoftPowerBIMgmt  • Az.Accounts, Az.Resources`n`nEnsure you have Global Administrator permissions and PowerShell 7+ for best compatibility."
    $prereqText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $prereqText.Size = New-Object System.Drawing.Size(720, 70)
    $prereqText.Location = New-Object System.Drawing.Point(20, 25)
    $prereqText.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $prereqGroupBox.Controls.Add($prereqText)

    $mainPanel.Controls.Add($prereqGroupBox)

    # Action buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "🚀 Start Discovery"
    $runButton.Size = New-Object System.Drawing.Size(150, 40)
    $runButton.Location = New-Object System.Drawing.Point(470, 510)
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($runButton)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = "✅ Check Prerequisites"
    $validateButton.Size = New-Object System.Drawing.Size(150, 40)
    $validateButton.Location = New-Object System.Drawing.Point(310, 510)
    $validateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $validateButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $validateButton.ForeColor = [System.Drawing.Color]::White
    $validateButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($validateButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "❌ Exit"
    $exitButton.Size = New-Object System.Drawing.Size(100, 40)
    $exitButton.Location = New-Object System.Drawing.Point(640, 510)
    $exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($exitButton)

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

    $validateButton.Add_Click({
        # Update configuration
        $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
        $Global:ScriptConfig.OpCo = $opcoTextBox.Text
        $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
        $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Use the helper function to set SharePoint Admin URL
        $Global:ScriptConfig.SharePointAdminSite = Get-SharePointAdminUrl -AdminEmail $Global:ScriptConfig.GlobalAdminAccount

        # Check what modules are needed and available
        $RequiredModules = @()
        $AvailableModules = @()
        $MissingModules = @()
        
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='ExchangeOnlineManagement'; Purpose='Exchange and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='PnP.PowerShell'; Purpose='SharePoint and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            $RequiredModules += @{Name='MicrosoftPowerBIMgmt'; Purpose='PowerBI discovery'}
        }
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            $RequiredModules += @{Name='Az.Accounts'; Purpose='Azure discovery'}
            $RequiredModules += @{Name='Az.Resources'; Purpose='Azure service principal discovery'}
        }

        if ($RequiredModules.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No modules are required since all discovery modules are disabled.", "Prerequisites Check", "OK", "Information")
            return
        }

        foreach ($module in $RequiredModules) {
            if (Get-Module -ListAvailable -Name $module.Name) {
                $AvailableModules += "$($module.Name) (for $($module.Purpose))"
            } else {
                $MissingModules += "$($module.Name) (for $($module.Purpose))"
            }
        }

        $resultMessage = ""
        
        if ($AvailableModules.Count -gt 0) {
            $resultMessage += "✅ Available Modules:`n$($AvailableModules -join "`n")`n`n"
        }
        
        if ($MissingModules.Count -gt 0) {
            $resultMessage += "❌ Missing Modules:`n$($MissingModules -join "`n")`n`n"
            $resultMessage += "To install missing modules, run PowerShell as Administrator:`n"
            foreach ($module in $RequiredModules) {
                if (-not (Get-Module -ListAvailable -Name $module.Name)) {
                    $resultMessage += "Install-Module $($module.Name) -Force`n"
                }
            }
            $resultMessage += "`nThe tool will automatically skip modules that aren't available."
        } else {
            $resultMessage += "🎉 All required modules are installed and ready!"
        }

        [System.Windows.Forms.MessageBox]::Show($resultMessage, "Prerequisites Check", "OK", "Information")
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
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Construct SharePoint Admin URL properly
        $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, try to extract the main part
            $tenantName = $domain.Split('.')[0]
        }
        $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        Write-LogMessage "SharePoint Admin URL set to: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
                $tenantName = $domain.Replace('.onmicrosoft.com', '')
            } else {
                $tenantName = $domain.Split('.')[0]
            }
            $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        }
        Write-LogMessage "SharePoint Admin Site: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
        # Create working directory if it doesn't exist
        if (-not (Test-Path $Global:ScriptConfig.WorkingFolder)) {
            New-Item -ItemType Directory -Path $Global:ScriptConfig.WorkingFolder -Force | Out-Null
            Write-LogMessage "Created working directory: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        }
        
        # Setup log file
        $Global:ScriptConfig.LogFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        # Auto-adjust configuration based on available modules
        Write-LogMessage "Checking available modules and adjusting configuration..." "INFO"
        
        # Check and adjust each module
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'ExchangeOnlineManagement' -ModuleDescription 'Exchange and Teams discovery')) {
                Write-LogMessage "Exchange module not available - disabling Exchange and Teams discovery" "WARNING"
                $Global:ScriptConfig.SkipExchange = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires Exchange module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'PnP.PowerShell' -ModuleDescription 'SharePoint and Teams discovery')) {
                Write-LogMessage "PnP.PowerShell module not available - disabling SharePoint discovery" "WARNING"
                $Global:ScriptConfig.SkipSharePoint = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires SharePoint module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'MicrosoftPowerBIMgmt' -ModuleDescription 'PowerBI discovery')) {
                Write-LogMessage "PowerBI module not available - disabling PowerBI discovery" "WARNING"
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Accounts' -ModuleDescription 'Azure discovery')) {
                Write-LogMessage "Az module not available - disabling Azure discovery" "WARNING"
                $Global:ScriptConfig.SkipAzureAD = $true
            } else {
                # Also check for Az.Resources module which is needed for service principal operations
                if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Resources' -ModuleDescription 'Azure service principal discovery')) {
                    Write-LogMessage "Az.Resources module not available - some Azure discovery features may be limited" "WARNING"
                }
            }
        }
        
        # Check if any modules are still enabled
        if ($Global:ScriptConfig.SkipPowerBI -and $Global:ScriptConfig.SkipTeams -and 
            $Global:ScriptConfig.SkipSharePoint -and $Global:ScriptConfig.SkipExchange -and 
            $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "No discovery modules can run due to missing PowerShell modules" "ERROR"
            
            # Provide specific install instructions
            $missingModulesMessage = @"
No discovery modules can run because required PowerShell modules are not installed.

To fix this issue, open PowerShell 7+ as Administrator and run:

For Azure Discovery:
Install-Module Az -Force -AllowClobber
# Or just the required sub-modules:
Install-Module Az.Accounts, Az.Resources -Force -AllowClobber

For Exchange Discovery:  
Install-Module ExchangeOnlineManagement -Force -AllowClobber

For SharePoint Discovery:
Install-Module PnP.PowerShell -Force -AllowClobber

For PowerBI Discovery:
Install-Module MicrosoftPowerBIMgmt -Force -AllowClobber

After installing the required modules, restart PowerShell and run this tool again.
"@
            [System.Windows.Forms.MessageBox]::Show($missingModulesMessage, "Install Required Modules", "OK", "Information")
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
        
        # Get credentials if MFA is not enabled
        $Credentials = $null
        if (-not $Global:ScriptConfig.UseMFA) {
            Write-LogMessage "Getting credentials..." "INFO"
            
            # Show confirmation dialog with the account being used
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Is this correct? Click No to go back and fix the account details.
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Authentication Details", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled authentication - returning to main form" "INFO"
                return $false
            }
            
            $Credentials = Get-Credential -UserName $Global:ScriptConfig.GlobalAdminAccount -Message "Enter password for $($Global:ScriptConfig.GlobalAdminAccount)"
            if (-not $Credentials) {
                Write-LogMessage "Credentials are required when MFA is not enabled" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Credentials are required when MFA is not enabled", "Authentication Error", "OK", "Error")
                return $false
            }
            
            # Validate that the entered username matches our expected account
            if ($Credentials.UserName -ne $Global:ScriptConfig.GlobalAdminAccount) {
                $mismatchMessage = @"
Username mismatch detected!

Expected: $($Global:ScriptConfig.GlobalAdminAccount)
Entered: $($Credentials.UserName)

Please use the correct administrator account or update the configuration.
"@
                Write-LogMessage "Username mismatch detected" "ERROR"
                [System.Windows.Forms.MessageBox]::Show($mismatchMessage, "Authentication Error", "OK", "Error")
                return $false
            }
        } else {
            # For MFA, still show confirmation but explain browser authentication
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Since MFA is enabled, you will be prompted to sign in through your browser 
for each service (Exchange, SharePoint, PowerBI, Azure AD).

Make sure you:
• Have access to your authenticator app or phone
• Are ready to complete MFA challenges
• Don't close browser windows during authentication

Is this correct and are you ready to proceed?
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm MFA Authentication", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled MFA authentication - returning to main form" "INFO"
                return $false
            }
            
            Write-LogMessage "MFA authentication confirmed - will use browser-based authentication" "INFO"
        }
        
        # Connect to services with retry logic
        Write-LogMessage "Attempting to connect to services..." "INFO"
        Update-ProgressDisplay "Authenticating with Microsoft 365 services..."
        
        if (-not (Connect-ServicesWithValidation -Credentials $Credentials -MFAEnabled $Global:ScriptConfig.UseMFA)) {
            Write-LogMessage "Service connection failed or was cancelled by user" "ERROR"
            
            # Show final termination message
            $terminationMessage = @"
Discovery process terminated due to authentication failure.

The script cannot continue without valid service connections.

You can:
• Check your account credentials and permissions
• Verify MFA settings
• Ensure all required PowerShell modules are installed
• Try running the script again

Click OK to return to the main screen.
"@
            [System.Windows.Forms.MessageBox]::Show($terminationMessage, "Discovery Terminated", "OK", "Information")
            return $false
        }
        
        Write-LogMessage "Service connections established successfully" "SUCCESS"
        Update-ProgressDisplay "Authentication successful - starting discovery..."
        
        # Show connection success summary
        $connectedServices = @()
        if (-not $Global:ScriptConfig.SkipExchange) { $connectedServices += "Exchange Online" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $connectedServices += "SharePoint Online" }
        if (-not $Global:ScriptConfig.SkipPowerBI) { $connectedServices += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $connectedServices += "Azure" }
        
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Connected services: $($connectedServices -join ', ')" "SUCCESS"
        }
        
        # Run discovery modules
        Invoke-PowerBIDiscovery
        Invoke-TeamsDiscovery
        Invoke-SharePointDiscovery
        Invoke-ExchangeDiscovery
        Invoke-AzureADDiscovery
        
        # Generate summary report
        Write-LogMessage "Generating summary report..." "INFO"
        $SummaryFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $Summary = @"
Microsoft 365 and Azure Discovery Summary
=========================================
Organization: $($Global:ScriptConfig.OpCo)
Discovery Date: $(Get-Date)
Working Folder: $($Global:ScriptConfig.WorkingFolder)
Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Discovery Modules Executed:
- PowerBI Discovery: $(if (-not $Global:ScriptConfig.SkipPowerBI) { "✓ Executed" } else { "✗ Skipped" })
- Teams & Groups Discovery: $(if (-not $Global:ScriptConfig.SkipTeams) { "✓ Executed" } else { "✗ Skipped" })
- SharePoint & OneDrive Discovery: $(if (-not $Global:ScriptConfig.SkipSharePoint) { "✓ Executed" } else { "✗ Skipped" })
- Exchange Discovery: $(if (-not $Global:ScriptConfig.SkipExchange) { "✓ Executed" } else { "✗ Skipped" })
- Azure Discovery: $(if (-not $Global:ScriptConfig.SkipAzureAD) { "✓ Executed" } else { "✗ Skipped" })

Log File: $($Global:ScriptConfig.LogFile)
Summary File: $SummaryFile

Generated Files:
$(Get-ChildItem $Global:ScriptConfig.WorkingFolder -Filter "*$($Global:ScriptConfig.OpCo)*" | ForEach-Object { "- $($_.Name)" })
"@
        $Summary | Out-File $SummaryFile
        
        Write-LogMessage "=== Discovery completed successfully ===" "SUCCESS"
        Write-LogMessage "All reports saved to: $($Global:ScriptConfig.WorkingFolder)" "SUCCESS"
        Write-LogMessage "Summary report: $SummaryFile" "SUCCESS"
        
        return $true
    }
    catch {
        Write-LogMessage "Discovery failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Clean up connections
        Disconnect-AllServices
        Write-LogMessage "=== Discovery process completed ===" "INFO"
    }
}

#endregion

#region GUI Functions

function Show-ProgressForm {
    $Global:ScriptConfig.ProgressForm = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.ProgressForm.Text = "Discovery in Progress"
    $Global:ScriptConfig.ProgressForm.Size = New-Object System.Drawing.Size(500, 250)
    $Global:ScriptConfig.ProgressForm.StartPosition = "CenterParent"
    $Global:ScriptConfig.ProgressForm.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.ProgressForm.MaximizeBox = $false
    $Global:ScriptConfig.ProgressForm.MinimizeBox = $false
    $Global:ScriptConfig.ProgressForm.BackColor = [System.Drawing.Color]::White
    $Global:ScriptConfig.ProgressForm.TopMost = $true
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery in Progress..."
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
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "❌ Cancel Discovery"
    $cancelButton.Size = New-Object System.Drawing.Size(130, 35)
    $cancelButton.Location = New-Object System.Drawing.Point(185, 170)
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
            $message = "Discovery completed successfully!`n`nReports saved to: $($Global:ScriptConfig.WorkingFolder)`n`nWould you like to open the folder?"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Discovery Complete", "YesNo", "Information")
            
            if ($result -eq "Yes") {
                Start-Process "explorer.exe" $Global:ScriptConfig.WorkingFolder
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Discovery failed. Please check the log file for details.", "Discovery Failed", "OK", "Error")
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
    $Global:ScriptConfig.Form.Text = "Microsoft 365 Discovery Tool"
    $Global:ScriptConfig.Form.Size = New-Object System.Drawing.Size(800, 700)
    $Global:ScriptConfig.Form.StartPosition = "CenterScreen"
    $Global:ScriptConfig.Form.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.Form.MaximizeBox = $false
    $Global:ScriptConfig.Form.MinimizeBox = $false
    $Global:ScriptConfig.Form.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)

    # Create a header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(780, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Header title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($titleLabel)

    # Header subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "All-in-one solution for comprehensive M365 environment discovery"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
    $subtitleLabel.Size = New-Object System.Drawing.Size(600, 25)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 45)
    $headerPanel.Controls.Add($subtitleLabel)

    $Global:ScriptConfig.Form.Controls.Add($headerPanel)

    # Main content panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Size = New-Object System.Drawing.Size(780, 570)
    $mainPanel.Location = New-Object System.Drawing.Point(10, 100)
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.BorderStyle = "FixedSingle"

    # Configuration Group Box
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = "🔧 Basic Configuration"
    $configGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $configGroupBox.Size = New-Object System.Drawing.Size(750, 150)
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

    # MFA Checkbox
    $mfaCheckBox = New-Object System.Windows.Forms.CheckBox
    $mfaCheckBox.Text = "Multi-Factor Authentication (MFA) Enabled"
    $mfaCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mfaCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $mfaCheckBox.Location = New-Object System.Drawing.Point(430, 105)
    $mfaCheckBox.Checked = $false
    $configGroupBox.Controls.Add($mfaCheckBox)

    $mainPanel.Controls.Add($configGroupBox)

    # Discovery Modules Group Box
    $modulesGroupBox = New-Object System.Windows.Forms.GroupBox
    $modulesGroupBox.Text = "📊 Discovery Modules (Uncheck to Skip)"
    $modulesGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $modulesGroupBox.Size = New-Object System.Drawing.Size(750, 200)
    $modulesGroupBox.Location = New-Object System.Drawing.Point(15, 175)
    $modulesGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Module checkboxes
    $powerBICheckBox = New-Object System.Windows.Forms.CheckBox
    $powerBICheckBox.Text = "📈 PowerBI Discovery (Workspaces, Reports, Datasets)"
    $powerBICheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $powerBICheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $powerBICheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $powerBICheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($powerBICheckBox)

    $teamsCheckBox = New-Object System.Windows.Forms.CheckBox
    $teamsCheckBox.Text = "👥 Teams & Groups Discovery"
    $teamsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $teamsCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $teamsCheckBox.Location = New-Object System.Drawing.Point(380, 30)
    $teamsCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($teamsCheckBox)

    $sharepointCheckBox = New-Object System.Windows.Forms.CheckBox
    $sharepointCheckBox.Text = "📁 SharePoint & OneDrive Discovery"
    $sharepointCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $sharepointCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $sharepointCheckBox.Location = New-Object System.Drawing.Point(20, 65)
    $sharepointCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($sharepointCheckBox)

    $exchangeCheckBox = New-Object System.Windows.Forms.CheckBox
    $exchangeCheckBox.Text = "📧 Exchange Discovery (Mailboxes, Archives)"
    $exchangeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exchangeCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $exchangeCheckBox.Location = New-Object System.Drawing.Point(380, 65)
    $exchangeCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($exchangeCheckBox)

    $azureADCheckBox = New-Object System.Windows.Forms.CheckBox
    $azureADCheckBox.Text = "🔑 Azure Discovery (Users, Groups, Apps)"
    $azureADCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $azureADCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $azureADCheckBox.Location = New-Object System.Drawing.Point(20, 100)
    $azureADCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($azureADCheckBox)

    # Progress information
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "💡 All modules are enabled by default. Uncheck modules you want to skip."
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $progressLabel.Size = New-Object System.Drawing.Size(700, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(20, 140)
    $progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $modulesGroupBox.Controls.Add($progressLabel)

    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Text = "⚠️ Required modules will be checked automatically before execution."
    $requirementsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $requirementsLabel.Size = New-Object System.Drawing.Size(700, 20)
    $requirementsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $requirementsLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
    $modulesGroupBox.Controls.Add($requirementsLabel)

    $mainPanel.Controls.Add($modulesGroupBox)

    # Prerequisites Group Box
    $prereqGroupBox = New-Object System.Windows.Forms.GroupBox
    $prereqGroupBox.Text = "📋 Prerequisites"
    $prereqGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $prereqGroupBox.Size = New-Object System.Drawing.Size(750, 100)
    $prereqGroupBox.Location = New-Object System.Drawing.Point(15, 385)
    $prereqGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    $prereqText = New-Object System.Windows.Forms.Label
    $prereqText.Text = "Required PowerShell modules (install with 'Install-Module ModuleName -Force'):`n• ExchangeOnlineManagement  • PnP.PowerShell  • MicrosoftPowerBIMgmt  • Az.Accounts, Az.Resources`n`nEnsure you have Global Administrator permissions and PowerShell 7+ for best compatibility."
    $prereqText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $prereqText.Size = New-Object System.Drawing.Size(720, 70)
    $prereqText.Location = New-Object System.Drawing.Point(20, 25)
    $prereqText.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $prereqGroupBox.Controls.Add($prereqText)

    $mainPanel.Controls.Add($prereqGroupBox)

    # Action buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "🚀 Start Discovery"
    $runButton.Size = New-Object System.Drawing.Size(150, 40)
    $runButton.Location = New-Object System.Drawing.Point(470, 510)
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($runButton)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = "✅ Check Prerequisites"
    $validateButton.Size = New-Object System.Drawing.Size(150, 40)
    $validateButton.Location = New-Object System.Drawing.Point(310, 510)
    $validateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $validateButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $validateButton.ForeColor = [System.Drawing.Color]::White
    $validateButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($validateButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "❌ Exit"
    $exitButton.Size = New-Object System.Drawing.Size(100, 40)
    $exitButton.Location = New-Object System.Drawing.Point(640, 510)
    $exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($exitButton)

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

    $validateButton.Add_Click({
        # Update configuration
        $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
        $Global:ScriptConfig.OpCo = $opcoTextBox.Text
        $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
        $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked

        # Check what modules are needed and available
        $RequiredModules = @()
        $AvailableModules = @()
        $MissingModules = @()
        
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='ExchangeOnlineManagement'; Purpose='Exchange and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='PnP.PowerShell'; Purpose='SharePoint and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            $RequiredModules += @{Name='MicrosoftPowerBIMgmt'; Purpose='PowerBI discovery'}
        }
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            $RequiredModules += @{Name='Az.Accounts'; Purpose='Azure discovery'}
            $RequiredModules += @{Name='Az.Resources'; Purpose='Azure service principal discovery'}
        }

        if ($RequiredModules.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No modules are required since all discovery modules are disabled.", "Prerequisites Check", "OK", "Information")
            return
        }

        foreach ($module in $RequiredModules) {
            if (Get-Module -ListAvailable -Name $module.Name) {
                $AvailableModules += "$($module.Name) (for $($module.Purpose))"
            } else {
                $MissingModules += "$($module.Name) (for $($module.Purpose))"
            }
        }

        $resultMessage = ""
        
        if ($AvailableModules.Count -gt 0) {
            $resultMessage += "✅ Available Modules:`n$($AvailableModules -join "`n")`n`n"
        }
        
        if ($MissingModules.Count -gt 0) {
            $resultMessage += "❌ Missing Modules:`n$($MissingModules -join "`n")`n`n"
            $resultMessage += "To install missing modules, run PowerShell as Administrator:`n"
            foreach ($module in $RequiredModules) {
                if (-not (Get-Module -ListAvailable -Name $module.Name)) {
                    $resultMessage += "Install-Module $($module.Name) -Force`n"
                }
            }
            $resultMessage += "`nThe tool will automatically skip modules that aren't available."
        } else {
            $resultMessage += "🎉 All required modules are installed and ready!"
        }

        [System.Windows.Forms.MessageBox]::Show($resultMessage, "Prerequisites Check", "OK", "Information")
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
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Construct SharePoint Admin URL properly
        $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, try to extract the main part
            $tenantName = $domain.Split('.')[0]
        }
        $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        Write-LogMessage "SharePoint Admin URL set to: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, we need to make an educated guess
            # In production, you might want to prompt the user for the tenant name
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
            if ($statusLabel) {
                $statusLabel.Text = $Message
                $Global:ScriptConfig.ProgressForm.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            # Ignore GUI update errors
        }
    }
    
    # Also update console
    if ($Message -match "Connecting to|Successfully connected|Failed to connect") {
        Write-Host $Message -ForegroundColor Cyan
    }
}

function Test-ModuleAndAdjustConfig {
    param([string]$ModuleName, [string]$ModuleDescription)
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-LogMessage "Module $ModuleName not found - $ModuleDescription will be skipped" "WARNING"
        return $false
    }
    
    # Special check for ExchangeOnlineManagement module version
    if ($ModuleName -eq 'ExchangeOnlineManagement') {
        $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($module.Version -lt '3.0.0') {
            Write-LogMessage "ExchangeOnlineManagement module version $($module.Version) is outdated. Version 3.0.0 or higher is required for browser authentication." "WARNING"
            Write-LogMessage "Please update the module using: Update-Module ExchangeOnlineManagement -Force" "WARNING"
            return $false
        }
    }
    
    # Try to import the module if it's available but not loaded
    if (-not (Get-Module -Name $ModuleName)) {
        try {
            Write-LogMessage "Importing module $ModuleName..." "INFO"
            Import-Module $ModuleName -Force -ErrorAction Stop
            Write-LogMessage "Successfully imported $ModuleName" "SUCCESS"
        } catch {
            Write-LogMessage "Failed to import $ModuleName`: $($_.Exception.Message)" "ERROR"
            return $false
        }
    }
    
    return $true
}

function Connect-ServicesWithValidation {
    param(
        [System.Management.Automation.PSCredential]$Credentials,
        [bool]$MFAEnabled
    )
    
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        Write-LogMessage "Connecting to services (Attempt $($retryCount + 1) of $maxRetries)..." "INFO"
        
        $connectedServices = @()
        $authenticationFailed = $false
        
        # Try Exchange Online
        if (-not $Global:ScriptConfig.SkipExchange) {
            Write-LogMessage "Connecting to Exchange Online..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, use UserPrincipalName to pre-fill the account
                    Connect-ExchangeOnline -UserPrincipalName $Global:ScriptConfig.GlobalAdminAccount -ShowBanner:$false
                } else {
                    Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false
                }
                
                $testResult = Get-OrganizationConfig -ErrorAction Stop
                Write-LogMessage "Exchange Online connected successfully" "SUCCESS"
                $connectedServices += "Exchange Online"
            }
            catch {
                Write-LogMessage "Exchange connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|0x80070520|cancelled|OAuth") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipExchange = $true
            }
        }
        
        # Try Azure AD (using Az module)
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "Connecting to Azure (Az module)..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, browser will open automatically
                    Connect-AzAccount
                } else {
                    Connect-AzAccount -Credential $Credentials
                }
                
                $testResult = Get-AzContext -ErrorAction Stop
                Write-LogMessage "Azure connected successfully" "SUCCESS"
                $connectedServices += "Azure"
            }
            catch {
                Write-LogMessage "Azure connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|cancelled|OAuth") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipAzureAD = $true
            }
        }
        
        # Try SharePoint Online (using PnP.PowerShell)
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipSharePoint) {
            Write-LogMessage "Connecting to SharePoint Online (PnP)..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, use interactive login
                    Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Interactive
                } else {
                    # Create PSCredential object for PnP
                    Connect-PnPOnline -Url $Global:ScriptConfig.SharePointAdminSite -Credential $Credentials
                }
                
                $testResult = Get-PnPContext -ErrorAction Stop
                Write-LogMessage "SharePoint Online connected successfully" "SUCCESS"
                $connectedServices += "SharePoint Online"
            }
            catch {
                Write-LogMessage "SharePoint connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|OAuth|cancelled") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipSharePoint = $true
            }
        }
        
        # Try PowerBI
        if (-not $authenticationFailed -and -not $Global:ScriptConfig.SkipPowerBI) {
            Write-LogMessage "Connecting to PowerBI..." "INFO"
            try {
                if ($MFAEnabled) {
                    # For MFA, don't pass any parameters to trigger browser auth
                    Connect-PowerBIServiceAccount
                } else {
                    Connect-PowerBIServiceAccount -Credential $Credentials
                }
                
                Write-LogMessage "PowerBI connected successfully" "SUCCESS"
                $connectedServices += "PowerBI"
            }
            catch {
                Write-LogMessage "PowerBI connection failed: $($_.Exception.Message)" "ERROR"
                if ($_.Exception.Message -match "authentication|credential|login|unauthorized|OAuth|cancelled") {
                    $authenticationFailed = $true
                }
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        # Check results
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Successfully connected to: $($connectedServices -join ', ')" "SUCCESS"
            return $true
        }
        
        # Handle retry
        if ($authenticationFailed) {
            $retryCount++
            
            if ($retryCount -lt $maxRetries) {
                $retryMessage = "Authentication failed. Attempt $retryCount of $maxRetries.`n`nDo you want to try again?"
                $retryResult = [System.Windows.Forms.MessageBox]::Show($retryMessage, "Authentication Failed", "YesNo", "Question")
                
                if ($retryResult -eq "No") {
                    return $false
                }
                
                Write-LogMessage "Retrying authentication..." "INFO"
                try {
                    Disconnect-AllServices
                } catch {
                    # Ignore disconnect errors
                }
                Start-Sleep -Seconds 2
                Continue
            } else {
                [System.Windows.Forms.MessageBox]::Show("Maximum authentication attempts reached.", "Authentication Failed", "OK", "Error")
                return $false
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No services could be connected.", "Connection Failed", "OK", "Error")
            return $false
        }
    }
    
    return $false
}

function Disconnect-AllServices {
    Write-LogMessage "Disconnecting from all services..." "INFO"
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
        }
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) { 
            try { Disconnect-PnPOnline } catch { }
        }
        if (Get-Command Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-PowerBIServiceAccount } catch { }
        }
        if (Get-Command Disconnect-AzAccount -ErrorAction SilentlyContinue) { 
            try { Disconnect-AzAccount -Confirm:$false } catch { }
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
        [string]$Description
    )
    
    try {
        Write-LogMessage "Exporting $Description to $FilePath..." "INFO"
        if ($Data -and $Data.Count -gt 0) {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            Write-LogMessage "Successfully exported $($Data.Count) records for $Description" "SUCCESS"
        } else {
            Write-LogMessage "No data found for $Description" "WARNING"
            "No data found" | Out-File $FilePath
        }
    }
    catch {
        Write-LogMessage "Failed to export $Description`: $($_.Exception.Message)" "ERROR"
        throw
    }
}

#endregion

#region Discovery Functions

function Invoke-PowerBIDiscovery {
    if ($Global:ScriptConfig.SkipPowerBI) {
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
            return
        }
        
        # PowerBI Activity Events
        Write-LogMessage "Getting PowerBI activity events..." "INFO"
        try {
            $OutputReport = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI.txt"
            Start-Transcript $OutputReport
            
            $StartDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddT00:00:00")
            $EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            
            Get-PowerBIActivityEvent -StartDateTime $StartDate -EndDateTime $EndDate | 
                Out-File (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_UsageActivity.txt")
            
            Stop-Transcript
        } catch {
            Write-LogMessage "Error getting PowerBI activity events: $($_.Exception.Message)" "WARNING"
            if (Get-Command Stop-Transcript -ErrorAction SilentlyContinue) { Stop-Transcript }
        }
        
        # PowerBI Workspaces
        Write-LogMessage "Getting PowerBI workspaces..." "INFO"
        try {
            $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
            Export-DataSafely -Data $Workspaces -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_PowerBI_Workspaces.csv") -Description "PowerBI Workspaces"
        } catch {
            Write-LogMessage "Error getting PowerBI workspaces: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "PowerBI discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "PowerBI discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-TeamsDiscovery {
    if ($Global:ScriptConfig.SkipTeams) {
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
            return
        }
        
        Write-LogMessage "Getting all Teams and M365 Groups..." "INFO"
        $AllTeamsAndGroups = Get-UnifiedGroup -ResultSize Unlimited
        $ArrayToExport = @()
        $Counter = 0
        $Total = $AllTeamsAndGroups.Count
        
        Write-LogMessage "Found $Total Teams and M365 Groups to process..." "INFO"
        
        foreach ($Team in $AllTeamsAndGroups) {
            $Counter++
            Write-LogMessage "Processing $($Team.PrimarySmtpAddress) ($Counter of $Total)" "INFO"
            
            try {
                $GroupType = if ($Team.ResourceProvisioningOptions -contains "Team") { "Team" } else { "M365 Group" }
                
                # Get mailbox statistics with error handling
                $MailboxStats = $null
                try {
                    $MailboxStats = Get-EXOMailboxStatistics $Team.ExchangeGuid -ErrorAction Stop
                } catch {
                    Write-LogMessage "Warning: Could not get mailbox stats for $($Team.PrimarySmtpAddress)" "WARNING"
                }
                
                # Get SharePoint site with error handling
                $SharePointSite = $null
                try {
                    if ($Team.SharePointSiteUrl -and !$Global:ScriptConfig.SkipSharePoint) {
                        # Use PnP to get the site
                        $SharePointSite = Get-PnPTenantSite -Identity $Team.SharePointSiteUrl -ErrorAction Stop
                    }
                } catch {
                    Write-LogMessage "Warning: Could not get SharePoint site for $($Team.PrimarySmtpAddress)" "WARNING"
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
                    SP_LockState = if ($SharePointSite) { $SharePointSite.LockState } else { "N/A" }
                    SP_StorageUsageCurrent = if ($SharePointSite) { $SharePointSite.StorageUsageCurrent } else { "N/A" }
                    SP_WebsCount = if ($SharePointSite) { $SharePointSite.WebsCount } else { "N/A" }
                    SP_Template = if ($SharePointSite) { $SharePointSite.Template } else { "N/A" }
                }
                $ArrayToExport += $TeamDetails
            }
            catch {
                Write-LogMessage "Error processing team $($Team.PrimarySmtpAddress): $($_.Exception.Message)" "WARNING"
            }
        }
        
        Export-DataSafely -Data $ArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_TeamsGroups.csv") -Description "Teams and M365 Groups"
        Write-LogMessage "Teams and Groups discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Teams discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-SharePointDiscovery {
    if ($Global:ScriptConfig.SkipSharePoint) {
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
            return
        }
        
        Write-LogMessage "Getting all SharePoint sites..." "INFO"
        # Get all sites using PnP cmdlets
        $SharePointSites = Get-PnPTenantSite -IncludeOneDriveSites
        
        Write-LogMessage "Processing OneDrive sites..." "INFO"
        $OneDriveSites = $SharePointSites | Where-Object { $_.Template -match "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner, ConditionalAccessPolicy
        
        Export-DataSafely -Data $OneDriveSites -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_OneDrive.csv") -Description "OneDrive Sites"
        
        Write-LogMessage "Processing SharePoint sites..." "INFO"
        $SharePointSitesFiltered = $SharePointSites | Where-Object { $_.Template -notmatch "SPSPERS" } | 
            Sort-Object Template, Url | 
            Select-Object Url, Template, Title, StorageUsageCurrent, LastContentModifiedDate, Status, LockState, WebsCount, LocaleId, Owner, ConditionalAccessPolicy
        
        Export-DataSafely -Data $SharePointSitesFiltered -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_SharePoint.csv") -Description "SharePoint Sites"
        
        Write-LogMessage "SharePoint and OneDrive discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "SharePoint discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-ExchangeDiscovery {
    if ($Global:ScriptConfig.SkipExchange) {
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
            return
        }
        
        Write-LogMessage "Getting all mailboxes..." "INFO"
        $AllMailboxes = Get-EXOMailbox -IncludeInactiveMailbox -ResultSize Unlimited -PropertySets All
        $MailboxArrayToExport = @()
        $Counter = 0
        $Total = $AllMailboxes.Count
        
        Write-LogMessage "Found $Total mailboxes to process..." "INFO"
        
        foreach ($Mailbox in $AllMailboxes) {
            $Counter++
            Write-LogMessage "Processing mailbox: $($Mailbox.UserPrincipalName) ($Counter of $Total)" "INFO"
            
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
                    Write-LogMessage "Warning: Could not get mailbox stats for $($Mailbox.UserPrincipalName)" "WARNING"
                }
                
                # Handle archive mailbox
                $ArchiveDetails = @{
                    HasArchive = "No"
                    DisplayName = ""
                    MailboxGuid = ""
                    ItemCount = ""
                    TotalItemSize = ""
                    DeletedItemCount = ""
                    TotalDeletedItemSize = ""
                }
                
                if ($Mailbox.ArchiveName) {
                    try {
                        $MailboxArchiveStats = if ($Mailbox.IsInactiveMailbox) {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive -IncludeSoftDeletedRecipients
                        } else {
                            Get-EXOMailboxStatistics -ExchangeGuid $Mailbox.ExchangeGuid -Archive
                        }
                        
                        $ArchiveDetails = @{
                            HasArchive = "Yes"
                            DisplayName = $MailboxArchiveStats.DisplayName
                            MailboxGuid = $MailboxArchiveStats.MailboxGuid
                            ItemCount = $MailboxArchiveStats.ItemCount
                            TotalItemSize = if ($MailboxArchiveStats.TotalItemSize) {
                                try {
                                    [math]::Round(($MailboxArchiveStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                                } catch { "N/A" }
                            } else { "N/A" }
                            DeletedItemCount = $MailboxArchiveStats.DeletedItemCount
                            TotalDeletedItemSize = if ($MailboxArchiveStats.TotalDeletedItemSize) {
                                try {
                                    [math]::Round(($MailboxArchiveStats.TotalDeletedItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                                } catch { "N/A" }
                            } else { "N/A" }
                        }
                    }
                    catch {
                        Write-LogMessage "Warning: Could not get archive stats for $($Mailbox.UserPrincipalName)" "WARNING"
                    }
                }
                
                $MailboxDetails = [PSCustomObject]@{
                    UserPrincipalName = $Mailbox.UserPrincipalName
                    DisplayName = $Mailbox.DisplayName
                    PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                    ExchangeGuid = $Mailbox.ExchangeGuid
                    RecipientType = $Mailbox.RecipientType
                    RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                    HasArchive = $ArchiveDetails.HasArchive
                    AccountDisabled = $Mailbox.AccountDisabled
                    LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                    RetentionHoldEnabled = $Mailbox.RetentionHoldEnabled
                    IsMailboxEnabled = $Mailbox.IsMailboxEnabled
                    IsInactiveMailbox = $Mailbox.IsInactiveMailbox
                    WhenSoftDeleted = $Mailbox.WhenSoftDeleted
                    ArchiveStatus = $Mailbox.ArchiveStatus
                    WhenMailboxCreated = $Mailbox.WhenMailboxCreated
                    RetentionPolicy = $Mailbox.RetentionPolicy
                    MB_DisplayName = if ($MailboxStats) { $MailboxStats.DisplayName } else { "N/A" }
                    MB_MailboxGuid = if ($MailboxStats) { $MailboxStats.MailboxGuid } else { "N/A" }
                    MB_ItemCount = if ($MailboxStats) { $MailboxStats.ItemCount } else { "N/A" }
                    MB_TotalItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalItemSize) {
                        try {
                            [math]::Round(($MailboxStats.TotalItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                        } catch { "N/A" }
                    } else { "N/A" }
                    Archive_DisplayName = $ArchiveDetails.DisplayName
                    Archive_MailboxGuid = $ArchiveDetails.MailboxGuid
                    Archive_ItemCount = $ArchiveDetails.ItemCount
                    Archive_TotalItemSizeMB = $ArchiveDetails.TotalItemSize
                    MaxSendSize = if ($Mailbox.MaxSendSize) { $Mailbox.MaxSendSize.Split(" (")[0] } else { "N/A" }
                    MaxReceiveSize = if ($Mailbox.MaxReceiveSize) { $Mailbox.MaxReceiveSize.Split(" (")[0] } else { "N/A" }
                    MB_DeletedItemCount = if ($MailboxStats) { $MailboxStats.DeletedItemCount } else { "N/A" }
                    MB_TotalDeletedItemSizeMB = if ($MailboxStats -and $MailboxStats.TotalDeletedItemSize) {
                        try {
                            [math]::Round(($MailboxStats.TotalDeletedItemSize.ToString().Split("(")[1]).Split(" bytes")[0].Replace(",", "") / 1MB, 2)
                        } catch { "N/A" }
                    } else { "N/A" }
                    Archive_DeletedItemCount = $ArchiveDetails.DeletedItemCount
                    Archive_TotalDeletedItemSizeMB = $ArchiveDetails.TotalDeletedItemSize
                }
                $MailboxArrayToExport += $MailboxDetails
            }
            catch {
                Write-LogMessage "Error processing mailbox $($Mailbox.UserPrincipalName): $($_.Exception.Message)" "WARNING"
            }
        }
        
        Export-DataSafely -Data $MailboxArrayToExport -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Mailboxes.csv") -Description "Exchange Mailboxes"
        Write-LogMessage "Exchange discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Exchange discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-AzureADDiscovery {
    if ($Global:ScriptConfig.SkipAzureAD) {
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
            return
        }
        
        # Enterprise Applications (Service Principals)
        Write-LogMessage "Getting enterprise applications..." "INFO"
        try {
            # Get service principals (enterprise apps)
            $EnterpriseApps = Get-AzADServicePrincipal | 
                Where-Object { $_.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp" -or $_.ServicePrincipalType -eq "Application" } | 
                Sort-Object DisplayName
            
            Export-DataSafely -Data ($EnterpriseApps | Select-Object Id, AccountEnabled, DisplayName, AppId, ServicePrincipalType, Tags) -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_EnterpriseApplications.csv") -Description "Enterprise Applications"
        } catch {
            Write-LogMessage "Error getting enterprise applications: $($_.Exception.Message)" "WARNING"
        }
        
        # Azure AD Users
        Write-LogMessage "Getting Azure AD users..." "INFO"
        try {
            $AzureADUsers = Get-AzADUser | 
                Select-Object UserPrincipalName, DisplayName, Id, UserType, AccountEnabled, Mail
            Export-DataSafely -Data $AzureADUsers -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Users.csv") -Description "Azure AD Users"
        } catch {
            Write-LogMessage "Error getting Azure AD users: $($_.Exception.Message)" "ERROR"
        }
        
        # Azure AD Groups
        Write-LogMessage "Getting Azure AD groups..." "INFO"
        try {
            $AzureADGroups = Get-AzADGroup | 
                Select-Object DisplayName, Id, MailEnabled, SecurityEnabled, Mail, Description
            Export-DataSafely -Data $AzureADGroups -FilePath (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_AzureAD_Groups.csv") -Description "Azure AD Groups"
        } catch {
            Write-LogMessage "Error getting Azure AD groups: $($_.Exception.Message)" "WARNING"
        }
        
        # DevOps Organizations discovery URL
        try {
            $DevOpsUrl = "https://app.vsaex.visualstudio.com/_apis/EnterpriseCatalog/Organizations?tenantId=$TenantId"
            Write-LogMessage "DevOps Organizations discovery URL: $DevOpsUrl" "INFO"
            
            # Save DevOps URL to file
            $DevOpsUrl | Out-File (Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_DevOps_Discovery_URL.txt")
        } catch {
            Write-LogMessage "Error generating DevOps URL: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "Azure discovery completed successfully" "SUCCESS"
    }
    catch {
        Write-LogMessage "Azure discovery failed: $($_.Exception.Message)" "ERROR"
    }
}

#endregion

#region Main Discovery Function

function Start-DiscoveryProcess {
    try {
        Write-LogMessage "=== Microsoft 365 and Azure Discovery Started ===" "INFO"
        Write-LogMessage "Working Folder: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        Write-LogMessage "Organization: $($Global:ScriptConfig.OpCo)" "INFO"
        Write-LogMessage "Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)" "INFO"
        Write-LogMessage "MFA Enabled: $($Global:ScriptConfig.UseMFA)" "INFO"
        
        # Ensure SharePoint Admin URL is set
        if ([string]::IsNullOrEmpty($Global:ScriptConfig.SharePointAdminSite)) {
            $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
            if ($domain -match '\.onmicrosoft\.com
        
        # Create working directory if it doesn't exist
        if (-not (Test-Path $Global:ScriptConfig.WorkingFolder)) {
            New-Item -ItemType Directory -Path $Global:ScriptConfig.WorkingFolder -Force | Out-Null
            Write-LogMessage "Created working directory: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        }
        
        # Setup log file
        $Global:ScriptConfig.LogFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        # Auto-adjust configuration based on available modules
        Write-LogMessage "Checking available modules and adjusting configuration..." "INFO"
        
        # Check and adjust each module
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'ExchangeOnlineManagement' -ModuleDescription 'Exchange and Teams discovery')) {
                Write-LogMessage "Exchange module not available - disabling Exchange and Teams discovery" "WARNING"
                $Global:ScriptConfig.SkipExchange = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires Exchange module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'PnP.PowerShell' -ModuleDescription 'SharePoint and Teams discovery')) {
                Write-LogMessage "PnP.PowerShell module not available - disabling SharePoint discovery" "WARNING"
                $Global:ScriptConfig.SkipSharePoint = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires SharePoint module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'MicrosoftPowerBIMgmt' -ModuleDescription 'PowerBI discovery')) {
                Write-LogMessage "PowerBI module not available - disabling PowerBI discovery" "WARNING"
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Accounts' -ModuleDescription 'Azure discovery')) {
                Write-LogMessage "Az module not available - disabling Azure discovery" "WARNING"
                $Global:ScriptConfig.SkipAzureAD = $true
            } else {
                # Also check for Az.Resources module which is needed for service principal operations
                if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Resources' -ModuleDescription 'Azure service principal discovery')) {
                    Write-LogMessage "Az.Resources module not available - some Azure discovery features may be limited" "WARNING"
                }
            }
        }
        
        # Check if any modules are still enabled
        if ($Global:ScriptConfig.SkipPowerBI -and $Global:ScriptConfig.SkipTeams -and 
            $Global:ScriptConfig.SkipSharePoint -and $Global:ScriptConfig.SkipExchange -and 
            $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "No discovery modules can run due to missing PowerShell modules" "ERROR"
            
            # Provide specific install instructions
            $missingModulesMessage = @"
No discovery modules can run because required PowerShell modules are not installed.

To fix this issue, open PowerShell 7+ as Administrator and run:

For Azure Discovery:
Install-Module Az -Force -AllowClobber
# Or just the required sub-modules:
Install-Module Az.Accounts, Az.Resources -Force -AllowClobber

For Exchange Discovery:  
Install-Module ExchangeOnlineManagement -Force -AllowClobber

For SharePoint Discovery:
Install-Module PnP.PowerShell -Force -AllowClobber

For PowerBI Discovery:
Install-Module MicrosoftPowerBIMgmt -Force -AllowClobber

After installing the required modules, restart PowerShell and run this tool again.
"@
            [System.Windows.Forms.MessageBox]::Show($missingModulesMessage, "Install Required Modules", "OK", "Information")
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
        
        # Get credentials if MFA is not enabled
        $Credentials = $null
        if (-not $Global:ScriptConfig.UseMFA) {
            Write-LogMessage "Getting credentials..." "INFO"
            
            # Show confirmation dialog with the account being used
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Is this correct? Click No to go back and fix the account details.
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Authentication Details", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled authentication - returning to main form" "INFO"
                return $false
            }
            
            $Credentials = Get-Credential -UserName $Global:ScriptConfig.GlobalAdminAccount -Message "Enter password for $($Global:ScriptConfig.GlobalAdminAccount)"
            if (-not $Credentials) {
                Write-LogMessage "Credentials are required when MFA is not enabled" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Credentials are required when MFA is not enabled", "Authentication Error", "OK", "Error")
                return $false
            }
            
            # Validate that the entered username matches our expected account
            if ($Credentials.UserName -ne $Global:ScriptConfig.GlobalAdminAccount) {
                $mismatchMessage = @"
Username mismatch detected!

Expected: $($Global:ScriptConfig.GlobalAdminAccount)
Entered: $($Credentials.UserName)

Please use the correct administrator account or update the configuration.
"@
                Write-LogMessage "Username mismatch detected" "ERROR"
                [System.Windows.Forms.MessageBox]::Show($mismatchMessage, "Authentication Error", "OK", "Error")
                return $false
            }
        } else {
            # For MFA, still show confirmation but explain browser authentication
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Since MFA is enabled, a browser window will open for EACH service requiring authentication:
• Exchange Online
• Azure 
• SharePoint Online  
• PowerBI

IMPORTANT:
✓ Complete authentication in EACH browser window that opens
✓ Do NOT close browser windows until authentication is complete
✓ Have your authenticator app or phone ready
✓ The script will wait for you to complete each authentication

If you cancel any authentication, that service will be skipped.

Is this correct and are you ready to proceed?
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm MFA Authentication", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled MFA authentication - returning to main form" "INFO"
                return $false
            }
            
            Write-LogMessage "MFA authentication confirmed - will use browser-based authentication" "INFO"
        }
        
        # Connect to services with retry logic
        Write-LogMessage "Attempting to connect to services..." "INFO"
        Update-ProgressDisplay "Authenticating with Microsoft 365 services..."
        
        if (-not (Connect-ServicesWithValidation -Credentials $Credentials -MFAEnabled $Global:ScriptConfig.UseMFA)) {
            Write-LogMessage "Service connection failed or was cancelled by user" "ERROR"
            
            # Show final termination message
            $terminationMessage = @"
Discovery process terminated due to authentication failure.

The script cannot continue without valid service connections.

You can:
• Check your account credentials and permissions
• Verify MFA settings
• Ensure all required PowerShell modules are installed
• Try running the script again

Click OK to return to the main screen.
"@
            [System.Windows.Forms.MessageBox]::Show($terminationMessage, "Discovery Terminated", "OK", "Information")
            return $false
        }
        
        Write-LogMessage "Service connections established successfully" "SUCCESS"
        Update-ProgressDisplay "Authentication successful - starting discovery..."
        
        # Show connection success summary
        $connectedServices = @()
        if (-not $Global:ScriptConfig.SkipExchange) { $connectedServices += "Exchange Online" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $connectedServices += "SharePoint Online" }
        if (-not $Global:ScriptConfig.SkipPowerBI) { $connectedServices += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $connectedServices += "Azure" }
        
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Connected services: $($connectedServices -join ', ')" "SUCCESS"
        }
        
        # Run discovery modules
        Invoke-PowerBIDiscovery
        Invoke-TeamsDiscovery
        Invoke-SharePointDiscovery
        Invoke-ExchangeDiscovery
        Invoke-AzureADDiscovery
        
        # Generate summary report
        Write-LogMessage "Generating summary report..." "INFO"
        $SummaryFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $Summary = @"
Microsoft 365 and Azure Discovery Summary
=========================================
Organization: $($Global:ScriptConfig.OpCo)
Discovery Date: $(Get-Date)
Working Folder: $($Global:ScriptConfig.WorkingFolder)
Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Discovery Modules Executed:
- PowerBI Discovery: $(if (-not $Global:ScriptConfig.SkipPowerBI) { "✓ Executed" } else { "✗ Skipped" })
- Teams & Groups Discovery: $(if (-not $Global:ScriptConfig.SkipTeams) { "✓ Executed" } else { "✗ Skipped" })
- SharePoint & OneDrive Discovery: $(if (-not $Global:ScriptConfig.SkipSharePoint) { "✓ Executed" } else { "✗ Skipped" })
- Exchange Discovery: $(if (-not $Global:ScriptConfig.SkipExchange) { "✓ Executed" } else { "✗ Skipped" })
- Azure Discovery: $(if (-not $Global:ScriptConfig.SkipAzureAD) { "✓ Executed" } else { "✗ Skipped" })

Log File: $($Global:ScriptConfig.LogFile)
Summary File: $SummaryFile

Generated Files:
$(Get-ChildItem $Global:ScriptConfig.WorkingFolder -Filter "*$($Global:ScriptConfig.OpCo)*" | ForEach-Object { "- $($_.Name)" })
"@
        $Summary | Out-File $SummaryFile
        
        Write-LogMessage "=== Discovery completed successfully ===" "SUCCESS"
        Write-LogMessage "All reports saved to: $($Global:ScriptConfig.WorkingFolder)" "SUCCESS"
        Write-LogMessage "Summary report: $SummaryFile" "SUCCESS"
        
        return $true
    }
    catch {
        Write-LogMessage "Discovery failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Clean up connections
        Disconnect-AllServices
        Write-LogMessage "=== Discovery process completed ===" "INFO"
    }
}

#endregion

#region GUI Functions

function Show-ProgressForm {
    $Global:ScriptConfig.ProgressForm = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.ProgressForm.Text = "Discovery in Progress"
    $Global:ScriptConfig.ProgressForm.Size = New-Object System.Drawing.Size(500, 250)
    $Global:ScriptConfig.ProgressForm.StartPosition = "CenterParent"
    $Global:ScriptConfig.ProgressForm.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.ProgressForm.MaximizeBox = $false
    $Global:ScriptConfig.ProgressForm.MinimizeBox = $false
    $Global:ScriptConfig.ProgressForm.BackColor = [System.Drawing.Color]::White
    $Global:ScriptConfig.ProgressForm.TopMost = $true
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery in Progress..."
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
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "❌ Cancel Discovery"
    $cancelButton.Size = New-Object System.Drawing.Size(130, 35)
    $cancelButton.Location = New-Object System.Drawing.Point(185, 170)
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
            $message = "Discovery completed successfully!`n`nReports saved to: $($Global:ScriptConfig.WorkingFolder)`n`nWould you like to open the folder?"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Discovery Complete", "YesNo", "Information")
            
            if ($result -eq "Yes") {
                Start-Process "explorer.exe" $Global:ScriptConfig.WorkingFolder
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Discovery failed. Please check the log file for details.", "Discovery Failed", "OK", "Error")
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
    $Global:ScriptConfig.Form.Text = "Microsoft 365 Discovery Tool"
    $Global:ScriptConfig.Form.Size = New-Object System.Drawing.Size(800, 700)
    $Global:ScriptConfig.Form.StartPosition = "CenterScreen"
    $Global:ScriptConfig.Form.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.Form.MaximizeBox = $false
    $Global:ScriptConfig.Form.MinimizeBox = $false
    $Global:ScriptConfig.Form.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)

    # Create a header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(780, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Header title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($titleLabel)

    # Header subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "All-in-one solution for comprehensive M365 environment discovery"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
    $subtitleLabel.Size = New-Object System.Drawing.Size(600, 25)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 45)
    $headerPanel.Controls.Add($subtitleLabel)

    $Global:ScriptConfig.Form.Controls.Add($headerPanel)

    # Main content panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Size = New-Object System.Drawing.Size(780, 570)
    $mainPanel.Location = New-Object System.Drawing.Point(10, 100)
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.BorderStyle = "FixedSingle"

    # Configuration Group Box
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = "🔧 Basic Configuration"
    $configGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $configGroupBox.Size = New-Object System.Drawing.Size(750, 150)
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

    # MFA Checkbox
    $mfaCheckBox = New-Object System.Windows.Forms.CheckBox
    $mfaCheckBox.Text = "Multi-Factor Authentication (MFA) Enabled"
    $mfaCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mfaCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $mfaCheckBox.Location = New-Object System.Drawing.Point(430, 105)
    $mfaCheckBox.Checked = $false
    $configGroupBox.Controls.Add($mfaCheckBox)

    $mainPanel.Controls.Add($configGroupBox)

    # Discovery Modules Group Box
    $modulesGroupBox = New-Object System.Windows.Forms.GroupBox
    $modulesGroupBox.Text = "📊 Discovery Modules (Uncheck to Skip)"
    $modulesGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $modulesGroupBox.Size = New-Object System.Drawing.Size(750, 200)
    $modulesGroupBox.Location = New-Object System.Drawing.Point(15, 175)
    $modulesGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Module checkboxes
    $powerBICheckBox = New-Object System.Windows.Forms.CheckBox
    $powerBICheckBox.Text = "📈 PowerBI Discovery (Workspaces, Reports, Datasets)"
    $powerBICheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $powerBICheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $powerBICheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $powerBICheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($powerBICheckBox)

    $teamsCheckBox = New-Object System.Windows.Forms.CheckBox
    $teamsCheckBox.Text = "👥 Teams & Groups Discovery"
    $teamsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $teamsCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $teamsCheckBox.Location = New-Object System.Drawing.Point(380, 30)
    $teamsCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($teamsCheckBox)

    $sharepointCheckBox = New-Object System.Windows.Forms.CheckBox
    $sharepointCheckBox.Text = "📁 SharePoint & OneDrive Discovery"
    $sharepointCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $sharepointCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $sharepointCheckBox.Location = New-Object System.Drawing.Point(20, 65)
    $sharepointCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($sharepointCheckBox)

    $exchangeCheckBox = New-Object System.Windows.Forms.CheckBox
    $exchangeCheckBox.Text = "📧 Exchange Discovery (Mailboxes, Archives)"
    $exchangeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exchangeCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $exchangeCheckBox.Location = New-Object System.Drawing.Point(380, 65)
    $exchangeCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($exchangeCheckBox)

    $azureADCheckBox = New-Object System.Windows.Forms.CheckBox
    $azureADCheckBox.Text = "🔑 Azure Discovery (Users, Groups, Apps)"
    $azureADCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $azureADCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $azureADCheckBox.Location = New-Object System.Drawing.Point(20, 100)
    $azureADCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($azureADCheckBox)

    # Progress information
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "💡 All modules are enabled by default. Uncheck modules you want to skip."
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $progressLabel.Size = New-Object System.Drawing.Size(700, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(20, 140)
    $progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $modulesGroupBox.Controls.Add($progressLabel)

    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Text = "⚠️ Required modules will be checked automatically before execution."
    $requirementsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $requirementsLabel.Size = New-Object System.Drawing.Size(700, 20)
    $requirementsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $requirementsLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
    $modulesGroupBox.Controls.Add($requirementsLabel)

    $mainPanel.Controls.Add($modulesGroupBox)

    # Prerequisites Group Box
    $prereqGroupBox = New-Object System.Windows.Forms.GroupBox
    $prereqGroupBox.Text = "📋 Prerequisites"
    $prereqGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $prereqGroupBox.Size = New-Object System.Drawing.Size(750, 100)
    $prereqGroupBox.Location = New-Object System.Drawing.Point(15, 385)
    $prereqGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    $prereqText = New-Object System.Windows.Forms.Label
    $prereqText.Text = "Required PowerShell modules (install with 'Install-Module ModuleName -Force'):`n• ExchangeOnlineManagement  • PnP.PowerShell  • MicrosoftPowerBIMgmt  • Az.Accounts, Az.Resources`n`nEnsure you have Global Administrator permissions and PowerShell 7+ for best compatibility."
    $prereqText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $prereqText.Size = New-Object System.Drawing.Size(720, 70)
    $prereqText.Location = New-Object System.Drawing.Point(20, 25)
    $prereqText.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $prereqGroupBox.Controls.Add($prereqText)

    $mainPanel.Controls.Add($prereqGroupBox)

    # Action buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "🚀 Start Discovery"
    $runButton.Size = New-Object System.Drawing.Size(150, 40)
    $runButton.Location = New-Object System.Drawing.Point(470, 510)
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($runButton)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = "✅ Check Prerequisites"
    $validateButton.Size = New-Object System.Drawing.Size(150, 40)
    $validateButton.Location = New-Object System.Drawing.Point(310, 510)
    $validateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $validateButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $validateButton.ForeColor = [System.Drawing.Color]::White
    $validateButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($validateButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "❌ Exit"
    $exitButton.Size = New-Object System.Drawing.Size(100, 40)
    $exitButton.Location = New-Object System.Drawing.Point(640, 510)
    $exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($exitButton)

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

    $validateButton.Add_Click({
        # Update configuration
        $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
        $Global:ScriptConfig.OpCo = $opcoTextBox.Text
        $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
        $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked

        # Check what modules are needed and available
        $RequiredModules = @()
        $AvailableModules = @()
        $MissingModules = @()
        
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='ExchangeOnlineManagement'; Purpose='Exchange and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='PnP.PowerShell'; Purpose='SharePoint and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            $RequiredModules += @{Name='MicrosoftPowerBIMgmt'; Purpose='PowerBI discovery'}
        }
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            $RequiredModules += @{Name='Az.Accounts'; Purpose='Azure discovery'}
            $RequiredModules += @{Name='Az.Resources'; Purpose='Azure service principal discovery'}
        }

        if ($RequiredModules.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No modules are required since all discovery modules are disabled.", "Prerequisites Check", "OK", "Information")
            return
        }

        foreach ($module in $RequiredModules) {
            if (Get-Module -ListAvailable -Name $module.Name) {
                $AvailableModules += "$($module.Name) (for $($module.Purpose))"
            } else {
                $MissingModules += "$($module.Name) (for $($module.Purpose))"
            }
        }

        $resultMessage = ""
        
        if ($AvailableModules.Count -gt 0) {
            $resultMessage += "✅ Available Modules:`n$($AvailableModules -join "`n")`n`n"
        }
        
        if ($MissingModules.Count -gt 0) {
            $resultMessage += "❌ Missing Modules:`n$($MissingModules -join "`n")`n`n"
            $resultMessage += "To install missing modules, run PowerShell as Administrator:`n"
            foreach ($module in $RequiredModules) {
                if (-not (Get-Module -ListAvailable -Name $module.Name)) {
                    $resultMessage += "Install-Module $($module.Name) -Force`n"
                }
            }
            $resultMessage += "`nThe tool will automatically skip modules that aren't available."
        } else {
            $resultMessage += "🎉 All required modules are installed and ready!"
        }

        [System.Windows.Forms.MessageBox]::Show($resultMessage, "Prerequisites Check", "OK", "Information")
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
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Construct SharePoint Admin URL properly
        $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, try to extract the main part
            $tenantName = $domain.Split('.')[0]
        }
        $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        Write-LogMessage "SharePoint Admin URL set to: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
                $tenantName = $domain.Replace('.onmicrosoft.com', '')
            } else {
                $tenantName = $domain.Split('.')[0]
            }
            $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        }
        Write-LogMessage "SharePoint Admin Site: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
        # Create working directory if it doesn't exist
        if (-not (Test-Path $Global:ScriptConfig.WorkingFolder)) {
            New-Item -ItemType Directory -Path $Global:ScriptConfig.WorkingFolder -Force | Out-Null
            Write-LogMessage "Created working directory: $($Global:ScriptConfig.WorkingFolder)" "INFO"
        }
        
        # Setup log file
        $Global:ScriptConfig.LogFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        # Auto-adjust configuration based on available modules
        Write-LogMessage "Checking available modules and adjusting configuration..." "INFO"
        
        # Check and adjust each module
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'ExchangeOnlineManagement' -ModuleDescription 'Exchange and Teams discovery')) {
                Write-LogMessage "Exchange module not available - disabling Exchange and Teams discovery" "WARNING"
                $Global:ScriptConfig.SkipExchange = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires Exchange module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'PnP.PowerShell' -ModuleDescription 'SharePoint and Teams discovery')) {
                Write-LogMessage "PnP.PowerShell module not available - disabling SharePoint discovery" "WARNING"
                $Global:ScriptConfig.SkipSharePoint = $true
                if (-not $Global:ScriptConfig.SkipTeams) {
                    $Global:ScriptConfig.SkipTeams = $true
                    Write-LogMessage "Teams discovery requires SharePoint module - Teams discovery disabled" "WARNING"
                }
            }
        }
        
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'MicrosoftPowerBIMgmt' -ModuleDescription 'PowerBI discovery')) {
                Write-LogMessage "PowerBI module not available - disabling PowerBI discovery" "WARNING"
                $Global:ScriptConfig.SkipPowerBI = $true
            }
        }
        
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Accounts' -ModuleDescription 'Azure discovery')) {
                Write-LogMessage "Az module not available - disabling Azure discovery" "WARNING"
                $Global:ScriptConfig.SkipAzureAD = $true
            } else {
                # Also check for Az.Resources module which is needed for service principal operations
                if (-not (Test-ModuleAndAdjustConfig -ModuleName 'Az.Resources' -ModuleDescription 'Azure service principal discovery')) {
                    Write-LogMessage "Az.Resources module not available - some Azure discovery features may be limited" "WARNING"
                }
            }
        }
        
        # Check if any modules are still enabled
        if ($Global:ScriptConfig.SkipPowerBI -and $Global:ScriptConfig.SkipTeams -and 
            $Global:ScriptConfig.SkipSharePoint -and $Global:ScriptConfig.SkipExchange -and 
            $Global:ScriptConfig.SkipAzureAD) {
            Write-LogMessage "No discovery modules can run due to missing PowerShell modules" "ERROR"
            
            # Provide specific install instructions
            $missingModulesMessage = @"
No discovery modules can run because required PowerShell modules are not installed.

To fix this issue, open PowerShell 7+ as Administrator and run:

For Azure Discovery:
Install-Module Az -Force -AllowClobber
# Or just the required sub-modules:
Install-Module Az.Accounts, Az.Resources -Force -AllowClobber

For Exchange Discovery:  
Install-Module ExchangeOnlineManagement -Force -AllowClobber

For SharePoint Discovery:
Install-Module PnP.PowerShell -Force -AllowClobber

For PowerBI Discovery:
Install-Module MicrosoftPowerBIMgmt -Force -AllowClobber

After installing the required modules, restart PowerShell and run this tool again.
"@
            [System.Windows.Forms.MessageBox]::Show($missingModulesMessage, "Install Required Modules", "OK", "Information")
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
        
        # Get credentials if MFA is not enabled
        $Credentials = $null
        if (-not $Global:ScriptConfig.UseMFA) {
            Write-LogMessage "Getting credentials..." "INFO"
            
            # Show confirmation dialog with the account being used
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Is this correct? Click No to go back and fix the account details.
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Authentication Details", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled authentication - returning to main form" "INFO"
                return $false
            }
            
            $Credentials = Get-Credential -UserName $Global:ScriptConfig.GlobalAdminAccount -Message "Enter password for $($Global:ScriptConfig.GlobalAdminAccount)"
            if (-not $Credentials) {
                Write-LogMessage "Credentials are required when MFA is not enabled" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Credentials are required when MFA is not enabled", "Authentication Error", "OK", "Error")
                return $false
            }
            
            # Validate that the entered username matches our expected account
            if ($Credentials.UserName -ne $Global:ScriptConfig.GlobalAdminAccount) {
                $mismatchMessage = @"
Username mismatch detected!

Expected: $($Global:ScriptConfig.GlobalAdminAccount)
Entered: $($Credentials.UserName)

Please use the correct administrator account or update the configuration.
"@
                Write-LogMessage "Username mismatch detected" "ERROR"
                [System.Windows.Forms.MessageBox]::Show($mismatchMessage, "Authentication Error", "OK", "Error")
                return $false
            }
        } else {
            # For MFA, still show confirmation but explain browser authentication
            $confirmMessage = @"
You are about to authenticate with:
Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Since MFA is enabled, you will be prompted to sign in through your browser 
for each service (Exchange, SharePoint, PowerBI, Azure AD).

Make sure you:
• Have access to your authenticator app or phone
• Are ready to complete MFA challenges
• Don't close browser windows during authentication

Is this correct and are you ready to proceed?
"@
            $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm MFA Authentication", "YesNo", "Question")
            
            if ($confirmResult -eq "No") {
                Write-LogMessage "User cancelled MFA authentication - returning to main form" "INFO"
                return $false
            }
            
            Write-LogMessage "MFA authentication confirmed - will use browser-based authentication" "INFO"
        }
        
        # Connect to services with retry logic
        Write-LogMessage "Attempting to connect to services..." "INFO"
        Update-ProgressDisplay "Authenticating with Microsoft 365 services..."
        
        if (-not (Connect-ServicesWithValidation -Credentials $Credentials -MFAEnabled $Global:ScriptConfig.UseMFA)) {
            Write-LogMessage "Service connection failed or was cancelled by user" "ERROR"
            
            # Show final termination message
            $terminationMessage = @"
Discovery process terminated due to authentication failure.

The script cannot continue without valid service connections.

You can:
• Check your account credentials and permissions
• Verify MFA settings
• Ensure all required PowerShell modules are installed
• Try running the script again

Click OK to return to the main screen.
"@
            [System.Windows.Forms.MessageBox]::Show($terminationMessage, "Discovery Terminated", "OK", "Information")
            return $false
        }
        
        Write-LogMessage "Service connections established successfully" "SUCCESS"
        Update-ProgressDisplay "Authentication successful - starting discovery..."
        
        # Show connection success summary
        $connectedServices = @()
        if (-not $Global:ScriptConfig.SkipExchange) { $connectedServices += "Exchange Online" }
        if (-not $Global:ScriptConfig.SkipSharePoint) { $connectedServices += "SharePoint Online" }
        if (-not $Global:ScriptConfig.SkipPowerBI) { $connectedServices += "PowerBI" }
        if (-not $Global:ScriptConfig.SkipAzureAD) { $connectedServices += "Azure" }
        
        if ($connectedServices.Count -gt 0) {
            Write-LogMessage "Connected services: $($connectedServices -join ', ')" "SUCCESS"
        }
        
        # Run discovery modules
        Invoke-PowerBIDiscovery
        Invoke-TeamsDiscovery
        Invoke-SharePointDiscovery
        Invoke-ExchangeDiscovery
        Invoke-AzureADDiscovery
        
        # Generate summary report
        Write-LogMessage "Generating summary report..." "INFO"
        $SummaryFile = Join-Path $Global:ScriptConfig.WorkingFolder "$($Global:ScriptConfig.OpCo)_Discovery_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $Summary = @"
Microsoft 365 and Azure Discovery Summary
=========================================
Organization: $($Global:ScriptConfig.OpCo)
Discovery Date: $(Get-Date)
Working Folder: $($Global:ScriptConfig.WorkingFolder)
Admin Account: $($Global:ScriptConfig.GlobalAdminAccount)
MFA Enabled: $($Global:ScriptConfig.UseMFA)

Discovery Modules Executed:
- PowerBI Discovery: $(if (-not $Global:ScriptConfig.SkipPowerBI) { "✓ Executed" } else { "✗ Skipped" })
- Teams & Groups Discovery: $(if (-not $Global:ScriptConfig.SkipTeams) { "✓ Executed" } else { "✗ Skipped" })
- SharePoint & OneDrive Discovery: $(if (-not $Global:ScriptConfig.SkipSharePoint) { "✓ Executed" } else { "✗ Skipped" })
- Exchange Discovery: $(if (-not $Global:ScriptConfig.SkipExchange) { "✓ Executed" } else { "✗ Skipped" })
- Azure Discovery: $(if (-not $Global:ScriptConfig.SkipAzureAD) { "✓ Executed" } else { "✗ Skipped" })

Log File: $($Global:ScriptConfig.LogFile)
Summary File: $SummaryFile

Generated Files:
$(Get-ChildItem $Global:ScriptConfig.WorkingFolder -Filter "*$($Global:ScriptConfig.OpCo)*" | ForEach-Object { "- $($_.Name)" })
"@
        $Summary | Out-File $SummaryFile
        
        Write-LogMessage "=== Discovery completed successfully ===" "SUCCESS"
        Write-LogMessage "All reports saved to: $($Global:ScriptConfig.WorkingFolder)" "SUCCESS"
        Write-LogMessage "Summary report: $SummaryFile" "SUCCESS"
        
        return $true
    }
    catch {
        Write-LogMessage "Discovery failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Clean up connections
        Disconnect-AllServices
        Write-LogMessage "=== Discovery process completed ===" "INFO"
    }
}

#endregion

#region GUI Functions

function Show-ProgressForm {
    $Global:ScriptConfig.ProgressForm = New-Object System.Windows.Forms.Form
    $Global:ScriptConfig.ProgressForm.Text = "Discovery in Progress"
    $Global:ScriptConfig.ProgressForm.Size = New-Object System.Drawing.Size(500, 250)
    $Global:ScriptConfig.ProgressForm.StartPosition = "CenterParent"
    $Global:ScriptConfig.ProgressForm.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.ProgressForm.MaximizeBox = $false
    $Global:ScriptConfig.ProgressForm.MinimizeBox = $false
    $Global:ScriptConfig.ProgressForm.BackColor = [System.Drawing.Color]::White
    $Global:ScriptConfig.ProgressForm.TopMost = $true
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery in Progress..."
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
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "❌ Cancel Discovery"
    $cancelButton.Size = New-Object System.Drawing.Size(130, 35)
    $cancelButton.Location = New-Object System.Drawing.Point(185, 170)
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
            $message = "Discovery completed successfully!`n`nReports saved to: $($Global:ScriptConfig.WorkingFolder)`n`nWould you like to open the folder?"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Discovery Complete", "YesNo", "Information")
            
            if ($result -eq "Yes") {
                Start-Process "explorer.exe" $Global:ScriptConfig.WorkingFolder
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Discovery failed. Please check the log file for details.", "Discovery Failed", "OK", "Error")
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
    $Global:ScriptConfig.Form.Text = "Microsoft 365 Discovery Tool"
    $Global:ScriptConfig.Form.Size = New-Object System.Drawing.Size(800, 700)
    $Global:ScriptConfig.Form.StartPosition = "CenterScreen"
    $Global:ScriptConfig.Form.FormBorderStyle = "FixedDialog"
    $Global:ScriptConfig.Form.MaximizeBox = $false
    $Global:ScriptConfig.Form.MinimizeBox = $false
    $Global:ScriptConfig.Form.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)

    # Create a header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(780, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Header title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🚀 Microsoft 365 Discovery Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($titleLabel)

    # Header subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "All-in-one solution for comprehensive M365 environment discovery"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
    $subtitleLabel.Size = New-Object System.Drawing.Size(600, 25)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 45)
    $headerPanel.Controls.Add($subtitleLabel)

    $Global:ScriptConfig.Form.Controls.Add($headerPanel)

    # Main content panel
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Size = New-Object System.Drawing.Size(780, 570)
    $mainPanel.Location = New-Object System.Drawing.Point(10, 100)
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.BorderStyle = "FixedSingle"

    # Configuration Group Box
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = "🔧 Basic Configuration"
    $configGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $configGroupBox.Size = New-Object System.Drawing.Size(750, 150)
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

    # MFA Checkbox
    $mfaCheckBox = New-Object System.Windows.Forms.CheckBox
    $mfaCheckBox.Text = "Multi-Factor Authentication (MFA) Enabled"
    $mfaCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mfaCheckBox.Size = New-Object System.Drawing.Size(250, 25)
    $mfaCheckBox.Location = New-Object System.Drawing.Point(430, 105)
    $mfaCheckBox.Checked = $false
    $configGroupBox.Controls.Add($mfaCheckBox)

    $mainPanel.Controls.Add($configGroupBox)

    # Discovery Modules Group Box
    $modulesGroupBox = New-Object System.Windows.Forms.GroupBox
    $modulesGroupBox.Text = "📊 Discovery Modules (Uncheck to Skip)"
    $modulesGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $modulesGroupBox.Size = New-Object System.Drawing.Size(750, 200)
    $modulesGroupBox.Location = New-Object System.Drawing.Point(15, 175)
    $modulesGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    # Module checkboxes
    $powerBICheckBox = New-Object System.Windows.Forms.CheckBox
    $powerBICheckBox.Text = "📈 PowerBI Discovery (Workspaces, Reports, Datasets)"
    $powerBICheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $powerBICheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $powerBICheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $powerBICheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($powerBICheckBox)

    $teamsCheckBox = New-Object System.Windows.Forms.CheckBox
    $teamsCheckBox.Text = "👥 Teams & Groups Discovery"
    $teamsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $teamsCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $teamsCheckBox.Location = New-Object System.Drawing.Point(380, 30)
    $teamsCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($teamsCheckBox)

    $sharepointCheckBox = New-Object System.Windows.Forms.CheckBox
    $sharepointCheckBox.Text = "📁 SharePoint & OneDrive Discovery"
    $sharepointCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $sharepointCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $sharepointCheckBox.Location = New-Object System.Drawing.Point(20, 65)
    $sharepointCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($sharepointCheckBox)

    $exchangeCheckBox = New-Object System.Windows.Forms.CheckBox
    $exchangeCheckBox.Text = "📧 Exchange Discovery (Mailboxes, Archives)"
    $exchangeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exchangeCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $exchangeCheckBox.Location = New-Object System.Drawing.Point(380, 65)
    $exchangeCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($exchangeCheckBox)

    $azureADCheckBox = New-Object System.Windows.Forms.CheckBox
    $azureADCheckBox.Text = "🔑 Azure Discovery (Users, Groups, Apps)"
    $azureADCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $azureADCheckBox.Size = New-Object System.Drawing.Size(350, 25)
    $azureADCheckBox.Location = New-Object System.Drawing.Point(20, 100)
    $azureADCheckBox.Checked = $true
    $modulesGroupBox.Controls.Add($azureADCheckBox)

    # Progress information
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "💡 All modules are enabled by default. Uncheck modules you want to skip."
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $progressLabel.Size = New-Object System.Drawing.Size(700, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(20, 140)
    $progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $modulesGroupBox.Controls.Add($progressLabel)

    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Text = "⚠️ Required modules will be checked automatically before execution."
    $requirementsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $requirementsLabel.Size = New-Object System.Drawing.Size(700, 20)
    $requirementsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $requirementsLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
    $modulesGroupBox.Controls.Add($requirementsLabel)

    $mainPanel.Controls.Add($modulesGroupBox)

    # Prerequisites Group Box
    $prereqGroupBox = New-Object System.Windows.Forms.GroupBox
    $prereqGroupBox.Text = "📋 Prerequisites"
    $prereqGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $prereqGroupBox.Size = New-Object System.Drawing.Size(750, 100)
    $prereqGroupBox.Location = New-Object System.Drawing.Point(15, 385)
    $prereqGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)

    $prereqText = New-Object System.Windows.Forms.Label
    $prereqText.Text = "Required PowerShell modules (install with 'Install-Module ModuleName -Force'):`n• ExchangeOnlineManagement  • PnP.PowerShell  • MicrosoftPowerBIMgmt  • Az.Accounts, Az.Resources`n`nEnsure you have Global Administrator permissions and PowerShell 7+ for best compatibility."
    $prereqText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $prereqText.Size = New-Object System.Drawing.Size(720, 70)
    $prereqText.Location = New-Object System.Drawing.Point(20, 25)
    $prereqText.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $prereqGroupBox.Controls.Add($prereqText)

    $mainPanel.Controls.Add($prereqGroupBox)

    # Action buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "🚀 Start Discovery"
    $runButton.Size = New-Object System.Drawing.Size(150, 40)
    $runButton.Location = New-Object System.Drawing.Point(470, 510)
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $runButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $runButton.ForeColor = [System.Drawing.Color]::White
    $runButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($runButton)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = "✅ Check Prerequisites"
    $validateButton.Size = New-Object System.Drawing.Size(150, 40)
    $validateButton.Location = New-Object System.Drawing.Point(310, 510)
    $validateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $validateButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $validateButton.ForeColor = [System.Drawing.Color]::White
    $validateButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($validateButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "❌ Exit"
    $exitButton.Size = New-Object System.Drawing.Size(100, 40)
    $exitButton.Location = New-Object System.Drawing.Point(640, 510)
    $exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $mainPanel.Controls.Add($exitButton)

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

    $validateButton.Add_Click({
        # Update configuration
        $Global:ScriptConfig.WorkingFolder = $workingFolderTextBox.Text
        $Global:ScriptConfig.OpCo = $opcoTextBox.Text
        $Global:ScriptConfig.GlobalAdminAccount = $adminAccountTextBox.Text
        $Global:ScriptConfig.UseMFA = $mfaCheckBox.Checked
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked

        # Check what modules are needed and available
        $RequiredModules = @()
        $AvailableModules = @()
        $MissingModules = @()
        
        if (-not $Global:ScriptConfig.SkipExchange -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='ExchangeOnlineManagement'; Purpose='Exchange and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipSharePoint -or -not $Global:ScriptConfig.SkipTeams) {
            $RequiredModules += @{Name='PnP.PowerShell'; Purpose='SharePoint and Teams discovery'}
        }
        if (-not $Global:ScriptConfig.SkipPowerBI) {
            $RequiredModules += @{Name='MicrosoftPowerBIMgmt'; Purpose='PowerBI discovery'}
        }
        if (-not $Global:ScriptConfig.SkipAzureAD) {
            $RequiredModules += @{Name='Az.Accounts'; Purpose='Azure discovery'}
            $RequiredModules += @{Name='Az.Resources'; Purpose='Azure service principal discovery'}
        }

        if ($RequiredModules.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No modules are required since all discovery modules are disabled.", "Prerequisites Check", "OK", "Information")
            return
        }

        foreach ($module in $RequiredModules) {
            if (Get-Module -ListAvailable -Name $module.Name) {
                $AvailableModules += "$($module.Name) (for $($module.Purpose))"
            } else {
                $MissingModules += "$($module.Name) (for $($module.Purpose))"
            }
        }

        $resultMessage = ""
        
        if ($AvailableModules.Count -gt 0) {
            $resultMessage += "✅ Available Modules:`n$($AvailableModules -join "`n")`n`n"
        }
        
        if ($MissingModules.Count -gt 0) {
            $resultMessage += "❌ Missing Modules:`n$($MissingModules -join "`n")`n`n"
            $resultMessage += "To install missing modules, run PowerShell as Administrator:`n"
            foreach ($module in $RequiredModules) {
                if (-not (Get-Module -ListAvailable -Name $module.Name)) {
                    $resultMessage += "Install-Module $($module.Name) -Force`n"
                }
            }
            $resultMessage += "`nThe tool will automatically skip modules that aren't available."
        } else {
            $resultMessage += "🎉 All required modules are installed and ready!"
        }

        [System.Windows.Forms.MessageBox]::Show($resultMessage, "Prerequisites Check", "OK", "Information")
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
        $Global:ScriptConfig.SkipPowerBI = -not $powerBICheckBox.Checked
        $Global:ScriptConfig.SkipTeams = -not $teamsCheckBox.Checked
        $Global:ScriptConfig.SkipSharePoint = -not $sharepointCheckBox.Checked
        $Global:ScriptConfig.SkipExchange = -not $exchangeCheckBox.Checked
        $Global:ScriptConfig.SkipAzureAD = -not $azureADCheckBox.Checked
        
        # Construct SharePoint Admin URL properly
        $domain = $Global:ScriptConfig.GlobalAdminAccount.Split('@')[1]
        if ($domain -match '\.onmicrosoft\.com
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI) {
            # Extract tenant name from onmicrosoft.com domain
            $tenantName = $domain.Replace('.onmicrosoft.com', '')
        } else {
            # For custom domains, try to extract the main part
            $tenantName = $domain.Split('.')[0]
        }
        $Global:ScriptConfig.SharePointAdminSite = "https://$tenantName-admin.sharepoint.com"
        Write-LogMessage "SharePoint Admin URL set to: $($Global:ScriptConfig.SharePointAdminSite)" "INFO"
        
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

📊 Discovery Modules:
$($selectedModules -join "`n")

⚠️ Make sure:
• You have Global Administrator permissions
• The admin account is correct
• Required PowerShell modules are installed

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

    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Global:ScriptConfig.Form.ShowDialog()
}

#endregion

# Start the GUI
New-DiscoveryGUI