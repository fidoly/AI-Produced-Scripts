<#
.SYNOPSIS
    Tenant-Accounts-Details.ps1 - Microsoft Tenant Sign-in Logs and Account Details Collector
    Version: 3.0.3
    
.DESCRIPTION
    This script collects a comprehensive set of sign-in logs and account properties from a Microsoft Entra tenant.
    The GUI allows for the selection of specific user properties to be collected, enabling fast, targeted reports.
    It includes a robust, automatic retry mechanism to handle Microsoft Graph API throttling and logic to assist in tenant decommission auditing.
    
.FEATURES
    - GUI allows for selective loading of user properties to control performance.
    - High-impact (slow) properties are clearly marked with an asterisk (*).
    - Includes a smart "Account Status" column to accurately distinguish shared mailboxes from deactivated users.
    - Uses the correct, efficient API endpoint for last sign-in data.
    - Automatic, resilient handling of API throttling with exponential backoff and retry.
    
.AUTHOR
    Designed by Phil Rangel
    
.NOTES
    Required Graph API Permissions:
    - User.Read.All
    - AuditLog.Read.All
    - Directory.Read.All
    - Application.Read.All
    - RoleManagement.Read.Directory
    - Policy.Read.All
    - Reports.Read.All
#>

# Script version
$ScriptVersion = "3.0.3" # REFINEMENT: Improved logic to differentiate Shared Mailboxes from Deactivated Users.
$ScriptName = "Tenant-Accounts-Details"

# Set strict mode for better error detection
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Module Management
function Test-RequiredModules {
    param (
        [string[]]$RequiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Applications',
            'Microsoft.Graph.Reports',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Identity.SignIns'
        )
    )
    
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    $modulesToInstall = @()
    
    foreach ($module in $RequiredModules) {
        $installed = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue
        if (-not $installed) {
            $modulesToInstall += $module
        }
    }
    
    if ($modulesToInstall.Count -gt 0) {
        Write-Host "The following modules need to be installed: $($modulesToInstall -join ', ')" -ForegroundColor Yellow
        $installChoice = [System.Windows.Forms.MessageBox]::Show(
            "The following modules need to be installed:`n`n$($modulesToInstall -join "`n")`n`nDo you want to install them now?",
            "Module Installation Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($installChoice -eq 'Yes') {
            foreach ($module in $modulesToInstall) {
                try {
                    Write-Host "Installing $module..." -ForegroundColor Yellow
                    Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
                    Write-Host "$module installed successfully" -ForegroundColor Green
                } catch { Write-Host "Failed to install ${module}: $($_.Exception.Message)" -ForegroundColor Red; return $false }
            }
        } else { return $false }
    }
    
    foreach ($module in $RequiredModules) {
        try {
            Import-Module $module -ErrorAction Stop
        } catch { Write-Host "Failed to load ${module}: $($_.Exception.Message)" -ForegroundColor Red; return $false }
    }
    return $true
}
#endregion

#region Windows Forms Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Global Variables
$script:TenantName = ""
$script:ExportPath = ""
$script:AuthMethod = "Interactive"
$script:PrincipalTypeToProcess = "Users"
$script:SelectedUserProperties = @()
$script:ProcessingCancelled = $false
$script:LogFile = ""
$script:StartTime = Get-Date
$script:ProcessedCount = 0
$script:TotalCount = 0
$script:Errors = @()
#endregion

#region Logging Functions
function Write-Log {
    param([string]$Message, [ValidateSet('Info', 'Warning', 'Error', 'Success')][string]$Level = 'Info')
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; $logEntry = "$timestamp [$Level] $Message"
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue }
    switch ($Level) {
        'Error'   { Write-Host $logEntry -ForegroundColor Red }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry -ForegroundColor Cyan }
    }
}
#endregion

#region GUI Functions
function Show-ConfigurationGUI {
    #region Form Setup
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$ScriptName v$ScriptVersion - Configuration"
    $form.Size = New-Object System.Drawing.Size(800, 850)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    
    $y = 20
    #endregion

    #region Main Controls
    $titleLabel = New-Object System.Windows.Forms.Label; $titleLabel.Text = "Microsoft Tenant Account Details Collector"; $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold); $titleLabel.Location = New-Object System.Drawing.Point(10, $y); $titleLabel.Size = New-Object System.Drawing.Size(600, 25); $form.Controls.Add($titleLabel); $y += 35
    $tenantLabel = New-Object System.Windows.Forms.Label; $tenantLabel.Text = "Tenant Name/ID:"; $tenantLabel.Location = New-Object System.Drawing.Point(10, $y); $tenantLabel.Size = New-Object System.Drawing.Size(150, 20); $form.Controls.Add($tenantLabel)
    $tenantTextBox = New-Object System.Windows.Forms.TextBox; $tenantTextBox.Location = New-Object System.Drawing.Point(170, $y); $tenantTextBox.Size = New-Object System.Drawing.Size(350, 20); $tenantTextBox.Text = "contoso.onmicrosoft.com"; $form.Controls.Add($tenantTextBox); $y += 35
    $pathLabel = New-Object System.Windows.Forms.Label; $pathLabel.Text = "Export Location:"; $pathLabel.Location = New-Object System.Drawing.Point(10, $y); $pathLabel.Size = New-Object System.Drawing.Size(150, 20); $form.Controls.Add($pathLabel)
    $pathTextBox = New-Object System.Windows.Forms.TextBox; $pathTextBox.Location = New-Object System.Drawing.Point(170, $y); $pathTextBox.Size = New-Object System.Drawing.Size(300, 20); $pathTextBox.Text = "C:\Temp"; $form.Controls.Add($pathTextBox)
    $browseButton = New-Object System.Windows.Forms.Button; $browseButton.Text = "Browse..."; $browseButton.Location = New-Object System.Drawing.Point(480, ($y - 1)); $browseButton.Size = New-Object System.Drawing.Size(75, 23); $form.Controls.Add($browseButton); $y += 35
    $authLabel = New-Object System.Windows.Forms.Label; $authLabel.Text = "Authentication Method:"; $authLabel.Location = New-Object System.Drawing.Point(10, $y); $authLabel.Size = New-Object System.Drawing.Size(150, 20); $form.Controls.Add($authLabel)
    $authComboBox = New-Object System.Windows.Forms.ComboBox; $authComboBox.Location = New-Object System.Drawing.Point(170, $y); $authComboBox.Size = New-Object System.Drawing.Size(200, 20); $authComboBox.DropDownStyle = 'DropDownList'; $authComboBox.Items.AddRange(@('Interactive Browser', 'Device Code Flow')); $authComboBox.SelectedIndex = 0; $form.Controls.Add($authComboBox); $y += 35
    $principalTypeLabel = New-Object System.Windows.Forms.Label; $principalTypeLabel.Text = "Principal Type to Process:"; $principalTypeLabel.Location = New-Object System.Drawing.Point(10, $y); $principalTypeLabel.Size = New-Object System.Drawing.Size(150, 20); $form.Controls.Add($principalTypeLabel)
    $principalTypeComboBox = New-Object System.Windows.Forms.ComboBox; $principalTypeComboBox.Location = New-Object System.Drawing.Point(170, $y); $principalTypeComboBox.Size = New-Object System.Drawing.Size(200, 20); $principalTypeComboBox.DropDownStyle = 'DropDownList'; $principalTypeComboBox.Items.AddRange(@('Users', 'Service Principals')); $principalTypeComboBox.SelectedIndex = 0; $form.Controls.Add($principalTypeComboBox); $y += 45
    #endregion
    
    #region User Property Selection
    $propertiesGroup = New-Object System.Windows.Forms.GroupBox; $propertiesGroup.Text = "User Properties to Collect"; $propertiesGroup.Location = New-Object System.Drawing.Point(10, $y); $propertiesGroup.Size = New-Object System.Drawing.Size(760, 450); $form.Controls.Add($propertiesGroup)
    
    $propertiesChecklist = New-Object System.Windows.Forms.CheckedListBox; $propertiesChecklist.Location = New-Object System.Drawing.Point(10, 25); $propertiesChecklist.Size = New-Object System.Drawing.Size(740, 380); $propertiesChecklist.CheckOnClick = $true; $propertiesChecklist.ColumnWidth = 240; $propertiesChecklist.MultiColumn = $true; $propertiesGroup.Controls.Add($propertiesChecklist)

    $selectAllButton = New-Object System.Windows.Forms.Button; $selectAllButton.Text = "Select All"; $selectAllButton.Location = New-Object System.Drawing.Point(10, 415); $selectAllButton.Size = New-Object System.Drawing.Size(100, 25); $propertiesGroup.Controls.Add($selectAllButton)
    $selectNoneButton = New-Object System.Windows.Forms.Button; $selectNoneButton.Text = "Select None"; $selectNoneButton.Location = New-Object System.Drawing.Point(120, 415); $selectNoneButton.Size = New-Object System.Drawing.Size(100, 25); $propertiesGroup.Controls.Add($selectNoneButton)
    
    $allProperties = @(
        "Given Name", "Surname", "Mail", "Mail Nickname", "Other Mails", "Proxy Addresses", "User Type", "Job Title", "Department", "Company Name", "Office Location", "Manager*", "Business Phones", "Mobile Phone", "Street Address", "City", "State", "Postal Code", "Country", "Usage Location", "Account Enabled", "Creation Date", "External User State", "Last Interactive Sign-in*", "Last Non-Interactive Sign-in*", "On-Premises Sync Enabled", "On-Premises Immutable ID", "On-Premises Last Sync", "On-Premises SAM Account Name", "On-Premises UPN", "On-Premises Distinguished Name", "Direct Reports Count*", "Assigned Licenses*", "Assigned Roles*", "MFA Methods*"
    )
    $propertiesChecklist.Items.AddRange($allProperties)

    $defaultProperties = @("Mail", "Account Enabled", "Creation Date", "Last Interactive Sign-in*", "Assigned Licenses*", "Surname", "Office Location")
    for ($i = 0; $i -lt $propertiesChecklist.Items.Count; $i++) {
        if ($propertiesChecklist.Items[$i] -in $defaultProperties) {
            $propertiesChecklist.SetItemChecked($i, $true)
        }
    }
    
    $y += 460
    #endregion

    #region Footer and Buttons
    $footerLabel = New-Object System.Windows.Forms.Label; $footerLabel.Text = "* This property will add considerable processing time."; $footerLabel.Location = New-Object System.Drawing.Point(10, $y); $footerLabel.Size = New-Object System.Drawing.Size(600, 20); $footerLabel.ForeColor = [System.Drawing.Color]::Red; $form.Controls.Add($footerLabel); $y += 30

    $startButton = New-Object System.Windows.Forms.Button; $startButton.Text = "Start Collection"; $startButton.Location = New-Object System.Drawing.Point(540, $y); $startButton.Size = New-Object System.Drawing.Size(120, 30); $startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $startButton.ForeColor = [System.Drawing.Color]::White; $startButton.FlatStyle = 'Flat'; $form.Controls.Add($startButton)
    $cancelButton = New-Object System.Windows.Forms.Button; $cancelButton.Text = "Cancel"; $cancelButton.Location = New-Object System.Drawing.Point(670, $y); $cancelButton.Size = New-Object System.Drawing.Size(100, 30); $cancelButton.FlatStyle = 'Flat'; $form.Controls.Add($cancelButton)
    #endregion
    
    #region Event Handlers
    $browseButton.Add_Click({ $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog; if ($folderBrowser.ShowDialog() -eq 'OK') { $pathTextBox.Text = $folderBrowser.SelectedPath } })
    
    $principalTypeComboBox.Add_SelectedIndexChanged({
        $isUsers = $principalTypeComboBox.SelectedItem -eq 'Users'
        $propertiesGroup.Enabled = $isUsers
    })

    $selectAllButton.Add_Click({ for ($i = 0; $i -lt $propertiesChecklist.Items.Count; $i++) { $propertiesChecklist.SetItemChecked($i, $true) } })
    $selectNoneButton.Add_Click({ for ($i = 0; $i -lt $propertiesChecklist.Items.Count; $i++) { $propertiesChecklist.SetItemChecked($i, $false) } })

    $startButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($tenantTextBox.Text)) { [System.Windows.Forms.MessageBox]::Show("Tenant Name/ID cannot be empty.", "Validation Error", 'OK', 'Warning'); return }
        if (-not (Test-Path $pathTextBox.Text)) {
            $createPath = [System.Windows.Forms.MessageBox]::Show("The export path does not exist. Do you want to create it?", "Path Not Found", 'YesNo', 'Question')
            if ($createPath -eq 'Yes') { try { New-Item -Path $pathTextBox.Text -ItemType Directory -Force | Out-Null } catch { [System.Windows.Forms.MessageBox]::Show("Failed to create directory: $_", "Error", 'OK', 'Error'); return } }
            else { return }
        }
        
        $script:TenantName = $tenantTextBox.Text; $script:ExportPath = $pathTextBox.Text
        $script:AuthMethod = if ($authComboBox.SelectedItem -eq 'Device Code Flow') { 'DeviceCode' } else { 'Interactive' }
        $script:PrincipalTypeToProcess = $principalTypeComboBox.SelectedItem
        $script:SelectedUserProperties = $propertiesChecklist.CheckedItems | ForEach-Object { $_ -replace '\*','' }

        $form.DialogResult = 'OK'; $form.Close()
    })
    $cancelButton.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() })
    
    $propertiesGroup.Enabled = $true
    return $form.ShowDialog()
    #endregion
}

function Show-ProgressWindow {
    try { $progressForm = New-Object System.Windows.Forms.Form; $progressForm.Text = "$ScriptName - Processing"; $progressForm.Size = New-Object System.Drawing.Size(600, 480); $progressForm.StartPosition = "CenterScreen"; $progressForm.FormBorderStyle = 'FixedDialog'; $progressForm.MaximizeBox = $false; $progressForm.ControlBox = $false; $statusLabel = New-Object System.Windows.Forms.Label; $statusLabel.Text = "Initializing..."; $statusLabel.Location = New-Object System.Drawing.Point(10, 20); $statusLabel.Size = New-Object System.Drawing.Size(560, 20); $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold); $progressForm.Controls.Add($statusLabel); $progressBar = New-Object System.Windows.Forms.ProgressBar; $progressBar.Location = New-Object System.Drawing.Point(10, 50); $progressBar.Size = New-Object System.Drawing.Size(560, 30); $progressBar.Style = 'Continuous'; $progressForm.Controls.Add($progressBar); $detailsTextBox = New-Object System.Windows.Forms.TextBox; $detailsTextBox.Location = New-Object System.Drawing.Point(10, 90); $detailsTextBox.Size = New-Object System.Drawing.Size(560, 230); $detailsTextBox.Multiline = $true; $detailsTextBox.ScrollBars = 'Vertical'; $detailsTextBox.ReadOnly = $true; $detailsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9); $progressForm.Controls.Add($detailsTextBox); $statsLabel = New-Object System.Windows.Forms.Label; $statsLabel.Text = "Statistics:"; $statsLabel.Location = New-Object System.Drawing.Point(10, 330); $statsLabel.Size = New-Object System.Drawing.Size(560, 80); $statsLabel.Font = New-Object System.Drawing.Font("Consolas", 9); $progressForm.Controls.Add($statsLabel); $cancelButton = New-Object System.Windows.Forms.Button; $cancelButton.Text = "Cancel"; $cancelButton.Location = New-Object System.Drawing.Point(470, 415); $cancelButton.Size = New-Object System.Drawing.Size(100, 25); $progressForm.Controls.Add($cancelButton); $cancelButton.Add_Click({ $script:ProcessingCancelled = $true; $cancelButton.Enabled = $false; $cancelButton.Text = "Cancelling..." }); $progressForm.Add_Shown({ $this.Activate() }); return @{ Form = $progressForm; StatusLabel = $statusLabel; ProgressBar = $progressBar; DetailsTextBox = $detailsTextBox; StatsLabel = $statsLabel; CancelButton = $cancelButton } } catch { Write-Log "Error creating progress window: $($_.Exception.Message)" -Level Error; return $null }
}

function Update-ProgressWindow {
    param($ProgressWindow, [string]$Status, [int]$PercentComplete, [string]$Details, [string]$Stats)
    if (-not $ProgressWindow) { return }
    if ($Status) { $ProgressWindow.StatusLabel.Text = $Status }
    if ($PercentComplete -ge 0) { $ProgressWindow.ProgressBar.Value = [Math]::Min($PercentComplete, 100) }
    if ($Details) { $ProgressWindow.DetailsTextBox.AppendText("$Details`r`n"); $ProgressWindow.DetailsTextBox.SelectionStart = $ProgressWindow.DetailsTextBox.Text.Length; $ProgressWindow.DetailsTextBox.ScrollToCaret() }
    if ($Stats) { $ProgressWindow.StatsLabel.Text = $Stats }
    [System.Windows.Forms.Application]::DoEvents()
}
#endregion

#region Core Functions

function Invoke-GraphApiWithRetry {
    param([Parameter(Mandatory=$true)][scriptblock]$ApiCall, [int]$MaxRetries = 5)
    $retryCount = 0
    do {
        try { return & $ApiCall }
        catch [Microsoft.Graph.PowerShell.Models.ErrorRecord] {
            $httpStatus = $_.Exception.HttpStatus
            if ($httpStatus -in @(429, 503) -and $retryCount -lt $MaxRetries) {
                $retryCount++; $retryAfterHeader = $_.Exception.Headers['Retry-After']; $waitTime = 10 * $retryCount
                if ($retryAfterHeader) { $waitTime = [int]$retryAfterHeader[0] }
                Write-Log "Throttled by Graph API (Status: $httpStatus). Waiting for $waitTime seconds before retry #$retryCount..." -Level Warning
                Update-ProgressWindow -ProgressWindow $progressWindow -Details "Throttled. Pausing for $waitTime seconds..."
                Start-Sleep -Seconds $waitTime
            } else { throw }
        } catch { throw }
    } while ($retryCount -lt $MaxRetries)
    throw "Maximum retry limit of $MaxRetries exceeded."
}

function Connect-ToMicrosoftGraph {
    param([string]$TenantId, [string]$AuthMethod = 'DeviceCode')
    Write-Log "Attempting to connect to Microsoft Graph using $AuthMethod authentication"
    try {
        $scopes = @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Application.Read.All", "RoleManagement.Read.Directory", "Policy.Read.All", "Reports.Read.All")
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        switch ($AuthMethod) {
            'DeviceCode'  { $oldInfoPreference = $InformationPreference; try { $InformationPreference = 'Continue'; Write-Host "`nWaiting for Device Code prompt from Microsoft..." -ForegroundColor Cyan; Connect-MgGraph -TenantId $TenantId -Scopes $scopes -UseDeviceCode } finally { $InformationPreference = $oldInfoPreference } }
            'Interactive' { Write-Host "`nStarting interactive browser authentication..." -ForegroundColor Cyan; Connect-MgGraph -TenantId $TenantId -Scopes $scopes }
        }
        Get-MgContext -ErrorAction Stop | Out-Null
        Write-Log "Successfully connected to tenant: $($script:TenantName)" -Level Success
        return $true
    } catch { Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error; $script:Errors += $_; return $false }
}

function Get-UsersToProcess { 
    param($ProgressWindow)
    Write-Log "Starting to collect Users from the tenant"
    Update-ProgressWindow -ProgressWindow $progressWindow -Status "Collecting Users..." -Details "Fetching primary user properties..."
    
    $propertyMap = @{
        'Given Name' = 'GivenName'; 'Surname' = 'Surname'; 'Mail' = 'Mail'; 'Mail Nickname' = 'MailNickname'; 'Other Mails' = 'OtherMails'; 'Proxy Addresses' = 'ProxyAddresses'; 'User Type' = 'UserType'; 'Job Title' = 'JobTitle'; 'Department' = 'Department'; 'Company Name' = 'CompanyName'; 'Office Location' = 'OfficeLocation'; 'Business Phones' = 'BusinessPhones'; 'Mobile Phone' = 'MobilePhone'; 'Street Address' = 'StreetAddress'; 'City' = 'City'; 'State' = 'State'; 'Postal Code' = 'PostalCode'; 'Country' = 'Country'; 'Usage Location' = 'UsageLocation'; 'Account Enabled' = 'AccountEnabled'; 'Creation Date' = 'CreatedDateTime'; 'External User State' = 'ExternalUserState'; 'On-Premises Sync Enabled' = 'OnPremisesSyncEnabled'; 'On-Premises Immutable ID' = 'OnPremisesImmutableId'; 'On-Premises Last Sync' = 'OnPremisesLastSyncDateTime'; 'On-Premises SAM Account Name' = 'OnPremisesSamAccountName'; 'On-Premises UPN' = 'OnPremisesUserPrincipalName'; 'On-Premises Distinguished Name' = 'OnPremisesDistinguishedName'
    }
    
    # FIX: Always fetch the properties needed for the Account Status logic, regardless of user selection.
    $baseProperties = @('Id', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'Surname', 'OfficeLocation')
    $selectedApiProperties = $script:SelectedUserProperties | ForEach-Object { $propertyMap[$_] } | Where-Object { $_ }
    $propertiesToSelect = ($baseProperties + $selectedApiProperties) | Select-Object -Unique

    $users = Invoke-GraphApiWithRetry -ApiCall { Get-MgUser -All -Property $propertiesToSelect }

    Write-Log "Found $($users.Count) total users" -Level Success
    $script:TotalCount = $users.Count
    return $users
}

function Process-User {
    param($User)
    $result = [ordered]@{
        'ID' = $User.Id
        'User Principal Name' = $User.UserPrincipalName ?? "N/A"
        'Display Name' = $User.DisplayName ?? "N/A"
        'Account Status' = 'N/A' 
    }

    foreach ($prop in $script:SelectedUserProperties) {
        $result[$prop] = "N/A"
        switch ($prop) {
            'Given Name'{ $result[$prop] = $User.GivenName ?? "N/A" }; 'Surname'{ $result[$prop] = $User.Surname ?? "N/A" }; 'Mail'{ $result[$prop] = $User.Mail ?? "N/A" }; 'Mail Nickname'{ $result[$prop] = $User.MailNickname ?? "N/A" }; 'Other Mails'{ $result[$prop] = if ($User.OtherMails) { $User.OtherMails -join '; ' } else { "N/A" } }; 'Proxy Addresses'{ $result[$prop] = if ($User.ProxyAddresses) { $User.ProxyAddresses -join '; ' } else { "N/A" } }; 'User Type'{ $result[$prop] = $User.UserType ?? "N/A" }; 'Job Title'{ $result[$prop] = $User.JobTitle ?? "N/A" }; 'Department'{ $result[$prop] = $User.Department ?? "N/A" }; 'Company Name'{ $result[$prop] = $User.CompanyName ?? "N/A" }; 'Office Location'{ $result[$prop] = $User.OfficeLocation ?? "N/A" }; 'Business Phones'{ $result[$prop] = if ($User.BusinessPhones) { $User.BusinessPhones -join '; ' } else { "N/A" } }; 'Mobile Phone'{ $result[$prop] = $User.MobilePhone ?? "N/A" }; 'Street Address'{ $result[$prop] = $User.StreetAddress ?? "N/A" }; 'City'{ $result[$prop] = $User.City ?? "N/A" }; 'State'{ $result[$prop] = $User.State ?? "N/A" }; 'Postal Code'{ $result[$prop] = $User.PostalCode ?? "N/A" }; 'Country'{ $result[$prop] = $User.Country ?? "N/A" }; 'Usage Location'{ $result[$prop] = $User.UsageLocation ?? "N/A" }; 'Account Enabled'{ $result[$prop] = $User.AccountEnabled }; 'Creation Date'{ $result[$prop] = if ($User.CreatedDateTime) { $User.CreatedDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { "N/A" } }; 'External User State'{ $result[$prop] = $User.ExternalUserState ?? "N/A" }; 'On-Premises Sync Enabled'{ $result[$prop] = if ($null -ne $User.OnPremisesSyncEnabled) { $User.OnPremisesSyncEnabled.ToString() } else { "N/A" } }; 'On-Premises Immutable ID'{ $result[$prop] = $User.OnPremisesImmutableId ?? "N/A" }; 'On-Premises Last Sync'{ $result[$prop] = if ($User.OnPremisesLastSyncDateTime) { $User.OnPremisesLastSyncDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { "N/A" } }; 'On-Premises SAM Account Name'{ $result[$prop] = $User.OnPremisesSamAccountName ?? "N/A" }; 'On-Premises UPN'{ $result[$prop] = $User.OnPremisesUserPrincipalName ?? "N/A" }; 'On-Premises Distinguished Name'{ $result[$prop] = $User.OnPremisesDistinguishedName ?? "N/A" }
        }
    }
    
    if ($script:SelectedUserProperties -match 'Last Interactive Sign-in|Last Non-Interactive Sign-in') {
        try {
            $signInActivity = Invoke-GraphApiWithRetry -ApiCall { Get-MgUserSignInActivity -UserId $User.Id }
            if ($signInActivity) {
                if ($script:SelectedUserProperties -contains 'Last Interactive Sign-in') { $result.'Last Interactive Sign-in' = if ($signInActivity.LastSignInDateTime) { $signInActivity.LastSignInDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { "N/A" } }
                if ($script:SelectedUserProperties -contains 'Last Non-Interactive Sign-in') { $result.'Last Non-Interactive Sign-in' = if ($signInActivity.LastNonInteractiveSignInDateTime) { $signInActivity.LastNonInteractiveSignInDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { "N/A" } }
            }
        } catch { Write-Log "Could not get SignInActivity for $($User.DisplayName). (License may be required)" -Level Warning }
    }

    $licenseStatus = "N/A" # Default license status
    if ($script:SelectedUserProperties -contains 'Assigned Licenses') {
        try { $licenses = Invoke-GraphApiWithRetry -ApiCall { Get-MgUserLicenseDetail -UserId $User.Id }; if ($licenses) { $licenseStatus = ($licenses.SkuPartNumber -join '; ') } } catch { Write-Log "Could not get licenses for $($User.DisplayName)" -Level Warning }
        $result.'Assigned Licenses' = $licenseStatus
    }
    
    if ($script:SelectedUserProperties -contains 'Assigned Roles') {
        try { $roles = Invoke-GraphApiWithRetry -ApiCall { Get-MgUserTransitiveMemberOf -UserId $User.Id -All }; $directoryRoles = $roles | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' }; if ($directoryRoles) { $result.'Assigned Roles' = $directoryRoles.AdditionalProperties.displayName -join '; ' } } catch { Write-Log "Could not get roles for $($User.DisplayName)" -Level Warning }
    }
    if ($script:SelectedUserProperties -contains 'MFA Methods') {
        try { $methods = Invoke-GraphApiWithRetry -ApiCall { Get-MgUserAuthenticationMethod -UserId $User.Id }; if ($methods) { $result.'MFA Methods' = ($methods.AdditionalProperties['@odata.type'].Split('.')[-1] -replace 'AuthenticationMethod', '' -join '; ') } } catch { Write-Log "Could not get MFA methods for $($User.DisplayName)" -Level Warning }
    }
    if ($script:SelectedUserProperties -contains 'Manager') {
        try { $manager = Invoke-GraphApiWithRetry -ApiCall { Get-MgUserManager -UserId $User.Id }; if ($manager) { $result.'Manager' = $manager.DisplayName } } catch { }
    }
    if ($script:SelectedUserProperties -contains 'Direct Reports Count') {
        try { Invoke-GraphApiWithRetry -ApiCall { Get-MgUserDirectReport -UserId $User.Id -ConsistencyLevel eventual -CountVariable reportCount } | Out-Null; $result.'Direct Reports Count' = $reportCount } catch { Write-Log "Could not get direct reports for $($User.DisplayName)" -Level Warning }
    }

    # FIX: Enhanced logic to determine Account Status for decommission scenarios.
    if ($User.AccountEnabled) {
        if ($licenseStatus -eq "N/A") { $result.'Account Status' = 'Enabled & Unlicensed' }
        else { $result.'Account Status' = 'Enabled User' }
    } else {
        if ($licenseStatus -eq "N/A" -and [string]::IsNullOrWhiteSpace($User.Surname) -and [string]::IsNullOrWhiteSpace($User.OfficeLocation)) {
            $result.'Account Status' = 'Shared Mailbox (Probable)'
        } else {
            $result.'Account Status' = 'Deactivated User'
        }
    }

    return [PSCustomObject]$result
}

function Get-ServicePrincipalsToProcess { param($ProgressWindow); Write-Log "Starting to collect Service Principals"; Update-ProgressWindow -ProgressWindow $progressWindow -Status "Collecting Service Principals..." -Details "Fetching SP accounts..."; $spProperties = "Id,AppId,DisplayName,ServicePrincipalType,AccountEnabled"; $sps = Invoke-GraphApiWithRetry -ApiCall { Get-MgServicePrincipal -All -Property $spProperties }; Write-Log "Found $($sps.Count) total SPs" -Level Success; $script:TotalCount = $sps.Count; return $sps }
function Process-ServicePrincipal { param($SP); $result = [PSCustomObject]@{ 'Display Name' = $SP.DisplayName; 'Application ID' = $SP.AppId; 'Object ID' = $SP.Id; 'SP Type' = $SP.ServicePrincipalType; 'Account Enabled' = $SP.AccountEnabled; 'Owner(s)' = "N/A"; 'Credential Expirations' = "N/A" }; try { $owners = Invoke-GraphApiWithRetry -ApiCall { Get-MgServicePrincipalOwner -ServicePrincipalId $SP.Id }; if ($owners) { $result.'Owner(s)' = ($owners.AdditionalProperties.displayName -join '; ') } } catch {}; try { $spDetails = Invoke-GraphApiWithRetry -ApiCall { Get-MgServicePrincipal -ServicePrincipalId $SP.Id -Property passwordCredentials, keyCredentials }; $expirations = @(); $spDetails.PasswordCredentials | ForEach-Object { if ($_.EndDateTime) { $expirations += "Password: $($_.EndDateTime.ToString('yyyy-MM-dd'))" } }; $spDetails.KeyCredentials | ForEach-Object { if ($_.EndDateTime) { $expirations += "Certificate: $($_.EndDateTime.ToString('yyyy-MM-dd'))" } }; $result.'Credential Expirations' = $expirations -join '; ' } catch {}; return $result }
function Export-Results { param([array]$Results, [string]$OutputPath, [string]$FilePrefix); try { $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $fileName = "$($FilePrefix)_$($timestamp).csv"; $fullPath = Join-Path -Path $OutputPath -ChildPath $fileName; $Results | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8; Write-Log "Results exported to: $fullPath" -Level Success; return $fullPath } catch { Write-Log "Failed to export results: $($_.Exception.Message)" -Level Error; throw } }
#endregion

#region Main Execution
function Main {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        Tenant-Accounts-Details.ps1 v$ScriptVersion               ║" -ForegroundColor Cyan
    Write-Host "║        Microsoft Tenant Account Details Collector              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    ""
    if (-not (Test-RequiredModules)) { Write-Host "Required modules not available. Exiting." -ForegroundColor Red; return }
    if ((Show-ConfigurationGUI) -ne 'OK') { Write-Host "Operation cancelled by user." -ForegroundColor Yellow; return }
    $tenantNameClean = $script:TenantName -replace '[^a-zA-Z0-9-]', '_'
    $script:LogFile = Join-Path -Path $script:ExportPath -ChildPath "activity_log_$($tenantNameClean).log"
    Write-Log "=== Starting Collection for Principal Type: $($script:PrincipalTypeToProcess) ==="
    if (-not (Connect-ToMicrosoftGraph -TenantId $script:TenantName -AuthMethod $script:AuthMethod)) { [System.Windows.Forms.MessageBox]::Show("Failed to authenticate. Check log for details.", "Auth Failed", 'OK', 'Error'); return }
    $progressWindow = Show-ProgressWindow
    if (-not $progressWindow) { Write-Log "Failed to create progress window. Exiting." -Level Error; return }
    $progressWindow.Form.Show()
    try {
        $principals = @()
        switch ($script:PrincipalTypeToProcess) {
            "Users"                { $principals = Get-UsersToProcess -ProgressWindow $progressWindow }
            "Service Principals"   { $principals = Get-ServicePrincipalsToProcess -ProgressWindow $progressWindow }
        }
        if ($principals.Count -eq 0) { Write-Log "No principals found." -Level Warning; Update-ProgressWindow -ProgressWindow $progressWindow -Status "No principals found."; Start-Sleep -Seconds 3; return }
        $results = @()
        $batchCount = [Math]::Ceiling($principals.Count / 100)
        for ($i = 0; $i -lt $batchCount; $i++) {
            if ($script:ProcessingCancelled) { Write-Log "Processing cancelled by user" -Level Warning; break }
            $startIndex = $i * 100
            $endIndex = [Math]::Min(($startIndex + 100 - 1), ($principals.Count - 1))
            $batch = $principals[$startIndex..$endIndex]
            
            foreach ($principal in $batch) {
                if ($script:ProcessingCancelled) { break }
                $script:ProcessedCount++
                try {
                    if($script:PrincipalTypeToProcess -eq "Users") {
                        $result = Process-User -User $principal
                    } else {
                        $result = Process-ServicePrincipal -SP $principal
                    }
                    if ($result) { $results += $result }
                } catch {
                    Write-Log "Failed to process principal '$($principal.DisplayName)': $($_.Exception.Message)" -Level Error
                    $script:Errors += $_
                }
                $percentComplete = [Math]::Round(($script:ProcessedCount / $script:TotalCount) * 100)
                $elapsed = (Get-Date) - $script:StartTime
                $rate = if ($elapsed.TotalSeconds -gt 0) { $script:ProcessedCount / $elapsed.TotalSeconds } else { 0 }
                $remaining = if ($rate -gt 0) { ($script:TotalCount - $script:ProcessedCount) / $rate } else { 0 }
                $eta = (Get-Date).AddSeconds($remaining)
                $stats = "Processed: $($script:ProcessedCount)/$($script:TotalCount)`nElapsed: $($elapsed.ToString('hh\:mm\:ss'))`nETA: $($eta.ToString('HH:mm:ss'))`nErrors: $($script:Errors.Count)"
                Update-ProgressWindow -ProgressWindow $progressWindow -PercentComplete $percentComplete -Details "Processing: $($principal.DisplayName)" -Stats $stats
            }
        }
        if ($results.Count -gt 0) {
            Update-ProgressWindow -ProgressWindow $progressWindow -Status "Exporting results..." -PercentComplete 95
            $outputFile = Export-Results -Results $results -OutputPath $script:ExportPath -FilePrefix "TenantDetails_$($script:PrincipalTypeToProcess.Replace(' ',''))"
            Update-ProgressWindow -ProgressWindow $progressWindow -Details "Exported results: $outputFile"
            Update-ProgressWindow -ProgressWindow $progressWindow -Status "Collection completed successfully!" -PercentComplete 100
        } else {
            Update-ProgressWindow -ProgressWindow $progressWindow -Status "No results to export" -PercentComplete 100
        }
    } catch {
        $errorMsg = "A critical error occurred during the main processing loop: $($_.Exception.Message)"
        Write-Log $errorMsg -Level Error
        Update-ProgressWindow -ProgressWindow $progressWindow -Status "An error occurred!" -Details $errorMsg
    } finally {
        $summary = "`nSUMMARY:`nTotal Processed: $($script:ProcessedCount)`nErrors: $($script:Errors.Count)`nTime: $(((Get-Date) - $script:StartTime).ToString('g'))`nLog File: $($script:LogFile)"
        Write-Log $summary -Level Success
        Update-ProgressWindow -ProgressWindow $progressWindow -Details $summary
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue; Write-Log "Disconnected from Microsoft Graph" } catch {}
        if ($progressWindow -and $progressWindow.Form) {
            $progressWindow.CancelButton.Visible = $false
            $closeButton = New-Object System.Windows.Forms.Button; $closeButton.Text = "Close"; $closeButton.Location = $progressWindow.CancelButton.Location; $closeButton.Size = $progressWindow.CancelButton.Size
            $closeButton.Add_Click({ $progressWindow.Form.Close() })
            $progressWindow.Form.Controls.Add($closeButton)
            while ($progressWindow.Form.Visible) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100 }
        }
    }
}
Main