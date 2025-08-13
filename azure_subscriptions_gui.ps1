<#
.SYNOPSIS
    Azure Subscription Management GUI
    
.DESCRIPTION
    A graphical user interface for managing Azure subscriptions, exporting data,
    and managing role assignments across multiple tenants.
    
.NOTES
    Requires: Az.Accounts, Az.Resources, Microsoft.Graph modules
    Windows PowerShell or PowerShell 7 with Windows Forms support
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import required modules function
function Test-RequiredModules {
    $RequiredModules = @('Az.Accounts', 'Az.Resources')
    $MissingModules = @()
    
    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            $MissingModules += $Module
        }
    }
    
    return $MissingModules
}

# Create the main form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Azure Subscription Management Tool"
$Form.Size = New-Object System.Drawing.Size(800, 700)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Header Panel
$HeaderPanel = New-Object System.Windows.Forms.Panel
$HeaderPanel.Size = New-Object System.Drawing.Size(780, 60)
$HeaderPanel.Location = New-Object System.Drawing.Point(10, 10)
$HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)

$HeaderLabel = New-Object System.Windows.Forms.Label
$HeaderLabel.Text = "Azure Subscription Management Tool"
$HeaderLabel.Size = New-Object System.Drawing.Size(760, 30)
$HeaderLabel.Location = New-Object System.Drawing.Point(20, 15)
$HeaderLabel.ForeColor = [System.Drawing.Color]::White
$HeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)

$HeaderSubLabel = New-Object System.Windows.Forms.Label
$HeaderSubLabel.Text = "Export subscriptions, resources, and permissions across multiple tenants"
$HeaderSubLabel.Size = New-Object System.Drawing.Size(760, 15)
$HeaderSubLabel.Location = New-Object System.Drawing.Point(20, 40)
$HeaderSubLabel.ForeColor = [System.Drawing.Color]::White
$HeaderSubLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$HeaderPanel.Controls.Add($HeaderLabel)
$HeaderPanel.Controls.Add($HeaderSubLabel)

# Configuration Panel
$ConfigPanel = New-Object System.Windows.Forms.GroupBox
$ConfigPanel.Text = "Configuration"
$ConfigPanel.Size = New-Object System.Drawing.Size(760, 180)
$ConfigPanel.Location = New-Object System.Drawing.Point(20, 80)

# Tenant ID
$TenantLabel = New-Object System.Windows.Forms.Label
$TenantLabel.Text = "Tenant ID:"
$TenantLabel.Size = New-Object System.Drawing.Size(120, 20)
$TenantLabel.Location = New-Object System.Drawing.Point(20, 30)

$TenantTextBox = New-Object System.Windows.Forms.TextBox
$TenantTextBox.Size = New-Object System.Drawing.Size(300, 25)
$TenantTextBox.Location = New-Object System.Drawing.Point(140, 28)
$TenantTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)

$TenantHelpLabel = New-Object System.Windows.Forms.Label
$TenantHelpLabel.Text = "Required: The Azure tenant ID to connect to"
$TenantHelpLabel.Size = New-Object System.Drawing.Size(280, 15)
$TenantHelpLabel.Location = New-Object System.Drawing.Point(450, 32)
$TenantHelpLabel.ForeColor = [System.Drawing.Color]::Gray
$TenantHelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# Subscription ID
$SubscriptionLabel = New-Object System.Windows.Forms.Label
$SubscriptionLabel.Text = "Subscription ID:"
$SubscriptionLabel.Size = New-Object System.Drawing.Size(120, 20)
$SubscriptionLabel.Location = New-Object System.Drawing.Point(20, 60)

$SubscriptionTextBox = New-Object System.Windows.Forms.TextBox
$SubscriptionTextBox.Size = New-Object System.Drawing.Size(300, 25)
$SubscriptionTextBox.Location = New-Object System.Drawing.Point(140, 58)
$SubscriptionTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)

$SubscriptionHelpLabel = New-Object System.Windows.Forms.Label
$SubscriptionHelpLabel.Text = "Optional: Target specific subscription (leave blank for all)"
$SubscriptionHelpLabel.Size = New-Object System.Drawing.Size(280, 15)
$SubscriptionHelpLabel.Location = New-Object System.Drawing.Point(450, 62)
$SubscriptionHelpLabel.ForeColor = [System.Drawing.Color]::Gray
$SubscriptionHelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# Admin User
$AdminUserLabel = New-Object System.Windows.Forms.Label
$AdminUserLabel.Text = "Admin User:"
$AdminUserLabel.Size = New-Object System.Drawing.Size(120, 20)
$AdminUserLabel.Location = New-Object System.Drawing.Point(20, 90)

$AdminUserTextBox = New-Object System.Windows.Forms.TextBox
$AdminUserTextBox.Size = New-Object System.Drawing.Size(300, 25)
$AdminUserTextBox.Location = New-Object System.Drawing.Point(140, 88)
$AdminUserTextBox.Text = "wppadmin@SETPDX.onmicrosoft.com"

$AdminUserHelpLabel = New-Object System.Windows.Forms.Label
$AdminUserHelpLabel.Text = "User to grant Reader access (if role assignment enabled)"
$AdminUserHelpLabel.Size = New-Object System.Drawing.Size(280, 15)
$AdminUserHelpLabel.Location = New-Object System.Drawing.Point(450, 92)
$AdminUserHelpLabel.ForeColor = [System.Drawing.Color]::Gray
$AdminUserHelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# Output Path
$OutputLabel = New-Object System.Windows.Forms.Label
$OutputLabel.Text = "Output Path:"
$OutputLabel.Size = New-Object System.Drawing.Size(120, 20)
$OutputLabel.Location = New-Object System.Drawing.Point(20, 120)

$OutputTextBox = New-Object System.Windows.Forms.TextBox
$OutputTextBox.Size = New-Object System.Drawing.Size(250, 25)
$OutputTextBox.Location = New-Object System.Drawing.Point(140, 118)
$OutputTextBox.Text = [Environment]::CurrentDirectory

$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Browse..."
$BrowseButton.Size = New-Object System.Drawing.Size(75, 25)
$BrowseButton.Location = New-Object System.Drawing.Point(400, 118)

# Skip Role Assignment Checkbox
$SkipRoleCheckBox = New-Object System.Windows.Forms.CheckBox
$SkipRoleCheckBox.Text = "Skip Role Assignment (Recommended)"
$SkipRoleCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$SkipRoleCheckBox.Location = New-Object System.Drawing.Point(140, 150)
$SkipRoleCheckBox.Checked = $true

# Add controls to Configuration Panel
$ConfigPanel.Controls.AddRange(@(
    $TenantLabel, $TenantTextBox, $TenantHelpLabel,
    $SubscriptionLabel, $SubscriptionTextBox, $SubscriptionHelpLabel,
    $AdminUserLabel, $AdminUserTextBox, $AdminUserHelpLabel,
    $OutputLabel, $OutputTextBox, $BrowseButton,
    $SkipRoleCheckBox
))

# Action Panel
$ActionPanel = New-Object System.Windows.Forms.GroupBox
$ActionPanel.Text = "Actions"
$ActionPanel.Size = New-Object System.Drawing.Size(760, 60)
$ActionPanel.Location = New-Object System.Drawing.Point(20, 270)

$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.Text = "🔐 Connect to Azure"
$ConnectButton.Size = New-Object System.Drawing.Size(150, 30)
$ConnectButton.Location = New-Object System.Drawing.Point(20, 20)
$ConnectButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$ConnectButton.ForeColor = [System.Drawing.Color]::White
$ConnectButton.FlatStyle = "Flat"

$RunButton = New-Object System.Windows.Forms.Button
$RunButton.Text = "▶️ Run Export"
$RunButton.Size = New-Object System.Drawing.Size(150, 30)
$RunButton.Location = New-Object System.Drawing.Point(180, 20)
$RunButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$RunButton.ForeColor = [System.Drawing.Color]::White
$RunButton.FlatStyle = "Flat"
$RunButton.Enabled = $false

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Text = "⏹️ Cancel"
$CancelButton.Size = New-Object System.Drawing.Size(100, 30)
$CancelButton.Location = New-Object System.Drawing.Point(340, 20)
$CancelButton.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
$CancelButton.ForeColor = [System.Drawing.Color]::White
$CancelButton.FlatStyle = "Flat"
$CancelButton.Enabled = $false

$DisconnectButton = New-Object System.Windows.Forms.Button
$DisconnectButton.Text = "🔓 Disconnect"
$DisconnectButton.Size = New-Object System.Drawing.Size(120, 30)
$DisconnectButton.Location = New-Object System.Drawing.Point(450, 20)
$DisconnectButton.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)
$DisconnectButton.ForeColor = [System.Drawing.Color]::White
$DisconnectButton.FlatStyle = "Flat"
$DisconnectButton.Enabled = $false

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready - Connect to Azure to begin"
$StatusLabel.Size = New-Object System.Drawing.Size(180, 20)
$StatusLabel.Location = New-Object System.Drawing.Point(580, 25)
$StatusLabel.ForeColor = [System.Drawing.Color]::Gray

$ActionPanel.Controls.AddRange(@($ConnectButton, $RunButton, $CancelButton, $DisconnectButton, $StatusLabel))

# Progress Panel
$ProgressPanel = New-Object System.Windows.Forms.GroupBox
$ProgressPanel.Text = "Progress"
$ProgressPanel.Size = New-Object System.Drawing.Size(760, 80)
$ProgressPanel.Location = New-Object System.Drawing.Point(20, 340)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Size = New-Object System.Drawing.Size(720, 25)
$ProgressBar.Location = New-Object System.Drawing.Point(20, 25)
$ProgressBar.Style = "Continuous"

$ProgressLabel = New-Object System.Windows.Forms.Label
$ProgressLabel.Text = "Waiting to start..."
$ProgressLabel.Size = New-Object System.Drawing.Size(720, 20)
$ProgressLabel.Location = New-Object System.Drawing.Point(20, 55)
$ProgressLabel.ForeColor = [System.Drawing.Color]::Blue

$ProgressPanel.Controls.AddRange(@($ProgressBar, $ProgressLabel))

# Log Panel
$LogPanel = New-Object System.Windows.Forms.GroupBox
$LogPanel.Text = "Execution Log"
$LogPanel.Size = New-Object System.Drawing.Size(760, 200)
$LogPanel.Location = New-Object System.Drawing.Point(20, 430)

$LogTextBox = New-Object System.Windows.Forms.RichTextBox
$LogTextBox.Size = New-Object System.Drawing.Size(720, 170)
$LogTextBox.Location = New-Object System.Drawing.Point(20, 20)
$LogTextBox.ReadOnly = $true
$LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$LogTextBox.BackColor = [System.Drawing.Color]::Black
$LogTextBox.ForeColor = [System.Drawing.Color]::White

$LogPanel.Controls.Add($LogTextBox)

# Add all panels to form
$Form.Controls.AddRange(@($HeaderPanel, $ConfigPanel, $ActionPanel, $ProgressPanel, $LogPanel))

# Global variables
$Script:IsConnected = $false
$Script:CurrentContext = $null
$Script:CancelRequested = $false

# Event Handlers

# Browse Button Click
$BrowseButton.Add_Click({
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select output folder for CSV files"
    $FolderDialog.SelectedPath = $OutputTextBox.Text
    
    if ($FolderDialog.ShowDialog() -eq "OK") {
        $OutputTextBox.Text = $FolderDialog.SelectedPath
    }
})

# Function to add log entry
function Add-LogEntry {
    param(
        [string]$Message,
        [string]$Type = "Info"  # Info, Success, Warning, Error
    )
    
    $Timestamp = Get-Date -Format "HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message`r`n"
    
    $LogTextBox.SelectionStart = $LogTextBox.TextLength
    
    switch ($Type) {
        "Success" { $LogTextBox.SelectionColor = [System.Drawing.Color]::Green }
        "Warning" { $LogTextBox.SelectionColor = [System.Drawing.Color]::Yellow }
        "Error" { $LogTextBox.SelectionColor = [System.Drawing.Color]::Red }
        default { $LogTextBox.SelectionColor = [System.Drawing.Color]::White }
    }
    
    $LogTextBox.AppendText($LogEntry)
    $LogTextBox.SelectionColor = [System.Drawing.Color]::White
    $LogTextBox.ScrollToCaret()
    $Form.Refresh()
}

# Function to update progress
function Update-Progress {
    param(
        [int]$Percentage,
        [string]$Message
    )
    
    $ProgressBar.Value = [Math]::Min(100, [Math]::Max(0, $Percentage))
    $ProgressLabel.Text = $Message
    $Form.Refresh()
}

# Connect Button Click
$ConnectButton.Add_Click({
    if (-not $TenantTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Tenant ID", "Missing Information", "OK", "Warning")
        return
    }
    
    try {
        $ConnectButton.Enabled = $false
        $StatusLabel.Text = "Connecting to Azure..."
        $StatusLabel.ForeColor = [System.Drawing.Color]::Blue
        Add-LogEntry "Connecting to Azure tenant: $($TenantTextBox.Text)"
        
        # Check modules first
        $MissingModules = Test-RequiredModules
        if ($MissingModules.Count -gt 0) {
            Add-LogEntry "Missing required modules: $($MissingModules -join ', ')" "Error"
            [System.Windows.Forms.MessageBox]::Show("Missing required modules: $($MissingModules -join ', ')`n`nPlease install these modules first.", "Missing Modules", "OK", "Error")
            return
        }
        
        # Import modules
        Add-LogEntry "Loading Azure modules..."
        $Form.Refresh()
        Import-Module Az.Accounts -Force
        Import-Module Az.Resources -Force
        
        # Show browser authentication guidance
        Add-LogEntry "Opening browser for authentication..." "Info"
        Add-LogEntry "⚠ If browser opens with wrong profile, close it and click 'Retry Connection'" "Warning"
        $Form.Refresh()
        
        # Connect to Azure with timeout and better error handling
        Add-LogEntry "Waiting for browser authentication to complete..."
        $Form.Refresh()
        
        if ($TenantTextBox.Text.Trim()) {
            $ConnectionResult = Connect-AzAccount -TenantId $TenantTextBox.Text.Trim() -ErrorAction Stop
        } else {
            $ConnectionResult = Connect-AzAccount -ErrorAction Stop
        }
        
        # Verify connection was successful
        if (-not $ConnectionResult) {
            throw "Authentication was cancelled or failed"
        }
        
        $Script:CurrentContext = Get-AzContext
        if (-not $Script:CurrentContext) {
            throw "Failed to establish Azure context after authentication"
        }
        
        $Script:IsConnected = $true
        
        $StatusLabel.Text = "Connected: $($Script:CurrentContext.Account.Id)"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Green
        $RunButton.Enabled = $true
        $DisconnectButton.Enabled = $true
        $ConnectButton.Text = "🔄 Retry Connection"
        
        Add-LogEntry "✅ Successfully connected as: $($Script:CurrentContext.Account.Id)" "Success"
        Add-LogEntry "✅ Tenant: $($Script:CurrentContext.Tenant.Id)" "Success"
        Add-LogEntry "Ready to run export operations!" "Info"
        
    } catch [Microsoft.Azure.PowerShell.Authenticators.AuthenticationFailedException] {
        # Handle authentication specific failures
        Add-LogEntry "❌ Authentication failed or was cancelled" "Error"
        Add-LogEntry "This usually happens when:" "Warning"
        Add-LogEntry "  • Browser opened with wrong user profile" "Warning"
        Add-LogEntry "  • Authentication window was closed before completing" "Warning"
        Add-LogEntry "  • Network connectivity issues" "Warning"
        
        $StatusLabel.Text = "Authentication failed"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
        
        $Result = [System.Windows.Forms.MessageBox]::Show(
            "Azure authentication failed or was cancelled.`n`n" +
            "Common causes:`n" +
            "• Browser opened with wrong user profile`n" +
            "• Authentication window was closed early`n" +
            "• Network connectivity issues`n`n" +
            "Solutions:`n" +
            "• Ensure you're using the correct browser profile`n" +
            "• Complete the full authentication process`n" +
            "• Check your network connection`n`n" +
            "Would you like to try connecting again?",
            "Authentication Failed", 
            "YesNo", 
            "Question"
        )
        
        if ($Result -eq "Yes") {
            Add-LogEntry "🔄 User requested retry - attempting connection again..." "Info"
            # Clear any existing context
            try { Disconnect-AzAccount -ErrorAction SilentlyContinue } catch { }
            $Script:IsConnected = $false
            $Script:CurrentContext = $null
            # Recursively call the connect function
            $ConnectButton.PerformClick()
            return
        } else {
            Add-LogEntry "User chose not to retry authentication" "Warning"
        }
        
    } catch [System.Management.Automation.CommandNotFoundException] {
        Add-LogEntry "❌ Azure PowerShell modules not properly loaded" "Error"
        $StatusLabel.Text = "Module error"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "Azure PowerShell modules are not properly installed or loaded.`n`n" +
            "Please install the required modules:`n" +
            "Install-Module -Name Az.Accounts -Force`n" +
            "Install-Module -Name Az.Resources -Force",
            "Missing Modules", 
            "OK", 
            "Error"
        )
        
    } catch [System.OperationCanceledException] {
        Add-LogEntry "❌ Authentication was cancelled by user" "Warning"
        $StatusLabel.Text = "Authentication cancelled"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Orange
        
        $Result = [System.Windows.Forms.MessageBox]::Show(
            "Authentication was cancelled.`n`n" +
            "To successfully connect:`n" +
            "• Allow the browser to fully load`n" +
            "• Sign in with the correct account`n" +
            "• Complete the entire authentication process`n`n" +
            "Try again?",
            "Authentication Cancelled", 
            "YesNo", 
            "Question"
        )
        
        if ($Result -eq "Yes") {
            Add-LogEntry "🔄 Retrying authentication..." "Info"
            $ConnectButton.PerformClick()
            return
        }
        
    } catch [System.TimeoutException] {
        Add-LogEntry "❌ Authentication timed out" "Error"
        $StatusLabel.Text = "Connection timeout"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "Authentication timed out.`n`n" +
            "This may be due to:`n" +
            "• Slow network connection`n" +
            "• Browser not responding`n" +
            "• Firewall blocking authentication`n`n" +
            "Please try again with a stable network connection.",
            "Connection Timeout", 
            "OK", 
            "Warning"
        )
        
    } catch {
        # Handle any other unexpected errors
        $ErrorMessage = $_.Exception.Message
        Add-LogEntry "❌ Unexpected connection error: $ErrorMessage" "Error"
        $StatusLabel.Text = "Connection error"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
        
        # Check for specific known error patterns
        if ($ErrorMessage -match "AADSTS|AAD" -or $ErrorMessage -match "tenant") {
            $SuggestedAction = "Please verify the Tenant ID is correct and you have access to this tenant."
        } elseif ($ErrorMessage -match "browser|web") {
            $SuggestedAction = "Browser-related issue. Try using a different browser or clearing browser cache."
        } elseif ($ErrorMessage -match "network|connection") {
            $SuggestedAction = "Network connectivity issue. Check your internet connection and firewall settings."
        } else {
            $SuggestedAction = "Please check the error details in the log and try again."
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Connection failed with error:`n$ErrorMessage`n`n" +
            "Suggested action:`n$SuggestedAction`n`n" +
            "Check the log for more details.",
            "Connection Error", 
            "OK", 
            "Error"
        )
    } finally {
        $ConnectButton.Enabled = $true
        if (-not $Script:IsConnected) {
            $ConnectButton.Text = "🔐 Connect to Azure"
            $StatusLabel.Text = "Ready - Click Connect to Azure"
            $StatusLabel.ForeColor = [System.Drawing.Color]::Gray
            $RunButton.Enabled = $false
            $DisconnectButton.Enabled = $false
        }
        $Form.Refresh()
    }
})

# Disconnect Button Click
$DisconnectButton.Add_Click({
    try {
        Add-LogEntry "Disconnecting from Azure..." "Info"
        Disconnect-AzAccount -ErrorAction Stop | Out-Null
        
        $Script:IsConnected = $false
        $Script:CurrentContext = $null
        
        $StatusLabel.Text = "Disconnected"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Gray
        $RunButton.Enabled = $false
        $DisconnectButton.Enabled = $false
        $ConnectButton.Text = "🔐 Connect to Azure"
        
        Add-LogEntry "✅ Successfully disconnected from Azure" "Success"
        
    } catch {
        Add-LogEntry "⚠ Disconnect error (this is usually harmless): $($_.Exception.Message)" "Warning"
        # Even if disconnect fails, reset the UI state
        $Script:IsConnected = $false
        $Script:CurrentContext = $null
        $StatusLabel.Text = "Disconnected"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Gray
        $RunButton.Enabled = $false
        $DisconnectButton.Enabled = $false
        $ConnectButton.Text = "🔐 Connect to Azure"
    }
})

# Run Button Click
$RunButton.Add_Click({
    if (-not $Script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Azure first", "Not Connected", "OK", "Warning")
        return
    }
    
    $Script:CancelRequested = $false
    $RunButton.Enabled = $false
    $CancelButton.Enabled = $true
    $ConnectButton.Enabled = $false
    $DisconnectButton.Enabled = $false
    
    try {
        Add-LogEntry "Starting Azure subscription export process..." "Info"
        Update-Progress 0 "Initializing..."
        
        # Get parameters
        $TenantId = $TenantTextBox.Text.Trim()
        $SubscriptionId = if ($SubscriptionTextBox.Text.Trim()) { $SubscriptionTextBox.Text.Trim() } else { $null }
        $AdminUser = $AdminUserTextBox.Text.Trim()
        $OutputPath = $OutputTextBox.Text.Trim()
        $SkipRoleAssignment = $SkipRoleCheckBox.Checked
        
        # Validate output path
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Add-LogEntry "Created output directory: $OutputPath" "Info"
        }
        
        # Discover subscriptions
        Update-Progress 10 "Discovering subscriptions..."
        Add-LogEntry "Discovering Azure subscriptions..."
        
        if ($SubscriptionId) {
            $AzureSubscriptions = @(Get-AzSubscription -SubscriptionId $SubscriptionId)
            Add-LogEntry "Target subscription: $($AzureSubscriptions[0].Name)" "Info"
        } else {
            $AzureSubscriptions = Get-AzSubscription
            Add-LogEntry "Found $($AzureSubscriptions.Count) subscription(s)" "Info"
        }
        
        # Filter enabled subscriptions
        $EnabledSubscriptions = $AzureSubscriptions | Where-Object { $_.State -eq 'Enabled' }
        if ($EnabledSubscriptions.Count -ne $AzureSubscriptions.Count) {
            Add-LogEntry "Processing $($EnabledSubscriptions.Count) enabled subscriptions (skipping $($AzureSubscriptions.Count - $EnabledSubscriptions.Count) disabled)" "Warning"
            $AzureSubscriptions = $EnabledSubscriptions
        }
        
        if ($Script:CancelRequested) { return }
        
        # Export subscriptions list
        Update-Progress 20 "Exporting subscriptions list..."
        Add-LogEntry "Exporting subscriptions list..."
        
        $TenantIdForFile = $Script:CurrentContext.Tenant.Id
        if ($SubscriptionId) {
            $SubscriptionsFileName = "Azure_Subscriptions_$TenantIdForFile" + "_Sub_$SubscriptionId.csv"
        } else {
            $SubscriptionsFileName = "Azure_Subscriptions_$TenantIdForFile.csv"
        }
        
        $SubscriptionsPath = Join-Path $OutputPath $SubscriptionsFileName
        $AzureSubscriptions | Select-Object Id, Name, State, TenantId | Export-Csv $SubscriptionsPath -NoTypeInformation
        Add-LogEntry "✓ Subscriptions exported to: $SubscriptionsFileName" "Success"
        
        if ($Script:CancelRequested) { return }
        
        # Role Assignment (if not skipped)
        $StepSize = if ($SkipRoleAssignment) { 40 } else { 25 }
        $CurrentProgress = 30
        
        if (-not $SkipRoleAssignment) {
            Update-Progress $CurrentProgress "Processing role assignments..."
            Add-LogEntry "Granting Reader access to: $AdminUser"
            
            foreach ($Subscription in $AzureSubscriptions) {
                if ($Script:CancelRequested) { return }
                
                try {
                    Set-AzContext -SubscriptionId $Subscription.Id -TenantId $Subscription.TenantId | Out-Null
                    
                    $ExistingAssignment = Get-AzRoleAssignment -SignInName $AdminUser -RoleDefinitionName "Reader" -Scope "/subscriptions/$($Subscription.Id)" -ErrorAction SilentlyContinue
                    
                    if (-not $ExistingAssignment) {
                        New-AzRoleAssignment -SignInName $AdminUser -RoleDefinitionName "Reader" -Scope "/subscriptions/$($Subscription.Id)" | Out-Null
                        Add-LogEntry "✓ Granted Reader access for: $($Subscription.Name)" "Success"
                    } else {
                        Add-LogEntry "⚠ Reader role already exists for: $($Subscription.Name)" "Warning"
                    }
                } catch {
                    Add-LogEntry "✗ Role assignment failed for $($Subscription.Name): $($_.Exception.Message)" "Error"
                }
            }
            $CurrentProgress += $StepSize
        } else {
            Add-LogEntry "⏭ Skipping role assignment (as requested)" "Info"
        }
        
        if ($Script:CancelRequested) { return }
        
        # Export Resources
        Update-Progress $CurrentProgress "Exporting resources..."
        Add-LogEntry "Exporting Azure resources..."
        
        $ExportArray = @()
        $ResourceCount = 0
        
        foreach ($Subscription in $AzureSubscriptions) {
            if ($Script:CancelRequested) { return }
            
            try {
                Add-LogEntry "Processing resources for: $($Subscription.Name)"
                Set-AzContext -SubscriptionId $Subscription.Id -TenantId $Subscription.TenantId | Out-Null
                
                $Resources = Get-AzResource
                foreach ($Resource in $Resources) {
                    $ExportArray += $Resource | Select-Object *, @{Name = 'SubscriptionName'; Expression = { $Subscription.Name } }
                }
                
                $ResourceCount += $Resources.Count
                Add-LogEntry "✓ Found $($Resources.Count) resources in $($Subscription.Name)" "Success"
            } catch {
                Add-LogEntry "✗ Resource export failed for $($Subscription.Name): $($_.Exception.Message)" "Error"
            }
        }
        
        if ($SubscriptionId) {
            $ResourcesFileName = "Azure_Resources_$TenantIdForFile" + "_Sub_$SubscriptionId.csv"
        } else {
            $ResourcesFileName = "Azure_Resources_$TenantIdForFile.csv"
        }
        
        $ResourcesPath = Join-Path $OutputPath $ResourcesFileName
        $ExportArray | Export-Csv $ResourcesPath -NoTypeInformation
        Add-LogEntry "✓ $ResourceCount resources exported to: $ResourcesFileName" "Success"
        
        $CurrentProgress += $StepSize
        if ($Script:CancelRequested) { return }
        
        # Export Permissions
        Update-Progress $CurrentProgress "Exporting permissions..."
        Add-LogEntry "Exporting role assignments and permissions..."
        
        $PermissionsArray = @()
        $PermissionCount = 0
        
        foreach ($Subscription in $AzureSubscriptions) {
            if ($Script:CancelRequested) { return }
            
            try {
                Add-LogEntry "Processing permissions for: $($Subscription.Name)"
                Set-AzContext -SubscriptionId $Subscription.Id -TenantId $Subscription.TenantId | Out-Null
                
                $Roles = Get-AzRoleAssignment -IncludeClassicAdministrators
                
                foreach ($Role in $Roles) {
                    $PermissionDetails = [PSCustomObject]@{
                        SubscriptionID     = $Subscription.Id
                        SubscriptionName   = $Subscription.Name
                        ObjectType         = $Role.ObjectType
                        RoleDefinitionName = $Role.RoleDefinitionName
                        DisplayName        = $Role.DisplayName
                        SignInName         = $Role.SignInName
                        ObjectId           = $Role.ObjectId
                        Scope              = $Role.Scope
                        CreatedOn          = $Role.CreatedOn
                        CreatedBy          = $Role.CreatedBy
                    }
                    $PermissionsArray += $PermissionDetails
                }
                
                $PermissionCount += $Roles.Count
                Add-LogEntry "✓ Found $($Roles.Count) role assignments in $($Subscription.Name)" "Success"
            } catch {
                Add-LogEntry "✗ Permission export failed for $($Subscription.Name): $($_.Exception.Message)" "Error"
            }
        }
        
        if ($SubscriptionId) {
            $PermissionsFileName = "Azure_Permissions_$TenantIdForFile" + "_Sub_$SubscriptionId.csv"
        } else {
            $PermissionsFileName = "Azure_Permissions_$TenantIdForFile.csv"
        }
        
        $PermissionsPath = Join-Path $OutputPath $PermissionsFileName
        $PermissionsArray | Export-Csv $PermissionsPath -NoTypeInformation
        Add-LogEntry "✓ $PermissionCount permissions exported to: $PermissionsFileName" "Success"
        
        # Complete
        Update-Progress 100 "Export completed successfully!"
        Add-LogEntry "🎉 Export process completed successfully!" "Success"
        Add-LogEntry "📁 Output location: $OutputPath" "Info"
        
        [System.Windows.Forms.MessageBox]::Show("Export completed successfully!`n`nFiles saved to: $OutputPath", "Export Complete", "OK", "Information")
        
    } catch {
        Add-LogEntry "💥 Export failed: $($_.Exception.Message)" "Error"
        Update-Progress 0 "Export failed"
        [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Export Failed", "OK", "Error")
    } finally {
        $RunButton.Enabled = $true
        $CancelButton.Enabled = $false
        $ConnectButton.Enabled = $true
        $DisconnectButton.Enabled = $true
        $Script:CancelRequested = $false
    }
})

# Cancel Button Click
$CancelButton.Add_Click({
    $Script:CancelRequested = $true
    Add-LogEntry "🛑 Cancellation requested by user" "Warning"
    Update-Progress 0 "Cancelling..."
})

# Form closing event
$Form.Add_FormClosing({
    if ($Script:IsConnected) {
        $Result = [System.Windows.Forms.MessageBox]::Show("Do you want to disconnect from Azure before closing?", "Disconnect from Azure", "YesNoCancel", "Question")
        if ($Result -eq "Cancel") {
            $_.Cancel = $true
            return
        }
        if ($Result -eq "Yes") {
            try {
                Disconnect-AzAccount -ErrorAction SilentlyContinue
                Add-LogEntry "Disconnected from Azure" "Info"
            } catch {
                # Ignore disconnect errors
            }
        }
    }
})

# Check modules on startup
$MissingModules = Test-RequiredModules
if ($MissingModules.Count -gt 0) {
    Add-LogEntry "⚠ Missing required modules detected:" "Warning"
    foreach ($Module in $MissingModules) {
        Add-LogEntry "  - $Module" "Warning"
    }
    Add-LogEntry "Please install missing modules before connecting to Azure" "Warning"
}

Add-LogEntry "🚀 Azure Subscription Management Tool loaded" "Info"
Add-LogEntry "💡 Authentication Tip: If browser opens with wrong profile, close it and retry" "Info"
Add-LogEntry "Ready to connect to Azure tenant..." "Info"

# Show the form
[System.Windows.Forms.Application]::EnableVisualStyles()
$Form.ShowDialog()
