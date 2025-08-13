<#
.SYNOPSIS
    A script to disable, delete, or revoke sessions for user accounts based on their email domain.

.DESCRIPTION
    This script provides a graphical user interface (GUI) to specify an email domain.
    It finds all users (Guest or Member) with that domain and allows the administrator
    to either disable their accounts, delete them permanently, or just revoke their active
    sessions. All actions are logged to a specified file.

.NOTES
    Author: Gemini (Enhanced)
    Version: 2.1 (Revised on 2025-08-05)
    - Added "Revoke Sessions Only" option for already disabled accounts
    - Reorganized UI to accommodate three action options
    - Enhanced logic to handle session-only revocation
    - Improved status reporting for different actions
#>

#region Module Check
# Check if the Microsoft.Graph.Users module is installed. If not, prompt the user to install it.
Write-Host "Checking for required PowerShell module 'Microsoft.Graph.Users'..." -ForegroundColor Yellow
if (-not (Get-Module -Name Microsoft.Graph.Users -ListAvailable)) {
    Write-Host "'Microsoft.Graph.Users' module not found." -ForegroundColor Red
    $installChoice = Read-Host "Would you like to try and install it now? (Y/N)"
    if ($installChoice -eq 'Y') {
        try {
            Write-Host "Installing module... This may take a few moments." -ForegroundColor Green
            Install-Module Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            Write-Host "Module installed successfully. Please re-run the script." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install the module. Please install it manually by running 'Install-Module Microsoft.Graph.Users'" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Script cannot continue without the required module." -ForegroundColor Red
    }
    return
}
Write-Host "Module found." -ForegroundColor Green
#endregion

#region GUI Creation
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Guest User Management'
$form.Size = New-Object System.Drawing.Size(460, 420)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Location = '20, 20'
$labelDomain.Size = '280, 20'
$labelDomain.Text = 'Enter the guest email domain: (e.g., example.com)'
$form.Controls.Add($labelDomain)

$textBoxDomain = New-Object System.Windows.Forms.TextBox
$textBoxDomain.Location = '20, 45'
$textBoxDomain.Size = '400, 20'
$form.Controls.Add($textBoxDomain)

$labelLogFile = New-Object System.Windows.Forms.Label
$labelLogFile.Location = '20, 85'
$labelLogFile.Size = '280, 20'
$labelLogFile.Text = 'Log file location:'
$form.Controls.Add($labelLogFile)

$textBoxLogFile = New-Object System.Windows.Forms.TextBox
$textBoxLogFile.Location = '20, 110'
$textBoxLogFile.Size = '315, 20'
$textBoxLogFile.Text = "C:\Temp\GuestUserLog-$(Get-Date -Format 'yyyyMMdd-HHmm').log"
$form.Controls.Add($textBoxLogFile)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = '340, 108'
$buttonBrowse.Size = '80, 25'
$buttonBrowse.Text = 'Browse...'
$buttonBrowse.add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = 'Log Files (*.log)|*.log|All Files (*.*)|*.*'
    $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxLogFile.Text = $saveFileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowse)

$groupBoxAction = New-Object System.Windows.Forms.GroupBox
$groupBoxAction.Location = '20, 150'
$groupBoxAction.Size = '400, 120'
$groupBoxAction.Text = 'Select Action'
$form.Controls.Add($groupBoxAction)

$radioRevokeOnly = New-Object System.Windows.Forms.RadioButton
$radioRevokeOnly.Text = 'Revoke Sessions Only'
$radioRevokeOnly.Location = '20, 25'
$radioRevokeOnly.Size = '360, 20'
$radioRevokeOnly.Checked = $true
$radioRevokeOnly.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular)
$radioRevokeOnly.ForeColor = [System.Drawing.Color]::DarkBlue
$groupBoxAction.Controls.Add($radioRevokeOnly)

$labelRevokeOnly = New-Object System.Windows.Forms.Label
$labelRevokeOnly.Text = '   (Sign out users without changing account status)'
$labelRevokeOnly.Location = '40, 45'
$labelRevokeOnly.Size = '340, 15'
$labelRevokeOnly.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 7.5, [System.Drawing.FontStyle]::Italic)
$labelRevokeOnly.ForeColor = [System.Drawing.Color]::DarkGray
$groupBoxAction.Controls.Add($labelRevokeOnly)

$radioDisable = New-Object System.Windows.Forms.RadioButton
$radioDisable.Text = 'Disable Users + Revoke Sessions'
$radioDisable.Location = '20, 65'
$radioDisable.Size = '200, 20'
$groupBoxAction.Controls.Add($radioDisable)

$radioDelete = New-Object System.Windows.Forms.RadioButton
$radioDelete.Text = 'DELETE Users'
$radioDelete.Location = '20, 90'
$radioDelete.Size = '120, 20'
$radioDelete.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold)
$radioDelete.ForeColor = [System.Drawing.Color]::Red
$groupBoxAction.Controls.Add($radioDelete)

# Additional options group
$groupBoxOptions = New-Object System.Windows.Forms.GroupBox
$groupBoxOptions.Location = '20, 280'
$groupBoxOptions.Size = '400, 50'
$groupBoxOptions.Text = 'Additional Options'
$form.Controls.Add($groupBoxOptions)

$checkBoxSkipDisabled = New-Object System.Windows.Forms.CheckBox
$checkBoxSkipDisabled.Location = '20, 20'
$checkBoxSkipDisabled.Size = '360, 20'
$checkBoxSkipDisabled.Text = 'Skip already disabled accounts (for Disable action only)'
$checkBoxSkipDisabled.Checked = $true
$checkBoxSkipDisabled.Enabled = $false
$groupBoxOptions.Controls.Add($checkBoxSkipDisabled)

# Update checkbox state based on radio selection
$radioRevokeOnly.add_Click({
    $checkBoxSkipDisabled.Enabled = $false
})
$radioDisable.add_Click({
    $checkBoxSkipDisabled.Enabled = $true
})
$radioDelete.add_Click({
    $checkBoxSkipDisabled.Enabled = $false
})

$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Location = '170, 350'
$buttonExecute.Size = '100, 30'
$buttonExecute.Text = 'Execute'
$buttonExecute.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $buttonExecute
$form.Controls.Add($buttonExecute)

$result = $form.ShowDialog()
#endregion

#region Main Logic
if ($result -eq 'OK') {
    $domain = $textBoxDomain.Text.Trim()
    $logFilePath = $textBoxLogFile.Text.Trim()
    
    # Determine action based on radio selection
    if ($radioRevokeOnly.Checked) {
        $action = 'RevokeOnly'
        $actionDescription = 'revoke sessions for'
    }
    elseif ($radioDisable.Checked) {
        $action = 'Disable'
        $actionDescription = 'disable and revoke sessions for'
    }
    else {
        $action = 'Delete'
        $actionDescription = 'delete'
    }
    
    $skipAlreadyDisabled = $checkBoxSkipDisabled.Checked -and $action -eq 'Disable'

    if ([string]::IsNullOrWhiteSpace($domain) -or [string]::IsNullOrWhiteSpace($logFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("Domain and Log File Path cannot be empty.", "Error", "OK", "Error")
        return
    }

    try {
        $logDirectory = [System.IO.Path]::GetDirectoryName($logFilePath)
        if (-not (Test-Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }
        "Log started at $(Get-Date)" | Out-File -FilePath $logFilePath -Append
        "Action: $action | Skip Already Disabled: $skipAlreadyDisabled" | Out-File -FilePath $logFilePath -Append
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Invalid log file path: $($_.Exception.Message)", "Error", "OK", "Error")
        return
    }

    # Build confirmation message
    $confirmationMessage = "⚠️ Are you sure you want to $actionDescription all users with the domain '@$domain'?"
    
    if ($action -eq 'RevokeOnly') {
        $confirmationMessage += "`n`n🔐 This will ONLY revoke active sessions (sign out users) without changing account status."
        $confirmationMessage += "`nThis is useful for already disabled accounts that still have active sessions."
    }
    elseif ($action -eq 'Disable') {
        $confirmationMessage += "`n`n🔐 This will disable accounts AND revoke all active sessions."
        if ($skipAlreadyDisabled) {
            $confirmationMessage += "`nAlready disabled accounts will be skipped."
        }
    }
    elseif ($action -eq 'Delete') {
        $confirmationMessage += "`n`n⚠️ This will PERMANENTLY DELETE the user accounts!"
    }
    
    $confirmationResult = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirm Action", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirmationResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        "User cancelled the operation." | Out-File -FilePath $logFilePath -Append
        [System.Windows.Forms.MessageBox]::Show("Operation cancelled by user.", "Cancelled", "OK", "Information")
        return
    }

    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        
        # Determine required scopes
        $requiredScopes = if ($action -eq 'RevokeOnly') {
            @("User.Read.All", "User.RevokeSessions.All")
        }
        else {
            @("User.ReadWrite.All", "User.RevokeSessions.All")
        }
        
        try {
            # First try to connect with existing session
            $context = Get-MgContext
            if ($context) {
                Write-Host "Found existing connection. Reconnecting with required scopes..." -ForegroundColor Yellow
                Disconnect-MgGraph | Out-Null
            }
            
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
            
            # Verify connection
            $context = Get-MgContext
            if (-not $context) {
                throw "Failed to establish connection to Microsoft Graph"
            }
            
            Write-Host "✅ Successfully connected to Microsoft Graph as '$($context.Account)'." -ForegroundColor Green
            "Successfully connected to Microsoft Graph as '$($context.Account)'" | Out-File -FilePath $logFilePath -Append
        }
        catch {
            Write-Host "Standard authentication failed. Trying device code authentication..." -ForegroundColor Yellow
            try {
                Connect-MgGraph -Scopes $requiredScopes -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
                $context = Get-MgContext
                Write-Host "✅ Successfully connected to Microsoft Graph as '$($context.Account)'." -ForegroundColor Green
                "Successfully connected to Microsoft Graph as '$($context.Account)'" | Out-File -FilePath $logFilePath -Append
            }
            catch {
                throw "Authentication failed: $($_.Exception.Message)"
            }
        }

        # For guest accounts, the domain is embedded in the UPN with #EXT# format
        # Example: john_contoso.com#EXT#@tenant.onmicrosoft.com
        # We need to search for users where UPN contains the domain pattern
        
        Write-Host "Searching for users with domain '@$domain'..." -ForegroundColor Yellow
        Write-Host "Note: Looking for both member and guest accounts" -ForegroundColor Gray
        
        try {
            # For guest accounts, we need to look for the domain within the UPN
            # The domain will appear as domain.com#EXT# or _domain.com#EXT#
            $searchPattern = $domain.Replace(".", "_")
            
            # Try multiple search strategies
            Write-Host "Attempting to retrieve users with advanced filter..." -ForegroundColor Gray
            
            # First try: Get all users and filter for guests with matching domain
            try {
                # Get all users with necessary properties
                Write-Host "Retrieving all users from directory..." -ForegroundColor Gray
                $allUsers = Get-MgUser -All -Property Id,UserPrincipalName,Mail,AccountEnabled,UserType,DisplayName -ErrorAction Stop
                
                Write-Host "Filtering for domain '@$domain'..." -ForegroundColor Gray
                
                # Filter for users that match the domain pattern
                # This handles both guest format (domain#EXT#) and regular member format
                $guestUsers = $allUsers | Where-Object { 
                    # Check for guest account pattern (domain appears before #EXT#)
                    ($_.UserPrincipalName -like "*$searchPattern#EXT#*") -or 
                    ($_.UserPrincipalName -like "*$domain#EXT#*") -or
                    # Also check regular email/UPN ending for member accounts
                    ($_.UserPrincipalName -like "*@$domain") -or
                    ($_.Mail -like "*@$domain")
                }
                
                # Log what we found
                $guestCount = ($guestUsers | Where-Object { $_.UserPrincipalName -like "*#EXT#*" }).Count
                $memberCount = $guestUsers.Count - $guestCount
                
                if ($guestUsers.Count -gt 0) {
                    Write-Host "Found $($guestUsers.Count) user(s): $guestCount guest(s), $memberCount member(s)" -ForegroundColor Green
                    "Found $($guestUsers.Count) total users: $guestCount guests, $memberCount members" | Out-File -FilePath $logFilePath -Append
                }
            }
            catch {
                Write-Host "Failed to retrieve users: $($_.Exception.Message)" -ForegroundColor Red
                throw
            }
        }
        catch {
            $errorMessage = "Failed to search for users: $($_.Exception.Message)"
            Write-Host $errorMessage -ForegroundColor Red
            $errorMessage | Out-File -FilePath $logFilePath -Append
            throw
        }

        if (-not $guestUsers) {
            $msg = "No users found with the domain '@$domain'."
            Write-Host $msg -ForegroundColor Green
            $msg | Out-File -FilePath $logFilePath -Append
            [System.Windows.Forms.MessageBox]::Show($msg, "Complete", "OK", "Information")
            return
        }

        $totalFound = $guestUsers.Count
        $processedCount = 0
        $sessionsRevokedCount = 0
        $skippedCount = 0
        $alreadyDisabledCount = 0
        
        Write-Host "$totalFound user(s) found. Starting the process..." -ForegroundColor Cyan
        "Found $totalFound user(s) to process." | Out-File -FilePath $logFilePath -Append

        foreach ($user in $guestUsers) {
            $userStatus = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            $userType = if ($user.UserPrincipalName -like "*#EXT#*") { "Guest" } else { "Member" }
            
            # Handle different action scenarios
            if ($action -eq 'RevokeOnly') {
                # Revoke sessions for all users regardless of account status
                try {
                    Write-Host "Processing $userType user '$($user.UserPrincipalName)' (Currently: $userStatus)..."
                    Write-Host "  Revoking active sessions..." -ForegroundColor Cyan
                    
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)/revokeSignInSessions" | Out-Null
                    
                    $revokeMessage = "Sessions revoked for $userType '$($user.UserPrincipalName)' (Status: $userStatus) on $(Get-Date)"
                    Write-Host "  ✅ $revokeMessage" -ForegroundColor Green
                    $revokeMessage | Out-File -FilePath $logFilePath -Append
                    $sessionsRevokedCount++
                    $processedCount++
                    
                    if (-not $user.AccountEnabled) {
                        $alreadyDisabledCount++
                    }
                }
                catch {
                    $errorMessage = "❌ ERROR revoking sessions for $userType '$($user.UserPrincipalName)': $($_.Exception.Message)"
                    Write-Host $errorMessage -ForegroundColor Red
                    $errorMessage | Out-File -FilePath $logFilePath -Append
                }
            }
            elseif ($action -eq 'Disable') {
                # Check if already disabled and skip if requested
                if (-not $user.AccountEnabled -and $skipAlreadyDisabled) {
                    $skipMessage = "Skipping $userType user '$($user.UserPrincipalName)' as it is already disabled."
                    Write-Host $skipMessage -ForegroundColor Yellow
                    $skipMessage | Out-File -FilePath $logFilePath -Append
                    $skippedCount++
                    continue
                }
                
                try {
                    Write-Host "Processing $userType user '$($user.UserPrincipalName)' (Currently: $userStatus)..."
                    
                    # First revoke sessions
                    try {
                        Write-Host "  Revoking active sessions..." -ForegroundColor Cyan
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)/revokeSignInSessions" | Out-Null
                        $revokeMessage = "  Sessions revoked for $userType '$($user.UserPrincipalName)'"
                        Write-Host $revokeMessage -ForegroundColor Green
                        $revokeMessage | Out-File -FilePath $logFilePath -Append
                        $sessionsRevokedCount++
                        Start-Sleep -Milliseconds 500
                    }
                    catch {
                        $errorMessage = "  ⚠️ Warning: Could not revoke sessions: $($_.Exception.Message)"
                        Write-Host $errorMessage -ForegroundColor Yellow
                        $errorMessage | Out-File -FilePath $logFilePath -Append
                    }
                    
                    # Then disable if not already disabled
                    if ($user.AccountEnabled) {
                        Update-MgUser -UserId $user.Id -AccountEnabled:$false
                        $logMessage = "$userType user '$($user.UserPrincipalName)' disabled on $(Get-Date)"
                    }
                    else {
                        $logMessage = "$userType user '$($user.UserPrincipalName)' was already disabled, sessions revoked on $(Get-Date)"
                    }
                    
                    Write-Host "  ✅ $logMessage" -ForegroundColor Green
                    $logMessage | Out-File -FilePath $logFilePath -Append
                    $processedCount++
                }
                catch {
                    $errorMessage = "❌ ERROR processing $userType user '$($user.UserPrincipalName)': $($_.Exception.Message)"
                    Write-Host $errorMessage -ForegroundColor Red
                    $errorMessage | Out-File -FilePath $logFilePath -Append
                }
            }
            else { # Delete action
                try {
                    Write-Host "Processing $userType user '$($user.UserPrincipalName)' for deletion..."
                    
                    # Optionally revoke sessions before deletion
                    try {
                        Write-Host "  Revoking active sessions before deletion..." -ForegroundColor Cyan
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)/revokeSignInSessions" | Out-Null
                        $sessionsRevokedCount++
                        Start-Sleep -Milliseconds 500
                    }
                    catch {
                        # Session revocation failure shouldn't stop deletion
                        Write-Host "  ⚠️ Could not revoke sessions, proceeding with deletion..." -ForegroundColor Yellow
                    }
                    
                    Remove-MgUser -UserId $user.Id
                    $logMessage = "$userType user '$($user.UserPrincipalName)' deleted on $(Get-Date)"
                    Write-Host "  ✅ $logMessage" -ForegroundColor Green
                    $logMessage | Out-File -FilePath $logFilePath -Append
                    $processedCount++
                }
                catch {
                    $errorMessage = "❌ ERROR deleting $userType user '$($user.UserPrincipalName)': $($_.Exception.Message)"
                    Write-Host $errorMessage -ForegroundColor Red
                    $errorMessage | Out-File -FilePath $logFilePath -Append
                }
            }
            
            # Add delay to avoid API throttling
            Start-Sleep -Milliseconds 800
        }
        
        # Build final summary message
        $finalMessage = "Process complete.`n`nSummary:"
        $finalMessage += "`n• Total users found: $totalFound"
        
        if ($action -eq 'RevokeOnly') {
            $finalMessage += "`n• Sessions revoked: $sessionsRevokedCount"
            if ($alreadyDisabledCount -gt 0) {
                $finalMessage += "`n• Already disabled accounts processed: $alreadyDisabledCount"
            }
        }
        elseif ($action -eq 'Disable') {
            $finalMessage += "`n• Users processed: $processedCount"
            $finalMessage += "`n• Sessions revoked: $sessionsRevokedCount"
            if ($skippedCount -gt 0) {
                $finalMessage += "`n• Skipped (already disabled): $skippedCount"
            }
        }
        else { # Delete
            $finalMessage += "`n• Users deleted: $processedCount"
            if ($sessionsRevokedCount -gt 0) {
                $finalMessage += "`n• Sessions revoked before deletion: $sessionsRevokedCount"
            }
        }
        
        $finalMessage += "`n`nLog file saved to:`n$logFilePath"
        
        $summaryLog = "`nProcessing complete at $(Get-Date)"
        $summaryLog += "`nTotal: $totalFound | Processed: $processedCount | Sessions Revoked: $sessionsRevokedCount"
        if ($skippedCount -gt 0) {
            $summaryLog += " | Skipped: $skippedCount"
        }
        $summaryLog | Out-File -FilePath $logFilePath -Append
        
        [System.Windows.Forms.MessageBox]::Show($finalMessage, "Success", "OK", "Information")

    }
    catch {
        $errorMessage = "A critical error occurred: $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFilePath -Append
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Critical Error", "OK", "Error")
    }
    finally {
        if (Get-MgContext) {
            Write-Host "Disconnecting from Microsoft Graph."
            Disconnect-MgGraph
        }
    }
}
#endregion