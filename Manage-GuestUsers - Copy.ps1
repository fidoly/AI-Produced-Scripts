<#
.SYNOPSIS
    A script to disable or delete user accounts based on their email domain.

.DESCRIPTION
    This script provides a graphical user interface (GUI) to specify an email domain.
    It finds all users (Guest or Member) with that domain and allows the administrator
    to either disable their accounts or delete them permanently. All actions are logged
    to a specified file.

.NOTES
    Author: Gemini
    Version: 1.5 (Revised on 2025-07-31)
    - Added a counter for processed accounts.
    - Added logic to skip users that are already disabled.
    - Added a minor delay to each loop to mitigate API throttling.
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
$form.Size = New-Object System.Drawing.Size(460, 320)
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
$groupBoxAction.Size = '400, 60'
$groupBoxAction.Text = 'Select Action'
$form.Controls.Add($groupBoxAction)

$radioDisable = New-Object System.Windows.Forms.RadioButton
$radioDisable.Text = 'Disable Users'
$radioDisable.Location = '20, 25'
$radioDisable.Size = '120, 20'
$radioDisable.Checked = $true
$groupBoxAction.Controls.Add($radioDisable)

$radioDelete = New-Object System.Windows.Forms.RadioButton
$radioDelete.Text = 'DELETE Users'
$radioDelete.Location = '240, 25'
$radioDelete.Size = '120, 20'
$radioDelete.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold)
$radioDelete.ForeColor = [System.Drawing.Color]::Red
$groupBoxAction.Controls.Add($radioDelete)

$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Location = '170, 230'
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
    $action = if ($radioDisable.Checked) { 'Disable' } else { 'Delete' }
    $actionVerb = if ($action -eq 'Disable') { 'disabled' } else { 'deleted' }

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
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Invalid log file path: $($_.Exception.Message)", "Error", "OK", "Error")
        return
    }

    $confirmationMessage = "⚠️ Are you sure you want to $($action.ToUpper()) all users with the domain '@$domain'?"
    $confirmationResult = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirm Action", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirmationResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        "User cancelled the operation." | Out-File -FilePath $logFilePath -Append
        [System.Windows.Forms.MessageBox]::Show("Operation cancelled by user.", "Cancelled", "OK", "Information")
        return
    }

    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.ReadWrite.All" -UseDeviceAuthentication
        Write-Host "✅ Successfully connected to Microsoft Graph as '$((Get-MgContext).Account)'." -ForegroundColor Green
        "Successfully connected to Microsoft Graph as '$((Get-MgContext).Account)'" | Out-File -FilePath $logFilePath -Append

        $filterString = "endsWith(mail, '@$domain')"
        
        Write-Host "Searching for users with domain '@$domain'..."
        # Add -Count and -Headers to support the 'endsWith' advanced query operator.
        $guestUsers = Get-MgUser -Filter $filterString -Count "userCount" -Headers @{'ConsistencyLevel'='eventual'} -All -ErrorAction Stop

        if (-not $guestUsers) {
            $msg = "No users found with the domain '@$domain'."
            Write-Host $msg -ForegroundColor Green
            $msg | Out-File -FilePath $logFilePath -Append
            [System.Windows.Forms.MessageBox]::Show($msg, "Complete", "OK", "Information")
            return
        }

        $totalFound = $guestUsers.Count
        $processedCount = 0
        Write-Host "$totalFound user(s) found. Starting the process..." -ForegroundColor Cyan
        "Found $totalFound user(s) to process." | Out-File -FilePath $logFilePath -Append

        foreach ($user in $guestUsers) {
            # Check if the account is already disabled and the action is 'Disable'
            if ($action -eq 'Disable' -and $user.AccountEnabled -eq $false) {
                $skipMessage = "Skipping user '$($user.UserPrincipalName)' as it is already disabled."
                Write-Host $skipMessage -ForegroundColor Yellow
                $skipMessage | Out-File -FilePath $logFilePath -Append
                continue # Skip to the next user in the loop
            }

            try {
                Write-Host "Processing user '$($user.UserPrincipalName)'..."
                if ($action -eq 'Disable') {
                    Update-MgUser -UserId $user.Id -AccountEnabled:$false
                }
                else {
                    Remove-MgUser -UserId $user.Id
                }
                $logMessage = "User '$($user.UserPrincipalName)' $actionVerb on $(Get-Date)"
                Write-Host $logMessage -ForegroundColor Green
                $logMessage | Out-File -FilePath $logFilePath -Append
                $processedCount++ # Increment counter on success
            }
            catch {
                $errorMessage = "❌ ERROR processing user '$($user.UserPrincipalName)': $($_.Exception.Message)"
                Write-Host $errorMessage -ForegroundColor Red
                $errorMessage | Out-File -FilePath $logFilePath -Append
            }
            finally {
                # Add a small delay to avoid API throttling
                Start-Sleep -Milliseconds 800
            }
        }
        
        $finalMessage = "Process complete.`nFound: $totalFound`nProcessed: $processedCount`n`nPlease check the log file for details:`n$logFilePath"
        "Processing complete. Processed $processedCount of $totalFound users." | Out-File -FilePath $logFilePath -Append
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