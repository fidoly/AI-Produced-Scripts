#Requires -Version 5.1
<#
.SYNOPSIS
    GUI-based tool to export Entra ID (Azure AD) user details with customizable field selection.

.DESCRIPTION
    This script provides a graphical interface to select which user properties to export,
    significantly improving performance by only querying selected data.
    Includes options for basic info, sign-in data, manager details, and more.

.PARAMETER None
    Script launches a GUI for all parameter input.

.EXAMPLE
    .\entra-user-select-details.ps1

.NOTES
    Version:        3.1.6
    Author:         System Administrator (Revised by Gemini)
    Purpose/Change:
        - v3.1.6: Replaced entire job-based auth with robust Device Code Flow to fix all timeout/cleanup errors.
        - v3.1.5: Simplified job cleanup (failed).
        - v3.1.4: Changed job cleanup to use -Id parameter (failed).
        - v3.1.3: Reduced auth timeout to 1 minute (failed).
        - v3.1.2: Implemented a 5-minute timeout for authentication (failed).
        - v3.1.1: FIXED REGRESSION. Restored search on 'mail' attribute to find Guest accounts.
#>

# Script Version
$ScriptVersion = "3.1.6"

# Add Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:SelectedFields = @{}
$script:Domain = ""
$script:OutputPath = "C:\temp"

# Function to create the GUI
function Show-ExportGUI {
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Entra ID User Export Tool v$ScriptVersion"
    $form.Size = New-Object System.Drawing.Size(520, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # --- Top Panel for Inputs ---
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Location = New-Object System.Drawing.Point(10, 10)
    $topPanel.Size = New-Object System.Drawing.Size(480, 120)
    $form.Controls.Add($topPanel)

    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Entra ID User Details Export"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $topPanel.Controls.Add($titleLabel)

    # Domain input
    $domainLabel = New-Object System.Windows.Forms.Label
    $domainLabel.Location = New-Object System.Drawing.Point(10, 50)
    $domainLabel.Size = New-Object System.Drawing.Size(100, 20)
    $domainLabel.Text = "Domain:"
    $domainLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $topPanel.Controls.Add($domainLabel)

    $domainTextBox = New-Object System.Windows.Forms.TextBox
    $domainTextBox.Location = New-Object System.Drawing.Point(120, 50)
    $domainTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $domainTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $domainTextBox.Text = ""
    $topPanel.Controls.Add($domainTextBox)

    # Output path
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Location = New-Object System.Drawing.Point(10, 85)
    $outputLabel.Size = New-Object System.Drawing.Size(100, 20)
    $outputLabel.Text = "Output Folder:"
    $outputLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $topPanel.Controls.Add($outputLabel)

    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Location = New-Object System.Drawing.Point(120, 85)
    $outputTextBox.Size = New-Object System.Drawing.Size(270, 20)
    $outputTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $outputTextBox.Text = $script:OutputPath
    $topPanel.Controls.Add($outputTextBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(400, 84)
    $browseButton.Size = New-Object System.Drawing.Size(60, 23)
    $browseButton.Text = "Browse"
    $browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($browseButton)

    # --- Scrollable Panel for Checkboxes ---
    $fieldsPanel = New-Object System.Windows.Forms.Panel
    $fieldsPanel.Location = New-Object System.Drawing.Point(10, 140)
    $fieldsPanel.Size = New-Object System.Drawing.Size(480, 420)
    $fieldsPanel.BorderStyle = "FixedSingle"
    $fieldsPanel.AutoScroll = $true
    $form.Controls.Add($fieldsPanel)

    $yPos = 10
    $checkboxes = @()

    # Helper function to add a checkbox item
    function Add-CheckboxItem {
        param($parent, $y, $text, $tag, $checked, $fontStyle, $description, $foreColor)

        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Location = New-Object System.Drawing.Point(10, $y)
        $checkbox.Size = New-Object System.Drawing.Size(430, 22)
        $checkbox.Text = $text
        $checkbox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]$fontStyle)
        $checkbox.Checked = $checked
        $checkbox.Tag = $tag
        if ($foreColor) { $checkbox.ForeColor = $foreColor }
        $parent.Controls.Add($checkbox)

        $currentY = $y + 22

        if ($description) {
            $descLabel = New-Object System.Windows.Forms.Label
            $descLabel.Location = New-Object System.Drawing.Point(30, $currentY)
            $descLabel.Size = New-Object System.Drawing.Size(400, 20)
            $descLabel.Text = $description
            $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $descLabel.ForeColor = [System.Drawing.Color]::DimGray
            $parent.Controls.Add($descLabel)
            $currentY += 25
        } else {
            $currentY += 5
        }

        return $checkbox, $currentY
    }

    # Field selection label
    $fieldLabel = New-Object System.Windows.Forms.Label
    $fieldLabel.Location = New-Object System.Drawing.Point(10, $yPos)
    $fieldLabel.Size = New-Object System.Drawing.Size(200, 20)
    $fieldLabel.Text = "Select Fields to Export:"
    $fieldLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $fieldsPanel.Controls.Add($fieldLabel)
    $yPos += 30

    # Basic Information (always checked)
    $basicCheckbox = New-Object System.Windows.Forms.CheckBox
    $basicCheckbox.Location = New-Object System.Drawing.Point(10, $yPos)
    $basicCheckbox.Size = New-Object System.Drawing.Size(430, 20)
    $basicCheckbox.Text = "Basic Information (Always included)"
    $basicCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $basicCheckbox.Checked = $true
    $basicCheckbox.Enabled = $false
    $fieldsPanel.Controls.Add($basicCheckbox)
    $yPos += 20

    $basicDesc = New-Object System.Windows.Forms.Label
    $basicDesc.Location = New-Object System.Drawing.Point(30, $yPos)
    $basicDesc.Size = New-Object System.Drawing.Size(400, 20)
    $basicDesc.Text = "Display Name, User Principal Name, Email, Object ID"
    $basicDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $basicDesc.ForeColor = [System.Drawing.Color]::DimGray
    $fieldsPanel.Controls.Add($basicDesc)
    $yPos += 30

    # --- Add Checkboxes using the helper function ---
    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Account Status & Dates" -tag "AccountStatus" -checked $true -fontStyle "Regular" -description "Enabled/Disabled, Created Date, Account Type"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Last Interactive Sign-In (Slower)" -tag "InteractiveSignIn" -checked $false -fontStyle "Regular" -description "Date of the last successful user-driven sign-in." -foreColor ([System.Drawing.Color]::DarkRed)
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Last Non-Interactive Sign-In (Slower)" -tag "NonInteractiveSignIn" -checked $false -fontStyle "Regular" -description "Date of the last successful application-driven sign-in." -foreColor ([System.Drawing.Color]::DarkRed)
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Organization Information" -tag "Organization" -checked $true -fontStyle "Regular" -description "Job Title, Department, Office, Company Name"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Manager Information (Slower)" -tag "Manager" -checked $false -fontStyle "Regular" -description "User's direct manager's display name." -foreColor ([System.Drawing.Color]::DarkRed)
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Contact Information" -tag "Contact" -checked $false -fontStyle "Regular" -description "Phone Numbers, Address"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Licensing Information" -tag "Licensing" -checked $false -fontStyle "Regular" -description "Assigned license names (e.g., E3, E5)."
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Employee ID" -tag "EmployeeID" -checked $false -fontStyle "Regular"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Employee Type" -tag "EmployeeType" -checked $false -fontStyle "Regular"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Usage Location" -tag "UsageLocation" -checked $false -fontStyle "Regular"
    $checkboxes += $cb

    $cb, $yPos = Add-CheckboxItem -parent $fieldsPanel -y $yPos -text "Proxy Addresses" -tag "ProxyAddresses" -checked $false -fontStyle "Regular"
    $checkboxes += $cb

    # --- Bottom Panel for buttons and status ---
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Location = New-Object System.Drawing.Point(10, 570)
    $bottomPanel.Size = New-Object System.Drawing.Size(480, 80)
    $form.Controls.Add($bottomPanel)

    # Select All / Clear All buttons
    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Location = New-Object System.Drawing.Point(10, 10)
    $selectAllButton.Size = New-Object System.Drawing.Size(80, 25)
    $selectAllButton.Text = "Select All"
    $selectAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $bottomPanel.Controls.Add($selectAllButton)

    $clearAllButton = New-Object System.Windows.Forms.Button
    $clearAllButton.Location = New-Object System.Drawing.Point(100, 10)
    $clearAllButton.Size = New-Object System.Drawing.Size(80, 25)
    $clearAllButton.Text = "Clear All"
    $clearAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $bottomPanel.Controls.Add($clearAllButton)

    # Time estimate label
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Location = New-Object System.Drawing.Point(220, 10)
    $timeLabel.Size = New-Object System.Drawing.Size(250, 25)
    $timeLabel.Text = "Estimated time: ~2 minutes"
    $timeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $timeLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    $timeLabel.TextAlign = "MiddleRight"
    $bottomPanel.Controls.Add($timeLabel)

    # Run/Cancel buttons
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(270, 45)
    $runButton.Size = New-Object System.Drawing.Size(90, 30)
    $runButton.Text = "Run Export"
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $runButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $bottomPanel.Controls.Add($runButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(370, 45)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $bottomPanel.Controls.Add($cancelButton)

    # --- Event handlers ---
    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select output folder"
        $folderBrowser.SelectedPath = $outputTextBox.Text

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputTextBox.Text = $folderBrowser.SelectedPath
        }
    })

    # Update time estimate based on selections
    $updateTimeEstimate = {
        $estimate = 2  # Base time
        if ($checkboxes | Where-Object { $_.Tag -like "*SignIn" -and $_.Checked }) { $estimate += 15 }
        if (($checkboxes | Where-Object { $_.Tag -eq "Manager" -and $_.Checked }).Checked) { $estimate += 10 }
        if (($checkboxes | Where-Object { $_.Tag -eq "Licensing" -and $_.Checked }).Checked) { $estimate += 3 }

        $timeLabel.Text = "Estimated time: ~$estimate minutes"
        $timeLabel.ForeColor = if ($estimate -gt 10) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DarkGreen }
    }

    # Function to manage CheckChanged events
    $suspendEvents = $false
    $changeCheckboxState = {
        param($checkState)
        $suspendEvents = $true
        foreach ($cb in $checkboxes) {
            $cb.Checked = $checkState
        }
        $suspendEvents = $false
        # Manually trigger the update once after all changes
        & $updateTimeEstimate
    }

    $selectAllButton.Add_Click({ & $changeCheckboxState $true })
    $clearAllButton.Add_Click({ & $changeCheckboxState $false })

    foreach ($cb in $checkboxes) {
        $cb.Add_CheckedChanged({ if (!$suspendEvents) { & $updateTimeEstimate }})
    }

    # Show the form
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([string]::IsNullOrWhiteSpace($domainTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a domain.", "Validation Error", "OK", "Warning")
            return $false
        }

        $script:Domain = $domainTextBox.Text
        $script:OutputPath = $outputTextBox.Text

        foreach ($cb in $checkboxes) {
            $script:SelectedFields[$cb.Tag] = $cb.Checked
        }

        return $true
    }

    return $false
}


# Function to install required modules
function Install-RequiredModule {
    param (
        [string]$ModuleName
    )

    Write-Host "🔍 Checking for $ModuleName module..." -ForegroundColor Yellow

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "📦 Installing $ModuleName module..." -ForegroundColor Yellow
        Write-Host "⏳ This may take a few minutes on first installation..." -ForegroundColor Gray

        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -Repository PSGallery
            Write-Host "✅ $ModuleName module installed successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Failed to install $ModuleName module: $_" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "✅ $ModuleName module is already installed." -ForegroundColor Green
    }

    return $true
}

# Function to show progress
function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity
    )
    $percentComplete = [Math]::Min([Math]::Round(($Current / $Total) * 100, 2), 100)
    if ($percentComplete -ge 0) {
        Write-Progress -Activity $Activity -Status "$Current of $Total users processed" -PercentComplete $percentComplete
    }
}

# --- Main Execution ---
try {
    Clear-Host
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          Entra ID User Details Export Tool v$ScriptVersion          ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Show-ExportGUI)) {
        Write-Host "❌ Export cancelled by user." -ForegroundColor Yellow
        return
    }

    Write-Host "✅ Configuration complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📋 Selected options:" -ForegroundColor Cyan
    Write-Host "   • Domain: $($script:Domain)"
    Write-Host "   • Output Path: $($script:OutputPath)"
    Write-Host "   • Fields selected:"

    $fieldCount = 0
    $script:SelectedFields.GetEnumerator() | ForEach-Object {
        if ($_.Value) {
            Write-Host "     ✓ $($_.Key)" -ForegroundColor Green
            $fieldCount++
        }
    }
    if ($fieldCount -eq 0) { Write-Host "     ✓ Basic Information Only" -ForegroundColor Green }

    Write-Host ""

    # Install and import required modules
    if (!(Install-RequiredModule -ModuleName "Microsoft.Graph")) { throw "Failed to install Microsoft Graph module" }

    # Build scopes and import modules based on selections
    $scopes = @("User.Read.All")
    $modulesToImport = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")

    if ($script:SelectedFields["InteractiveSignIn"] -or $script:SelectedFields["NonInteractiveSignIn"]) {
        $scopes += "AuditLog.Read.All", "Reports.Read.All"
        $modulesToImport += "Microsoft.Graph.Reports"
    }
    if ($script:SelectedFields["Manager"]) {
        $scopes += "Directory.Read.All"
    }

    Write-Host "📥 Importing required Microsoft Graph modules..." -ForegroundColor Yellow
    $importStart = Get-Date
    foreach($module in ($modulesToImport | Get-Unique)) {
        Import-Module $module -ErrorAction Stop
    }
    $importTime = [Math]::Round(((Get-Date) - $importStart).TotalSeconds, 2)
    Write-Host "✅ Modules imported successfully in $importTime seconds!" -ForegroundColor Green; Write-Host ""

    # Create output directory
    if (!(Test-Path -Path $script:OutputPath)) {
        Write-Host "📁 Creating directory: $($script:OutputPath)" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $script:OutputPath -Force | Out-Null
    }

    # Generate filename
    $fileName = "User Details for $($script:Domain) - $(Get-Date -Format 'yyyy-MM-dd').csv"
    $fullPath = Join-Path -Path $script:OutputPath -ChildPath $fileName

    # --- Connect to Microsoft Graph with Device Code Flow ---
    Write-Host "🔐 Authenticating..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication
        Write-Host "✅ Successfully authenticated!" -ForegroundColor Green
    } catch {
        throw "Authentication failed. $_"
    }
    
    Write-Host ""
    Write-Host "🔍 Retrieving user data for domain: $($script:Domain)" -ForegroundColor Cyan

    # Build property list based on selections
    $properties = @("Id", "DisplayName", "UserPrincipalName", "Mail")
    if ($script:SelectedFields["AccountStatus"]) { $properties += "AccountEnabled", "CreatedDateTime", "UserType" }
    if ($script:SelectedFields["Organization"]) { $properties += "JobTitle", "Department", "OfficeLocation", "CompanyName" }
    if ($script:SelectedFields["Contact"]) { $properties += "MobilePhone", "BusinessPhones", "StreetAddress", "City", "State", "Country", "PostalCode" }
    if ($script:SelectedFields["Licensing"]) { $properties += "AssignedLicenses" }
    if ($script:SelectedFields["EmployeeID"]) { $properties += "EmployeeId" }
    if ($script:SelectedFields["EmployeeType"]) { $properties += "EmployeeType" }
    if ($script:SelectedFields["UsageLocation"]) { $properties += "UsageLocation" }
    if ($script:SelectedFields["ProxyAddresses"]) { $properties += "ProxyAddresses" }
    if ($script:SelectedFields["Manager"]) { $properties += "Manager" }
    if ($script:SelectedFields["InteractiveSignIn"] -or $script:SelectedFields["NonInteractiveSignIn"]) { $properties += "SignInActivity" }

    try {
        $filter = "endsWith(mail,'@$($script:Domain)') or endsWith(userPrincipalName,'@$($script:Domain)')"
        $users = Get-MgUser -Filter $filter -All -Property ($properties | Get-Unique) -ConsistencyLevel eventual
        Write-Host "✅ Successfully retrieved $($users.Count) users" -ForegroundColor Green
    } catch { throw "Failed to retrieve users: $_" }

    if ($users.Count -eq 0) {
        Write-Host "⚠️  No users found with domain: $($script:Domain)" -ForegroundColor Yellow
        return
    }

    # Process users
    Write-Host ""; Write-Host "🔄 Processing user details..." -ForegroundColor Cyan

    $processedUsers = @()
    $totalUsers = $users.Count
    foreach ($i in 1..$totalUsers) {
        $user = $users[$i-1]
        Show-Progress -Current $i -Total $totalUsers -Activity "Processing user details"

        try {
            $props = [ordered]@{
                'Display Name' = $user.DisplayName
                'User Principal Name' = $user.UserPrincipalName
                'Email' = $user.Mail
                'Object ID' = $user.Id
            }

            if ($script:SelectedFields["AccountStatus"]) {
                $props['Account Enabled'] = if ($null -ne $user.AccountEnabled) { if ($user.AccountEnabled) { "Enabled" } else { "Disabled" } } else { "Unknown" }
                $props['Created DateTime'] = if ($user.CreatedDateTime) { [DateTime]::Parse($user.CreatedDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
                $props['Account Type'] = $user.UserType
            }

            if ($script:SelectedFields["InteractiveSignIn"] -or $script:SelectedFields["NonInteractiveSignIn"]) {
                 $signInActivity = $user.SignInActivity
                if ($script:SelectedFields["InteractiveSignIn"]) {
                    $props['Last Interactive Sign-In'] = if ($signInActivity.lastSignInDateTime) { [DateTime]::Parse($signInActivity.lastSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
                }
                if ($script:SelectedFields["NonInteractiveSignIn"]) {
                    $props['Last Non-Interactive Sign-In'] = if ($signInActivity.lastNonInteractiveSignInDateTime) { [DateTime]::Parse($signInActivity.lastNonInteractiveSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
                }
            }

            if ($script:SelectedFields["Organization"]) {
                $props['Job Title'] = $user.JobTitle
                $props['Department'] = $user.Department
                $props['Office Location'] = $user.OfficeLocation
                $props['Company Name'] = $user.CompanyName
            }

            if ($script:SelectedFields["Manager"]) {
                $props['Manager'] = if ($user.Manager.AdditionalProperties.displayName) { $user.Manager.AdditionalProperties.displayName } else { "" }
            }

            if ($script:SelectedFields["Contact"]) {
                $props['Mobile Phone'] = $user.MobilePhone
                $props['Business Phone'] = $user.BusinessPhones -join '; '
                $props['Street Address'] = $user.StreetAddress
                $props['City'] = $user.City
                $props['State'] = $user.State
                $props['Country'] = $user.Country
                $props['Postal Code'] = $user.PostalCode
            }

            if ($script:SelectedFields["Licensing"]) {
                $licenseNames = $user.AssignedLicenses | ForEach-Object {
                    switch ($_.SkuId) {
                        "06ebc4ee-1bb5-47dd-8120-11324bc54e06" { "E5" }
                        "c7df2760-2c81-4ef7-b578-5b5392b571df" { "E5" }
                        "26d45bd9-adf1-46cd-a9e1-51e9a5524128" { "E3" }
                        "189a915c-fe4f-4ffa-bde4-85b9628d07a0" { "Developer E5" }
                        "b05e124f-c7cc-45a0-a6aa-8cf78c946968" { "F1" }
                        "4b585984-651b-448a-9e53-3b10f069cf7f" { "F3" }
                        default { $_.SkuId }
                    }
                }
                $props['License Details'] = $licenseNames -join '; '
            }

            if ($script:SelectedFields["EmployeeID"]) { $props['Employee ID'] = $user.EmployeeId }
            if ($script:SelectedFields["EmployeeType"]) { $props['Employee Type'] = $user.EmployeeType }
            if ($script:SelectedFields["UsageLocation"]) { $props['Usage Location'] = $user.UsageLocation }
            if ($script:SelectedFields["ProxyAddresses"]) { $props['Proxy Addresses'] = $user.ProxyAddresses -join '; ' }

            $processedUsers += [PSCustomObject]$props
        }
        catch {
            Write-Host "`n⚠️  Warning: Could not process user $($user.UserPrincipalName): $_" -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity "Processing user details" -Completed; Write-Host ""

    # Export to CSV
    Write-Host "💾 Exporting data to CSV..." -ForegroundColor Cyan
    try {
        $processedUsers | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Successfully exported $($processedUsers.Count) users to:`n   📄 $fullPath" -ForegroundColor Green
    } catch { throw "Failed to export CSV: $_" }

    # Summary
    Write-Host "","╔═══════════════════════════════════════════════════════════════╗" -Separator "`n" -ForegroundColor Green
    Write-Host "║                    ✅ Export Complete!                         ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green; Write-Host ""

    # Open file location
    $openFolder = Read-Host "Would you like to open the output folder? (Y/N)"
    if ($openFolder -match 'Y') {
        Invoke-Item $script:OutputPath
    }
}
catch {
    Write-Host ""
    Write-Host "❌ Script Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "📋 Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
}
finally {
    if (Get-MgContext) {
        Write-Host ""; Write-Host "🔒 Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Disconnect-MgGraph
        Write-Host "✅ Disconnected successfully." -ForegroundColor Green
    }
    Write-Host ""; Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}