<#
.SYNOPSIS
    A PowerShell script with a GUI to discover shared channels, their members, and recent sign-in activity.

.DESCRIPTION
    This script provides a user-friendly Windows Forms interface to find all shared channels across all Teams
    in a tenant. It reports on each member of the shared channel, identifies if they are external by checking
    against all verified tenant domains, and queries Azure AD sign-in logs to find their last Teams-related 
    activity since a specified date. The process can be stopped gracefully by the user.

    The results can be exported to a CSV file.

    It requires PowerShell 7 and the Microsoft.Graph.Authentication, Microsoft.Graph.Teams, 
    Microsoft.Graph.Groups, and Microsoft.Graph.Reports modules.

.AUTHOR
    Gemini
.DATE
    2024-08-12
.MODIFIED
    2024-08-12 - Added a "Stop Process" button and cancellation logic to gracefully stop the script during execution.
#>

#region PRE-REQUISITES AND MODULE CHECK

# Check for required modules
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Reports')
$missingModules = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    $moduleList = $missingModules -join ", "
    Write-Error "Missing required module(s): $moduleList. Please install them by running: Install-Module -Name $moduleList"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Missing required module(s): $moduleList. Please install them and restart the script.", "Error", "OK", "Error")
    }
    catch {}
    exit
}

# Import necessary modules
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Teams -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    Import-Module Microsoft.Graph.Reports -ErrorAction Stop
}
catch {
    Write-Error "Failed to import required modules. Please ensure they are installed correctly."
    exit
}

#endregion

#region GUI DEFINITION (WINDOWS FORMS)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Shared Channels User & Activity Report"
$form.Size = New-Object System.Drawing.Size(1100, 750)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = 'Dpi'

# --- Font ---
$font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Font = $font

# --- Global variables ---
$Global:SharedChannelData = $null
$Global:IsCancellationRequested = $false

# --- Status Strip ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. Configure options and click 'Fetch Details'."
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# --- Main Container ---
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = 'Fill'
$mainPanel.ColumnCount = 2
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 280)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainPanel.RowCount = 1
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($mainPanel)

# --- Left Panel (for options) ---
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = 'Fill'
$mainPanel.Controls.Add($leftPanel, 0, 0)

# --- Options GroupBox ---
$optionsGroupBox = New-Object System.Windows.Forms.GroupBox
$optionsGroupBox.Text = "Report Options"
$optionsGroupBox.Anchor = 'Top, Left, Right'
$optionsGroupBox.Size = New-Object System.Drawing.Size(260, 100)
$optionsGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$leftPanel.Controls.Add($optionsGroupBox)

# --- Date Picker Label & Control ---
$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Text = "Show Sign-ins Since:"
$dateLabel.Location = New-Object System.Drawing.Point(10, 30)
$dateLabel.Size = New-Object System.Drawing.Size(240, 20)
$optionsGroupBox.Controls.Add($dateLabel)

$sinceDatePicker = New-Object System.Windows.Forms.DateTimePicker
$sinceDatePicker.Location = New-Object System.Drawing.Point(10, 55)
$sinceDatePicker.Size = New-Object System.Drawing.Size(240, 25)
$sinceDatePicker.Value = (Get-Date).AddDays(-30) # Default to 30 days ago
$optionsGroupBox.Controls.Add($sinceDatePicker)

# --- Buttons Panel ---
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Anchor = 'Bottom, Left, Right'
$buttonPanel.Size = New-Object System.Drawing.Size(260, 150)
$buttonPanel.Location = New-Object System.Drawing.Point(10, 540)
$leftPanel.Controls.Add($buttonPanel)

# --- Fetch Button ---
$fetchButton = New-Object System.Windows.Forms.Button
$fetchButton.Text = "Fetch Details"
$fetchButton.Size = New-Object System.Drawing.Size(240, 40)
$fetchButton.Location = New-Object System.Drawing.Point(10, 10)
$fetchButton.BackColor = [System.Drawing.Color]::FromArgb(2, 117, 216)
$fetchButton.ForeColor = [System.Drawing.Color]::White
$fetchButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($fetchButton)

# --- Stop Button (initially hidden) ---
$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop Process"
$stopButton.Size = New-Object System.Drawing.Size(240, 40)
$stopButton.Location = New-Object System.Drawing.Point(10, 10)
$stopButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69) # Red color
$stopButton.ForeColor = [System.Drawing.Color]::White
$stopButton.FlatStyle = 'Flat'
$stopButton.Visible = $false
$buttonPanel.Controls.Add($stopButton)

# --- Export Button ---
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export to CSV"
$exportButton.Size = New-Object System.Drawing.Size(240, 40)
$exportButton.Location = New-Object System.Drawing.Point(10, 55)
$exportButton.Enabled = $false
$buttonPanel.Controls.Add($exportButton)

# --- Close Button ---
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object System.Drawing.Size(240, 40)
$closeButton.Location = New-Object System.Drawing.Point(10, 100)
$buttonPanel.Controls.Add($closeButton)

# --- Right Panel (for results) ---
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$mainPanel.Controls.Add($rightPanel, 1, 0)

# --- Results DataGridView ---
$resultsDataGridView = New-Object System.Windows.Forms.DataGridView
$resultsDataGridView.Dock = 'Fill'
$resultsDataGridView.ReadOnly = $true
$resultsDataGridView.AllowUserToAddRows = $false
$resultsDataGridView.AutoSizeColumnsMode = 'AllCells'
$resultsDataGridView.BackgroundColor = [System.Drawing.SystemColors]::Window
$rightPanel.Controls.Add($resultsDataGridView)

#endregion

#region EVENT HANDLERS

$fetchButton.Add_Click({
    # --- Reset cancellation flag and UI ---
    $Global:IsCancellationRequested = $false
    $fetchButton.Visible = $false
    $stopButton.Visible = $true
    $exportButton.Enabled = $false
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $resultsDataGridView.DataSource = $null
    $resultsDataGridView.Rows.Clear()
    $resultsDataGridView.Columns.Clear()
    $Global:SharedChannelData = @()

    # --- Connect to Graph ---
    try {
        $scopes = @("Team.ReadBasic.All", "ChannelMember.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Organization.Read.All")
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        
        $statusLabel.Text = "Successfully connected. Fetching tenant domains..."
        $form.Refresh()
    }
    catch {
        $statusLabel.Text = "Error: Failed to connect to Microsoft Graph."
        [System.Windows.Forms.MessageBox]::Show("Could not connect to Microsoft Graph. Please check permissions and internet connection.`n`nError: $($_.Exception.Message)", "Connection Failed", "OK", "Error")
        # Reset UI in case of connection failure
        $fetchButton.Visible = $true; $stopButton.Visible = $false; $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    # --- Setup DataGridView Columns ---
    $columns = @{
        TeamName      = "Team Name"
        ChannelName   = "Channel Name"
        UserName      = "User Name"
        UserPrincipal = "User Principal"
        Role          = "Role"
        IsExternal    = "Is External"
        LastSignIn    = "Last Teams Sign-In"
        Application   = "Sign-In Application"
    }
    foreach ($col in $columns.GetEnumerator()) {
        $resultsDataGridView.Columns.Add($col.Name, $col.Value) | Out-Null
    }

    # --- Fetch Data ---
    try {
        $verifiedDomains = (Get-MgOrganization).VerifiedDomains.Name
        if (-not $verifiedDomains) {
            [System.Windows.Forms.MessageBox]::Show("Could not retrieve verified domains for the tenant. Please check permissions.", "Error", "OK", "Error")
            throw "Could not retrieve tenant domains."
        }

        $allTeams = Get-MgTeam -All -ErrorAction Stop
        $statusLabel.Text = "Found $($allTeams.Count) teams. Searching for shared channels..."
        $form.Refresh()

        $teamProgress = 0
        foreach ($team in $allTeams) {
            if ($Global:IsCancellationRequested) { break } # Check for cancellation
            $teamProgress++
            $statusLabel.Text = "Processing Team $teamProgress of $($allTeams.Count): $($team.DisplayName)"

            $sharedChannels = Get-MgTeamChannel -TeamId $team.Id -Filter "membershipType eq 'shared'" -ErrorAction SilentlyContinue
            if (-not $sharedChannels) { continue }

            foreach ($channel in $sharedChannels) {
                if ($Global:IsCancellationRequested) { break } # Check for cancellation
                $statusLabel.Text = "Team $($team.DisplayName) | Channel: $($channel.DisplayName)"
                $channelMembers = Get-MgTeamChannelMember -TeamId $team.Id -ChannelId $channel.Id -All -ErrorAction SilentlyContinue

                foreach ($member in $channelMembers) {
                    if ($Global:IsCancellationRequested) { break } # Check for cancellation
                    
                    $userDomain = $member.Email.Split('@')[1]
                    $isExternal = -not ($verifiedDomains -contains $userDomain)
                    
                    $roleText = "Member"
                    if ($member.Roles -contains 'owner') { $roleText = "Owner" }
                    
                    $sinceDate = $sinceDatePicker.Value
                    $filterString = "userPrincipalName eq '$($member.Email)' and createdDateTime ge $($sinceDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")) and appDisplayName eq 'Microsoft Teams'"
                    $signIn = Get-MgAuditLogSignIn -Filter $filterString -Top 1 -ErrorAction SilentlyContinue

                    $resultRow = @{
                        TeamName      = $team.DisplayName; ChannelName   = $channel.DisplayName
                        UserName      = $member.DisplayName; UserPrincipal = $member.Email
                        Role          = $roleText; IsExternal    = $isExternal
                        LastSignIn    = if ($signIn) { $signIn.CreatedDateTime } else { "No recent sign-in" }
                        Application   = if ($signIn) { $signIn.AppDisplayName } else { "N/A" }
                    }
                    
                    $Global:SharedChannelData += New-Object PSObject -Property $resultRow
                    $resultsDataGridView.Rows.Add($resultRow.Values) | Out-Null
                }
            }
        }
        
        if ($Global:IsCancellationRequested) {
            $statusLabel.Text = "Process stopped by user."
        } else {
            $statusLabel.Text = "Done. Report complete."
            if ($Global:SharedChannelData.Count -eq 0) {
                $statusLabel.Text = "Done. No shared channels found in this tenant."
            }
        }
        
        if ($Global:SharedChannelData.Count -gt 0) {
            $exportButton.Enabled = $true
        }
    }
    catch {
        $statusLabel.Text = "An error occurred during data fetching."
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        # --- Reset UI ---
        $fetchButton.Visible = $true
        $stopButton.Visible = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$stopButton.Add_Click({
    $statusLabel.Text = "Stopping process..."
    $Global:IsCancellationRequested = $true
})

$exportButton.Add_Click({
    if ($null -eq $Global:SharedChannelData -or $Global:SharedChannelData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "No Data", "OK", "Information")
        return
    }

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV File (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Shared Channels Report"
    $saveFileDialog.FileName = "SharedChannelsReport_$(Get-Date -Format 'yyyy-MM-dd').csv"

    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $Global:SharedChannelData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Data successfully exported to $($saveFileDialog.FileName)", "Export Complete", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export data. Error: $($_.Exception.Message)", "Export Failed", "OK", "Error")
        }
    }
})

$closeButton.Add_Click({
    $Global:IsCancellationRequested = $true
    $form.Close()
})

#endregion

#region SHOW FORM
$form.Add_FormClosing({
    if (Get-MgContext) {
        Write-Host "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph
    }
})

[void]$form.ShowDialog()

#endregion
