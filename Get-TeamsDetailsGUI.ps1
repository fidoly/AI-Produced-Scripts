<#
.SYNOPSIS
    A PowerShell script with a GUI to fetch and display details for Microsoft Teams in a tenant.

.DESCRIPTION
    This script provides a user-friendly Windows Forms interface to connect to Microsoft Graph
    and retrieve specified details for all Teams. Users can select which properties they
    want to see, view the results in a grid, and export the data to a CSV file.

    It requires PowerShell 7 and the Microsoft.Graph.Teams and Microsoft.Graph.Authentication modules.

.AUTHOR
    Gemini
.DATE
    2024-08-12
.MODIFIED
    2024-08-12 - Replaced deprecated StatusBar with StatusStrip for modern PowerShell compatibility.
#>

#region PRE-REQUISITES AND MODULE CHECK

# Check for required modules
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Groups')
$missingModules = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    $moduleList = $missingModules -join ", "
    Write-Error "Missing required module(s): $moduleList. Please install them by running: Install-Module -Name $moduleList"
    # Display a message box if GUI context is available
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Missing required module(s): $moduleList. Please install them and restart the script.", "Error", "OK", "Error")
    }
    catch {
        # Fallback for non-GUI environments
    }
    exit
}

# Import necessary modules
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Teams -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
}
catch {
    Write-Error "Failed to import required modules. Please ensure they are installed correctly."
    exit
}

#endregion

#region GUI DEFINITION (WINDOWS FORMS)

# Load necessary assemblies for the GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Teams Details Retriever"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.MinimumSize = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'Sizable'
$form.AutoScaleMode = 'Dpi'

# --- Font ---
$font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Font = $font

# --- Global variable to store results for export ---
$Global:TeamsData = $null

# --- Status Strip (replaces deprecated StatusBar) ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. Click 'Fetch Details' to begin."
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# --- Main Container (TableLayoutPanel for responsive layout) ---
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = 'Fill'
$mainPanel.ColumnCount = 2
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 250)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainPanel.RowCount = 1
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($mainPanel)

# --- Left Panel (for options) ---
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = 'Fill'
$mainPanel.Controls.Add($leftPanel, 0, 0)

# --- Properties GroupBox ---
$propertiesGroupBox = New-Object System.Windows.Forms.GroupBox
$propertiesGroupBox.Text = "Select Details to Gather"
$propertiesGroupBox.Anchor = 'Top, Bottom, Left, Right'
$propertiesGroupBox.Size = New-Object System.Drawing.Size(230, 450)
$propertiesGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$leftPanel.Controls.Add($propertiesGroupBox)

# --- Properties CheckedListBox ---
$propertiesCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$propertiesCheckedListBox.Dock = 'Fill'
$propertiesCheckedListBox.CheckOnClick = $true
$propertiesCheckedListBox.Items.AddRange(@(
    "Description",
    "Visibility",
    "IsArchived",
    "Mail",
    "Owners",
    "Members (Count)",
    "Guests (Count)",
    "Channels (Count)",
    "AllowCreateUpdateChannels",
    "AllowDeleteChannels",
    "AllowGiphy",
    "GiphyContentRating"
))
# Pre-check some common properties
$propertiesCheckedListBox.SetItemChecked(0, $true) # Description
$propertiesCheckedListBox.SetItemChecked(1, $true) # Visibility
$propertiesCheckedListBox.SetItemChecked(5, $true) # Members (Count)
$propertiesCheckedListBox.SetItemChecked(4, $true) # Owners

$propertiesGroupBox.Controls.Add($propertiesCheckedListBox)

# --- Buttons Panel (Bottom Left) ---
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Anchor = 'Bottom, Left, Right'
$buttonPanel.Size = New-Object System.Drawing.Size(230, 150)
$buttonPanel.Location = New-Object System.Drawing.Point(10, 470)
$leftPanel.Controls.Add($buttonPanel)

# --- Fetch Button ---
$fetchButton = New-Object System.Windows.Forms.Button
$fetchButton.Text = "Fetch Teams Details"
$fetchButton.Size = New-Object System.Drawing.Size(210, 40)
$fetchButton.Location = New-Object System.Drawing.Point(10, 10)
$fetchButton.BackColor = [System.Drawing.Color]::FromArgb(2, 117, 216)
$fetchButton.ForeColor = [System.Drawing.Color]::White
$fetchButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($fetchButton)

# --- Export Button ---
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export to CSV"
$exportButton.Size = New-Object System.Drawing.Size(210, 40)
$exportButton.Location = New-Object System.Drawing.Point(10, 55)
$exportButton.Enabled = $false # Disabled until data is fetched
$buttonPanel.Controls.Add($exportButton)

# --- Close Button ---
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object System.Drawing.Size(210, 40)
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
$resultsDataGridView.AutoSizeColumnsMode = 'Fill'
$resultsDataGridView.BackgroundColor = [System.Drawing.SystemColors]::Window
$rightPanel.Controls.Add($resultsDataGridView)

#endregion

#region EVENT HANDLERS

$fetchButton.Add_Click({
    # --- Disable UI elements during fetch ---
    $fetchButton.Enabled = $false
    $exportButton.Enabled = $false
    $fetchButton.Text = "Fetching..."
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $resultsDataGridView.DataSource = $null
    $resultsDataGridView.Rows.Clear()
    $resultsDataGridView.Columns.Clear()
    $Global:TeamsData = @()

    # --- Connect to Graph ---
    try {
        # Scopes required for the selected properties
        $scopes = @(
            "Team.ReadBasic.All", "TeamSettings.Read.All", "Channel.ReadBasic.All",
            "User.Read.All", "GroupMember.Read.All", "Directory.Read.All"
        )
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        $statusLabel.Text = "Successfully connected. Fetching Teams list..."
        $form.Refresh()
    }
    catch {
        $statusLabel.Text = "Error: Failed to connect to Microsoft Graph."
        [System.Windows.Forms.MessageBox]::Show("Could not connect to Microsoft Graph. Please check your permissions and internet connection.`n`nError: $($_.Exception.Message)", "Connection Failed", "OK", "Error")
        # --- Re-enable UI ---
        $fetchButton.Enabled = $true
        $fetchButton.Text = "Fetch Teams Details"
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    # --- Get Selected Properties ---
    $selectedProperties = $propertiesCheckedListBox.CheckedItems | ForEach-Object { $_.ToString() }
    if ($selectedProperties.Count -eq 0) {
        $statusLabel.Text = "Ready."
        [System.Windows.Forms.MessageBox]::Show("Please select at least one detail to gather.", "No Selection", "OK", "Warning")
        $fetchButton.Enabled = $true
        $fetchButton.Text = "Fetch Teams Details"
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    # --- Setup DataGridView Columns ---
    $resultsDataGridView.Columns.Add("DisplayName", "Team Name") | Out-Null
    foreach ($prop in $selectedProperties) {
        $resultsDataGridView.Columns.Add($prop.Replace(" ", ""), $prop) | Out-Null
    }

    # --- Fetch Data ---
    try {
        $allTeams = Get-MgTeam -All -ErrorAction Stop
        $statusLabel.Text = "Found $($allTeams.Count) teams. Gathering details..."
        $form.Refresh()

        $progress = 0
        foreach ($team in $allTeams) {
            $progress++
            $statusLabel.Text = "Processing team $progress of $($allTeams.Count): $($team.DisplayName)"

            $teamDetails = [ordered]@{
                DisplayName = $team.DisplayName
            }

            # Use a more efficient call to get multiple settings at once
            $fullTeamObject = Get-MgTeam -TeamId $team.Id -Property "memberSettings,messagingSettings,funSettings"

            foreach ($prop in $selectedProperties) {
                $propValue = switch ($prop) {
                    "Description" { $team.Description }
                    "Visibility" { $team.Visibility }
                    "IsArchived" { $team.IsArchived }
                    "Mail" { (Get-MgGroup -GroupId $team.Id -Property "mail").Mail }
                    "Owners" { (Get-MgGroupOwner -GroupId $team.Id -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserPrincipalName) -join "; " }
                    "Members (Count)" { (Get-MgGroupMember -GroupId $team.Id -All -ErrorAction SilentlyContinue).Count }
                    "Guests (Count)" { (Get-MgGroupMember -GroupId $team.Id -All -ErrorAction SilentlyContinue | Where-Object { $_.UserPrincipalName -like "*#EXT#*" }).Count }
                    "Channels (Count)" { (Get-MgTeamChannel -TeamId $team.Id -All -ErrorAction SilentlyContinue).Count }
                    "AllowCreateUpdateChannels" { $fullTeamObject.MemberSettings.AllowCreateUpdateChannels }
                    "AllowDeleteChannels" { $fullTeamObject.MemberSettings.AllowDeleteChannels }
                    "AllowGiphy" { $fullTeamObject.FunSettings.AllowGiphy }
                    "GiphyContentRating" { $fullTeamObject.FunSettings.GiphyContentRating }
                    default { "N/A" }
                }
                $teamDetails[$prop] = $propValue
            }
            
            # Add to global data for export
            $Global:TeamsData += New-Object PSObject -Property $teamDetails
            
            # Add row to DataGridView
            $rowValues = @($teamDetails.DisplayName) + ($selectedProperties | ForEach-Object { $teamDetails[$_] })
            $resultsDataGridView.Rows.Add($rowValues) | Out-Null
        }
        
        $statusLabel.Text = "Done. Successfully processed $($allTeams.Count) teams."
        if ($Global:TeamsData.Count -gt 0) {
            $exportButton.Enabled = $true
        }
    }
    catch {
        $statusLabel.Text = "An error occurred during data fetching."
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        # --- Re-enable UI ---
        $fetchButton.Enabled = $true
        $fetchButton.Text = "Fetch Teams Details"
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$exportButton.Add_Click({
    if ($null -eq $Global:TeamsData -or $Global:TeamsData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "No Data", "OK", "Information")
        return
    }

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV File (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Teams Report"
    $saveFileDialog.FileName = "TeamsReport_$(Get-Date -Format 'yyyy-MM-dd').csv"

    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $Global:TeamsData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Data successfully exported to $($saveFileDialog.FileName)", "Export Complete", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export data. Error: $($_.Exception.Message)", "Export Failed", "OK", "Error")
        }
    }
})

$closeButton.Add_Click({
    $form.Close()
})

#endregion

#region SHOW FORM
# Disconnect from Graph when the form is closed to clean up the session.
$form.Add_FormClosing({
    if (Get-MgContext) {
        Write-Host "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph
    }
})

# Show the form
[void]$form.ShowDialog()

#endregion
