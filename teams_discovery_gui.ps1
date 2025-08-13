#requires -Version 7.0
<#!
  Teams Discovery (GUI)
  -------------------------------------------------
  Purpose: Help admins capture Teams/Graph discovery data without editing code.
  Runs on Windows PowerShell 7+ with WinForms.

  What it can export (pick & choose):
   - Usage reports (D7/D30/D90):
       • User Activity Counts
       • User Activity User Detail
       • Team Activity Detail
       • Device Usage User Detail
   - Inventory (Teams list, owners/members/guests, archived, visibility)
   - Channels (per team, membership type)
   - External settings (federation / external access / client config)
   - Voice (Phone numbers, PSTN calls last 30 days)

  Auth modes:
   - App-only (certificate): requires TenantId, AppId, cert with private key
   - Browser (delegated): interactive sign-in

  Tip: Run with -STA to ensure WinForms behaves well, e.g.:
       pwsh -STA -File .\Teams_Discovery_GUI.ps1

  Notes in this revision:
   - Fixed spacing & anchoring so controls don’t overlap and resize nicely.
   - Guarded Select-MgProfile (removed error if cmdlet isn’t present in your Graph version).
   - Added compatibility for both old/new Graph Reports cmdlet names (Team/Teams variants).
   - Made certificate picker show friendly names; value is the thumbprint.
   - PSTN export tries v1.0 first, then falls back to beta automatically.
!#>

# --- Safety checks (Windows + WinForms) ---
if ($IsLinux -or $IsMacOS) {
  Write-Host "This GUI requires Windows." -ForegroundColor Yellow
  return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'

# ------------------ Utility: Logging ------------------
$global:TxtLog = $null
function Write-UILog {
  param([string]$Message,[ValidateSet('INFO','WARN','ERROR')]$Level='INFO')
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "[$ts][$Level] $Message"
  Write-Host $line
  if ($global:TxtLog) {
    $global:TxtLog.AppendText($line + [Environment]::NewLine)
    $global:TxtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents() | Out-Null
  }
}

# ------------------ Helper: Command resolver ------------------
function Resolve-Cmd {
  param([string[]]$Names)
  foreach ($n in $Names) {
    $cmd = Get-Command $n -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Name }
  }
  return $null
}

# ------------------ Modules / Prereqs ------------------
function Ensure-Module {
  param([string]$Name,[string]$MinVersion='0.0.0')
  $have = Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge [Version]$MinVersion }
  if (-not $have) {
    Write-UILog "Installing module $Name (min $MinVersion)..."
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber -MinimumVersion $MinVersion
  }
  Import-Module $Name -ErrorAction Stop
}

function Ensure-Prereqs {
  Write-UILog "Preparing required modules (this can take a minute the first time)..."
  Ensure-Module -Name 'Microsoft.Graph' -MinVersion '2.0.0'     # Graph SDK
  Ensure-Module -Name 'MicrosoftTeams'  -MinVersion '4.9.0'     # App auth supported >= 4.7.1
  Write-UILog "Modules ready."
}

# ------------------ Auth ------------------
function Connect-GraphApp {
  param([string]$TenantId,[string]$AppId,[string]$Thumb)
  Write-UILog "Connecting to Microsoft Graph (app-only)..."
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name 'v1.0' | Out-Null
  }
  Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome
  Write-UILog "Graph connected as app $AppId."
}

function Connect-GraphDelegated {
  param([switch]$Voice)
  Write-UILog "Connecting to Microsoft Graph (interactive)..."
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name 'v1.0' | Out-Null
  }
  $scopes = @('Reports.Read.All','Group.Read.All','Directory.Read.All','Team.ReadBasic.All','Channel.ReadBasic.All')
  if ($Voice) { $scopes += 'CallRecords.Read.All' }
  Connect-MgGraph -Scopes $scopes -NoWelcome
  Write-UILog "Graph connected (delegated)."
}

function Connect-TeamsApp {
  param([string]$TenantId,[string]$AppId,[string]$Thumb)
  Write-UILog "Connecting to Microsoft Teams PowerShell (app-only)..."
  Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $AppId -CertificateThumbprint $Thumb | Out-Null
  Write-UILog "Teams PS connected (app-only)."
}

function Connect-TeamsDelegated {
  Write-UILog "Connecting to Microsoft Teams PowerShell (interactive)..."
  Connect-MicrosoftTeams | Out-Null
  Write-UILog "Teams PS connected (delegated)."
}

# ------------------ Cert Picker ------------------
function Get-CertCandidates {
  $cands = @()
  foreach ($locName in 'CurrentUser','LocalMachine') {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', $locName)
    try {
      $store.Open('ReadOnly')
      $now = Get-Date
      foreach ($c in $store.Certificates) {
        if ($c.HasPrivateKey -and $c.NotAfter -gt $now) {
          $cands += [pscustomobject]@{
            Location   = $locName
            Subject    = $c.Subject
            Thumbprint = ($c.Thumbprint -replace '\s','').ToUpper()
            NotAfter   = $c.NotAfter
            Display    = "[$locName] $($c.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,$false)) — $($c.Thumbprint.Substring(0,8))... (exp $(Get-Date $c.NotAfter -f 'yyyy-MM-dd'))"
          }
        }
      }
    } finally { $store.Close() }
  }
  $cands | Sort-Object NotAfter -Descending
}

# ------------------ Output Folders ------------------
function New-Directories {
  param([string]$Root)
  foreach ($name in 'Reports','Inventory','Settings','Voice') {
    $dir = Join-Path $Root $name
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
  }
}

# ------------------ Data Collection ------------------
function Invoke-ReportCmd {
  param(
    [string[]]$Candidates,
    [hashtable]$Params
  )
  $cmd = Resolve-Cmd -Names $Candidates
  if (-not $cmd) { throw "Required Graph Reports cmdlet not found: $($Candidates -join ', ')" }
  & $cmd @Params
}

function Get-TeamsReports {
  param(
    [string[]]$Periods,
    [bool]$DoCounts,
    [bool]$DoUserDetail,
    [bool]$DoTeamDetail,
    [bool]$DoDeviceDetail,
    [string]$OutRoot
  )

  $reportMap = @{
    Counts        = 'Teams-UserActivity-Counts-{0}.csv'
    UserDetail    = 'Teams-UserActivity-UserDetail-{0}.csv'
    TeamDetail    = 'Teams-TeamActivity-Detail-{0}.csv'
    DeviceDetail  = 'Teams-DeviceUsage-UserDetail-{0}.csv'
  }

  foreach ($p in $Periods) {
    Write-UILog "Downloading usage reports for $p..."
    if ($DoCounts) {
      $path = Join-Path $OutRoot ("Reports/" + ($reportMap.Counts -f $p))
      Invoke-ReportCmd -Candidates @('Get-MgReportTeamsUserActivityCounts','Get-MgReportTeamUserActivityCount') -Params @{ Period=$p; OutFile=$path }
      Write-UILog "Saved: $(Split-Path $path -Leaf)"
    }
    if ($DoUserDetail) {
      $path = Join-Path $OutRoot ("Reports/" + ($reportMap.UserDetail -f $p))
      Invoke-ReportCmd -Candidates @('Get-MgReportTeamsUserActivityUserDetail','Get-MgReportTeamUserActivityUserDetail') -Params @{ Period=$p; OutFile=$path }
      Write-UILog "Saved: $(Split-Path $path -Leaf)"
    }
    if ($DoTeamDetail) {
      $path = Join-Path $OutRoot ("Reports/" + ($reportMap.TeamDetail -f $p))
      Invoke-ReportCmd -Candidates @('Get-MgReportTeamsTeamActivityDetail','Get-MgReportTeamActivityDetail') -Params @{ Period=$p; OutFile=$path }
      Write-UILog "Saved: $(Split-Path $path -Leaf)"
    }
    if ($DoDeviceDetail) {
      $path = Join-Path $OutRoot ("Reports/" + ($reportMap.DeviceDetail -f $p))
      Invoke-ReportCmd -Candidates @('Get-MgReportTeamsDeviceUsageUserDetail','Get-MgReportTeamDeviceUsageUserDetail') -Params @{ Period=$p; OutFile=$path }
      Write-UILog "Saved: $(Split-Path $path -Leaf)"
    }
    [System.Windows.Forms.Application]::DoEvents() | Out-Null
  }
  Write-UILog "Usage reports complete."
}

function Build-TeamsInventory {
  param([string]$OutRoot,[bool]$IncludeChannels=$true)

  Write-UILog "Building Teams inventory..."
  $teamsGroups = Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property "id,displayName,visibility,resourceProvisioningOptions"

  $inventory = foreach ($g in $teamsGroups) {
    # team properties
    $isArchived = $null
    try {
      $t = Get-MgTeam -TeamId $g.Id -ErrorAction Stop
      $isArchived = $t.IsArchived
    } catch { $isArchived = $null }

    # owners
    $owners = Get-MgGroupOwner -GroupId $g.Id -All
    $ownerUsers = $owners | Where-Object { $_.OdataType -eq '#microsoft.graph.user' }
    $ownerCount = @($ownerUsers).Count

    # members + guests
    $members = Get-MgGroupMember -GroupId $g.Id -All
    $userMembers = $members | Where-Object { $_.OdataType -eq '#microsoft.graph.user' }
    $memberCount = @($userMembers).Count
    $guestCount  = @($userMembers | Where-Object { $_.AdditionalProperties['userType'] -eq 'Guest' }).Count

    # channel counts (best-effort)
    $std=$null; $priv=$null; $shrd=$null
    try {
      $chs = Get-MgTeamChannel -TeamId $g.Id -All
      $std  = @($chs | Where-Object { $_.MembershipType -eq 'standard' }).Count
      $priv = @($chs | Where-Object { $_.MembershipType -eq 'private'  }).Count
      $shrd = @($chs | Where-Object { $_.MembershipType -eq 'shared'   }).Count
    } catch { $std=$priv=$shrd=$null }

    [pscustomobject]@{
      TeamId           = $g.Id
      TeamName         = $g.DisplayName
      Visibility       = $g.Visibility
      Archived         = $isArchived
      Owners           = $ownerCount
      Members          = $memberCount
      Guests           = $guestCount
      ChannelsStandard = $std
      ChannelsPrivate  = $priv
      ChannelsShared   = $shrd
    }
  }

  $invPath = Join-Path $OutRoot 'Inventory/Teams-Inventory.csv'
  $inventory | Sort-Object TeamName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $invPath
  Write-UILog "Teams inventory -> $(Split-Path $invPath -Leaf)"

  if ($IncludeChannels) {
    Write-UILog "Listing channels per team..."
    $rows = foreach ($g in $teamsGroups) {
      try {
        $chs = Get-MgTeamChannel -TeamId $g.Id -All
        foreach ($c in $chs) {
          [pscustomobject]@{
            TeamId         = $g.Id
            TeamName       = $g.DisplayName
            ChannelId      = $c.Id
            ChannelName    = $c.DisplayName
            MembershipType = $c.MembershipType
          }
        }
      } catch { }
      [System.Windows.Forms.Application]::DoEvents() | Out-Null
    }
    $chPath = Join-Path $OutRoot 'Inventory/Teams-Channels.csv'
    $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $chPath
    Write-UILog "Channels list -> $(Split-Path $chPath -Leaf)"
  }

  Write-UILog "Inventory complete."
}

function Export-ExternalSettings {
  param([string]$OutRoot)
  Write-UILog "Exporting external access / client configuration..."
  $fed = $null; $eap = $null; $tcc = $null
  try { $fed = Get-CsTenantFederationConfiguration } catch { Write-UILog "Federation config not accessible: $($_.Exception.Message)" 'WARN' }
  try { $eap = Get-CsExternalAccessPolicy }          catch { Write-UILog "External access policies not accessible: $($_.Exception.Message)" 'WARN' }
  try { $tcc = Get-CsTeamsClientConfiguration }      catch { Write-UILog "Client configuration not accessible: $($_.Exception.Message)" 'WARN' }

  if ($fed) { $fed | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/ExternalAccess-Federation.json') -Encoding UTF8 }
  if ($eap) { $eap | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/ExternalAccess-Policies.json') -Encoding UTF8 }
  if ($tcc) { $tcc | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/TeamsClientConfiguration.json') -Encoding UTF8 }
  Write-UILog "Settings saved under Settings/."
}

function Export-PhoneNumbers {
  param([string]$OutRoot)
  Write-UILog "Exporting Teams phone numbers..."
  $top=1000; $skip=0
  $all=@()
  do {
    $page = Get-CsPhoneNumberAssignment -Top $top -Skip $skip
    $count = @($page).Count
    if ($count -gt 0) {
      $all += $page
      $skip += $top
      Write-UILog "Retrieved $count (total $($all.Count))..."
    }
    [System.Windows.Forms.Application]::DoEvents() | Out-Null
  } while ($count -eq $top)
  $path = Join-Path $OutRoot 'Voice/PhoneNumbers.csv'
  $all | Select-Object TelephoneNumber,NumberType,ActivationState,AssignmentType,AssignedPstnTargetId,AssignedPstnTargetName,CountryCode,City |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path
  Write-UILog "Phone numbers -> $(Split-Path $path -Leaf)"
}

function Export-PstnCallsLast30Days {
  param([string]$OutRoot)
  Write-UILog "Exporting PSTN call log (last 30 days)..."
  $to = (Get-Date).ToUniversalTime().ToString("s") + "Z"
  $from = (Get-Date).AddDays(-30).ToUniversalTime().ToString("s") + "Z"
  $uri = "/communications/callRecords/getPstnCalls(fromDateTime=$from,toDateTime=$to)?`$top=999"
  $rows = @()
  do {
    try {
      $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
    } catch {
      # Some tenants require beta for this endpoint
      $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ApiVersion 'beta'
    }
    $rows += $resp.value
    $uri = $resp.'@odata.nextLink'
    Write-UILog "Fetched $($rows.Count) PSTN rows so far..."
    [System.Windows.Forms.Application]::DoEvents() | Out-Null
  } while ($uri)

  $path = Join-Path $OutRoot 'Voice/PSTN-Calls-Last30Days.csv'
  $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path
  Write-UILog "PSTN log -> $(Split-Path $path -Leaf)"
}

# ------------------ GUI ------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Teams Discovery (GUI)'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1160, 800)
$form.MinimumSize = New-Object System.Drawing.Size(1160, 800)
$form.Padding = '10,10,10,10'

# Auth group
$grpAuth = New-Object System.Windows.Forms.GroupBox
$grpAuth.Text = 'Authentication'
$grpAuth.Location = New-Object System.Drawing.Point(10,10)
$grpAuth.Size = New-Object System.Drawing.Size(540, 210)
$grpAuth.Anchor = 'Top,Left'

$rbApp = New-Object System.Windows.Forms.RadioButton
$rbApp.Text = 'App-only (certificate)'
$rbApp.Location = New-Object System.Drawing.Point(15,25)
$rbApp.Checked = $true

$rbBrowser = New-Object System.Windows.Forms.RadioButton
$rbBrowser.Text = 'Browser (interactive)'
$rbBrowser.Location = New-Object System.Drawing.Point(270,25)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = 'Tenant ID (GUID)'
$lblTenant.Location = New-Object System.Drawing.Point(15,60)
$lblTenant.AutoSize = $true

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Location = New-Object System.Drawing.Point(15,80)
$txtTenant.Size = New-Object System.Drawing.Size(500,24)
$txtTenant.Anchor = 'Top,Left,Right'

$lblAppId = New-Object System.Windows.Forms.Label
$lblAppId.Text = 'App (Client) ID'
$lblAppId.Location = New-Object System.Drawing.Point(15,110)
$lblAppId.AutoSize = $true

$txtAppId = New-Object System.Windows.Forms.TextBox
$txtAppId.Location = New-Object System.Drawing.Point(15,130)
$txtAppId.Size = New-Object System.Drawing.Size(500,24)
$txtAppId.Anchor = 'Top,Left,Right'

$lblCert = New-Object System.Windows.Forms.Label
$lblCert.Text = 'Certificate'
$lblCert.Location = New-Object System.Drawing.Point(15,160)
$lblCert.AutoSize = $true

$cmbCert = New-Object System.Windows.Forms.ComboBox
$cmbCert.DropDownStyle = 'DropDownList'
$cmbCert.Location = New-Object System.Drawing.Point(85,156)
$cmbCert.Size = New-Object System.Drawing.Size(360,24)
$cmbCert.Anchor = 'Top,Left,Right'
$cmbCert.DisplayMember = 'Display'
$cmbCert.ValueMember   = 'Thumbprint'

$btnRefreshCert = New-Object System.Windows.Forms.Button
$btnRefreshCert.Text = 'Refresh'
$btnRefreshCert.Location = New-Object System.Drawing.Point(455,154)
$btnRefreshCert.Size = New-Object System.Drawing.Size(60,28)

$grpAuth.Controls.AddRange(@($rbApp,$rbBrowser,$lblTenant,$txtTenant,$lblAppId,$txtAppId,$lblCert,$cmbCert,$btnRefreshCert))

# Periods group
$grpPeriods = New-Object System.Windows.Forms.GroupBox
$grpPeriods.Text = 'Report Periods'
$grpPeriods.Location = New-Object System.Drawing.Point(10,230)
$grpPeriods.Size = New-Object System.Drawing.Size(540, 90)
$grpPeriods.Anchor = 'Top,Left'

$cbD7 = New-Object System.Windows.Forms.CheckBox;  $cbD7.Text='D7';   $cbD7.Location=New-Object System.Drawing.Point(15,38);  $cbD7.Checked=$true
$cbD30= New-Object System.Windows.Forms.CheckBox; $cbD30.Text='D30'; $cbD30.Location=New-Object System.Drawing.Point(85,38); $cbD30.Checked=$true
$cbD90= New-Object System.Windows.Forms.CheckBox; $cbD90.Text='D90'; $cbD90.Location=New-Object System.Drawing.Point(160,38);$cbD90.Checked=$true
$grpPeriods.Controls.AddRange(@($cbD7,$cbD30,$cbD90))

# Reports group (pick which reports)
$grpReports = New-Object System.Windows.Forms.GroupBox
$grpReports.Text = 'Usage Reports (CSV)'
$grpReports.Location = New-Object System.Drawing.Point(10,330)
$grpReports.Size = New-Object System.Drawing.Size(540, 160)
$grpReports.Anchor = 'Top,Left'

$cbCounts = New-Object System.Windows.Forms.CheckBox; $cbCounts.Text='User Activity Counts'; $cbCounts.Location=New-Object System.Drawing.Point(15,38); $cbCounts.Checked=$true
$cbUserDet= New-Object System.Windows.Forms.CheckBox; $cbUserDet.Text='User Activity User Detail'; $cbUserDet.Location=New-Object System.Drawing.Point(15,68); $cbUserDet.Checked=$true
$cbTeamDet= New-Object System.Windows.Forms.CheckBox; $cbTeamDet.Text='Team Activity Detail'; $cbTeamDet.Location=New-Object System.Drawing.Point(15,98); $cbTeamDet.Checked=$true
$cbDevDet = New-Object System.Windows.Forms.CheckBox; $cbDevDet.Text='Device Usage User Detail'; $cbDevDet.Location=New-Object System.Drawing.Point(270,38); $cbDevDet.Checked=$true
$grpReports.Controls.AddRange(@($cbCounts,$cbUserDet,$cbTeamDet,$cbDevDet))

# Discovery group (inventory/settings/voice)
$grpDiscover = New-Object System.Windows.Forms.GroupBox
$grpDiscover.Text = 'Discovery Options'
$grpDiscover.Location = New-Object System.Drawing.Point(10,500)
$grpDiscover.Size = New-Object System.Drawing.Size(540, 170)
$grpDiscover.Anchor = 'Top,Left'

$cbInventory = New-Object System.Windows.Forms.CheckBox; $cbInventory.Text='Teams Inventory'; $cbInventory.Location=New-Object System.Drawing.Point(15,38); $cbInventory.Checked=$true
$cbChannels  = New-Object System.Windows.Forms.CheckBox; $cbChannels.Text='Include Channels'; $cbChannels.Location=New-Object System.Drawing.Point(160,38); $cbChannels.Checked=$true
$cbExtSets   = New-Object System.Windows.Forms.CheckBox; $cbExtSets.Text='External Settings (Federation/Client)'; $cbExtSets.Location=New-Object System.Drawing.Point(15,68); $cbExtSets.Checked=$true
$cbVoiceNums = New-Object System.Windows.Forms.CheckBox; $cbVoiceNums.Text='Voice: Phone Numbers'; $cbVoiceNums.Location=New-Object System.Drawing.Point(15,98); $cbVoiceNums.Checked=$false
$cbVoicePstn = New-Object System.Windows.Forms.CheckBox; $cbVoicePstn.Text='Voice: PSTN Calls (last 30 days)'; $cbVoicePstn.Location=New-Object System.Drawing.Point(200,98); $cbVoicePstn.Checked=$false
$grpDiscover.Controls.AddRange(@($cbInventory,$cbChannels,$cbExtSets,$cbVoiceNums,$cbVoicePstn))

# Output group
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = 'Output'
$grpOut.Location = New-Object System.Drawing.Point(560,10)
$grpOut.Size = New-Object System.Drawing.Size(580, 110)
$grpOut.Anchor = 'Top,Right'

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = 'Save CSVs to folder'
$lblOut.Location = New-Object System.Drawing.Point(15,30)
$lblOut.AutoSize = $true

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(15,55)
$txtOut.Size = New-Object System.Drawing.Size(480,24)
$txtOut.Anchor = 'Top,Left,Right'
$defaultOut = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("TeamsDiscovery-{0:yyyyMMdd_HHmm}" -f (Get-Date))
$txtOut.Text = $defaultOut

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(505,53)
$btnBrowse.Size = New-Object System.Drawing.Size(65,28)
$btnBrowse.Anchor = 'Top,Right'

$grpOut.Controls.AddRange(@($lblOut,$txtOut,$btnBrowse))

# Log area
$txtLog = New-Object System.Windows.Forms.TextBox
$global:TxtLog = $txtLog
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(560,130)
$txtLog.Size = New-Object System.Drawing.Size(580, 580)
$txtLog.Anchor = 'Top,Bottom,Right'

# Buttons
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(965,720)
$btnRun.Size = New-Object System.Drawing.Size(80,30)
$btnRun.Anchor = 'Bottom,Right'

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Close'
$btnCancel.Location = New-Object System.Drawing.Point(1055,720)
$btnCancel.Size = New-Object System.Drawing.Size(80,30)
$btnCancel.Anchor = 'Bottom,Right'

$form.AcceptButton = $btnRun
$form.CancelButton = $btnCancel

$form.Controls.AddRange(@($grpAuth,$grpPeriods,$grpReports,$grpDiscover,$grpOut,$txtLog,$btnRun,$btnCancel))

# --- UI behavior ---
$refreshCerts = {
  $cmbCert.Items.Clear()
  $cands = Get-CertCandidates
  foreach ($c in $cands) { [void]$cmbCert.Items.Add($c) }
  if ($cmbCert.Items.Count -gt 0) { $cmbCert.SelectedIndex = 0 }
}

$btnRefreshCert.Add_Click($refreshCerts)

$rbToggle = {
  $isApp = $rbApp.Checked
  $txtTenant.Enabled = $isApp
  $txtAppId.Enabled  = $isApp
  $cmbCert.Enabled   = $isApp
  $btnRefreshCert.Enabled = $isApp
}

$rbApp.Add_CheckedChanged($rbToggle)
$rbBrowser.Add_CheckedChanged($rbToggle)

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = 'Choose a folder to save the CSV/JSON outputs'
  $dlg.ShowNewFolderButton = $true
  if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.SelectedPath }
})

$btnCancel.Add_Click({ $form.Close() })

# --- Runner ---
$btnRun.Add_Click({
  try {
    $btnRun.Enabled = $false
    Write-UILog "Starting Teams discovery..."
    Ensure-Prereqs

    # Resolve periods
    $periods = @()
    if ($cbD7.Checked)  { $periods += 'D7' }
    if ($cbD30.Checked) { $periods += 'D30' }
    if ($cbD90.Checked) { $periods += 'D90' }
    if (-not $periods) { [void][System.Windows.Forms.MessageBox]::Show('Choose at least one period (D7/D30/D90).','Validation','OK','Warning'); return }

    # Output folder
    $outRoot = $txtOut.Text
    if (-not $outRoot) { [void][System.Windows.Forms.MessageBox]::Show('Select an output folder.','Validation','OK','Warning'); return }
    if (-not (Test-Path $outRoot)) { New-Item -ItemType Directory -Path $outRoot | Out-Null }
    New-Directories -Root $outRoot

    # Connect
    if ($rbApp.Checked) {
      $tenant = $txtTenant.Text.Trim()
      $appId  = $txtAppId.Text.Trim()
      $certObj = $cmbCert.SelectedItem
      if (-not $tenant -or -not $appId -or -not $certObj) {
        [void][System.Windows.Forms.MessageBox]::Show('For App-only, provide TenantId, AppId, and select a certificate.','Validation','OK','Warning'); return
      }
      Connect-GraphApp  -TenantId $tenant -AppId $appId -Thumb $certObj.Thumbprint
      Connect-TeamsApp  -TenantId $tenant -AppId $appId -Thumb $certObj.Thumbprint
    } else {
      $needVoice = ($cbVoicePstn.Checked)
      Connect-GraphDelegated -Voice:$needVoice
      Connect-TeamsDelegated
    }

    # Usage reports
    if ($cbCounts.Checked -or $cbUserDet.Checked -or $cbTeamDet.Checked -or $cbDevDet.Checked) {
      Get-TeamsReports -Periods $periods -DoCounts:$($cbCounts.Checked) -DoUserDetail:$($cbUserDet.Checked) -DoTeamDetail:$($cbTeamDet.Checked) -DoDeviceDetail:$($cbDevDet.Checked) -OutRoot $outRoot
    }

    # Inventory / Channels
    if ($cbInventory.Checked) {
      Build-TeamsInventory -OutRoot $outRoot -IncludeChannels:$($cbChannels.Checked)
    }

    # External settings
    if ($cbExtSets.Checked) { Export-ExternalSettings -OutRoot $outRoot }

    # Voice
    if ($cbVoiceNums.Checked) { try { Export-PhoneNumbers -OutRoot $outRoot } catch { Write-UILog "Phone numbers failed: $($_.Exception.Message)" 'WARN' } }
    if ($cbVoicePstn.Checked) { try { Export-PstnCallsLast30Days -OutRoot $outRoot } catch { Write-UILog "PSTN export failed: $($_.Exception.Message)" 'WARN' } }

    Write-UILog "Done. Output folder: $outRoot"
    [void][System.Windows.Forms.MessageBox]::Show("Finished!`n`nOutput folder:`n$outRoot","Teams Discovery",'OK','Information')
  }
  catch {
    Write-UILog $_.Exception.Message 'ERROR'
    [void][System.Windows.Forms.MessageBox]::Show("Error:`n$($_.Exception.Message)",'Error','OK','Error')
  }
  finally {
    $btnRun.Enabled = $true
  }
})

# Init state
& $refreshCerts
& $rbToggle

# Show UI
[void]$form.ShowDialog()
