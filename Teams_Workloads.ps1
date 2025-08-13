#requires -Version 7.0
<#
  Teams Discovery (GUI) — Device Code Only
  -------------------------------------------------
  Purpose: One-click snapshot of Teams usage & inventory with **device code auth only**.
  No app registrations, no certificates, no pop-up browser auth.

  What it exports (pick & choose):
   - Usage reports (D7/D30/D90):
       • User Activity Counts
       • User Activity User Detail
       • Team Activity Detail
       • Device Usage User Detail
   - Inventory (Teams list, owners/members/guests, archived, visibility)
   - Channels (per team, membership type)
   - External settings (federation / external access / client config)
   - Voice (Phone numbers, PSTN calls last 30 days)

  How auth works now:
   - **Microsoft Graph:** Connects with `-UseDeviceCode` and required scopes.
   - **Microsoft Teams PS:** Connects with `-UseDeviceAuthentication`.
   - Optional Tenant ID helps when signed into multiple orgs.

  Run it:
    pwsh -STA -File .\Teams_Discovery_GUI.ps1

  Notes in this revision:
   - Removed all app/cert/browser flows — **device code only**.
   - Simplified UI (Tenant ID + options + output folder).
   - Handles the Graph progress glitch; continues if CSV is saved.
   - Verifies Teams connection with Get-CsTenant and logs tenant name/ID.
#>

# --- Windows + WinForms ---
if ($IsLinux -or $IsMacOS) { Write-Host 'This GUI requires Windows.' -ForegroundColor Yellow; return }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$ErrorActionPreference = 'Stop'

# ------------------ Logging ------------------
$global:TxtLog = $null
function Write-UILog { param([string]$Message,[ValidateSet('INFO','WARN','ERROR')]$Level='INFO')
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "[$ts][$Level] $Message"
  Write-Host $line
  if ($global:TxtLog) { $global:TxtLog.AppendText($line + [Environment]::NewLine); $global:TxtLog.ScrollToCaret(); [System.Windows.Forms.Application]::DoEvents() | Out-Null }
}

# ------------------ Helpers ------------------
function Resolve-Cmd { param([string[]]$Names)
  foreach ($n in $Names) { $cmd = Get-Command $n -ErrorAction SilentlyContinue; if ($cmd) { return $cmd.Name } }
  return $null
}

# ------------------ Modules ------------------
function Ensure-Module { param([string]$Name,[string]$MinVersion='0.0.0')
  $have = Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge [Version]$MinVersion }
  if (-not $have) { Write-UILog "Installing module $Name (min $MinVersion)..."; Install-Module $Name -Scope CurrentUser -Force -AllowClobber -MinimumVersion $MinVersion }
  Import-Module $Name -ErrorAction Stop
}
function Ensure-Prereqs {
  Write-UILog 'Preparing required modules (first run may take a minute)...'
  Ensure-Module -Name 'Microsoft.Graph' -MinVersion '2.0.0'
  Ensure-Module -Name 'MicrosoftTeams'  -MinVersion '4.9.0'
  Write-UILog 'Modules ready.'
}

# ------------------ Auth (Device Code only) ------------------
function Connect-GraphDeviceCode { param([string]$TenantId,[switch]$NeedVoice)
  Write-UILog 'Connecting to Microsoft Graph (device code)...'
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' | Out-Null }
  $scopes = @('Reports.Read.All','Group.Read.All','Directory.Read.All','Team.ReadBasic.All','Channel.ReadBasic.All')
  if ($NeedVoice) { $scopes += 'CallRecords.Read.All' }

  # Build args dynamically to support older Graph versions
  $args = @{ Scopes = $scopes; NoWelcome = $true }
  if ($TenantId) { $args['TenantId'] = $TenantId }
  if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('UseDeviceCode')) { $args['UseDeviceCode'] = $true }
  else { Write-UILog 'WARNING: Your Graph module lacks -UseDeviceCode. Falling back to default sign-in.' 'WARN' }

  Connect-MgGraph @args
  Write-UILog 'Graph connected.'
}

function Connect-TeamsDeviceCode { param([string]$TenantId)
  Write-UILog 'Connecting to Microsoft Teams PowerShell (device code)...'
  if ($TenantId) { Connect-MicrosoftTeams -UseDeviceAuthentication -TenantId $TenantId -ShowBanner:$false | Out-Null }
  else { Connect-MicrosoftTeams -UseDeviceAuthentication -ShowBanner:$false | Out-Null }
  try { $t = Get-CsTenant -ErrorAction Stop; Write-UILog ("Teams PS connected to tenant: {0} [{1}]" -f ($t.DisplayName ?? '(unknown name)'), ($t.TenantId ?? '(unknown id)')) }
  catch { Write-UILog "Teams PS connected, but Get-CsTenant failed: $($_.Exception.Message)" 'WARN' }
}

# ------------------ Output Folders ------------------
function New-Directories { param([string]$Root)
  foreach ($name in 'Reports','Inventory','Settings','Voice') { $dir = Join-Path $Root $name; if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null } }
}

# ------------------ Reports ------------------
function Invoke-ReportCmd { param([string[]]$Candidates,[hashtable]$Params)
  $cmd = Resolve-Cmd -Names $Candidates; if (-not $cmd) { throw "Missing Graph Reports cmdlet: $($Candidates -join ', ')" }
  $oldPP = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
  try { & $cmd @Params }
  catch {
    $msg = $_.Exception.Message; $outFile = $Params['OutFile']
    $known = ($msg -match 'PercentComplete' -and $msg -match '2147483647')
    if ($known -and $outFile -and (Test-Path $outFile) -and ((Get-Item $outFile).Length -gt 0)) { Write-UILog ("Graph progress glitch; file saved: {0}" -f (Split-Path $outFile -Leaf)) 'WARN' }
    else { throw }
  }
  finally { $ProgressPreference = $oldPP }
}

function Get-TeamsReports { param([string[]]$Periods,[bool]$DoCounts,[bool]$DoUserDetail,[bool]$DoTeamDetail,[bool]$DoDeviceDetail,[string]$OutRoot)
  $reportMap = @{ Counts='Teams-UserActivity-Counts-{0}.csv'; UserDetail='Teams-UserActivity-UserDetail-{0}.csv'; TeamDetail='Teams-TeamActivity-Detail-{0}.csv'; DeviceDetail='Teams-DeviceUsage-UserDetail-{0}.csv' }
  foreach ($p in $Periods) {
    Write-UILog "Downloading usage reports for $p..."
    if ($DoCounts)     { $path = Join-Path $OutRoot ("Reports/" + ($reportMap.Counts -f $p));     Invoke-ReportCmd -Candidates @('Get-MgReportTeamsUserActivityCounts','Get-MgReportTeamUserActivityCount') -Params @{ Period=$p; OutFile=$path }; Write-UILog "Saved: $(Split-Path $path -Leaf)" }
    if ($DoUserDetail) { $path = Join-Path $OutRoot ("Reports/" + ($reportMap.UserDetail -f $p)); Invoke-ReportCmd -Candidates @('Get-MgReportTeamsUserActivityUserDetail','Get-MgReportTeamUserActivityUserDetail') -Params @{ Period=$p; OutFile=$path }; Write-UILog "Saved: $(Split-Path $path -Leaf)" }
    if ($DoTeamDetail) { $path = Join-Path $OutRoot ("Reports/" + ($reportMap.TeamDetail -f $p)); Invoke-ReportCmd -Candidates @('Get-MgReportTeamsTeamActivityDetail','Get-MgReportTeamActivityDetail') -Params @{ Period=$p; OutFile=$path }; Write-UILog "Saved: $(Split-Path $path -Leaf)" }
    if ($DoDeviceDetail){ $path = Join-Path $OutRoot ("Reports/" + ($reportMap.DeviceDetail -f $p)); Invoke-ReportCmd -Candidates @('Get-MgReportTeamsDeviceUsageUserDetail','Get-MgReportTeamDeviceUsageUserDetail') -Params @{ Period=$p; OutFile=$path }; Write-UILog "Saved: $(Split-Path $path -Leaf)" }
    [System.Windows.Forms.Application]::DoEvents() | Out-Null
  }
  Write-UILog 'Usage reports complete.'
}

# ------------------ Inventory / Settings / Voice ------------------
function Build-TeamsInventory { param([string]$OutRoot,[bool]$IncludeChannels=$true)
  Write-UILog 'Building Teams inventory...'
  $teamsGroups = Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property 'id,displayName,visibility,resourceProvisioningOptions'
  $inventory = foreach ($g in $teamsGroups) {
    $isArchived = $null; try { $t = Get-MgTeam -TeamId $g.Id -ErrorAction Stop; $isArchived = $t.IsArchived } catch { $isArchived = $null }
    $owners = Get-MgGroupOwner -GroupId $g.Id -All; $ownerUsers = $owners | Where-Object { $_.OdataType -eq '#microsoft.graph.user' }; $ownerCount = @($ownerUsers).Count
    $members = Get-MgGroupMember -GroupId $g.Id -All; $userMembers = $members | Where-Object { $_.OdataType -eq '#microsoft.graph.user' }; $memberCount = @($userMembers).Count; $guestCount = @($userMembers | Where-Object { $_.AdditionalProperties['userType'] -eq 'Guest' }).Count
    $std=$null;$priv=$null;$shrd=$null; try { $chs = Get-MgTeamChannel -TeamId $g.Id -All; $std=@($chs|?{$_.MembershipType -eq 'standard'}).Count; $priv=@($chs|?{$_.MembershipType -eq 'private'}).Count; $shrd=@($chs|?{$_.MembershipType -eq 'shared'}).Count } catch { $std=$priv=$shrd=$null }
    [pscustomobject]@{ TeamId=$g.Id; TeamName=$g.DisplayName; Visibility=$g.Visibility; Archived=$isArchived; Owners=$ownerCount; Members=$memberCount; Guests=$guestCount; ChannelsStandard=$std; ChannelsPrivate=$priv; ChannelsShared=$shrd }
  }
  $invPath = Join-Path $OutRoot 'Inventory/Teams-Inventory.csv'; $inventory | Sort-Object TeamName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $invPath; Write-UILog ("Teams inventory -> {0}" -f (Split-Path $invPath -Leaf))
  if ($IncludeChannels) {
    Write-UILog 'Listing channels per team...'
    $rows = foreach ($g in $teamsGroups) { try { $chs = Get-MgTeamChannel -TeamId $g.Id -All; foreach ($c in $chs) { [pscustomobject]@{ TeamId=$g.Id; TeamName=$g.DisplayName; ChannelId=$c.Id; ChannelName=$c.DisplayName; MembershipType=$c.MembershipType } } } catch { } ; [System.Windows.Forms.Application]::DoEvents() | Out-Null }
    $chPath = Join-Path $OutRoot 'Inventory/Teams-Channels.csv'; $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $chPath; Write-UILog ("Channels list -> {0}" -f (Split-Path $chPath -Leaf))
  }
  Write-UILog 'Inventory complete.'
}

function Export-ExternalSettings { param([string]$OutRoot)
  Write-UILog 'Exporting external access / client configuration...'
  $fed=$null;$eap=$null;$tcc=$null
  try { $fed = Get-CsTenantFederationConfiguration } catch { Write-UILog "Federation config not accessible: $($_.Exception.Message)" 'WARN' }
  try { $eap = Get-CsExternalAccessPolicy }          catch { Write-UILog "External access policies not accessible: $($_.Exception.Message)" 'WARN' }
  try { $tcc = Get-CsTeamsClientConfiguration }      catch { Write-UILog "Client configuration not accessible: $($_.Exception.Message)" 'WARN' }
  if ($fed) { $fed | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/ExternalAccess-Federation.json') -Encoding UTF8 }
  if ($eap) { $eap | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/ExternalAccess-Policies.json') -Encoding UTF8 }
  if ($tcc) { $tcc | ConvertTo-Json -Depth 6 | Out-File (Join-Path $OutRoot 'Settings/TeamsClientConfiguration.json') -Encoding UTF8 }
  Write-UILog 'Settings saved under Settings/.'
}

function Export-PhoneNumbers { param([string]$OutRoot)
  Write-UILog 'Exporting Teams phone numbers...'
  $top=1000;$skip=0;$all=@(); do { $page = Get-CsPhoneNumberAssignment -Top $top -Skip $skip; $count=@($page).Count; if ($count -gt 0){$all+=$page;$skip+=$top;Write-UILog "Retrieved $count (total $($all.Count))..."}; [System.Windows.Forms.Application]::DoEvents() | Out-Null } while ($count -eq $top)
  $path = Join-Path $OutRoot 'Voice/PhoneNumbers.csv'; $all | Select-Object TelephoneNumber,NumberType,ActivationState,AssignmentType,AssignedPstnTargetId,AssignedPstnTargetName,CountryCode,City | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path; Write-UILog ("Phone numbers -> {0}" -f (Split-Path $path -Leaf))
}

function Export-PstnCallsLast30Days { param([string]$OutRoot)
  Write-UILog 'Exporting PSTN call log (last 30 days)...'
  $to=(Get-Date).ToUniversalTime().ToString('s')+'Z'; $from=(Get-Date).AddDays(-30).ToUniversalTime().ToString('s')+'Z'; $uri="/communications/callRecords/getPstnCalls(fromDateTime=$from,toDateTime=$to)?`$top=999"; $rows=@()
  do { try { $resp=Invoke-MgGraphRequest -Method GET -Uri $uri } catch { $resp=Invoke-MgGraphRequest -Method GET -Uri $uri -ApiVersion 'beta' }; $rows+=$resp.value; $uri=$resp.'@odata.nextLink'; Write-UILog "Fetched $($rows.Count) PSTN rows so far..."; [System.Windows.Forms.Application]::DoEvents() | Out-Null } while ($uri)
  $path = Join-Path $OutRoot 'Voice/PSTN-Calls-Last30Days.csv'; $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path; Write-UILog ("PSTN log -> {0}" -f (Split-Path $path -Leaf))
}

# ------------------ GUI ------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Teams Discovery (Device Code)'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(980, 720)
$form.MinimumSize = New-Object System.Drawing.Size(980, 720)
$form.Padding = '10,10,10,10'

# Auth box (just Tenant ID + instructions)
$grpAuth = New-Object System.Windows.Forms.GroupBox; $grpAuth.Text='Authentication'; $grpAuth.Location=New-Object System.Drawing.Point(10,10); $grpAuth.Size=New-Object System.Drawing.Size(470,120); $grpAuth.Anchor='Top,Left'
$lblTenant = New-Object System.Windows.Forms.Label; $lblTenant.Text='Tenant ID (GUID - optional)'; $lblTenant.Location=New-Object System.Drawing.Point(15,25); $lblTenant.AutoSize=$true
$txtTenant = New-Object System.Windows.Forms.TextBox; $txtTenant.Location=New-Object System.Drawing.Point(15,45); $txtTenant.Size=New-Object System.Drawing.Size(440,24); $txtTenant.Anchor='Top,Left,Right'
$lblAuthInfo = New-Object System.Windows.Forms.Label; $lblAuthInfo.Text='On Run, you''ll get two device codes in the console: first Graph, then Teams. Open https://microsoft.com/devicelogin and enter each code.'; $lblAuthInfo.Location=New-Object System.Drawing.Point(15,75); $lblAuthInfo.Size=New-Object System.Drawing.Size(440,32)
$grpAuth.Controls.AddRange(@($lblTenant,$txtTenant,$lblAuthInfo))

# Periods
$grpPeriods = New-Object System.Windows.Forms.GroupBox; $grpPeriods.Text='Report Periods'; $grpPeriods.Location=New-Object System.Drawing.Point(10,140); $grpPeriods.Size=New-Object System.Drawing.Size(470,70); $grpPeriods.Anchor='Top,Left'
$cbD7=New-Object System.Windows.Forms.CheckBox; $cbD7.Text='D7'; $cbD7.Location=New-Object System.Drawing.Point(15,30); $cbD7.Checked=$true
$cbD30=New-Object System.Windows.Forms.CheckBox; $cbD30.Text='D30'; $cbD30.Location=New-Object System.Drawing.Point(80,30); $cbD30.Checked=$true
$cbD90=New-Object System.Windows.Forms.CheckBox; $cbD90.Text='D90'; $cbD90.Location=New-Object System.Drawing.Point(145,30); $cbD90.Checked=$true
$grpPeriods.Controls.AddRange(@($cbD7,$cbD30,$cbD90))

# Usage Reports
$grpReports = New-Object System.Windows.Forms.GroupBox; $grpReports.Text='Usage Reports (CSV)'; $grpReports.Location=New-Object System.Drawing.Point(10,220); $grpReports.Size=New-Object System.Drawing.Size(470,140); $grpReports.Anchor='Top,Left'
$cbCounts=New-Object System.Windows.Forms.CheckBox; $cbCounts.Text='User Activity Counts'; $cbCounts.Location=New-Object System.Drawing.Point(15,30); $cbCounts.Checked=$true
$cbUserDet=New-Object System.Windows.Forms.CheckBox; $cbUserDet.Text='User Activity User Detail'; $cbUserDet.Location=New-Object System.Drawing.Point(15,60); $cbUserDet.Checked=$true
$cbTeamDet=New-Object System.Windows.Forms.CheckBox; $cbTeamDet.Text='Team Activity Detail'; $cbTeamDet.Location=New-Object System.Drawing.Point(15,90); $cbTeamDet.Checked=$true
$cbDevDet=New-Object System.Windows.Forms.CheckBox; $cbDevDet.Text='Device Usage User Detail'; $cbDevDet.Location=New-Object System.Drawing.Point(240,30); $cbDevDet.Checked=$true
$grpReports.Controls.AddRange(@($cbCounts,$cbUserDet,$cbTeamDet,$cbDevDet))

# Discovery Options
$grpDiscover = New-Object System.Windows.Forms.GroupBox; $grpDiscover.Text='Discovery Options'; $grpDiscover.Location=New-Object System.Drawing.Point(10,370); $grpDiscover.Size=New-Object System.Drawing.Size(470,160); $grpDiscover.Anchor='Top,Left'
$cbInventory=New-Object System.Windows.Forms.CheckBox; $cbInventory.Text='Teams Inventory'; $cbInventory.Location=New-Object System.Drawing.Point(15,30); $cbInventory.Checked=$true
$cbChannels=New-Object System.Windows.Forms.CheckBox; $cbChannels.Text='Include Channels'; $cbChannels.Location=New-Object System.Drawing.Point(150,30); $cbChannels.Checked=$true
$cbExtSets=New-Object System.Windows.Forms.CheckBox; $cbExtSets.Text='External Settings (Federation/Client)'; $cbExtSets.Location=New-Object System.Drawing.Point(15,60); $cbExtSets.Checked=$true
$cbVoiceNums=New-Object System.Windows.Forms.CheckBox; $cbVoiceNums.Text='Voice: Phone Numbers'; $cbVoiceNums.Location=New-Object System.Drawing.Point(15,90)
$cbVoicePstn=New-Object System.Windows.Forms.CheckBox; $cbVoicePstn.Text='Voice: PSTN Calls (last 30 days)'; $cbVoicePstn.Location=New-Object System.Drawing.Point(200,90)
$grpDiscover.Controls.AddRange(@($cbInventory,$cbChannels,$cbExtSets,$cbVoiceNums,$cbVoicePstn))

# Output
$grpOut = New-Object System.Windows.Forms.GroupBox; $grpOut.Text='Output'; $grpOut.Location=New-Object System.Drawing.Point(490,10); $grpOut.Size=New-Object System.Drawing.Size(470,90); $grpOut.Anchor='Top,Right'
$lblOut=New-Object System.Windows.Forms.Label; $lblOut.Text='Save CSVs to folder'; $lblOut.Location=New-Object System.Drawing.Point(15,25); $lblOut.AutoSize=$true
$txtOut=New-Object System.Windows.Forms.TextBox; $txtOut.Location=New-Object System.Drawing.Point(15,45); $txtOut.Size=New-Object System.Drawing.Size(360,24); $txtOut.Anchor='Top,Left,Right'; $txtOut.Text=(Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("TeamsDiscovery-{0:yyyyMMdd_HHmm}" -f (Get-Date)))
$btnBrowse=New-Object System.Windows.Forms.Button; $btnBrowse.Text='Browse...'; $btnBrowse.Location=New-Object System.Drawing.Point(380,43); $btnBrowse.Size=New-Object System.Drawing.Size(70,26); $btnBrowse.Anchor='Top,Right'
$grpOut.Controls.AddRange(@($lblOut,$txtOut,$btnBrowse))

# Log
$txtLog = New-Object System.Windows.Forms.TextBox; $global:TxtLog=$txtLog; $txtLog.Multiline=$true; $txtLog.ScrollBars='Vertical'; $txtLog.ReadOnly=$true; $txtLog.Location=New-Object System.Drawing.Point(490,110); $txtLog.Size=New-Object System.Drawing.Size(470, 520); $txtLog.Anchor='Top,Bottom,Right'

# Buttons
$btnRun=New-Object System.Windows.Forms.Button; $btnRun.Text='Run'; $btnRun.Location=New-Object System.Drawing.Point(800,640); $btnRun.Size=New-Object System.Drawing.Size(70,28); $btnRun.Anchor='Bottom,Right'
$btnClose=New-Object System.Windows.Forms.Button; $btnClose.Text='Close'; $btnClose.Location=New-Object System.Drawing.Point(870,640); $btnClose.Size=New-Object System.Drawing.Size(70,28); $btnClose.Anchor='Bottom,Right'

$form.AcceptButton=$btnRun; $form.CancelButton=$btnClose
$form.Controls.AddRange(@($grpAuth,$grpPeriods,$grpReports,$grpDiscover,$grpOut,$txtLog,$btnRun,$btnClose))

$btnBrowse.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description='Choose a folder to save the CSV/JSON outputs'; $dlg.ShowNewFolderButton=$true; if ($dlg.ShowDialog()-eq 'OK'){ $txtOut.Text=$dlg.SelectedPath } })
$btnClose.Add_Click({ $form.Close() })

# ------------------ Runner ------------------
$btnRun.Add_Click({
  try {
    $btnRun.Enabled=$false
    Write-UILog 'Starting Teams discovery...'
    Ensure-Prereqs

    # Periods
    $periods=@(); if ($cbD7.Checked){$periods+='D7'}; if ($cbD30.Checked){$periods+='D30'}; if ($cbD90.Checked){$periods+='D90'}
    if (-not $periods){ [void][System.Windows.Forms.MessageBox]::Show('Choose at least one period (D7/D30/D90).','Validation','OK','Warning'); return }

    # Output folder
    $outRoot=$txtOut.Text; if (-not $outRoot){ [void][System.Windows.Forms.MessageBox]::Show('Select an output folder.','Validation','OK','Warning'); return }
    if (-not (Test-Path $outRoot)) { New-Item -ItemType Directory -Path $outRoot | Out-Null }
    New-Directories -Root $outRoot

    # Tenant ID (optional)
    $tenantText=$txtTenant.Text.Trim(); if ($tenantText -and $tenantText -notmatch '^[0-9a-fA-F-]{36}$'){ [void][System.Windows.Forms.MessageBox]::Show('Tenant ID must be a GUID or leave blank.','Validation','OK','Warning'); return }

    # Connect (device code for both)
    $needVoice = ($cbVoicePstn.Checked)
    Connect-GraphDeviceCode -TenantId $tenantText -NeedVoice:$needVoice
    Connect-TeamsDeviceCode -TenantId $tenantText

    # Usage reports
    if ($cbCounts.Checked -or $cbUserDet.Checked -or $cbTeamDet.Checked -or $cbDevDet.Checked) {
      Get-TeamsReports -Periods $periods -DoCounts:$($cbCounts.Checked) -DoUserDetail:$($cbUserDet.Checked) -DoTeamDetail:$($cbTeamDet.Checked) -DoDeviceDetail:$($cbDevDet.Checked) -OutRoot $outRoot
    }

    # Inventory / Channels
    if ($cbInventory.Checked) { Build-TeamsInventory -OutRoot $outRoot -IncludeChannels:$($cbChannels.Checked) }

    # External settings
    if ($cbExtSets.Checked) { Export-ExternalSettings -OutRoot $outRoot }

    # Voice
    if ($cbVoiceNums.Checked) { try { Export-PhoneNumbers -OutRoot $outRoot } catch { Write-UILog "Phone numbers failed: $($_.Exception.Message)" 'WARN' } }
    if ($cbVoicePstn.Checked) { try { Export-PstnCallsLast30Days -OutRoot $outRoot } catch { Write-UILog "PSTN export failed: $($_.Exception.Message)" 'WARN' } }

    Write-UILog ("Done. Output folder: {0}" -f $outRoot)
    [void][System.Windows.Forms.MessageBox]::Show("Finished!`n`nOutput folder:`n$outRoot","Teams Discovery",'OK','Information')
  }
  catch { Write-UILog $_.Exception.Message 'ERROR'; [void][System.Windows.Forms.MessageBox]::Show("Error:`n$($_.Exception.Message)",'Error','OK','Error') }
  finally { $btnRun.Enabled=$true }
})

# Show UI
[void]$form.ShowDialog()
