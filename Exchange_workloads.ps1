# Name:            Exchange_Workloads.ps1
# Version:         1.3.0
# Build with:      GPT-5 Thinking
# Designed by:     Phil Rangel
# Changes:
#   1.3.0 (2025-08-11)
#       - Post-login tenant confirmation: shows account UPN + accepted domains; prompts to continue
#       - Authentication UI spacing polished; anchors so fields don’t wrap on resize
#   1.2.0 (2025-08-11)
#       - Added "Browser (force picker)" auth mode: tries -UseWebLogin or -DisableWAM; falls back to Device
#       - Spacing/wording tweaks in auth section
#   1.1.0 (2025-08-11)
#       - Auth mode selector: Auto (SSO), Browser, Device (+ optional UPN)
#   1.0.1 (2025-08-11)
#       - Separate transcript vs operational logs (no file lock), PID in filenames; formatting fixes; full UI restored
#   1.0.0 (2025-08-11)
#       - Initial release (PS7 advisory, module hygiene, delegated admin, forwarding discovery, retries, incremental CSV, summary)

<# 
.SYNOPSIS
  Exchange Online Discovery (GUI + Console fallback), PS7-ready, resilient logging/error handling.

.DESCRIPTION
  - Connects to Exchange Online with selectable auth mode:
      * Auto (SSO if available)
      * Browser (account picker; optional UPN)
      * Browser (force picker)
      * Device code (always interactive)
  - Confirms the connected context (UPN + accepted domains) before discovery (GUI Yes/No; console Y/N).
  - Supports delegated admin (Tenant Domain for -DelegatedOrganization).
  - Discovers all mailboxes (incl. inactive) + stats (primary & archive).
  - Captures forwarding: mailbox props, optional inbox rules, transport rules; org AutoForward policy.
  - Outputs: CSV (per-mailbox), TXT (summary), optional CSV (transport rules), logs (operational, transcript, errors).
  - WinForms UI: output path, tenant domain, auth mode, UPN, progress bar, ETA, live log, copy/paste summary.
  - Console-only mode via -ConsoleOnly.
  - Retries with exponential backoff; continues on errors.

.PARAMETER ConsoleOnly
  Use console mode (no GUI).

.PARAMETER SkipInboxRuleScan
  Skip per-mailbox Inbox rule scans (faster). Mailbox property + transport rule forwarding still collected.

.PARAMETER TopN
  How many “largest mailboxes” to show in the summary (default 10; 1..1000).

.PARAMETER ExportTransportRules
  Export transport rules to CSV.

.PARAMETER AuthMode
  Authentication mode: Auto | Browser | BrowserForce | Device (default Auto).

.PARAMETER AuthUpn
  Optional UPN to steer browser login (Browser/BrowserForce).

.NOTES
  Author: Phil Rangel
  Date:   2025-08-11
#>

[CmdletBinding()]
param(
  [switch]$ConsoleOnly,
  [switch]$SkipInboxRuleScan,
  [int]$TopN = 10,
  [switch]$ExportTransportRules,
  [ValidateSet('Auto','Browser','BrowserForce','Device')]
  [string]$AuthMode = 'Auto',
  [string]$AuthUpn
)

#------------------------------ Global Setup ----------------------------------
$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date
$script:SessionIsPS7 = $PSVersionTable.PSVersion.Major -ge 7
$script:AppName = "EXO Discovery"
$script:WorkRoot = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath 'EXO-Discovery'
New-Item -ItemType Directory -Force -Path $script:WorkRoot | Out-Null
$stamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$pidTag = $PID

# Separate files to avoid transcript lock conflicts
$script:TranscriptFile = Join-Path $script:WorkRoot ("EXO-Discovery-{0}-{1}.transcript.log" -f $stamp,$pidTag)
$script:LogFile        = Join-Path $script:WorkRoot ("EXO-Discovery-{0}-{1}.operational.log" -f $stamp,$pidTag)
$script:ErrFile        = Join-Path $script:WorkRoot ("EXO-Discovery-{0}-{1}.errors.log" -f $stamp,$pidTag)

$script:SummaryFile = $null
$script:CsvFile = $null
$script:RulesCsv = $null
$script:UiLogBox = $null
$script:UiProgressBar = $null
$script:UiStartBtn = $null
$script:UiSummaryBox = $null
$script:UiEtaLabel = $null
$script:UiIsActive = $false  # set true after UI loads to know when we can show MessageBoxes

Start-Transcript -Path $script:TranscriptFile -Append | Out-Null

function Write-Log {
  param([string]$Message, [ValidateSet('INFO','WARN','ERROR','DEBUG')]$Level='INFO')
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[{0}][{1}] {2}" -f $ts,$Level,$Message
  Write-Host $line
  Add-Content -Path $script:LogFile -Value $line
  if ($Level -eq 'ERROR') { Add-Content -Path $script:ErrFile -Value $line }
  if ($script:UiLogBox) {
    try {
      $script:UiLogBox.AppendText("$line`r`n")
      $script:UiLogBox.SelectionStart = $script:UiLogBox.Text.Length
      $script:UiLogBox.ScrollToCaret()
    } catch {}
  }
}

function Advise-PSVersion {
  if (-not $script:SessionIsPS7) {
    Write-Log ("Detected PowerShell {0}. PowerShell 7+ is recommended for Graph/EXO performance and reliability." -f $PSVersionTable.PSVersion) 'WARN'
  } else {
    Write-Log ("PowerShell {0} detected. Good to go." -f $PSVersionTable.PSVersion) 'INFO'
  }
}

function Ensure-PSGalleryTrusted {
  $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
  if (-not $repo) { Register-PSRepository -Default -ErrorAction Stop }
  if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
  }
}

function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name, [string]$MinVersion='3.0.0')
  $loaded = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
  if (-not $loaded -or [version]$loaded.Version -lt [version]$MinVersion) {
    Write-Log ("Installing/updating {0} ({1}+)..." -f $Name, $MinVersion) 'INFO'
    Ensure-PSGalleryTrusted
    Install-Module -Name $Name -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
  }
  Import-Module $Name -ErrorAction Stop
  $ver = (Get-Module $Name).Version.ToString()
  Write-Log ("{0} loaded (v{1})." -f $Name, $ver) 'INFO'
}

function Invoke-WithRetry {
  param(
    [Parameter(Mandatory)][scriptblock]$ScriptBlock,
    [int]$MaxAttempts = 6,
    [int]$BaseDelaySeconds = 5,
    [string]$OpName = 'operation'
  )
  $attempt = 0
  do {
    try {
      $attempt++
      return & $ScriptBlock
    } catch {
      $msg = $_.Exception.Message
      Write-Log ("Attempt {0}/{1} failed for {2}: {3}" -f $attempt, $MaxAttempts, $OpName, $msg) 'WARN'
      Add-Content -Path $script:ErrFile -Value ("{0} - {1} failed: {2}" -f (Get-Date), $OpName, $msg)
      if ($attempt -ge $MaxAttempts) { throw }
      $delay = [int]([math]::Min(300, $BaseDelaySeconds * [math]::Pow(2, $attempt-1))) + (Get-Random -Minimum 0 -Maximum 3)
      Start-Sleep -Seconds $delay
    }
  } while ($true)
}

function Bytes-ToMB {
  param($o)
  try {
    if ($o -and $o.ToString().Contains("bytes")) {
      $b = ($o.ToString().Split('(')[1]).Split(' bytes')[0].Replace(',','')
      return [int]([double]$b/1MB)
    } elseif ($o -and $o.ToString() -match '^\d+$') {
      return [int]([double]$o/1MB)
    } else { return $null }
  } catch { return $null }
}

function Safe-GetInboxRules {
  param([string]$MailboxUPN)
  try {
    Invoke-WithRetry -OpName ("Get-InboxRule {0}" -f $MailboxUPN) -ScriptBlock {
      Get-InboxRule -Mailbox $MailboxUPN -ErrorAction Stop
    }
  } catch {
    Write-Log ("Get-InboxRule failed for {0}: {1}" -f $MailboxUPN, $_.Exception.Message) 'ERROR'; @()
  }
}

function Safe-GetMailboxStats {
  param([Guid]$ExchangeGuid, [switch]$Archive, [switch]$IncludeSoftDeleted)
  Invoke-WithRetry -OpName 'Get-EXOMailboxStatistics' -ScriptBlock {
    $params = @{ ExchangeGuid = $ExchangeGuid; ErrorAction = 'Stop' }
    if ($Archive) { $params['Archive'] = $true }
    if ($IncludeSoftDeleted) { $params['IncludeSoftDeletedRecipients'] = $true }
    Get-EXOMailboxStatistics @params
  }
}

function Connect-EXO {
  param(
    [string]$DelegatedOrganization,
    [ValidateSet('Auto','Browser','BrowserForce','Device')]
    [string]$Mode = 'Auto',
    [string]$Upn
  )
  $params = @{ ShowBanner = $false; ErrorAction = 'Stop' }
  if ($DelegatedOrganization) { $params['DelegatedOrganization'] = $DelegatedOrganization }

  switch ($Mode) {
    'Device'  { $params['Device'] = $true }
    'Browser' { if ($Upn) { $params['UserPrincipalName'] = $Upn } }
    'BrowserForce' {
      $cmd = Get-Command Connect-ExchangeOnline -ErrorAction Stop
      $hasUseWebLogin = $cmd.Parameters.ContainsKey('UseWebLogin')
      $hasDisableWAM  = $cmd.Parameters.ContainsKey('DisableWAM')
      if ($hasUseWebLogin) { $params['UseWebLogin'] = $true }
      elseif ($hasDisableWAM) { $params['DisableWAM'] = $true }
      else { Write-Log "BrowserForce not supported by this module version; falling back to Device flow." 'WARN'; $params['Device'] = $true }
      if ($Upn) { $params['UserPrincipalName'] = $Upn }
    }
    default { } # Auto
  }

  Invoke-WithRetry -OpName 'Connect-ExchangeOnline' -ScriptBlock {
    Connect-ExchangeOnline @params | Out-Null
  }
  Write-Log ("Connected to Exchange Online (AuthMode={0}{1})." -f $Mode, $(if($Upn){"; UPN=$Upn"})) 'INFO'
}

function Get-ConnectedContext {
  # Returns UPN + a short list of accepted domains for visual confirmation
  $upn = $null; $org = $null; $domains = @()
  try {
    $ci = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
    if ($ci) {
      $upn = $ci.UserPrincipalName
      $org = $ci.Organization
    }
  } catch {}
  try {
    $ads = Get-AcceptedDomain -ErrorAction Stop | Sort-Object -Property Default -Descending
    foreach ($ad in $ads | Select-Object -First 4) {
      $dn = $null
      if ($ad.PSObject.Properties['DomainName']) { $dn = $ad.DomainName.ToString() } else { $dn = $ad.Name }
      if ($ad.Default) { $dn = "$dn (Default)" }
      $domains += $dn
    }
  } catch {}
  [PSCustomObject]@{
    UserUPN = $upn
    Organization = $org
    AcceptedDomains = ($domains -join ', ')
  }
}

function Resolve-RecipientToSMTP {
  param($obj)
  if (-not $obj) { return $null }
  try {
    if ($obj -is [string]) { return $obj }
    if ($obj.PrimarySmtpAddress) { return $obj.PrimarySmtpAddress.ToString() }
    if ($obj.ExternalEmailAddress) { return $obj.ExternalEmailAddress.ToString() }
    if ($obj.DisplayName) { return $obj.DisplayName }
    return $obj.ToString()
  } catch { return $obj.ToString() }
}

function Build-CSVRow {
  param($Mailbox, $Stats, $ArchStats, [string[]]$InboxRuleRecipients, [bool]$HasInboxForward)
  [PSCustomObject]@{
    UserPrincipalName               = $Mailbox.UserPrincipalName
    DisplayName                     = $Mailbox.DisplayName
    PrimarySmtpAddress              = $Mailbox.PrimarySmtpAddress
    RecipientType                   = $Mailbox.RecipientType
    RecipientTypeDetails            = $Mailbox.RecipientTypeDetails
    IsInactiveMailbox               = $Mailbox.IsInactiveMailbox
    AccountDisabled                 = $Mailbox.AccountDisabled
    LitigationHoldEnabled           = $Mailbox.LitigationHoldEnabled
    RetentionHoldEnabled            = $Mailbox.RetentionHoldEnabled
    HasArchive                      = [string]([bool]$Mailbox.ArchiveName)
    WhenMailboxCreated              = $Mailbox.WhenMailboxCreated
    RetentionPolicy                 = $Mailbox.RetentionPolicy
    MaxSendSize                     = ($Mailbox.MaxSendSize -split ' \(')[0]
    MaxReceiveSize                  = ($Mailbox.MaxReceiveSize -split ' \(')[0]

    MB_DisplayName                  = $Stats.DisplayName
    MB_ItemCount                    = $Stats.ItemCount
    MB_TotalItemSizeMB              = (Bytes-ToMB $Stats.TotalItemSize)
    MB_DeletedItemCount             = $Stats.DeletedItemCount
    MB_TotalDeletedItemSizeMB       = (Bytes-ToMB $Stats.TotalDeletedItemSize)

    Archive_DisplayName             = $ArchStats.DisplayName
    Archive_ItemCount               = $ArchStats.ItemCount
    Archive_TotalItemSizeMB         = (Bytes-ToMB $ArchStats.TotalItemSize)
    Archive_DeletedItemCount        = $ArchStats.DeletedItemCount
    Archive_TotalDeletedItemSizeMB  = (Bytes-ToMB $ArchStats.TotalDeletedItemSize)

    ForwardingSmtpAddress           = (Resolve-RecipientToSMTP $Mailbox.ForwardingSmtpAddress)
    ForwardingAddress               = (Resolve-RecipientToSMTP $Mailbox.ForwardingAddress)
    DeliverToMailboxAndForward      = [bool]$Mailbox.DeliverToMailboxAndForward
    InboxRuleForwarding             = [bool]$HasInboxForward
    InboxRuleForwardingDetails      = ($InboxRuleRecipients -join '; ')
  }
}

function Update-Progress {
  param([int]$Current, [int]$Total)
  $pct = if ($Total -gt 0) { [int](($Current/$Total)*100) } else { 0 }
  $elapsed = (Get-Date) - $script:StartTime
  $rate = if ($Current -gt 0) { $elapsed.TotalSeconds / $Current } else { $null }
  $remaining = if ($rate) { [TimeSpan]::FromSeconds($rate * ($Total - $Current)) } else { [TimeSpan]::Zero }
  $eta = (Get-Date).Add($remaining)
  Write-Progress -Activity "Processing mailboxes..." -Status ("{0} of {1} ({2}%)" -f $Current, $Total, $pct) -PercentComplete $pct
  if ($script:UiProgressBar) { $script:UiProgressBar.Value = [math]::Min(100, [math]::Max(0,$pct)) }
  if ($script:UiEtaLabel) { $script:UiEtaLabel.Text = ("Elapsed: {0} | ETA: {1}" -f ([string]$elapsed), $eta.ToString('yyyy-MM-dd HH:mm')) }
}

function Summarize-And-Save {
  param($Rows, $OrgPolicy, $TransportRules, [int]$TopNLocal)

  $total = $Rows.Count
  $byType = $Rows | Group-Object RecipientTypeDetails | Sort-Object Count -Descending
  $inactive = ($Rows | Where-Object { $_.IsInactiveMailbox }).Count
  $archCount = ($Rows | Where-Object { $_.HasArchive -eq 'True' -or $_.HasArchive -eq 'Yes' }).Count
  $holdLit = ($Rows | Where-Object { $_.LitigationHoldEnabled }).Count
  $holdRet = ($Rows | Where-Object { $_.RetentionHoldEnabled }).Count
  $primaryGB = [Math]::Round((($Rows | Measure-Object -Property MB_TotalItemSizeMB -Sum).Sum)/1024,2)
  $archiveGB = [Math]::Round((($Rows | Measure-Object -Property Archive_TotalItemSizeMB -Sum).Sum)/1024,2)
  $fwMailbox = ($Rows | Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress }).Count
  $fwRules = ($Rows | Where-Object { $_.InboxRuleForwarding }).Count
  $oldest = ($Rows | Sort-Object WhenMailboxCreated | Select-Object -First 1).WhenMailboxCreated
  $newest = ($Rows | Sort-Object WhenMailboxCreated -Descending | Select-Object -First 1).WhenMailboxCreated

  $TopNLocal = [math]::Max(1,[math]::Min(1000,$TopNLocal))
  $topNRows = $Rows | Sort-Object MB_TotalItemSizeMB -Descending | Select-Object DisplayName,PrimarySmtpAddress,MB_TotalItemSizeMB -First $TopNLocal

  $sb = New-Object System.Text.StringBuilder
  $null = $sb.AppendLine("Exchange Online Discovery Summary")
  $null = $sb.AppendLine(("Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
  $null = $sb.AppendLine(("CSV: {0}" -f $script:CsvFile))
  $null = $sb.AppendLine(("Operational Log: {0}" -f $script:LogFile))
  $null = $sb.AppendLine(("Transcript: {0}" -f $script:TranscriptFile))
  if ($script:RulesCsv) { $null = $sb.AppendLine(("Transport Rules CSV: {0}" -f $script:RulesCsv)) }
  $null = $sb.AppendLine()
  $null = $sb.AppendLine("Counts")
  $null = $sb.AppendLine("------")
  $null = $sb.AppendLine(("Total mailboxes: {0}" -f $total))
  foreach ($g in $byType) { $null = $sb.AppendLine(("  {0}: {1}" -f $g.Name, $g.Count)) }
  $null = $sb.AppendLine(("Inactive mailboxes: {0}" -f $inactive))
  $null = $sb.AppendLine(("Archive enabled: {0}" -f $archCount))
  $null = $sb.AppendLine(("Litigation Hold enabled: {0}" -f $holdLit))
  $null = $sb.AppendLine(("Retention Hold enabled: {0}" -f $holdRet))
  $null = $sb.AppendLine(("Total primary size (GB): {0}" -f $primaryGB))
  $null = $sb.AppendLine(("Total archive size (GB): {0}" -f $archiveGB))
  $null = $sb.AppendLine()
  $null = $sb.AppendLine("Forwarding")
  $null = $sb.AppendLine("----------")
  $null = $sb.AppendLine(("Mailbox-level forwarding count: {0}" -f $fwMailbox))
  $null = $sb.AppendLine(("Inbox-rule forwarding count: {0}" -f $fwRules))
  $null = $sb.AppendLine(("Tenant AutoForwardEnabled: {0}" -f $($OrgPolicy.AutoForwardEnabled)))
  if ($TransportRules -and $TransportRules.Count -gt 0) {
    $null = $sb.AppendLine("Transport rules affecting forwarding/redirect:")
    foreach ($r in $TransportRules) { $null = $sb.AppendLine(("  [{0}] {1} [{2}]" -f $r.Priority, $r.Name, $r.State)) }
  } else {
    $null = $sb.AppendLine("No transport rules detected that explicitly redirect/route outbound.")
  }
  $null = $sb.AppendLine()
  $null = $sb.AppendLine("Mailbox creation range")
  $null = $sb.AppendLine("----------------------")
  $null = $sb.AppendLine(("Oldest: {0}" -f $oldest))
  $null = $sb.AppendLine(("Newest: {0}" -f $newest))
  $null = $sb.AppendLine()
  $null = $sb.AppendLine(("Top {0} largest (by primary, MB)" -f $TopNLocal))
  $null = $sb.AppendLine("---------------------------------------")
  foreach ($t in $topNRows) {
    $null = $sb.AppendLine(("  {0} <{1}> : {2} MB" -f $t.DisplayName, $t.PrimarySmtpAddress, $t.MB_TotalItemSizeMB))
  }

  $text = $sb.ToString()
  if ($script:SummaryFile) { Set-Content -Path $script:SummaryFile -Value $text -Encoding UTF8 }
  if ($script:UiSummaryBox) { $script:UiSummaryBox.Text = $text }
  Write-Log ("Summary saved to {0}" -f $script:SummaryFile) 'INFO'
}

function Confirm-Context {
  param([PSCustomObject]$Ctx, [string]$TenantDomain, [switch]$IsUI)
  $connected = ("Connected as: {0}`nOrganization: {1}`nAccepted domains: {2}" -f `
    ($Ctx.UserUPN ? $Ctx.UserUPN : '<unknown>'),
    ($TenantDomain ? $TenantDomain : ($Ctx.Organization ? $Ctx.Organization : '<unknown>')),
    ($Ctx.AcceptedDomains ? $Ctx.AcceptedDomains : '<unknown>'))
  Write-Log $connected 'INFO'
  if ($IsUI) {
    $res = [System.Windows.Forms.MessageBox]::Show($connected + "`n`nProceed with discovery?", "Confirm tenant", `
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
  } else {
    $ans = Read-Host ($connected + "`nProceed? (Y/N)")
    return ($ans -match '^[Yy]')
  }
}

function Run-Discovery {
  param(
    [string]$TenantDomain,
    [string]$OutputFolder,
    [switch]$SkipRulesScan,
    [int]$TopNLocal,
    [switch]$DoExportTransportRules,
    [ValidateSet('Auto','Browser','BrowserForce','Device')]
    [string]$AuthModeLocal = 'Auto',
    [string]$AuthUpnLocal
  )

  Advise-PSVersion
  Ensure-Module -Name ExchangeOnlineManagement -MinVersion '3.0.0'

  if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
  $base = "EXO-Discovery-$stamp-$pidTag"
  $script:CsvFile = Join-Path $OutputFolder "$base.csv"
  $script:SummaryFile = Join-Path $OutputFolder "$base-summary.txt"
  if ($DoExportTransportRules) { $script:RulesCsv = Join-Path $OutputFolder "$base-transport-rules.csv" } else { $script:RulesCsv = $null }

  Write-Log ("Output CSV: {0}" -f $script:CsvFile) 'INFO'
  Write-Log ("Summary TXT: {0}" -f $script:SummaryFile) 'INFO'
  if ($script:RulesCsv) { Write-Log ("Transport Rules CSV: {0}" -f $script:RulesCsv) 'INFO' }
  Write-Log "This process can take up to 30+ minutes on large tenants." 'INFO'

  Connect-EXO -DelegatedOrganization $TenantDomain -Mode $AuthModeLocal -Upn $AuthUpnLocal

  # Confirm we’re in the right tenant before doing heavy work
  $ctx = Get-ConnectedContext
  $proceed = Confirm-Context -Ctx $ctx -TenantDomain $TenantDomain -IsUI:$script:UiIsActive
  if (-not $proceed) {
    Write-Log "User cancelled after context confirmation. Disconnecting..." 'WARN'
    try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
    return
  }

  try { $null = Get-OrganizationConfig -ErrorAction Stop } catch { Write-Log ("Get-OrganizationConfig failed: {0}" -f $_.Exception.Message) 'WARN' }

  $orgPolicy = Get-OrgForwardingPolicy
  $transportRules = Get-TransportForwardingRules
  if ($script:RulesCsv -and $transportRules) {
    try {
      $transportRules | Export-Csv -Path $script:RulesCsv -NoTypeInformation
      Write-Log ("Exported transport rules ({0}) to {1}" -f $transportRules.Count, $script:RulesCsv) 'INFO'
    } catch { Write-Log ("Failed to export transport rules: {0}" -f $_.Exception.Message) 'ERROR' }
  }

  $rows = New-Object System.Collections.Generic.List[object]
  $all = Get-EXO-Mailboxes
  $total = ($all | Measure-Object).Count
  Write-Log ("Mailboxes discovered: {0}" -f $total) 'INFO'

  # Durable header
  $null = [PSCustomObject]@{
    UserPrincipalName=''; DisplayName=''; PrimarySmtpAddress=''; RecipientType='';
    RecipientTypeDetails=''; IsInactiveMailbox=''; AccountDisabled='';
    LitigationHoldEnabled=''; RetentionHoldEnabled=''; HasArchive='';
    WhenMailboxCreated=''; RetentionPolicy=''; MaxSendSize=''; MaxReceiveSize='';
    MB_DisplayName=''; MB_ItemCount=''; MB_TotalItemSizeMB=''; MB_DeletedItemCount='';
    MB_TotalDeletedItemSizeMB=''; Archive_DisplayName=''; Archive_ItemCount='';
    Archive_TotalItemSizeMB=''; Archive_DeletedItemCount=''; Archive_TotalDeletedItemSizeMB='';
    ForwardingSmtpAddress=''; ForwardingAddress=''; DeliverToMailboxAndForward='';
    InboxRuleForwarding=''; InboxRuleForwardingDetails=''
  } | Export-Csv -Path $script:CsvFile -NoTypeInformation

  $i = 0
  foreach ($mbx in $all) {
    $i++
    Update-Progress -Current $i -Total $total
    Write-Log ("Processing {0}" -f $mbx.UserPrincipalName) 'INFO'

    try {
      $includeSoft = $mbx.IsInactiveMailbox
      $stats = Safe-GetMailboxStats -ExchangeGuid $mbx.ExchangeGuid -IncludeSoftDeleted:$includeSoft
      $archStats = if ($mbx.ArchiveName) {
        Safe-GetMailboxStats -ExchangeGuid $mbx.ExchangeGuid -Archive -IncludeSoftDeleted:$includeSoft
      } else {
        [PSCustomObject]@{ DisplayName=$null; ItemCount=$null; TotalItemSize=$null; DeletedItemCount=$null; TotalDeletedItemSize=$null }
      }

      $fwdRecipients = @(); $hasFwd = $false
      if (-not $SkipRulesScan) {
        $rules = Safe-GetInboxRules -MailboxUPN $mbx.UserPrincipalName
        foreach ($r in $rules) {
          $recips = @()
          if ($r.ForwardTo)             { $recips += ($r.ForwardTo             | ForEach-Object { Resolve-RecipientToSMTP $_ }) }
          if ($r.ForwardAsAttachmentTo) { $recips += ($r.ForwardAsAttachmentTo | ForEach-Object { Resolve-RecipientToSMTP $_ }) }
          if ($r.RedirectTo)            { $recips += ($r.RedirectTo            | ForEach-Object { Resolve-RecipientToSMTP $_ }) }
          if ($recips.Count -gt 0) { $hasFwd = $true; $fwdRecipients += $recips }
        }
      }

      $row = Build-CSVRow -Mailbox $mbx -Stats $stats -ArchStats $archStats -InboxRuleRecipients $fwdRecipients -HasInboxForward:$hasFwd
      $rows.Add($row) | Out-Null
      $row | Export-Csv -Path $script:CsvFile -NoTypeInformation -Append
    } catch {
      Write-Log ("Error processing {0}: {1}" -f $mbx.UserPrincipalName, $_.Exception.Message) 'ERROR'
      [PSCustomObject]@{
        UserPrincipalName = $mbx.UserPrincipalName
        DisplayName = $mbx.DisplayName
        PrimarySmtpAddress = $mbx.PrimarySmtpAddress
        RecipientType = $mbx.RecipientType
        RecipientTypeDetails = $mbx.RecipientTypeDetails
        IsInactiveMailbox = $mbx.IsInactiveMailbox
        AccountDisabled = $mbx.AccountDisabled
        LitigationHoldEnabled = $mbx.LitigationHoldEnabled
        RetentionHoldEnabled = $mbx.RetentionHoldEnabled
        HasArchive = [string]([bool]$mbx.ArchiveName)
        WhenMailboxCreated = $mbx.WhenMailboxCreated
        RetentionPolicy = $mbx.RetentionPolicy
        MaxSendSize = ($mbx.MaxSendSize -split ' \(')[0]
        MaxReceiveSize = ($mbx.MaxReceiveSize -split ' \(')[0]
        MB_DisplayName = 'ERROR'
        MB_ItemCount = $null
        MB_TotalItemSizeMB = $null
        MB_DeletedItemCount = $null
        MB_TotalDeletedItemSizeMB = $null
        Archive_DisplayName = $null
        Archive_ItemCount = $null
        Archive_TotalItemSizeMB = $null
        Archive_DeletedItemCount = $null
        Archive_TotalDeletedItemSizeMB = $null
        ForwardingSmtpAddress = (Resolve-RecipientToSMTP $mbx.ForwardingSmtpAddress)
        ForwardingAddress = (Resolve-RecipientToSMTP $mbx.ForwardingAddress)
        DeliverToMailboxAndForward = [bool]$mbx.DeliverToMailboxAndForward
        InboxRuleForwarding = $false
        InboxRuleForwardingDetails = ''
      } | Export-Csv -Path $script:CsvFile -NoTypeInformation -Append
      continue
    }
  }

  Summarize-And-Save -Rows $rows -OrgPolicy $orgPolicy -TransportRules $transportRules -TopNLocal $TopNLocal

  try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
  Write-Log "Discovery complete." 'INFO'
  Update-Progress -Current $total -Total $total
}

#------------------------------ UI (WinForms) ---------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-UI {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "$script:AppName"
  $form.Size = New-Object System.Drawing.Size(900, 860)
  $form.StartPosition = 'CenterScreen'

  $lblTenant = New-Object System.Windows.Forms.Label
  $lblTenant.Text = "Tenant Domain (optional for delegated admin):"
  $lblTenant.Location = New-Object System.Drawing.Point(10,15)
  $lblTenant.AutoSize = $true
  $form.Controls.Add($lblTenant)

  $txtTenant = New-Object System.Windows.Forms.TextBox
  $txtTenant.Location = New-Object System.Drawing.Point(10,35)
  $txtTenant.Width = 400
  $txtTenant.Anchor = 'Top,Left'
  $form.Controls.Add($txtTenant)

  $lblOut = New-Object System.Windows.Forms.Label
  $lblOut.Text = "Output Folder:"
  $lblOut.Location = New-Object System.Drawing.Point(10,70)
  $lblOut.AutoSize = $true
  $form.Controls.Add($lblOut)

  $txtOut = New-Object System.Windows.Forms.TextBox
  $txtOut.Location = New-Object System.Drawing.Point(10,90)
  $txtOut.Width = 750
  $txtOut.Text = $script:WorkRoot
  $txtOut.Anchor = 'Top,Left,Right'
  $form.Controls.Add($txtOut)

  $btnBrowse = New-Object System.Windows.Forms.Button
  $btnBrowse.Text = "Browse..."
  $btnBrowse.Location = New-Object System.Drawing.Point(770,88)
  $btnBrowse.Anchor = 'Top,Right'
  $btnBrowse.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') { $txtOut.Text = $fbd.SelectedPath }
  })
  $form.Controls.Add($btnBrowse)

  $lblWarn = New-Object System.Windows.Forms.Label
  $lblWarn.Location = New-Object System.Drawing.Point(10,125)
  $lblWarn.AutoSize = $true
  $lblWarn.ForeColor = [System.Drawing.Color]::DarkOrange
  $lblWarn.Text = if ($script:SessionIsPS7) { "PowerShell $($PSVersionTable.PSVersion) detected." } else { "Warning: PowerShell $($PSVersionTable.PSVersion) — PS 7+ is recommended." }
  $form.Controls.Add($lblWarn)

  # Authentication group (polished spacing)
  $grpAuth = New-Object System.Windows.Forms.GroupBox
  $grpAuth.Text = "Authentication"
  $grpAuth.Location = New-Object System.Drawing.Point(10,150)
  $grpAuth.Size = New-Object System.Drawing.Size(860, 115)
  $grpAuth.Anchor = 'Top,Left,Right'
  $form.Controls.Add($grpAuth)

  $rbAuto = New-Object System.Windows.Forms.RadioButton
  $rbAuto.Text = "Auto (SSO if available)"
  $rbAuto.Location = New-Object System.Drawing.Point(10,25)
  $rbAuto.AutoSize = $true
  $rbAuto.Checked = $true
  $grpAuth.Controls.Add($rbAuto)

  $rbBrowser = New-Object System.Windows.Forms.RadioButton
  $rbBrowser.Text = "Browser (account picker)"
  $rbBrowser.Location = New-Object System.Drawing.Point(220,25)
  $rbBrowser.AutoSize = $true
  $grpAuth.Controls.Add($rbBrowser)

  $rbBrowserForce = New-Object System.Windows.Forms.RadioButton
  $rbBrowserForce.Text = "Browser (force picker)"
  $rbBrowserForce.Location = New-Object System.Drawing.Point(440,25)
  $rbBrowserForce.AutoSize = $true
  $grpAuth.Controls.Add($rbBrowserForce)

  $rbDevice = New-Object System.Windows.Forms.RadioButton
  $rbDevice.Text = "Device code (always prompts)"
  $rbDevice.Location = New-Object System.Drawing.Point(640,25)
  $rbDevice.AutoSize = $true
  $grpAuth.Controls.Add($rbDevice)

  $lblUpn = New-Object System.Windows.Forms.Label
  $lblUpn.Text = "UPN (optional, used with Browser/Force):"
  $lblUpn.Location = New-Object System.Drawing.Point(10,60)
  $lblUpn.AutoSize = $true
  $grpAuth.Controls.Add($lblUpn)

  $txtUpn = New-Object System.Windows.Forms.TextBox
  $txtUpn.Location = New-Object System.Drawing.Point(260,58)
  $txtUpn.Width = 580
  $txtUpn.Enabled = $false
  $txtUpn.Anchor = 'Top,Left,Right'
  $grpAuth.Controls.Add($txtUpn)

  # Enable UPN only if Browser or BrowserForce selected
  $rbBrowser.Add_CheckedChanged({ $txtUpn.Enabled = ($rbBrowser.Checked -or $rbBrowserForce.Checked) })
  $rbBrowserForce.Add_CheckedChanged({ $txtUpn.Enabled = ($rbBrowser.Checked -or $rbBrowserForce.Checked) })

  # Options row
  $chkSkipRules = New-Object System.Windows.Forms.CheckBox
  $chkSkipRules.Text = "Skip Inbox-rule scans (faster)"
  $chkSkipRules.Checked = $false
  $chkSkipRules.Location = New-Object System.Drawing.Point(10,275)
  $chkSkipRules.AutoSize = $true
  $form.Controls.Add($chkSkipRules)

  $chkExportRules = New-Object System.Windows.Forms.CheckBox
  $chkExportRules.Text = "Export transport rules to CSV"
  $chkExportRules.Checked = $true
  $chkExportRules.Location = New-Object System.Drawing.Point(240,275)
  $chkExportRules.AutoSize = $true
  $form.Controls.Add($chkExportRules)

  $lblTopN = New-Object System.Windows.Forms.Label
  $lblTopN.Text = "Top-N largest mailboxes in summary:"
  $lblTopN.Location = New-Object System.Drawing.Point(500,275)
  $lblTopN.AutoSize = $true
  $form.Controls.Add($lblTopN)

  $numTopN = New-Object System.Windows.Forms.NumericUpDown
  $numTopN.Location = New-Object System.Drawing.Point(740,273)
  $numTopN.Minimum = 1
  $numTopN.Maximum = 1000
  $numTopN.Value = [decimal]$TopN
  $form.Controls.Add($numTopN)

  $pb = New-Object System.Windows.Forms.ProgressBar
  $pb.Location = New-Object System.Drawing.Point(10,305)
  $pb.Width = 860
  $pb.Style = 'Continuous'
  $pb.Anchor = 'Top,Left,Right'
  $form.Controls.Add($pb)
  $script:UiProgressBar = $pb

  $lblEta = New-Object System.Windows.Forms.Label
  $lblEta.Location = New-Object System.Drawing.Point(10,335)
  $lblEta.AutoSize = $true
  $lblEta.Text = "Elapsed: 00:00:00 | ETA: --"
  $form.Controls.Add($lblEta)
  $script:UiEtaLabel = $lblEta

  $log = New-Object System.Windows.Forms.TextBox
  $log.Location = New-Object System.Drawing.Point(10,360)
  $log.Size = New-Object System.Drawing.Size(860, 245)
  $log.Multiline = $true
  $log.ScrollBars = 'Vertical'
  $log.ReadOnly = $true
  $log.Anchor = 'Top,Left,Right'
  $form.Controls.Add($log)
  $script:UiLogBox = $log

  $btnStart = New-Object System.Windows.Forms.Button
  $btnStart.Text = "Start Discovery"
  $btnStart.Location = New-Object System.Drawing.Point(10,615)
  $form.Controls.Add($btnStart)
  $script:UiStartBtn = $btnStart

  $lblSummary = New-Object System.Windows.Forms.Label
  $lblSummary.Text = "Discovery Summary (copy/paste):"
  $lblSummary.Location = New-Object System.Drawing.Point(10,650)
  $lblSummary.AutoSize = $true
  $form.Controls.Add($lblSummary)

  $summary = New-Object System.Windows.Forms.TextBox
  $summary.Location = New-Object System.Drawing.Point(10,670)
  $summary.Size = New-Object System.Drawing.Size(860, 150)
  $summary.Multiline = $true
  $summary.ScrollBars = 'Vertical'
  $summary.ReadOnly = $false
  $summary.Anchor = 'Top,Left,Right'
  $form.Controls.Add($summary)
  $script:UiSummaryBox = $summary

  $btnStart.Add_Click({
    try {
      $btnStart.Enabled = $false
      Write-Log "Starting discovery..." 'INFO'

      $mode = 'Auto'
      if ($rbBrowser.Checked)      { $mode = 'Browser' }
      if ($rbBrowserForce.Checked) { $mode = 'BrowserForce' }
      if ($rbDevice.Checked)       { $mode = 'Device' }

      Run-Discovery -TenantDomain $txtTenant.Text -OutputFolder $txtOut.Text `
        -SkipRulesScan:$($chkSkipRules.Checked) `
        -TopNLocal ([int]$numTopN.Value) `
        -DoExportTransportRules:$($chkExportRules.Checked) `
        -AuthModeLocal $mode `
        -AuthUpnLocal $($txtUpn.Text)

      [System.Windows.Forms.MessageBox]::Show("Discovery complete.`nCSV: $script:CsvFile`nSummary: $script:SummaryFile`nOperational Log: $script:LogFile`nTranscript: $script:TranscriptFile" + $(if($script:RulesCsv){"`nTransport Rules: $script:RulesCsv"}),"Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
      Write-Log ("Fatal error: {0}" -f $_.Exception.Message) 'ERROR'
      [System.Windows.Forms.MessageBox]::Show(("A fatal error occurred. See log:`n{0}" -f $script:LogFile),"Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
      $btnStart.Enabled = $true
    }
  })

  $form.Add_Shown({ $script:UiIsActive = $true; $form.Activate() })
  [void]$form.ShowDialog()
}

#------------------------------ Entrypoint ------------------------------------
try {
  if ($ConsoleOnly) {
    Write-Log "Console mode selected." 'INFO'
    Advise-PSVersion
    $out = Read-Host ("Output folder (default: {0})" -f $script:WorkRoot)
    if ([string]::IsNullOrWhiteSpace($out)) { $out = $script:WorkRoot }
    $tenant = Read-Host "Tenant domain for delegated admin (optional, e.g. contoso.onmicrosoft.com)"
    $auth = Read-Host "Auth mode? Auto/Browser/BrowserForce/Device (default Auto)"
    if ([string]::IsNullOrWhiteSpace($auth)) { $auth = 'Auto' }
    $upn  = $null
    if ($auth -match '^(Browser|BrowserForce)$') { $upn = Read-Host "Optional UPN for browser login (press Enter to skip)" }
    $skip = Read-Host "Skip Inbox-rule scans? (Y/N, default N)"
    $expR = Read-Host "Export transport rules to CSV? (Y/N, default Y)"
    $nTop = Read-Host ("Top-N largest in summary? (default {0})" -f $TopN)

    if ([string]::IsNullOrWhiteSpace($nTop)) { $nTop = $TopN } else { $nTop = [int]$nTop }
    $doSkip = $false; if ($skip -match '^[Yy]') { $doSkip = $true }
    $doExport = $true; if ($expR -match '^[Nn]') { $doExport = $false }

    Run-Discovery -TenantDomain $tenant -OutputFolder $out `
      -SkipRulesScan:$doSkip -TopNLocal $nTop -DoExportTransportRules:$doExport `
      -AuthModeLocal $auth -AuthUpnLocal $upn

    Write-Host ("CSV: {0}" -f $script:CsvFile)
    Write-Host ("Summary: {0}" -f $script:SummaryFile)
    if ($script:RulesCsv) { Write-Host ("Transport Rules: {0}" -f $script:RulesCsv) }
    Write-Host ("Operational Log: {0}" -f $script:LogFile)
    Write-Host ("Transcript: {0}" -f $script:TranscriptFile)
  } else {
    Show-UI
  }
} catch {
  Write-Log ("Unhandled exception: {0}" -f $_.Exception.Message) 'ERROR'
} finally {
  try { Stop-Transcript | Out-Null } catch {}
}
