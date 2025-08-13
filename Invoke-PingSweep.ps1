@'
# Invoke-PingSweep.ps1
# Ping sweep (CIDR or Base/Start/End) with latency ms, optional progress, and safe concurrency via thread jobs.

[CmdletBinding()]
param(
  [string]$CIDR,                 # e.g. 74.115.234.0/24
  [string]$Base,                 # e.g. 74.115.234
  [int]$Start = 1,
  [int]$End   = 255,
  [int]$TimeoutMs = 500,
  [int]$Concurrency = 64,        # number of concurrent pings (jobs)
  [switch]$All,                  # include DOWN hosts
  [switch]$Progress,             # show a progress bar (sequential)
  [string]$Csv                   # optional CSV export path
)

# ---- guard rails ----
if ($CIDR -and ($Base -or $PSBoundParameters.ContainsKey('Start') -or $PSBoundParameters.ContainsKey('End'))) {
  throw "Use either -CIDR OR -Base/-Start/-End, not both."
}
if (-not $CIDR -and -not $Base) {
  throw "Provide -CIDR (e.g. 192.168.1.0/24) OR -Base/-Start/-End (e.g. -Base 192.168.1 -Start 10 -End 200)."
}
if ($Base) {
  if ($End -lt $Start) { throw "-End must be >= -Start." }
  $octs = $Base.Split('.')
  if ($octs.Count -ne 3 -or ($octs | Where-Object { $_ -notmatch "^\d+$" -or [int]$_ -lt 0 -or [int]$_ -gt 255 })) {
    throw "-Base must look like 192.168.1 (three octets 0–255)."
  }
}

# ---- helpers ----
function ConvertTo-UInt32 { param([System.Net.IPAddress]$Ip) $b=$Ip.GetAddressBytes(); [array]::Reverse($b); [BitConverter]::ToUInt32($b,0) }
function ConvertFrom-UInt32 { param([uint32]$U) $b=[BitConverter]::GetBytes($U); [array]::Reverse($b); [System.Net.IPAddress]::new($b).ToString() }
function Parse-CIDR {
  param([string]$cidr)
  $ipStr,$prefixStr = $cidr -split '/'
  $ipObj = $null
  if (-not [System.Net.IPAddress]::TryParse($ipStr, [ref]$ipObj) -or
      $ipObj.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { throw "CIDR IP must be valid IPv4." }
  $prefix = [int]$prefixStr
  if ($prefix -lt 0 -or $prefix -gt 32) { throw "CIDR prefix must be 0–32." }
  $ipU = ConvertTo-UInt32 $ipObj
  $mask = if ($prefix -eq 0) { [uint32]0 } else { ([uint32]0..($prefix-1) | ForEach-Object { }) ; $m=[uint32]0; for($i=0;$i -lt $prefix;$i++){$m=($m -shl 1) -bor 1}; $m -shl (32-$prefix) }
  $network   = $ipU -band $mask
  $broadcast = $network -bor (-bnot $mask)
  if ($prefix -ge 31) { $start=$network; $end=$broadcast } else { $start=$network+1; $end=$broadcast-1 }
  for ($u=$start; $u -le $end; $u++) { ConvertFrom-UInt32 $u }
}

# ---- build target list ----
$addresses   = if ($CIDR) { Parse-CIDR $CIDR } else { $Start..$End | ForEach-Object { "$Base.$_" } }
$targetCount = $addresses.Count

# ---- worker (runs inside jobs too) ----
$jobScript = {
  param($ip,$timeout,$includeAll)
  try {
    $p = [System.Net.NetworkInformation.Ping]::new()
    $reply = $p.Send($ip, $timeout)
    if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
      [pscustomobject]@{ IP=$ip; Status='Up'; LatencyMs=[int]$reply.RoundtripTime }
    } elseif ($includeAll) {
      [pscustomobject]@{ IP=$ip; Status='Down'; LatencyMs=$null }
    }
  } catch {
    if ($includeAll) { [pscustomobject]@{ IP=$ip; Status='Down'; LatencyMs=$null } }
  }
}

Write-Host "Scanning $targetCount host(s) with concurrency $Concurrency..."

$results = @()
if ($Concurrency -le 1) {
  # Sequential (optional progress)
  $i = 0
  foreach ($ip in $addresses) {
    $i++
    if ($Progress) {
      $pct = [int](($i / $targetCount) * 100)
      Write-Progress -Activity "Pinging hosts" -Status "$i of $targetCount ($ip)" -PercentComplete $pct
    }
    $r = & $jobScript $ip $TimeoutMs $All
    if ($r) { $results += $r }
  }
  if ($Progress) { Write-Progress -Activity "Pinging hosts" -Completed }
} else {
  # Concurrent via thread jobs (fallback to process jobs if thread jobs unavailable)
  $useThreadJobs = [bool](Get-Command Start-ThreadJob -ErrorAction SilentlyContinue)
  $jobs = @()
  foreach ($ip in $addresses) {
    $jobs += if ($useThreadJobs) {
      Start-ThreadJob -ScriptBlock $jobScript -ArgumentList $ip,$TimeoutMs,$All
    } else {
      Start-Job       -ScriptBlock $jobScript -ArgumentList $ip,$TimeoutMs,$All
    }
    while ($jobs.Count -ge $Concurrency) {
      $done = Wait-Job -Job $jobs -Any
      $results += Receive-Job -Job $done
      Remove-Job  -Job $done
      $jobs = $jobs | Where-Object { $_.State -eq 'Running' }
    }
  }
  if ($jobs.Count) {
    Wait-Job -Job $jobs
    $results += Receive-Job -Job $jobs
    Remove-Job  -Job $jobs
  }
}

# ---- sort numerically by IP and output
$results = $results | Where-Object { $_ } | Sort-Object {
  $o=$_.IP.Split('.') | ForEach-Object {[int]$_}
  ($o[0] -shl 24) -bor ($o[1] -shl 16) -bor ($o[2] -shl 8) -bor $o[3]
}

$results | Format-Table -AutoSize

if ($Csv) {
  $results | Export-Csv -NoTypeInformation -Path $Csv
  Write-Host "Saved results to $Csv"
}
'@ | Set-Content -Encoding utf8 -NoNewline -Path .\Invoke-PingSweep.ps1
