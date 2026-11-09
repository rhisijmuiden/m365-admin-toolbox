# M365 Admin Toolbox: Security & Compliance helper (compliance search and purge).
# Runs as a long-lived Windows PowerShell process per tenant. The app writes one JSON request per line
# to stdin ({ id, op, args }) and reads one marked JSON line per reply from stdout. Only the Tb-* functions
# below can be called, with arguments splatted from the request, so nothing is ever evaluated as code.

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding $false
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$MARK = '@@TB@@'
$script:Connected = $false

function Out-Reply($obj) {
  [Console]::Out.WriteLine($MARK + (ConvertTo-Json -InputObject $obj -Depth 8 -Compress))
  [Console]::Out.Flush()
}

function ConvertTo-Hash($o) {
  if ($null -eq $o) { return $null }
  if ($o -is [System.Management.Automation.PSCustomObject]) {
    $h = @{}
    foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = ConvertTo-Hash $p.Value }
    return $h
  }
  if ($o -is [System.Collections.IEnumerable] -and $o -isnot [string]) { return , @($o | ForEach-Object { ConvertTo-Hash $_ }) }
  return $o
}

function Iso($d) { if ($d) { ([datetime]$d).ToUniversalTime().ToString('o') } else { $null } }

# EXO returns sizes as "1.5 GB (1,610,612,736 bytes)" or "Unlimited".
function Bytes($v) {
  $s = "$v"
  if ($s -match '\(([\d,\.\s]+) bytes\)') { return [int64]($matches[1] -replace '[^\d]', '') }
  return $null
}

function Strs($v) { , @($v | Where-Object { $_ } | ForEach-Object { "$_" }) }

function Smtp($addrs) { , @($addrs | ForEach-Object { "$_" } | Where-Object { $_ -match '^smtp:' }) }

# ---------- session ----------

function Tb-Init {
  $m = Get-Module -ListAvailable ExchangeOnlineManagement | Where-Object { $_.Version -ge [version]'3.2.0' } | Sort-Object Version -Descending | Select-Object -First 1
  $installed = $false
  if (-not $m) {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge '2.8.5.201' })) {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module ExchangeOnlineManagement -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -MinimumVersion 3.2.0
    $installed = $true
    $m = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
  }
  Import-Module $m.Path -DisableNameChecking
  @{ version = "$($m.Version)"; installed = $installed }
}

# Which token audience and session type Connect-IPPSSession accepts with -AccessToken differs between
# module versions, so the variants are tried in order and the first that connects is kept.
function Tb-Connect([string[]]$Tokens, [string]$Org) {
  if ($script:Connected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null; $script:Connected = $false }
  $errors = @()
  foreach ($tok in @($Tokens | Where-Object { $_ })) {
    foreach ($searchOnly in $true, $false) {
      try {
        $p = @{ AccessToken = $tok; Organization = $Org; ShowBanner = $false }
        if ($searchOnly) { $p.EnableSearchOnlySession = $true }
        Connect-IPPSSession @p -WarningAction SilentlyContinue | Out-Null
        if (-not (Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue)) { throw 'The compliance search cmdlets are not available for this account' }
        $script:Connected = $true
        return @{ ok = $true; searchOnly = $searchOnly }
      } catch {
        $errors += $_.Exception.Message
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
      }
    }
  }
  throw ('Could not connect to Security & Compliance PowerShell: ' + ((@($errors) | Select-Object -Unique -First 2) -join ' | '))
}

# ---------- compliance search ----------

function Tb-StartSearch([string]$Name, [string]$Query, [string[]]$Locations) {
  $loc = if ($Locations -and $Locations.Count) { $Locations } else { 'All' }
  New-ComplianceSearch -Name $Name -ExchangeLocation $loc -ContentMatchQuery $Query -Confirm:$false | Out-Null
  Start-ComplianceSearch -Identity $Name
  @{ name = $Name }
}

function Tb-RestartSearch([string]$Name) { Start-ComplianceSearch -Identity $Name; @{ name = $Name } }

# SuccessResults lists each mailbox with hits as "Location: x, Item count: n, Total size: s".
function Tb-SearchStatus([string]$Name) {
  $s = Get-ComplianceSearch -Identity $Name
  $hits = @()
  foreach ($m in [regex]::Matches("$($s.SuccessResults)", 'Location:\s*([^,]+),\s*Item count:\s*(\d+),\s*Total size:\s*(\d+)')) {
    if ([int]$m.Groups[2].Value -gt 0) { $hits += @{ location = $m.Groups[1].Value.Trim(); items = [int]$m.Groups[2].Value; size = [int64]$m.Groups[3].Value } }
  }
  @{ status = "$($s.Status)"; items = $s.Items; size = $s.Size; mailboxes = $hits; errors = "$($s.Errors)"; jobProgress = $s.JobProgress }
}

function Tb-Purge([string]$Name, [string]$Type) {
  if ($Type -ne 'SoftDelete' -and $Type -ne 'HardDelete') { throw "Unknown purge type $Type" }
  $a = New-ComplianceSearchAction -SearchName $Name -Purge -PurgeType $Type -Confirm:$false -Force
  @{ action = "$($a.Name)" }
}

function Tb-ActionStatus([string]$Action) {
  $a = Get-ComplianceSearchAction -Identity $Action
  @{ status = "$($a.Status)"; results = "$($a.Results)"; errors = "$($a.Errors)" }
}

function Tb-RemoveSearch([string]$Name) {
  Get-ComplianceSearchAction -ErrorAction SilentlyContinue | Where-Object { "$($_.SearchName)" -eq $Name } | ForEach-Object { Remove-ComplianceSearchAction -Identity "$($_.Identity)" -Confirm:$false -ErrorAction SilentlyContinue }
  Remove-ComplianceSearch -Identity $Name -Confirm:$false
  @{ ok = $true }
}

# ---------- loop ----------

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  if (-not $line.Trim()) { continue }
  $req = $null
  try {
    $req = ConvertFrom-Json $line
    $op = "$($req.op)"
    if ($op -notmatch '^Tb-[A-Za-z]+$' -or -not (Get-Command $op -CommandType Function -ErrorAction SilentlyContinue)) { throw "Unknown operation $op" }
    $a = ConvertTo-Hash $req.args
    if ($null -eq $a) { $a = @{} }
    $data = & $op @a
    Out-Reply @{ id = $req.id; ok = $true; data = $data }
  } catch {
    $msg = $_.Exception.Message
    $code = $null
    if ($msg -match 'is not recognized as the name of a cmdlet|isn''t within your current write scope|UnAuthorized|AccessDenied|Access Denied|not have permission') { $code = 'EXO_DENIED' }
    elseif ($msg -match '^Could not connect to Security') { $code = 'SCC_CONNECT' }
    Out-Reply @{ id = $(if ($req) { $req.id } else { $null }); ok = $false; error = $msg; code = $code }
  }
}
