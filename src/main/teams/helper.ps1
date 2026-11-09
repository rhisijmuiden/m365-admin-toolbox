# M365 Admin Toolbox: Microsoft Teams helper.
# Same protocol as the Exchange helper: one JSON request per line on stdin ({ id, op, args }) and one marked
# JSON line per reply on stdout. Only the Tb-* functions below can be called, with arguments splatted from
# the request, so nothing is ever evaluated as code. Teams, channels and members go through Graph in the
# app itself; this helper covers what only the MicrosoftTeams module can do: policies, external and guest
# access, and phone numbers.

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

function Strs($v) { , @($v | Where-Object { $_ } | ForEach-Object { "$_" }) }

# Policy names come back as "Tag:Name" in some module versions; an empty value means the Global policy.
function PolicyName($v) { $s = "$v" -replace '^Tag:', ''; if ($s) { $s } else { $null } }

# ---------- session ----------

function Tb-Init {
  $m = Get-Module -ListAvailable MicrosoftTeams | Where-Object { $_.Version.Major -ge 5 } | Sort-Object Version -Descending | Select-Object -First 1
  $installed = $false
  if (-not $m) {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge '2.8.5.201' })) {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module MicrosoftTeams -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -MinimumVersion 5.0.0
    $installed = $true
    $m = Get-Module -ListAvailable MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
  }
  Import-Module $m.Path -DisableNameChecking
  @{ version = "$($m.Version)"; installed = $installed }
}

function Tb-Connect([string]$GraphToken, [string]$AdminToken) {
  if ($script:Connected) { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null }
  Connect-MicrosoftTeams -AccessTokens @($GraphToken, $AdminToken) | Out-Null
  $script:Connected = $true
  @{ ok = $true }
}

# ---------- policies ----------

$PolicyTypes = 'TeamsMeetingPolicy', 'TeamsMessagingPolicy', 'TeamsAppSetupPolicy', 'TeamsCallingPolicy', 'OnlineVoiceRoutingPolicy'

function Assert-Type([string]$Type) { if ($PolicyTypes -notcontains $Type) { throw "Policy type $Type is not allowed" } }

function CsUserOut($u) {
  $p = @{}
  foreach ($t in $PolicyTypes) { $p[$t] = PolicyName $u.$t }
  @{
    id              = "$($u.Identity)"
    upn             = "$($u.UserPrincipalName)"
    displayName     = "$($u.DisplayName)"
    enabled         = [bool]$u.AccountEnabled
    policies        = $p
    lineUri         = if ($u.LineUri) { "$($u.LineUri)" } else { $null }
    enterpriseVoice = [bool]$u.EnterpriseVoiceEnabled
    phoneSystem     = @($u.FeatureTypes | Where-Object { "$_" -eq 'PhoneSystem' }).Count -gt 0
  }
}

function Get-Users {
  try { $list = Get-CsOnlineUser -AccountType User } catch { $list = Get-CsOnlineUser }
  , @($list | Where-Object { $_.UserPrincipalName } | ForEach-Object { CsUserOut $_ })
}

function Tb-Policies {
  $catalog = @{}
  foreach ($t in $PolicyTypes) {
    $catalog[$t] = @(& "Get-Cs$t" | ForEach-Object {
        $id = "$($_.Identity)" -replace '^Tag:', ''
        @{ name = $id; description = "$($_.Description)"; isGlobal = ($id -eq 'Global') }
      })
  }
  $groups = @()
  try {
    $groups = @(Get-CsGroupPolicyAssignment | ForEach-Object {
        @{ groupId = "$($_.GroupId)"; type = "$($_.PolicyType)"; policy = PolicyName $_.PolicyName; rank = $_.Priority }
      })
  } catch {}
  @{ catalog = $catalog; users = Get-Users; groups = $groups }
}

function Tb-User([string]$Identity) { CsUserOut (Get-CsOnlineUser -Identity $Identity) }

# Grants one policy to several users. An empty Policy resets them to the Global (org-wide default) policy.
function Tb-GrantUsers([string]$Type, [string]$Policy, [string[]]$Users) {
  Assert-Type $Type
  $name = if ($Policy) { $Policy } else { $null }
  $failed = @()
  $done = @()
  foreach ($u in @($Users | Where-Object { $_ })) {
    try {
      & "Grant-Cs$Type" -Identity $u -PolicyName $name
      $done += Tb-User $u
    } catch { $failed += @{ user = $u; error = $_.Exception.Message } }
  }
  @{ users = $done; failed = $failed }
}

function Tb-GrantGroup([string]$Type, [string]$Policy, [string]$GroupId, [int]$Rank) {
  Assert-Type $Type
  New-CsGroupPolicyAssignment -GroupId $GroupId -PolicyType $Type -PolicyName $Policy -Rank $Rank | Out-Null
  @{ ok = $true }
}

function Tb-RemoveGroupPolicy([string]$Type, [string]$GroupId) {
  Assert-Type $Type
  Remove-CsGroupPolicyAssignment -GroupId $GroupId -PolicyType $Type | Out-Null
  @{ ok = $true }
}

# ---------- external and guest access ----------

function Tb-Access {
  $f = Get-CsTenantFederationConfiguration
  $ad = $f.AllowedDomains
  $allowed = Strs @($ad.AllowedDomain | ForEach-Object { $_.Domain })
  $allowAll = (-not $allowed.Count) -and ("$ad" -match 'AllowAllKnownDomains' -or ($ad -and $ad.GetType().Name -match 'AllowAll'))
  $blocked = Strs @($f.BlockedDomains | ForEach-Object { if ($_.Domain) { $_.Domain } else { $_ } })
  $c = Get-CsTeamsClientConfiguration -Identity Global
  $gc = Get-CsTeamsGuestCallingConfiguration -Identity Global
  $gm = Get-CsTeamsGuestMeetingConfiguration -Identity Global
  $gs = Get-CsTeamsGuestMessagingConfiguration -Identity Global
  @{
    federation = @{
      allowFederated  = [bool]$f.AllowFederatedUsers
      allowAll        = [bool]$allowAll
      allowed         = $allowed
      blocked         = $blocked
      consumer        = [bool]$f.AllowTeamsConsumer
      consumerInbound = [bool]$f.AllowTeamsConsumerInbound
    }
    guest      = @{
      allowGuest     = [bool]$c.AllowGuestUser
      privateCalling = [bool]$gc.AllowPrivateCalling
      ipVideo        = [bool]$gm.AllowIPVideo
      meetNow        = [bool]$gm.AllowMeetNow
      screenSharing  = "$($gm.ScreenSharingMode)"
      chat           = [bool]$gs.AllowUserChat
      editMessages   = [bool]$gs.AllowUserEditMessage
      deleteMessages = [bool]$gs.AllowUserDeleteMessage
      giphy          = [bool]$gs.AllowGiphy
    }
  }
}

function Tb-SetFederation([bool]$AllowFederated, [bool]$AllowAll, [string[]]$Allowed, [string[]]$Blocked, [bool]$Consumer, [bool]$ConsumerInbound) {
  $p = @{ AllowFederatedUsers = $AllowFederated; AllowTeamsConsumer = $Consumer }
  if ($Consumer) { $p.AllowTeamsConsumerInbound = $ConsumerInbound }
  if ($AllowAll) { $p.AllowedDomains = New-CsEdgeAllowAllKnownDomains }
  else {
    $doms = @($Allowed | Where-Object { $_ } | ForEach-Object { New-CsEdgeDomainPattern -Domain $_ })
    if (-not $doms.Count) { throw 'Add at least one allowed domain, or allow all domains' }
    $p.AllowedDomains = New-CsEdgeAllowList -AllowedDomain $doms
  }
  Set-CsTenantFederationConfiguration @p
  # The block list is changed with Add/Remove of domain patterns, the form every module version accepts.
  $cur = Strs @((Get-CsTenantFederationConfiguration).BlockedDomains | ForEach-Object { if ($_.Domain) { $_.Domain } else { $_ } })
  $want = Strs $Blocked
  foreach ($d in $want) { if ($cur -notcontains $d) { Set-CsTenantFederationConfiguration -BlockedDomains @{ Add = (New-CsEdgeDomainPattern -Domain $d) } } }
  foreach ($d in $cur) { if ($want -notcontains $d) { Set-CsTenantFederationConfiguration -BlockedDomains @{ Remove = (New-CsEdgeDomainPattern -Domain $d) } } }
  Tb-Access
}

function Tb-SetGuest($Params) {
  $g = $Params
  if ($g.ContainsKey('allowGuest')) { Set-CsTeamsClientConfiguration -Identity Global -AllowGuestUser ([bool]$g.allowGuest) }
  if ($g.ContainsKey('privateCalling')) { Set-CsTeamsGuestCallingConfiguration -Identity Global -AllowPrivateCalling ([bool]$g.privateCalling) }
  $m = @{}
  if ($g.ContainsKey('ipVideo')) { $m.AllowIPVideo = [bool]$g.ipVideo }
  if ($g.ContainsKey('meetNow')) { $m.AllowMeetNow = [bool]$g.meetNow }
  if ($g.ContainsKey('screenSharing')) {
    if (@('Disabled', 'SingleApplication', 'EntireScreen') -notcontains $g.screenSharing) { throw 'Invalid screen sharing mode' }
    $m.ScreenSharingMode = "$($g.screenSharing)"
  }
  if ($m.Count) { Set-CsTeamsGuestMeetingConfiguration -Identity Global @m }
  $s = @{}
  if ($g.ContainsKey('chat')) { $s.AllowUserChat = [bool]$g.chat }
  if ($g.ContainsKey('editMessages')) { $s.AllowUserEditMessage = [bool]$g.editMessages }
  if ($g.ContainsKey('deleteMessages')) { $s.AllowUserDeleteMessage = [bool]$g.deleteMessages }
  if ($g.ContainsKey('giphy')) { $s.AllowGiphy = [bool]$g.giphy }
  if ($s.Count) { Set-CsTeamsGuestMessagingConfiguration -Identity Global @s }
  Tb-Access
}

# ---------- phone numbers ----------

function Tb-Phone {
  $numbers = @()
  $skip = 0
  do {
    $page = @(Get-CsPhoneNumberAssignment -Top 500 -Skip $skip)
    $numbers += @($page | ForEach-Object {
        @{
          number     = "$($_.TelephoneNumber)"
          type       = "$($_.NumberType)"
          activation = "$($_.ActivationState)"
          capability = Strs $_.Capability
          assignedTo = if ($_.AssignedPstnTargetId) { "$($_.AssignedPstnTargetId)" } else { $null }
          status     = "$($_.PstnAssignmentStatus)"
          city       = if ($_.City) { "$($_.City)" } else { $null }
        }
      })
    $skip += 500
  } while ($page.Count -eq 500)
  $users = @(Get-Users | Where-Object { $_.phoneSystem -or $_.lineUri })
  @{
    numbers = $numbers
    users   = $users
    calling = @(Get-CsTeamsCallingPolicy | ForEach-Object { "$($_.Identity)" -replace '^Tag:', '' })
    routing = @(Get-CsOnlineVoiceRoutingPolicy | ForEach-Object { "$($_.Identity)" -replace '^Tag:', '' })
  }
}

$NumberTypes = 'CallingPlan', 'DirectRouting', 'OperatorConnect'

function Tb-AssignNumber([string]$User, [string]$Number, [string]$Type, [string]$LocationId) {
  if ($NumberTypes -notcontains $Type) { throw "Unknown number type $Type" }
  $p = @{ Identity = $User; PhoneNumber = $Number; PhoneNumberType = $Type }
  if ($LocationId) { $p.LocationId = $LocationId }
  Set-CsPhoneNumberAssignment @p
  Tb-User $User
}

function Tb-UnassignNumber([string]$User, [string]$Number, [string]$Type) {
  if ($NumberTypes -notcontains $Type) { throw "Unknown number type $Type" }
  Remove-CsPhoneNumberAssignment -Identity $User -PhoneNumber $Number -PhoneNumberType $Type
  Tb-User $User
}

function Tb-EnterpriseVoice([string]$User, [bool]$Enabled) {
  Set-CsPhoneNumberAssignment -Identity $User -EnterpriseVoiceEnabled $Enabled
  Tb-User $User
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
    if ($msg -match 'is not recognized as the name of a cmdlet|Forbidden|Unauthorized|not authorized|Access Denied|AccessDenied|does not have permission|insufficient privileges') { $code = 'TEAMS_DENIED' }
    Out-Reply @{ id = $(if ($req) { $req.id } else { $null }); ok = $false; error = $msg; code = $code }
  }
}
