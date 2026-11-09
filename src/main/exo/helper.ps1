# M365 Admin Toolbox: Exchange Online helper.
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
  $m = Get-Module -ListAvailable ExchangeOnlineManagement | Where-Object { $_.Version.Major -ge 3 } | Sort-Object Version -Descending | Select-Object -First 1
  $installed = $false
  if (-not $m) {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge '2.8.5.201' })) {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module ExchangeOnlineManagement -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -MinimumVersion 3.0.0
    $installed = $true
    $m = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
  }
  Import-Module $m.Path -DisableNameChecking
  @{ version = "$($m.Version)"; installed = $installed }
}

function Tb-Connect([string]$Token, [string]$Org) {
  if ($script:Connected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
  Connect-ExchangeOnline -AccessToken $Token -Organization $Org -ShowBanner:$false -SkipLoadingFormatData | Out-Null
  $script:Connected = $true
  @{ ok = $true }
}

# ---------- mailboxes ----------

function MailboxOut($m) {
  @{
    id                         = "$($m.Guid)"
    userId                     = "$($m.ExternalDirectoryObjectId)"
    displayName                = "$($m.DisplayName)"
    alias                      = "$($m.Alias)"
    upn                        = "$($m.UserPrincipalName)"
    primarySmtp                = "$($m.PrimarySmtpAddress)"
    type                       = "$($m.RecipientTypeDetails)"
    addresses                  = Smtp $m.EmailAddresses
    quotaBytes                 = Bytes $m.ProhibitSendReceiveQuota
    archive                    = ("$($m.ArchiveStatus)" -eq 'Active') -or ($m.ArchiveGuid -and "$($m.ArchiveGuid)" -ne '00000000-0000-0000-0000-000000000000')
    forwardingSmtp             = if ($m.ForwardingSmtpAddress) { ("$($m.ForwardingSmtpAddress)" -replace '^smtp:', '') } else { $null }
    forwardingAddress          = if ($m.ForwardingAddress) { "$($m.ForwardingAddress)" } else { $null }
    deliverAndForward          = [bool]$m.DeliverToMailboxAndForward
    litigationHold             = [bool]$m.LitigationHoldEnabled
    retentionPolicy            = if ($m.RetentionPolicy) { "$($m.RetentionPolicy)" } else { $null }
    holds                      = Strs $m.InPlaceHolds
    hidden                     = [bool]$m.HiddenFromAddressListsEnabled
    capacity                   = $m.ResourceCapacity
    created                    = Iso $m.WhenCreated
  }
}

$MbxProps = 'ProhibitSendReceiveQuota', 'ArchiveStatus', 'ArchiveGuid', 'ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward',
'LitigationHoldEnabled', 'RetentionPolicy', 'InPlaceHolds', 'HiddenFromAddressListsEnabled', 'WhenCreated', 'ResourceCapacity', 'GrantSendOnBehalfTo'

function Tb-Mailboxes {
  $list = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum -Properties $MbxProps | ForEach-Object { MailboxOut $_ })
  $domains = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() })
  @{ mailboxes = $list; domains = $domains }
}

function Tb-Mailbox([string]$Identity) { MailboxOut (Get-EXOMailbox -Identity $Identity -PropertySets Minimum -Properties $MbxProps) }

# Live sizes, used when the Graph usage report is not available. Slower: one call per mailbox.
function Tb-MailboxStats {
  , @(Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus | ForEach-Object {
      $m = $_
      try {
        $s = Get-EXOMailboxStatistics -Identity "$($m.Guid)" -Properties LastUserActionTime
        $a = $null
        if ("$($m.ArchiveStatus)" -eq 'Active') { try { $a = Bytes (Get-EXOMailboxStatistics -Identity "$($m.Guid)" -Archive).TotalItemSize } catch {} }
        @{ id = "$($m.Guid)"; bytes = Bytes $s.TotalItemSize; items = $s.ItemCount; lastActivity = Iso $s.LastUserActionTime; archiveBytes = $a }
      } catch { @{ id = "$($m.Guid)"; bytes = $null; items = $null; lastActivity = $null; archiveBytes = $null } }
    })
}

$SetMailboxAllowed = 'ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward', 'Type', 'HiddenFromAddressListsEnabled', 'ResourceCapacity', 'DisplayName'

function Tb-SetMailbox([string]$Identity, $Params) {
  $p = @{}
  foreach ($k in $Params.Keys) { if ($SetMailboxAllowed -contains $k) { $p[$k] = $Params[$k] } else { throw "Parameter $k is not allowed" } }
  Set-Mailbox -Identity $Identity @p
  Tb-Mailbox $Identity
}

function Tb-EnableArchive([string]$Identity) { Enable-Mailbox -Identity $Identity -Archive | Out-Null; Tb-Mailbox $Identity }

function Tb-NewMailbox([string]$Kind, [string]$Name, [string]$Alias, [string]$Smtp, $Capacity, [string[]]$Members) {
  $p = @{ Name = $Name; DisplayName = $Name; Alias = $Alias; PrimarySmtpAddress = $Smtp }
  if ($Kind -eq 'shared') { $p.Shared = $true } elseif ($Kind -eq 'room') { $p.Room = $true } elseif ($Kind -eq 'equipment') { $p.Equipment = $true } else { throw "Unknown kind $Kind" }
  $m = New-Mailbox @p
  if ($Capacity -and $Kind -ne 'shared') { Set-Mailbox -Identity "$($m.Guid)" -ResourceCapacity ([int]$Capacity) }
  $failed = @()
  foreach ($u in @($Members | Where-Object { $_ })) {
    try {
      Add-MailboxPermission -Identity "$($m.Guid)" -User $u -AccessRights FullAccess -InheritanceType All -AutoMapping $true | Out-Null
      Add-RecipientPermission -Identity "$($m.Guid)" -Trustee $u -AccessRights SendAs -Confirm:$false | Out-Null
    } catch { $failed += @{ user = $u; error = $_.Exception.Message } }
  }
  @{ mailbox = (MailboxOut $m); failed = $failed }
}

# ---------- permissions ----------

function IsSystemTrustee([string]$u) { $u -match '^(NT AUTHORITY\\|S-1-5-|NAMPR|EURPR|APCPR|.*\\Discovery Management|.*\\Organization Management)' }

function Tb-Permissions([string]$Identity) {
  $full = @(Get-EXOMailboxPermission -Identity $Identity | Where-Object { -not $_.IsInherited -and -not $_.Deny -and "$($_.AccessRights)" -match 'FullAccess' -and -not (IsSystemTrustee "$($_.User)") } |
    ForEach-Object { "$($_.User)" })
  $sendAs = @(Get-EXORecipientPermission -Identity $Identity | Where-Object { "$($_.AccessRights)" -match 'SendAs' -and $_.AccessControlType -ne 'Deny' -and -not (IsSystemTrustee "$($_.Trustee)") } |
    ForEach-Object { "$($_.Trustee)" })
  $m = Get-EXOMailbox -Identity $Identity -Properties GrantSendOnBehalfTo
  $behalf = @($m.GrantSendOnBehalfTo | Where-Object { $_ } | ForEach-Object {
      $r = $null
      try { $r = Get-EXORecipient -Identity "$_" } catch {}
      if ($r) { "$($r.PrimarySmtpAddress)" } else { "$_" }
    })
  @{ fullAccess = $full; sendAs = $sendAs; sendOnBehalf = $behalf }
}

# Every mailbox where this user has rights. One call per mailbox, so it takes a while in large tenants.
function Tb-PermissionsForUser([string]$User) {
  $out = @()
  $sendAs = @{}
  Get-EXORecipientPermission -Trustee $User -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { "$($_.AccessRights)" -match 'SendAs' } |
    ForEach-Object { $sendAs["$($_.Identity)"] = $true }
  $me = Get-EXORecipient -Identity $User
  foreach ($m in Get-EXOMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo) {
    if ("$($m.Guid)" -eq "$($me.Guid)") { continue }
    $rights = @()
    try {
      if (Get-EXOMailboxPermission -Identity "$($m.Guid)" -User $User -ErrorAction SilentlyContinue | Where-Object { -not $_.IsInherited -and "$($_.AccessRights)" -match 'FullAccess' }) { $rights += 'FullAccess' }
    } catch {}
    if ($sendAs["$($m.Identity)"] -or $sendAs["$($m.Name)"] -or $sendAs["$($m.DisplayName)"]) { $rights += 'SendAs' }
    if (@($m.GrantSendOnBehalfTo | ForEach-Object { "$_" }) | Where-Object { $_ -eq "$($me.Name)" -or $_ -eq "$($me.Identity)" -or $_ -eq "$($me.DisplayName)" }) { $rights += 'SendOnBehalf' }
    if ($rights.Count) { $out += @{ id = "$($m.Guid)"; displayName = "$($m.DisplayName)"; primarySmtp = "$($m.PrimarySmtpAddress)"; type = "$($m.RecipientTypeDetails)"; rights = $rights } }
  }
  , $out
}

function Tb-Grant([string]$Identity, [string]$User, [string]$Right, [bool]$AutoMapping = $true) {
  switch ($Right) {
    'FullAccess' { Add-MailboxPermission -Identity $Identity -User $User -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping | Out-Null }
    'SendAs' { Add-RecipientPermission -Identity $Identity -Trustee $User -AccessRights SendAs -Confirm:$false | Out-Null }
    'SendOnBehalf' { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{ Add = $User } }
    default { throw "Unknown right $Right" }
  }
  Tb-Permissions $Identity
}

function Tb-Revoke([string]$Identity, [string]$User, [string]$Right) {
  switch ($Right) {
    'FullAccess' { Remove-MailboxPermission -Identity $Identity -User $User -AccessRights FullAccess -InheritanceType All -Confirm:$false | Out-Null }
    'SendAs' { Remove-RecipientPermission -Identity $Identity -Trustee $User -AccessRights SendAs -Confirm:$false | Out-Null }
    'SendOnBehalf' { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{ Remove = $User } }
    default { throw "Unknown right $Right" }
  }
  Tb-Permissions $Identity
}

# ---------- auto reply ----------

function Tb-AutoReply([string]$Identity) {
  $c = Get-MailboxAutoReplyConfiguration -Identity $Identity
  @{
    state            = "$($c.AutoReplyState)"
    start            = Iso $c.StartTime
    end              = Iso $c.EndTime
    internalMessage  = "$($c.InternalMessage)"
    externalMessage  = "$($c.ExternalMessage)"
    externalAudience = "$($c.ExternalAudience)"
  }
}

function Tb-SetAutoReply([string]$Identity, [string]$State, $Start, $End, [string]$InternalMessage, [string]$ExternalMessage, [string]$ExternalAudience) {
  $p = @{ Identity = $Identity; AutoReplyState = $State; InternalMessage = $InternalMessage; ExternalMessage = $ExternalMessage; ExternalAudience = $ExternalAudience }
  if ($State -eq 'Scheduled') { $p.StartTime = [datetime]$Start; $p.EndTime = [datetime]$End }
  Set-MailboxAutoReplyConfiguration @p
  Tb-AutoReply $Identity
}

# ---------- resources ----------

function Tb-CalendarProcessing([string]$Identity) {
  $c = Get-CalendarProcessing -Identity $Identity
  @{
    automateProcessing       = "$($c.AutomateProcessing)"
    allowConflicts           = [bool]$c.AllowConflicts
    bookingWindowInDays      = $c.BookingWindowInDays
    maximumDurationInMinutes = $c.MaximumDurationInMinutes
    allowRecurringMeetings   = [bool]$c.AllowRecurringMeetings
    deleteSubject            = [bool]$c.DeleteSubject
    addOrganizerToSubject    = [bool]$c.AddOrganizerToSubject
    processExternal          = [bool]$c.ProcessExternalMeetingMessages
  }
}

$CalAllowed = 'AutomateProcessing', 'AllowConflicts', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowRecurringMeetings', 'DeleteSubject', 'AddOrganizerToSubject', 'ProcessExternalMeetingMessages'

function Tb-SetCalendarProcessing([string]$Identity, $Params) {
  $p = @{}
  foreach ($k in $Params.Keys) { if ($CalAllowed -contains $k) { $p[$k] = $Params[$k] } else { throw "Parameter $k is not allowed" } }
  Set-CalendarProcessing -Identity $Identity @p
  Tb-CalendarProcessing $Identity
}

# ---------- distribution lists ----------

function GroupOut($g, [string]$kind) {
  @{
    id            = "$($g.Guid)"
    displayName   = "$($g.DisplayName)"
    alias         = "$($g.Alias)"
    primarySmtp   = "$($g.PrimarySmtpAddress)"
    kind          = $kind
    addresses     = Smtp $g.EmailAddresses
    owners        = Strs $g.ManagedBy
    allowExternal = -not [bool]$g.RequireSenderAuthenticationEnabled
    hidden        = [bool]$g.HiddenFromAddressListsEnabled
    filter        = if ($kind -eq 'dynamic') { "$($g.RecipientFilter)" } else { $null }
    created       = Iso $g.WhenCreated
  }
}

function Tb-Groups {
  $dl = @(Get-DistributionGroup -ResultSize Unlimited | ForEach-Object { GroupOut $_ ($(if ("$($_.RecipientTypeDetails)" -eq 'MailUniversalSecurityGroup') { 'security' } else { 'distribution' })) })
  $dyn = @(Get-DynamicDistributionGroup -ResultSize Unlimited | ForEach-Object { GroupOut $_ 'dynamic' })
  , @($dl + $dyn)
}

function Tb-Group([string]$Identity, [string]$Kind) {
  if ($Kind -eq 'dynamic') { GroupOut (Get-DynamicDistributionGroup -Identity $Identity) 'dynamic' }
  else { $g = Get-DistributionGroup -Identity $Identity; GroupOut $g ($(if ("$($g.RecipientTypeDetails)" -eq 'MailUniversalSecurityGroup') { 'security' } else { 'distribution' })) }
}

function RecipientOut($r) { @{ id = "$($r.Guid)"; displayName = "$($r.DisplayName)"; primarySmtp = "$($r.PrimarySmtpAddress)"; type = "$($r.RecipientTypeDetails)"; addresses = Smtp $r.EmailAddresses } }

function Tb-GroupMembers([string]$Identity, [string]$Kind) {
  $m = if ($Kind -eq 'dynamic') { Get-DynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited } else { Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited }
  , @($m | ForEach-Object { RecipientOut $_ })
}

function Tb-GroupDetail([string]$Identity, [string]$Kind) {
  $g = Tb-Group $Identity $Kind
  $owners = @($g.owners | ForEach-Object {
      $o = $null
      try { $o = Get-EXORecipient -Identity $_ -PropertySets Minimum } catch {}
      if ($o) { RecipientOut $o } else { @{ id = "$_"; displayName = "$_"; primarySmtp = ''; type = ''; addresses = @() } }
    })
  @{ group = $g; owners = $owners; members = (Tb-GroupMembers $Identity $Kind) }
}

function Tb-NewGroup([string]$Kind, [string]$Name, [string]$Alias, [string]$Smtp, [string[]]$Members, [string[]]$Owners, [bool]$AllowExternal,
  [string]$IncludedRecipients, [string[]]$Department, [string[]]$Company, [string[]]$City) {
  if ($Kind -eq 'dynamic') {
    $p = @{ Name = $Name; DisplayName = $Name; Alias = $Alias; PrimarySmtpAddress = $Smtp; IncludedRecipients = $IncludedRecipients }
    if ($Department) { $p.ConditionalDepartment = $Department }
    if ($Company) { $p.ConditionalCompany = $Company }
    if ($City) { $p.ConditionalCity = $City }
    $g = New-DynamicDistributionGroup @p
    Set-DynamicDistributionGroup -Identity "$($g.Guid)" -RequireSenderAuthenticationEnabled (-not $AllowExternal)
    if ($Owners) { Set-DynamicDistributionGroup -Identity "$($g.Guid)" -ManagedBy $Owners[0] }
    return GroupOut $g 'dynamic'
  }
  $p = @{ Name = $Name; DisplayName = $Name; Alias = $Alias; PrimarySmtpAddress = $Smtp; Type = $(if ($Kind -eq 'security') { 'Security' } else { 'Distribution' }) }
  if ($Owners) { $p.ManagedBy = $Owners; $p.CopyOwnerToMember = $false }
  if ($Members) { $p.Members = $Members }
  $g = New-DistributionGroup @p
  Set-DistributionGroup -Identity "$($g.Guid)" -RequireSenderAuthenticationEnabled (-not $AllowExternal) -BypassSecurityGroupManagerCheck
  GroupOut $g $Kind
}

function Tb-SetGroup([string]$Identity, [string]$Kind, $Owners, $AllowExternal, $DisplayName) {
  $p = @{ Identity = $Identity }
  if ($null -ne $Owners) { $p.ManagedBy = @($Owners) }
  if ($null -ne $AllowExternal) { $p.RequireSenderAuthenticationEnabled = -not [bool]$AllowExternal }
  if ($DisplayName) { $p.DisplayName = $DisplayName }
  if ($Kind -eq 'dynamic') { Set-DynamicDistributionGroup @p } else { Set-DistributionGroup @p -BypassSecurityGroupManagerCheck }
  Tb-Group $Identity $Kind
}

function Tb-AddMember([string]$Identity, [string]$Member) { Add-DistributionGroupMember -Identity $Identity -Member $Member -BypassSecurityGroupManagerCheck; Tb-GroupMembers $Identity 'distribution' }

function Tb-RemoveMember([string]$Identity, [string]$Member) { Remove-DistributionGroupMember -Identity $Identity -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false; Tb-GroupMembers $Identity 'distribution' }

function Tb-RemoveGroup([string]$Identity, [string]$Kind) {
  if ($Kind -eq 'dynamic') { Remove-DynamicDistributionGroup -Identity $Identity -Confirm:$false } else { Remove-DistributionGroup -Identity $Identity -BypassSecurityGroupManagerCheck -Confirm:$false }
  @{ ok = $true }
}

# ---------- addresses ----------

function Tb-Recipients {
  , @(Get-EXORecipient -ResultSize Unlimited -PropertySets Minimum | Where-Object { "$($_.RecipientTypeDetails)" -ne 'DiscoveryMailbox' } | ForEach-Object {
      @{ id = "$($_.Guid)"; displayName = "$($_.DisplayName)"; primarySmtp = "$($_.PrimarySmtpAddress)"; type = "$($_.RecipientTypeDetails)"; addresses = Smtp $_.EmailAddresses }
    })
}

# Sets the primary address and adds or removes aliases. Non-SMTP proxy addresses (X500, SIP, SPO) are kept.
function Tb-SetAddresses([string]$Identity, [string]$Type, [string]$Primary, [string[]]$Add, [string[]]$Remove) {
  $r = Get-EXORecipient -Identity $Identity -PropertySets Minimum
  if ($Type -eq 'GroupMailbox') {
    if ($Primary) { Set-UnifiedGroup -Identity $Identity -PrimarySmtpAddress $Primary }
    if ($Add -or $Remove) {
      $h = @{}
      if ($Add) { $h.Add = @($Add | ForEach-Object { "smtp:$_" }) }
      if ($Remove) { $h.Remove = @($Remove | ForEach-Object { "smtp:$_" }) }
      Set-UnifiedGroup -Identity $Identity -EmailAddresses $h
    }
  } else {
    $all = @($r.EmailAddresses | ForEach-Object { "$_" })
    $other = @($all | Where-Object { $_ -notmatch '^smtp:' })
    $smtp = @($all | Where-Object { $_ -match '^smtp:' } | ForEach-Object { $_.Substring(5) })
    $prim = if ($Primary) { $Primary } else { "$($r.PrimarySmtpAddress)" }
    $smtp = @($smtp + @($Add) + @($prim) | Where-Object { $_ } | Where-Object { @($Remove) -notcontains $_ } | Sort-Object -Unique)
    if (@($Remove) -contains $prim) { throw 'The primary address cannot be removed' }
    $list = @($other + @($smtp | ForEach-Object { if ($_ -ieq $prim) { "SMTP:$_" } else { "smtp:$_" } }))
    switch -Regex ($Type) {
      'Mailbox$' { Set-Mailbox -Identity $Identity -EmailAddresses $list }
      '^MailUniversal(Distribution|Security)Group$' { Set-DistributionGroup -Identity $Identity -EmailAddresses $list -BypassSecurityGroupManagerCheck }
      '^DynamicDistributionGroup$' { Set-DynamicDistributionGroup -Identity $Identity -EmailAddresses $list }
      '^MailUser$|^GuestMailUser$' { Set-MailUser -Identity $Identity -EmailAddresses $list }
      '^MailContact$' { Set-MailContact -Identity $Identity -EmailAddresses $list }
      default { throw "Addresses of a $Type cannot be changed here" }
    }
  }
  $n = Get-EXORecipient -Identity $Identity -PropertySets Minimum
  @{ id = "$($n.Guid)"; displayName = "$($n.DisplayName)"; primarySmtp = "$($n.PrimarySmtpAddress)"; type = "$($n.RecipientTypeDetails)"; addresses = Smtp $n.EmailAddresses }
}

# ---------- inbox rules ----------

$SuspectFolders = '(Deleted Items|Verwijderde items|RSS|Conversation History|Gespreksgeschiedenis|Archive|Archief|Junk|Ongewenste)'

function Tb-InboxRules {
  $domains = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() })
  $out = @()
  foreach ($m in Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Properties ForwardingSmtpAddress) {
    $rules = @()
    try { $rules = @(Get-InboxRule -Mailbox "$($m.Guid)") } catch { continue }
    foreach ($r in $rules) {
      $targets = @(@($r.ForwardTo) + @($r.ForwardAsAttachmentTo) + @($r.RedirectTo) | Where-Object { $_ } | ForEach-Object {
          $s = "$_"
          if ($s -match '\[SMTP:([^\]]+)\]') { $matches[1] } elseif ($s -match '^"?([^"]+@[^"\s]+)') { $matches[1] } else { $s }
        })
      $external = @($targets | Where-Object { $_ -match '@' -and $domains -notcontains ($_.Split('@')[-1].ToLower()) })
      $flags = @()
      if ($targets.Count) { $flags += 'forward' }
      if ($external.Count) { $flags += 'external' }
      if ($r.DeleteMessage -or $r.SoftDeleteMessage) { $flags += 'delete' }
      if ("$($r.MoveToFolder)" -match $SuspectFolders) { $flags += 'move' }
      if ($r.MarkAsRead -and ($flags -contains 'move' -or $flags -contains 'delete')) { $flags += 'hide' }
      $out += @{
        mailboxId   = "$($m.Guid)"
        mailbox     = "$($m.DisplayName)"
        mailboxSmtp = "$($m.PrimarySmtpAddress)"
        ruleId      = "$($r.RuleIdentity)"
        identity    = "$($r.Identity)"
        name        = "$($r.Name)"
        enabled     = [bool]$r.Enabled
        description = ("$($r.Description)" -replace '\s+', ' ').Trim()
        targets     = $targets
        flags       = $flags
      }
    }
  }
  , $out
}

function Tb-SetInboxRule([string]$Mailbox, [string]$RuleId, [string]$Action) {
  switch ($Action) {
    'disable' { Disable-InboxRule -Mailbox $Mailbox -Identity $RuleId -Confirm:$false -AlwaysDeleteOutlookRulesBlob }
    'enable' { Enable-InboxRule -Mailbox $Mailbox -Identity $RuleId -Confirm:$false -AlwaysDeleteOutlookRulesBlob }
    'remove' { Remove-InboxRule -Mailbox $Mailbox -Identity $RuleId -Confirm:$false -AlwaysDeleteOutlookRulesBlob }
    default { throw "Unknown action $Action" }
  }
  @{ ok = $true }
}

# ---------- mail flow (read-only) ----------

function Tb-MailFlow {
  @{
    rules   = @(Get-TransportRule -ResultSize Unlimited | ForEach-Object {
        @{ name = "$($_.Name)"; state = "$($_.State)"; mode = "$($_.Mode)"; priority = $_.Priority; description = ("$($_.Description)" -replace '\s+', ' ').Trim(); modified = Iso $_.WhenChanged }
      })
    domains = @(Get-AcceptedDomain | ForEach-Object { @{ name = "$($_.DomainName)"; type = "$($_.DomainType)"; isDefault = [bool]$_.Default } })
  }
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
    Out-Reply @{ id = $(if ($req) { $req.id } else { $null }); ok = $false; error = $msg; code = $code }
  }
}
