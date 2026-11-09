import type {
  AutoReply,
  CalendarSettings,
  DistGroup,
  DistGroupDetail,
  DistKind,
  InboxRule,
  MailFlow,
  Mailbox,
  MailboxList,
  MailboxPerms,
  MailboxSize,
  MailboxSizes,
  MailRight,
  NewDistGroup,
  NewMailbox,
  Recipient,
  UserAccess
} from '@shared/types'
import { NeedsConsent, tokenFor, tokenForScopes } from './auth'
import { all, call, GraphError } from './graph'
import { run } from './exo'

export const REPORT_SCOPES = ['Reports.Read.All']

export const listMailboxes = (tid: string) => run<MailboxList>(tid, 'Tb-Mailboxes')
export const getMailbox = (tid: string, id: string) => run<Mailbox>(tid, 'Tb-Mailbox', { Identity: id })

type ReportRow = {
  userPrincipalName: string
  storageUsedInBytes: number | null
  itemCount: number | null
  lastActivityDate: string | null
  prohibitSendReceiveQuotaInBytes: number | null
  reportRefreshDate?: string
  isDeleted?: boolean
}

/**
 * Sizes from the Graph mailbox usage report (one call, a day or two behind). When the report can't be
 * used (no consent, names concealed, no access) the sizes come live from Exchange, one call per mailbox.
 */
export async function mailboxSizes(tid: string, upns: string[], live = false): Promise<MailboxSizes> {
  let graphIssue: MailboxSizes['graphIssue']
  if (!live) {
    try {
      const token = await tokenForScopes(tid, REPORT_SCOPES)
      const rows = await all<ReportRow>(token, "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D7')?$format=application/json")
      const live = rows.filter((r) => !r.isDeleted)
      const known = new Set(upns.map((u) => u.toLowerCase()))
      // With "conceal user names in reports" on, UPNs come back as hashes and match nothing.
      if (live.length && !live.some((r) => known.has(r.userPrincipalName?.toLowerCase()))) graphIssue = 'concealed'
      else
        return {
          source: 'graph',
          reportDate: live[0]?.reportRefreshDate || null,
          sizes: live.map(
            (r): MailboxSize => ({
              upn: r.userPrincipalName.toLowerCase(),
              bytes: r.storageUsedInBytes,
              items: r.itemCount,
              lastActivity: r.lastActivityDate,
              quotaBytes: r.prohibitSendReceiveQuotaInBytes
            })
          )
        }
    } catch (e) {
      graphIssue = e instanceof NeedsConsent ? 'consent' : e instanceof GraphError && e.status === 403 ? 'denied' : 'error'
    }
  }
  const sizes = await run<MailboxSize[]>(tid, 'Tb-MailboxStats')
  return { source: 'exo', sizes, graphIssue }
}

type SetMailbox = Partial<{
  ForwardingSmtpAddress: string | null
  ForwardingAddress: string | null
  DeliverToMailboxAndForward: boolean
  Type: 'Shared' | 'Regular'
  HiddenFromAddressListsEnabled: boolean
  ResourceCapacity: number | null
  DisplayName: string
}>

export const setMailbox = (tid: string, id: string, params: SetMailbox) => run<Mailbox>(tid, 'Tb-SetMailbox', { Identity: id, Params: params })
export const enableArchive = (tid: string, id: string) => run<Mailbox>(tid, 'Tb-EnableArchive', { Identity: id })

/** Converts a user mailbox to shared and optionally removes all licenses, which a shared mailbox under 50 GB doesn't need. */
export async function convertToShared(tid: string, id: string, userId: string, removeLicenses: boolean) {
  const m = await setMailbox(tid, id, { Type: 'Shared' })
  let licensesRemoved = 0
  if (removeLicenses && userId) {
    const token = await tokenFor(tid)
    const u = await call<{ assignedLicenses: { skuId: string }[] }>(token, `/users/${userId}?$select=assignedLicenses`)
    if (u.assignedLicenses.length) {
      await call(token, `/users/${userId}/assignLicense`, {
        method: 'POST',
        body: JSON.stringify({ addLicenses: [], removeLicenses: u.assignedLicenses.map((l) => l.skuId) })
      })
      licensesRemoved = u.assignedLicenses.length
    }
  }
  return { mailbox: m, licensesRemoved }
}

export const newMailbox = (tid: string, m: NewMailbox) =>
  run<{ mailbox: Mailbox; failed: { user: string; error: string }[] }>(tid, 'Tb-NewMailbox', {
    Kind: m.kind,
    Name: m.name,
    Alias: m.alias,
    Smtp: m.smtp,
    Capacity: m.capacity ?? null,
    Members: m.members || []
  })

export const permissions = (tid: string, id: string) => run<MailboxPerms>(tid, 'Tb-Permissions', { Identity: id })
export const permissionsForUser = (tid: string, user: string) => run<UserAccess[]>(tid, 'Tb-PermissionsForUser', { User: user })
export const grant = (tid: string, id: string, user: string, right: MailRight, autoMapping = true) =>
  run<MailboxPerms>(tid, 'Tb-Grant', { Identity: id, User: user, Right: right, AutoMapping: autoMapping })
export const revoke = (tid: string, id: string, user: string, right: MailRight) => run<MailboxPerms>(tid, 'Tb-Revoke', { Identity: id, User: user, Right: right })

export const autoReply = (tid: string, id: string) => run<AutoReply>(tid, 'Tb-AutoReply', { Identity: id })
export const setAutoReply = (tid: string, id: string, a: AutoReply) =>
  run<AutoReply>(tid, 'Tb-SetAutoReply', {
    Identity: id,
    State: a.state,
    Start: a.start,
    End: a.end,
    InternalMessage: a.internalMessage,
    ExternalMessage: a.externalMessage,
    ExternalAudience: a.externalAudience
  })

export const calendar = (tid: string, id: string) => run<CalendarSettings>(tid, 'Tb-CalendarProcessing', { Identity: id })
export const setCalendar = (tid: string, id: string, c: CalendarSettings) =>
  run<CalendarSettings>(tid, 'Tb-SetCalendarProcessing', {
    Identity: id,
    Params: {
      AutomateProcessing: c.automateProcessing,
      AllowConflicts: c.allowConflicts,
      BookingWindowInDays: c.bookingWindowInDays,
      MaximumDurationInMinutes: c.maximumDurationInMinutes,
      AllowRecurringMeetings: c.allowRecurringMeetings,
      DeleteSubject: c.deleteSubject,
      AddOrganizerToSubject: c.addOrganizerToSubject,
      ProcessExternalMeetingMessages: c.processExternal
    }
  })

export const listDistGroups = (tid: string) => run<DistGroup[]>(tid, 'Tb-Groups')
export const distGroupDetail = (tid: string, id: string, kind: DistKind) => run<DistGroupDetail>(tid, 'Tb-GroupDetail', { Identity: id, Kind: kind })
export const newDistGroup = (tid: string, g: NewDistGroup) =>
  run<DistGroup>(tid, 'Tb-NewGroup', {
    Kind: g.kind,
    Name: g.name,
    Alias: g.alias,
    Smtp: g.smtp,
    Members: g.members,
    Owners: g.owners,
    AllowExternal: g.allowExternal,
    IncludedRecipients: g.includedRecipients || 'MailboxUsers',
    Department: g.department || [],
    Company: g.company || [],
    City: g.city || []
  })
export const setDistGroup = (tid: string, id: string, kind: DistKind, p: { owners?: string[]; allowExternal?: boolean; displayName?: string }) =>
  run<DistGroup>(tid, 'Tb-SetGroup', { Identity: id, Kind: kind, Owners: p.owners ?? null, AllowExternal: p.allowExternal ?? null, DisplayName: p.displayName ?? null })
export const addMember = (tid: string, id: string, member: string) => run<Recipient[]>(tid, 'Tb-AddMember', { Identity: id, Member: member })
export const removeMember = (tid: string, id: string, member: string) => run<Recipient[]>(tid, 'Tb-RemoveMember', { Identity: id, Member: member })
export const removeDistGroup = (tid: string, id: string, kind: DistKind) => run<void>(tid, 'Tb-RemoveGroup', { Identity: id, Kind: kind })

export const recipients = (tid: string) => run<Recipient[]>(tid, 'Tb-Recipients')
export const setAddresses = (tid: string, r: Recipient, change: { primary?: string; add?: string[]; remove?: string[] }) =>
  run<Recipient>(tid, 'Tb-SetAddresses', { Identity: r.id, Type: r.type, Primary: change.primary || '', Add: change.add || [], Remove: change.remove || [] })

export const inboxRules = (tid: string) => run<InboxRule[]>(tid, 'Tb-InboxRules')
export const setInboxRule = (tid: string, mailboxId: string, ruleId: string, action: 'disable' | 'enable' | 'remove') =>
  run<void>(tid, 'Tb-SetInboxRule', { Mailbox: mailboxId, RuleId: ruleId, Action: action })

export const mailFlow = (tid: string) => run<MailFlow>(tid, 'Tb-MailFlow')
