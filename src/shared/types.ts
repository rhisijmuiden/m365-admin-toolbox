export type TenantMode = 'gdap' | 'direct'
export type Lang = 'en' | 'nl'
export type Theme = 'system' | 'light' | 'dark'
export type Channel = 'prod' | 'test'

export interface Tenant {
  tenantId: string
  name: string
  defaultDomain: string
  domains: string[]
  mode: TenantMode
  homeAccountId: string
  username: string
  color: string
  addedAt: string
}

export interface Settings {
  lang: Lang
  theme: Theme
  channel: Channel
  clientId: string
  lastTenantId?: string
}

export interface User {
  id: string
  displayName: string
  userPrincipalName: string
  department: string | null
  accountEnabled: boolean
  skuIds: string[]
  lastSignIn: string | null
  createdDateTime: string | null
}

export interface Sku {
  skuId: string
  skuPartNumber: string
  name: string
  enabled: number
  suspended: number
  warning: number
  consumed: number
  available: number
  shortage: number
  capabilityStatus: string
}

export interface Group {
  id: string
  displayName: string
  description: string | null
  kind: 'm365' | 'security' | 'distribution' | 'mailSecurity'
  mail: string | null
  createdDateTime: string | null
}

export interface NewUser {
  givenName: string
  surname: string
  displayName: string
  mailNickname: string
  userPrincipalName: string
  password: string
  forceChangePasswordNextSignIn: boolean
  usageLocation: string
  department?: string
  skuIds: string[]
}

export interface NewGroup {
  displayName: string
  description?: string
  mailNickname: string
  kind: 'm365' | 'security'
  ownerId?: string
}

export interface LicenseChange {
  userId: string
  add: string[]
  remove: string[]
}

export interface UpdateState {
  status: 'idle' | 'checking' | 'none' | 'available' | 'downloading' | 'ready' | 'error' | 'dev'
  version?: string
  percent?: number
  error?: string
}

export interface AppInfo {
  version: string
  channel: Channel
  isPackaged: boolean
  electron: string
  repo: string
}

/** Result envelope so errors cross IPC with a usable code and message. */
export type Result<T> = { ok: true; data: T } | { ok: false; error: string; code?: string; status?: number }

export type PimKind = 'role' | 'group'

export interface PimItem {
  /** Role definition id (directory role) or group id (PIM for groups). */
  targetId: string
  kind: PimKind
  name: string
  state: 'eligible' | 'active' | 'permanent' | 'pending'
  start?: string | null
  end?: string | null
  requestId?: string
}

export interface PimData {
  /** "directory": roles in the tenant itself. "groups": PIM for groups in the partner tenant (GDAP). */
  mode: 'directory' | 'groups'
  homeTenantId: string
  eligible: PimItem[]
  active: PimItem[]
  pending: PimItem[]
  /** True when the group list was narrowed to the GDAP relationship with this customer. */
  filtered: boolean
}

export interface PimRequest {
  kind: PimKind
  targetId: string
  name: string
  justification: string
  hours: number
  ticket?: string
}

/* ---------- Exchange Online (phase 2) ---------- */

export type ExoPhase = 'module' | 'connecting' | 'ready'

export interface Mailbox {
  /** Exchange object GUID, used as -Identity for every cmdlet. */
  id: string
  /** Entra object id, for Graph (licenses). */
  userId: string
  displayName: string
  alias: string
  upn: string
  primarySmtp: string
  /** RecipientTypeDetails: UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox. */
  type: string
  /** SMTP proxy addresses with their prefix; "SMTP:" marks the primary one. */
  addresses: string[]
  quotaBytes: number | null
  archive: boolean
  forwardingSmtp: string | null
  forwardingAddress: string | null
  deliverAndForward: boolean
  litigationHold: boolean
  retentionPolicy: string | null
  holds: string[]
  hidden: boolean
  capacity: number | null
  created: string | null
}

export interface MailboxList {
  mailboxes: Mailbox[]
  /** Accepted domains, lower case; forwarding to anything else is external. */
  domains: string[]
}

export interface MailboxSize {
  /** Set for live Exchange statistics. */
  id?: string
  /** Set for the Graph usage report, which is keyed by UPN. */
  upn?: string
  bytes: number | null
  items: number | null
  lastActivity: string | null
  archiveBytes?: number | null
  quotaBytes?: number | null
}

export interface MailboxSizes {
  source: 'graph' | 'exo'
  sizes: MailboxSize[]
  /** Why the faster Graph report was not used. */
  graphIssue?: 'consent' | 'concealed' | 'denied' | 'error'
  /** Day the Graph report was generated; the report runs one or two days behind. */
  reportDate?: string | null
}

export type MailRight = 'FullAccess' | 'SendAs' | 'SendOnBehalf'

export interface MailboxPerms {
  fullAccess: string[]
  sendAs: string[]
  sendOnBehalf: string[]
}

export interface UserAccess {
  id: string
  displayName: string
  primarySmtp: string
  type: string
  rights: MailRight[]
}

export interface AutoReply {
  state: 'Disabled' | 'Enabled' | 'Scheduled'
  start: string | null
  end: string | null
  internalMessage: string
  externalMessage: string
  externalAudience: 'None' | 'Known' | 'All'
}

export interface CalendarSettings {
  automateProcessing: 'None' | 'AutoUpdate' | 'AutoAccept'
  allowConflicts: boolean
  bookingWindowInDays: number
  maximumDurationInMinutes: number
  allowRecurringMeetings: boolean
  deleteSubject: boolean
  addOrganizerToSubject: boolean
  processExternal: boolean
}

export interface NewMailbox {
  kind: 'shared' | 'room' | 'equipment'
  name: string
  alias: string
  smtp: string
  capacity?: number
  /** Users who get Full Access and Send As (shared mailboxes). */
  members?: string[]
}

export type DistKind = 'distribution' | 'security' | 'dynamic'

export interface DistGroup {
  id: string
  displayName: string
  alias: string
  primarySmtp: string
  kind: DistKind
  addresses: string[]
  /** ManagedBy as Exchange returns it (names). */
  owners: string[]
  allowExternal: boolean
  hidden: boolean
  filter: string | null
  created: string | null
}

export interface Recipient {
  id: string
  displayName: string
  primarySmtp: string
  type: string
  addresses: string[]
}

export interface DistGroupDetail {
  group: DistGroup
  owners: Recipient[]
  members: Recipient[]
}

export interface NewDistGroup {
  kind: DistKind
  name: string
  alias: string
  smtp: string
  members: string[]
  owners: string[]
  allowExternal: boolean
  /** Dynamic lists: MailboxUsers, AllRecipients, ... plus optional conditions. */
  includedRecipients?: string
  department?: string[]
  company?: string[]
  city?: string[]
}

export interface InboxRule {
  mailboxId: string
  mailbox: string
  mailboxSmtp: string
  ruleId: string
  identity: string
  name: string
  enabled: boolean
  description: string
  targets: string[]
  /** forward, external, delete, move (to Deleted Items, RSS and similar), hide (also marks as read). */
  flags: string[]
}

export interface MailFlow {
  rules: { name: string; state: string; mode: string; priority: number; description: string; modified: string | null }[]
  domains: { name: string; type: string; isDefault: boolean }[]
}
