import type {
  AppInfo,
  AutoReply,
  CalendarSettings,
  DistGroup,
  DistGroupDetail,
  DistKind,
  ExoPhase,
  InboxRule,
  MailFlow,
  Mailbox,
  MailboxList,
  MailboxPerms,
  MailboxSizes,
  MailRight,
  NewDistGroup,
  NewMailbox,
  Recipient,
  UserAccess,
  Group,
  LicenseChange,
  NewGroup,
  NewUser,
  PimData,
  PimKind,
  PimRequest,
  Result,
  Settings,
  Sku,
  Tenant,
  TenantMode,
  UpdateState,
  User
} from '../shared/types'

type R<T> = Promise<Result<T>>

export interface Api {
  appInfo(): R<AppInfo>
  getSettings(): R<Settings>
  setSettings(patch: Partial<Settings>): R<Settings>
  listTenants(): R<Tenant[]>
  addTenant(mode: TenantMode, target?: string): R<Tenant>
  reauthTenant(tenantId: string): R<Tenant>
  removeTenant(tenantId: string): R<Tenant[]>
  listUsers(tenantId: string): R<User[]>
  createUser(tenantId: string, u: NewUser): R<User>
  setSignIn(tenantId: string, userId: string, enabled: boolean): R<void>
  resetPassword(tenantId: string, userId: string, password: string, forceChange: boolean): R<void>
  createTap(tenantId: string, userId: string, lifetimeInMinutes: number, isUsableOnce: boolean): R<{ pass: string; lifetimeInMinutes: number; start: string }>
  revokeSessions(tenantId: string, userId: string): R<void>
  listSkus(tenantId: string): R<Sku[]>
  changeLicenses(tenantId: string, changes: LicenseChange[]): R<{ failed: { userId: string; error: string }[] }>
  listGroups(tenantId: string): R<Group[]>
  createGroup(tenantId: string, g: NewGroup): R<Group>
  reauthHome(tenantId: string): R<void>
  pimLoad(tenantId: string): R<PimData>
  pimActivate(tenantId: string, r: PimRequest): R<{ status: string }>
  pimExtend(tenantId: string, r: PimRequest): R<{ status: string }>
  pimDeactivate(tenantId: string, kind: PimKind, targetId: string): R<void>
  pimCancel(tenantId: string, kind: PimKind, requestId: string): R<void>
  onNav(cb: (n: { tenantId: string; view: string }) => void): () => void
  exo: ExoApi
  updateState(): R<UpdateState>
  checkUpdate(): R<UpdateState>
  downloadUpdate(): R<void>
  installUpdate(): R<void>
  onUpdate(cb: (s: UpdateState) => void): () => void
}

export interface MailboxPatch {
  ForwardingSmtpAddress?: string | null
  ForwardingAddress?: string | null
  DeliverToMailboxAndForward?: boolean
  Type?: 'Shared' | 'Regular'
  HiddenFromAddressListsEnabled?: boolean
  ResourceCapacity?: number | null
  DisplayName?: string
}

export interface ExoApi {
  signIn(tenantId: string): R<void>
  consentReports(tenantId: string): R<void>
  mailboxes(tenantId: string): R<MailboxList>
  mailbox(tenantId: string, id: string): R<Mailbox>
  sizes(tenantId: string, upns: string[], live?: boolean): R<MailboxSizes>
  setMailbox(tenantId: string, id: string, p: MailboxPatch): R<Mailbox>
  enableArchive(tenantId: string, id: string): R<Mailbox>
  convertToShared(tenantId: string, id: string, userId: string, removeLicenses: boolean): R<{ mailbox: Mailbox; licensesRemoved: number }>
  newMailbox(tenantId: string, m: NewMailbox): R<{ mailbox: Mailbox; failed: { user: string; error: string }[] }>
  permissions(tenantId: string, id: string): R<MailboxPerms>
  permissionsForUser(tenantId: string, user: string): R<UserAccess[]>
  grant(tenantId: string, id: string, user: string, right: MailRight, autoMapping?: boolean): R<MailboxPerms>
  revoke(tenantId: string, id: string, user: string, right: MailRight): R<MailboxPerms>
  autoReply(tenantId: string, id: string): R<AutoReply>
  setAutoReply(tenantId: string, id: string, a: AutoReply): R<AutoReply>
  calendar(tenantId: string, id: string): R<CalendarSettings>
  setCalendar(tenantId: string, id: string, c: CalendarSettings): R<CalendarSettings>
  distGroups(tenantId: string): R<DistGroup[]>
  distGroup(tenantId: string, id: string, kind: DistKind): R<DistGroupDetail>
  newDistGroup(tenantId: string, g: NewDistGroup): R<DistGroup>
  setDistGroup(tenantId: string, id: string, kind: DistKind, p: { owners?: string[]; allowExternal?: boolean; displayName?: string }): R<DistGroup>
  addMember(tenantId: string, id: string, member: string): R<Recipient[]>
  removeMember(tenantId: string, id: string, member: string): R<Recipient[]>
  removeDistGroup(tenantId: string, id: string, kind: DistKind): R<void>
  recipients(tenantId: string): R<Recipient[]>
  setAddresses(tenantId: string, r: Recipient, c: { primary?: string; add?: string[]; remove?: string[] }): R<Recipient>
  inboxRules(tenantId: string): R<InboxRule[]>
  setInboxRule(tenantId: string, mailboxId: string, ruleId: string, action: 'disable' | 'enable' | 'remove'): R<void>
  mailFlow(tenantId: string): R<MailFlow>
  onStatus(cb: (s: { tenantId: string; phase: ExoPhase }) => void): () => void
}

declare global {
  interface Window {
    api: Api
  }
}
