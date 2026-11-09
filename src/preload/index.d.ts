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
  User,
  CsUser,
  GroupPolicyAssignment,
  NewTeam,
  NewTeamResult,
  NumberType,
  PolicyType,
  Team,
  TeamActivityReport,
  TeamChannel,
  TeamGuest,
  TeamMember,
  TeamsAccess,
  TeamsFederation,
  TeamsGuestSettings,
  TeamsPhone,
  TeamsPolicies,
  TeamTemplate,
  TemplateChannel
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
  teams: TeamsApi
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

export interface TeamsApi {
  consent(tenantId: string): R<void>
  consentReports(tenantId: string): R<void>
  signIn(tenantId: string): R<void>
  list(tenantId: string): R<Team[]>
  get(tenantId: string, teamId: string): R<Team>
  activity(tenantId: string, teamIds: string[]): R<TeamActivityReport>
  create(tenantId: string, t: NewTeam): R<NewTeamResult>
  update(tenantId: string, teamId: string, p: { displayName?: string; description?: string; visibility?: 'private' | 'public' }): R<Team>
  archive(tenantId: string, teamId: string, archived: boolean, siteReadOnly?: boolean): R<Team>
  members(tenantId: string, teamId: string): R<TeamMember[]>
  addMember(tenantId: string, teamId: string, user: string, owner: boolean): R<TeamMember[]>
  setMemberRole(tenantId: string, teamId: string, membershipId: string, owner: boolean): R<TeamMember[]>
  removeMember(tenantId: string, teamId: string, membershipId: string): R<TeamMember[]>
  channels(tenantId: string, teamId: string): R<TeamChannel[]>
  addChannel(tenantId: string, teamId: string, c: TemplateChannel, owners: string[]): R<{ channel: TeamChannel; failed: { what: string; error: string }[] }>
  removeChannel(tenantId: string, teamId: string, channelId: string): R<void>
  channelMembers(tenantId: string, teamId: string, channelId: string): R<TeamMember[]>
  addChannelMember(tenantId: string, teamId: string, channelId: string, user: string, owner: boolean): R<TeamMember[]>
  setChannelMemberRole(tenantId: string, teamId: string, channelId: string, membershipId: string, owner: boolean): R<TeamMember[]>
  removeChannelMember(tenantId: string, teamId: string, channelId: string, membershipId: string): R<TeamMember[]>
  guests(tenantId: string, teams: { id: string; displayName: string }[]): R<TeamGuest[]>
  removeGuest(tenantId: string, teamId: string, userId: string): R<void>
  deleteGuest(tenantId: string, userId: string): R<void>
  templates(): R<TeamTemplate[]>
  saveTemplates(items: TeamTemplate[]): R<TeamTemplate[]>
  policies(tenantId: string): R<TeamsPolicies>
  grantUsers(tenantId: string, type: PolicyType, policy: string | null, users: string[]): R<{ users: CsUser[]; failed: { user: string; error: string }[] }>
  grantGroup(tenantId: string, a: GroupPolicyAssignment): R<void>
  removeGroupPolicy(tenantId: string, type: PolicyType, groupId: string): R<void>
  access(tenantId: string): R<TeamsAccess>
  setFederation(tenantId: string, f: TeamsFederation): R<TeamsAccess>
  setGuest(tenantId: string, g: Partial<TeamsGuestSettings>): R<TeamsAccess>
  phone(tenantId: string): R<TeamsPhone>
  assignNumber(tenantId: string, user: string, number: string, type: NumberType, locationId?: string): R<CsUser>
  unassignNumber(tenantId: string, user: string, number: string, type: NumberType): R<CsUser>
  enterpriseVoice(tenantId: string, user: string, enabled: boolean): R<CsUser>
  onStatus(cb: (s: { tenantId: string; phase: ExoPhase }) => void): () => void
}

declare global {
  interface Window {
    api: Api
  }
}
