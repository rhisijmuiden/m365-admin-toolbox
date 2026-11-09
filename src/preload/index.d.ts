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
  TemplateChannel,
  ApiResource,
  AppAssignment,
  AppCredential,
  AppGrants,
  AppPermissionState,
  AppReg,
  AppRole,
  AuthMethod,
  CaDraft,
  CaPolicy,
  CaState,
  CertFile,
  DirectoryRef,
  EnterpriseApp,
  EntraGuest,
  ExpiringCred,
  GdapStatus,
  GrantedPermission,
  Invite,
  MfaReport,
  NamedLocation,
  NewAppReg,
  PimApproval,
  PimAudit,
  PimGroupEntry,
  PimRoleSettings,
  ResourceAccess,
  RoleAssign,
  RoleMember,
  RolesData,
  AllowBlockItem,
  AllowBlockList,
  AnalyzedEmail,
  AuditFilter,
  AuditRecord,
  DefenderAction,
  NewAllowBlock,
  PurgeType,
  QuarantineFilter,
  QuarantineMsg,
  QuarantinePreview,
  ReportedMsg,
  RestrictedSender,
  RetractJob,
  RetractSpec,
  RiskDetection,
  RiskyUser,
  SecPolicies,
  SecPolicyKind,
  SecureMailboxResult,
  SignIn,
  TraceEvent,
  TraceFilter,
  TraceRow
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
  entra: EntraApi
  sec: SecApi
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

type OwnerKind = 'applications' | 'servicePrincipals'

export interface EntraApi {
  consent(tenantId: string): R<void>
  apps(tenantId: string): R<AppReg[]>
  app(tenantId: string, id: string): R<AppReg>
  createApp(tenantId: string, n: NewAppReg): R<{ app: AppReg; consentError?: string }>
  updateApp(tenantId: string, id: string, p: { displayName?: string; web?: string[]; spa?: string[]; publicClient?: string[]; permissions?: ResourceAccess[] }): R<AppReg>
  deleteApp(tenantId: string, id: string): R<void>
  setAppRoles(tenantId: string, id: string, roles: AppRole[]): R<AppReg>
  addSecret(tenantId: string, kind: 'app' | 'sp', objectId: string, displayName: string, months: number): R<{ secret: string; keyId: string; end: string }>
  removeCredential(tenantId: string, kind: 'app' | 'sp', objectId: string, c: Pick<AppCredential, 'keyId' | 'kind'>): R<void>
  pickCert(): R<CertFile | null>
  addCert(tenantId: string, appObjectId: string, cert: CertFile, displayName: string): R<AppReg>
  resource(tenantId: string, appId: string): R<ApiResource>
  permissions(tenantId: string, appObjectId: string): R<AppPermissionState>
  grantConsent(tenantId: string, appObjectId: string): R<AppPermissionState>
  removePermission(tenantId: string, appObjectId: string, resourceAppId: string, permId: string): R<AppPermissionState>
  owners(tenantId: string, kind: OwnerKind, id: string): R<DirectoryRef[]>
  addOwner(tenantId: string, kind: OwnerKind, id: string, userId: string): R<DirectoryRef[]>
  removeOwner(tenantId: string, kind: OwnerKind, id: string, userId: string): R<DirectoryRef[]>
  credentials(tenantId: string): R<ExpiringCred[]>
  enterpriseApps(tenantId: string): R<EnterpriseApp[]>
  updateEnterpriseApp(tenantId: string, id: string, p: { enabled?: boolean; assignmentRequired?: boolean; hidden?: boolean }): R<EnterpriseApp>
  assignments(tenantId: string, spId: string): R<AppAssignment[]>
  assign(tenantId: string, spId: string, principalId: string, appRoleId: string): R<AppAssignment[]>
  unassign(tenantId: string, spId: string, assignmentId: string): R<AppAssignment[]>
  grants(tenantId: string): R<AppGrants[]>
  revokeGrant(tenantId: string, spId: string, p: Pick<GrantedPermission, 'grantId' | 'kind' | 'value'>): R<void>
  caPolicies(tenantId: string): R<CaPolicy[]>
  createCa(tenantId: string, d: CaDraft): R<CaPolicy[]>
  setCaState(tenantId: string, id: string, state: CaState): R<CaPolicy[]>
  deleteCa(tenantId: string, id: string): R<CaPolicy[]>
  resolve(tenantId: string, ids: string[]): R<Record<string, string>>
  locations(tenantId: string): R<NamedLocation[]>
  saveLocation(tenantId: string, l: Omit<NamedLocation, 'modified' | 'id'> & { id?: string }): R<NamedLocation[]>
  deleteLocation(tenantId: string, id: string): R<NamedLocation[]>
  roles(tenantId: string): R<RolesData>
  assignRole(tenantId: string, a: RoleAssign): R<RolesData>
  removeRole(tenantId: string, m: Pick<RoleMember, 'id' | 'kind' | 'roleId'> & { principalId: string }): R<RolesData>
  gdap(tenantId: string): R<GdapStatus>
  approvals(tenantId: string): R<PimApproval[]>
  review(tenantId: string, a: Pick<PimApproval, 'approvalId' | 'kind'>, approve: boolean, justification: string): R<void>
  roleSettings(tenantId: string, roleId: string): R<PimRoleSettings>
  setRoleSettings(tenantId: string, roleId: string, s: Omit<PimRoleSettings, 'policyId'>): R<PimRoleSettings>
  pimGroup(tenantId: string, groupId: string): R<PimGroupEntry[]>
  pimGroupAssign(tenantId: string, groupId: string, a: { principalId: string; accessId: 'member' | 'owner'; kind: 'eligible' | 'active'; start?: string; end?: string | null; justification?: string }): R<PimGroupEntry[]>
  pimGroupRemove(tenantId: string, groupId: string, e: Pick<PimGroupEntry, 'accessId' | 'kind'> & { principalId: string }): R<PimGroupEntry[]>
  pimAudit(tenantId: string, days: number): R<PimAudit[]>
  mfa(tenantId: string): R<MfaReport>
  methods(tenantId: string, userId: string): R<AuthMethod[]>
  deleteMethod(tenantId: string, userId: string, m: Pick<AuthMethod, 'id' | 'type'>): R<void>
  resetMfa(tenantId: string, userId: string, revoke: boolean): R<{ failed: { method: string; error: string }[]; methods: AuthMethod[] }>
  guests(tenantId: string): R<EntraGuest[]>
  invite(tenantId: string, i: Invite): R<{ userId: string; redeemUrl: string }>
  deleteGuests(tenantId: string, ids: string[]): R<{ id: string; error: string }[]>
}

type Failed = R<{ failed: { id: string; error: string }[] }>

export interface SecApi {
  consent(tenantId: string, set: 'risk' | 'submissions' | 'defender'): R<void>
  quarantine(tenantId: string, f: QuarantineFilter): R<QuarantineMsg[]>
  quarantinePreview(tenantId: string, id: string): R<QuarantinePreview>
  quarantineRelease(tenantId: string, ids: string[], allowSender: boolean, reportFalsePositive: boolean): Failed
  quarantineDelete(tenantId: string, ids: string[]): Failed
  allowBlock(tenantId: string): R<AllowBlockItem[]>
  newAllowBlock(tenantId: string, n: NewAllowBlock): R<AllowBlockItem[]>
  removeAllowBlock(tenantId: string, listType: AllowBlockList, ids: string[]): R<AllowBlockItem[]>
  policies(tenantId: string): R<SecPolicies>
  setPolicy(tenantId: string, kind: SecPolicyKind, name: string, params: Record<string, unknown>): R<SecPolicies>
  setPolicyRule(tenantId: string, kind: SecPolicyKind, rule: string, enabled: boolean): R<SecPolicies>
  restricted(tenantId: string): R<RestrictedSender[]>
  releaseSender(tenantId: string, s: RestrictedSender): R<RestrictedSender[]>
  trace(tenantId: string, f: TraceFilter): R<TraceRow[]>
  traceDetail(tenantId: string, traceId: string, recipient: string): R<TraceEvent[]>
  auditStatus(tenantId: string): R<{ enabled: boolean }>
  enableAudit(tenantId: string): R<{ enabled: boolean }>
  audit(tenantId: string, f: AuditFilter): R<AuditRecord[]>
  secureMailbox(tenantId: string, mailbox: string, rules: boolean, forwarding: boolean): R<SecureMailboxResult>
  signIns(tenantId: string, userId: string): R<SignIn[]>
  riskyUsers(tenantId: string): R<RiskyUser[]>
  riskDetections(tenantId: string, days: number): R<RiskDetection[]>
  setRisk(tenantId: string, userIds: string[], action: 'dismiss' | 'confirmCompromised'): R<void>
  reported(tenantId: string, days: number): R<ReportedMsg[]>
  retractJobs(tenantId: string): R<RetractJob[]>
  startRetract(tenantId: string, s: RetractSpec): R<RetractJob>
  purgeRetract(jobId: string, type: PurgeType): R<RetractJob>
  discardRetract(jobId: string): R<void>
  analyzed(tenantId: string, s: RetractSpec): R<AnalyzedEmail[]>
  remediate(tenantId: string, emails: Pick<AnalyzedEmail, 'networkMessageId' | 'recipient'>[], action: DefenderAction, name: string): R<void>
  onStatus(cb: (s: { tenantId: string; phase: ExoPhase }) => void): () => void
}

declare global {
  interface Window {
    api: Api
  }
}
