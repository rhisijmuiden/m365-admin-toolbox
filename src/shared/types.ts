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

// ---------- Teams (phase 3) ----------

export interface TeamPerson {
  id: string
  displayName: string
  upn: string
}

export interface Team {
  id: string
  displayName: string
  description: string | null
  visibility: 'private' | 'public' | 'hiddenMembership' | string
  mail: string | null
  created: string | null
  owners: TeamPerson[]
  /** null when the team's settings could not be read. */
  archived: boolean | null
}

export interface TeamActivity {
  teamId: string
  lastActivity: string | null
  activeUsers: number | null
  guests: number | null
  messages: number | null
}

export interface TeamActivityReport {
  rows: TeamActivity[]
  reportDate: string | null
  /** Why no report rows are available: no consent, names concealed, no access, other error. */
  issue?: 'consent' | 'concealed' | 'denied' | 'error'
}

export interface TeamMember {
  /** conversationMember id, used to change or remove the membership. */
  id: string
  userId: string
  displayName: string
  email: string | null
  roles: string[]
}

export type ChannelType = 'standard' | 'private' | 'shared'

export interface TeamChannel {
  id: string
  displayName: string
  description: string | null
  type: ChannelType
  created: string | null
}

export interface TemplateChannel {
  name: string
  description?: string
  type: ChannelType
}

export interface TeamTemplate {
  id: string
  name: string
  description: string
  visibility: 'private' | 'public'
  channels: TemplateChannel[]
}

export interface NewTeam {
  displayName: string
  description: string
  visibility: 'private' | 'public'
  /** User ids or UPNs. At least one owner is required. */
  owners: string[]
  members: string[]
  channels: TemplateChannel[]
}

export interface NewTeamResult {
  team: Team
  /** Steps after the team itself (extra members, private and shared channels) that failed. */
  failed: { what: string; error: string }[]
}

export interface TeamGuest {
  teamId: string
  teamName: string
  userId: string
  displayName: string
  mail: string | null
  upn: string
  created: string | null
  state: string | null
  lastSignIn: string | null
}

export type PolicyType = 'TeamsMeetingPolicy' | 'TeamsMessagingPolicy' | 'TeamsAppSetupPolicy' | 'TeamsCallingPolicy' | 'OnlineVoiceRoutingPolicy'

export interface CsUser {
  id: string
  upn: string
  displayName: string
  enabled: boolean
  /** Directly assigned policy per type; null means the Global policy or one inherited from a group. */
  policies: Partial<Record<PolicyType, string | null>>
  lineUri: string | null
  enterpriseVoice: boolean
  phoneSystem: boolean
}

export interface GroupPolicyAssignment {
  groupId: string
  type: PolicyType
  policy: string | null
  rank: number
}

export interface TeamsPolicies {
  catalog: Record<PolicyType, { name: string; description: string; isGlobal: boolean }[]>
  users: CsUser[]
  groups: GroupPolicyAssignment[]
}

export interface TeamsFederation {
  allowFederated: boolean
  allowAll: boolean
  allowed: string[]
  blocked: string[]
  consumer: boolean
  consumerInbound: boolean
}

export interface TeamsGuestSettings {
  allowGuest: boolean
  privateCalling: boolean
  ipVideo: boolean
  meetNow: boolean
  screenSharing: string
  chat: boolean
  editMessages: boolean
  deleteMessages: boolean
  giphy: boolean
}

export interface TeamsAccess {
  federation: TeamsFederation
  guest: TeamsGuestSettings
}

export type NumberType = 'CallingPlan' | 'DirectRouting' | 'OperatorConnect'

export interface PhoneNumber {
  number: string
  type: string
  activation: string
  capability: string[]
  /** Object id of the user or resource account, null when free. */
  assignedTo: string | null
  status: string
  city: string | null
}

export interface TeamsPhone {
  numbers: PhoneNumber[]
  users: CsUser[]
  calling: string[]
  routing: string[]
}

// ---------- Entra ID (phase 4) ----------

export type CredKind = 'secret' | 'cert'

export interface AppCredential {
  keyId: string
  kind: CredKind
  displayName: string | null
  /** Last characters of a secret (hint) or the certificate thumbprint. */
  hint: string | null
  start: string | null
  end: string | null
}

export type SignInAudience = 'AzureADMyOrg' | 'AzureADMultipleOrgs' | 'AzureADandPersonalMicrosoftAccount' | 'PersonalMicrosoftAccount'

export interface ResourceAccess {
  resourceAppId: string
  access: { id: string; type: 'Scope' | 'Role' }[]
}

export interface AppRole {
  id: string
  displayName: string
  description: string
  value: string | null
  allowedMemberTypes: ('User' | 'Application')[]
  isEnabled: boolean
}

export interface AppReg {
  /** Object id of the application. */
  id: string
  appId: string
  displayName: string
  audience: SignInAudience | string
  created: string | null
  web: string[]
  spa: string[]
  publicClient: string[]
  permissions: ResourceAccess[]
  appRoles: AppRole[]
  credentials: AppCredential[]
}

export interface PermissionDef {
  id: string
  type: 'Scope' | 'Role'
  value: string
  name: string
  description: string
  /** Delegated scopes that need admin consent, and every application permission. */
  admin: boolean
}

export interface ApiResource {
  appId: string
  spId: string
  displayName: string
  permissions: PermissionDef[]
}

export interface AppPermissionState {
  /** Service principal of the app in this tenant, null when it has none yet. */
  spId: string | null
  resources: ApiResource[]
  /** Granted for the whole tenant: delegated scope values and app role ids per resource appId. */
  granted: Record<string, { scopes: string[]; roles: string[] }>
}

export interface NewAppReg {
  displayName: string
  audience: SignInAudience
  platform: 'none' | 'web' | 'spa' | 'publicClient'
  redirectUris: string[]
  permissions: ResourceAccess[]
  adminConsent: boolean
}

export interface DirectoryRef {
  id: string
  displayName: string
  /** UPN for users, mail for groups, appId for service principals. */
  detail: string | null
  type: 'user' | 'group' | 'servicePrincipal' | string
}

export interface CertFile {
  key: string
  thumbprint: string
  subject: string
  start: string
  end: string
  fileName: string
}

export interface ExpiringCred extends AppCredential {
  /** Object id of the application or service principal. */
  objectId: string
  appId: string
  appName: string
  source: 'app' | 'sp'
}

export interface EnterpriseApp {
  id: string
  appId: string
  displayName: string
  enabled: boolean
  assignmentRequired: boolean
  hidden: boolean
  publisher: string | null
  /** True for apps published by Microsoft itself. */
  microsoft: boolean
  /** True when the app registration lives in this tenant. */
  ownTenant: boolean
  type: string
  appRoles: AppRole[]
  homepage: string | null
}

export interface AppAssignment {
  id: string
  principalId: string
  principalName: string
  principalType: string
  appRoleId: string
  created: string | null
}

export interface GrantedPermission {
  /** oauth2PermissionGrant id or appRoleAssignment id, used to revoke. */
  grantId: string
  kind: 'delegated' | 'application'
  resource: string
  value: string
  /** Delegated: AllPrincipals (admin consent) or a single user's consent. */
  consent: 'admin' | 'user' | null
  risky: boolean
}

export interface AppGrants {
  spId: string
  appId: string
  displayName: string
  publisher: string | null
  microsoft: boolean
  enabled: boolean
  /** Users who gave their own consent (delegated grants with consentType Principal). */
  userConsents: number
  permissions: GrantedPermission[]
}

export type CaState = 'enabled' | 'disabled' | 'enabledForReportingButNotEnforced'

export interface CaPolicy {
  id: string
  displayName: string
  state: CaState
  created: string | null
  modified: string | null
  /** The full policy as Graph returns it, for the details view. */
  raw: Record<string, unknown>
}

export interface CaDraft {
  displayName: string
  state: CaState
  users: 'all' | 'guests' | 'admins' | 'selected'
  include: string[]
  exclude: string[]
  apps: 'all' | 'office365' | 'azureManagement'
  clientApps: 'all' | 'legacy'
  locations: 'any' | 'excludeTrusted' | 'selected'
  locationIds: string[]
  grant: 'block' | 'mfa' | 'compliant' | 'mfaOrCompliant'
}

export interface NamedLocation {
  id: string
  displayName: string
  kind: 'ip' | 'country'
  trusted: boolean
  ranges: string[]
  countries: string[]
  includeUnknown: boolean
  modified: string | null
}

export interface RoleDef {
  id: string
  templateId: string
  displayName: string
  description: string
  builtIn: boolean
  privileged: boolean
}

export interface RoleMember {
  /** roleAssignment id (permanent), or the schedule instance id. */
  id: string
  roleId: string
  principal: DirectoryRef
  kind: 'permanent' | 'active' | 'eligible'
  start: string | null
  end: string | null
}

export interface RolesData {
  roles: RoleDef[]
  members: RoleMember[]
  /** False when PIM schedules could not be read (no Entra ID P2). */
  pim: boolean
}

export interface RoleAssign {
  roleId: string
  principalId: string
  kind: 'permanent' | 'active' | 'eligible'
  start?: string
  end?: string | null
  justification?: string
}

export interface GdapRelationship {
  id: string
  displayName: string
  status: string
  start: string | null
  end: string | null
  autoExtend: string | null
  /** Role template ids in the relationship. */
  roles: string[]
  groups: { id: string; name: string; roles: string[] }[]
}

export interface GdapStatus {
  /** False for tenants added with an own admin login. */
  gdap: boolean
  relationships: GdapRelationship[]
  error?: string
}

export interface PimApproval {
  approvalId: string
  requestId: string
  kind: 'role' | 'group'
  targetId: string
  targetName: string
  principal: DirectoryRef
  justification: string | null
  created: string
  accessId?: string
  start: string | null
  end: string | null
}

export interface PimRoleSettings {
  policyId: string
  maxHours: number
  mfa: boolean
  justification: boolean
  ticket: boolean
  approval: boolean
  approvers: DirectoryRef[]
}

export interface PimGroupEntry {
  id: string
  principal: DirectoryRef
  accessId: 'member' | 'owner'
  kind: 'eligible' | 'active' | 'permanent'
  start: string | null
  end: string | null
}

export interface PimAudit {
  id: string
  date: string
  activity: string
  by: string | null
  target: string | null
  result: string
  reason: string | null
}

export interface AuthMethod {
  id: string
  type: string
  label: string | null
  created: string | null
}

export interface UserMfa {
  userId: string
  displayName: string
  upn: string
  methods: string[]
  mfaCapable: boolean
  defaultMethod: string | null
  isAdmin: boolean
}

export interface MfaReport {
  source: 'report' | 'methods'
  users: UserMfa[]
}

export interface EntraGuest {
  id: string
  displayName: string
  mail: string | null
  upn: string
  state: string | null
  stateChanged: string | null
  created: string | null
  lastSignIn: string | null
  enabled: boolean
  company: string | null
}

export interface Invite {
  email: string
  displayName: string
  message?: string
  send: boolean
}
