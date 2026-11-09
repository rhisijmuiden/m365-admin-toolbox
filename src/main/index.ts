import { app, BrowserWindow, ipcMain, nativeTheme, shell } from 'electron'
import { join } from 'path'
import type {
  AppInfo,
  AutoReply,
  CalendarSettings,
  DistKind,
  LicenseChange,
  MailRight,
  NewDistGroup,
  NewGroup,
  NewMailbox,
  NewUser,
  PimKind,
  PimRequest,
  Recipient,
  Result,
  Settings,
  GroupPolicyAssignment,
  NewTeam,
  NumberType,
  PolicyType,
  TeamsFederation,
  TeamsGuestSettings,
  TeamTemplate,
  TemplateChannel,
  Tenant,
  TenantMode,
  AppCredential,
  AppRole,
  AuthMethod,
  CaDraft,
  CaState,
  CertFile,
  GrantedPermission,
  Invite,
  NewAppReg,
  PimApproval,
  PimRoleSettings,
  RoleAssign,
  AllowBlockList,
  AnalyzedEmail,
  AuditFilter,
  DefenderAction,
  NewAllowBlock,
  PurgeType,
  QuarantineFilter,
  RestrictedSender,
  RetractSpec,
  SecPolicyKind,
  TraceFilter
} from '@shared/types'
import { consentScopes, exoSignIn, forget, homeTenantOf, signInInteractive, teamsSignIn } from './auth'
import * as exo from './exchange'
import * as entra from './entra'
import * as exoPs from './exo'
import * as graph from './graph'
import * as pim from './pim'
import * as sec from './security'
import { initPimWatch } from './pimWatch'
import { setPsWindow } from './ps'
import { settings, teamTemplates, tenants } from './store'
import * as teams from './teams'
import { applyChannel, checkForUpdates, downloadUpdate, initUpdater, installUpdate, updateState } from './updater'

const REPO = 'rhisijmuiden/m365-admin-toolbox'
const COLORS = ['#1f6feb', '#c2381f', '#0a6f75', '#5b5fc7', '#a15c00', '#1a7f4b', '#b4237a', '#3a475a']

let win: BrowserWindow | null = null

function createWindow() {
  const s = settings.get()
  nativeTheme.themeSource = s.theme
  win = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 980,
    minHeight: 640,
    show: false,
    title: 'M365 Admin Toolbox',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#0e141c' : '#f3f5f8',
    autoHideMenuBar: true,
    // Packaged builds take the icon from the exe; in dev point at the source PNG.
    icon: app.isPackaged ? undefined : join(__dirname, '../../build/icon.png'),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  })
  win.once('ready-to-show', () => win?.show())
  // Links from the UI (GitHub, docs) open in the default browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https:\/\//.test(url)) shell.openExternal(url)
    return { action: 'deny' }
  })
  if (!app.isPackaged && process.env.ELECTRON_RENDERER_URL) win.loadURL(process.env.ELECTRON_RENDERER_URL)
  else win.loadFile(join(__dirname, '../renderer/index.html'))
  initUpdater(win, s.channel)
  initPimWatch(win)
  setPsWindow(win)
}

/** Wraps a handler so errors reach the renderer as a Result instead of a rejected invoke. */
function handle<A extends unknown[], T>(name: string, fn: (...args: A) => Promise<T> | T) {
  ipcMain.handle(name, async (_e, ...args: unknown[]): Promise<Result<T>> => {
    try {
      return { ok: true, data: await fn(...(args as A)) }
    } catch (e) {
      const err = e as Error & { code?: string; status?: number }
      return { ok: false, error: err.message || String(e), code: err.code, status: err.status }
    }
  })
}

async function addTenant(mode: TenantMode, target?: string): Promise<Tenant> {
  const authority = mode === 'gdap' ? (target || '').trim() : 'organizations'
  if (!authority) throw new Error('Tenant domain or ID is required')
  const r = await signInInteractive(authority)
  if (!r.account) throw new Error('Sign-in returned no account')
  const org = await graph.organization(r.accessToken)
  const existing = tenants.get(r.tenantId)
  const t: Tenant = {
    tenantId: r.tenantId,
    name: org.name,
    defaultDomain: org.defaultDomain,
    domains: org.domains,
    mode,
    homeAccountId: r.account.homeAccountId,
    username: r.account.username,
    color: existing?.color || COLORS[tenants.list().length % COLORS.length],
    addedAt: existing?.addedAt || new Date().toISOString()
  }
  tenants.upsert(t)
  settings.set({ lastTenantId: t.tenantId })
  return t
}

/** Signs in again for a tenant whose refresh token expired or was revoked. */
async function reauth(tenantId: string): Promise<Tenant> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error('Unknown tenant')
  const r = await signInInteractive(tenantId, t.username)
  if (!r.account) throw new Error('Sign-in returned no account')
  if (r.tenantId !== tenantId) throw new Error('Signed in to a different tenant')
  const next = { ...t, homeAccountId: r.account.homeAccountId, username: r.account.username }
  tenants.upsert(next)
  return next
}

/** For GDAP tenants: sign in to the partner tenant, which PIM for groups needs (consent there too). */
async function reauthHome(tenantId: string) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error('Unknown tenant')
  await signInInteractive(homeTenantOf(tenantId), t.username)
}

function registerIpc() {
  handle('app:info', (): AppInfo => ({
    version: app.getVersion(),
    channel: settings.get().channel,
    isPackaged: app.isPackaged,
    electron: process.versions.electron,
    repo: REPO
  }))
  handle('settings:get', () => settings.get())
  handle('settings:set', (patch: Partial<Settings>) => {
    const next = settings.set(patch)
    if (patch.theme) nativeTheme.themeSource = next.theme
    if (patch.channel) applyChannel(next.channel)
    return next
  })
  handle('tenants:list', () => tenants.list())
  handle('tenants:add', (mode: TenantMode, target?: string) => addTenant(mode, target))
  handle('tenants:reauth', (id: string) => reauth(id))
  handle('tenants:remove', async (id: string) => {
    const t = tenants.get(id)
    tenants.remove(id)
    exoPs.closeSession(id)
    teams.closeSession(id)
    sec.closeSession(id)
    if (t) await forget(t.homeAccountId)
    return tenants.list()
  })
  handle('users:list', (id: string) => graph.listUsers(id))
  handle('users:create', (id: string, u: NewUser) => graph.createUser(id, u))
  handle('users:setSignIn', (id: string, userId: string, enabled: boolean) => graph.setSignIn(id, userId, enabled))
  handle('users:resetPassword', (id: string, userId: string, pw: string, force: boolean) =>
    graph.resetPassword(id, userId, pw, force)
  )
  handle('users:createTap', (id: string, userId: string, minutes: number, once: boolean) =>
    graph.createTap(id, userId, minutes, once)
  )
  handle('users:revokeSessions', (id: string, userId: string) => graph.revokeSessions(id, userId))
  handle('licenses:list', (id: string) => graph.listSkus(id))
  handle('licenses:change', (id: string, changes: LicenseChange[]) => graph.changeLicenses(id, changes))
  handle('groups:list', (id: string) => graph.listGroups(id))
  handle('groups:create', (id: string, g: NewGroup) => graph.createGroup(id, g))
  handle('tenants:reauthHome', (id: string) => reauthHome(id))
  handle('pim:load', (id: string) => pim.load(id))
  handle('pim:activate', (id: string, r: PimRequest) => pim.activate(id, r))
  handle('pim:extend', (id: string, r: PimRequest) => pim.extend(id, r))
  handle('pim:deactivate', (id: string, kind: PimKind, targetId: string) => pim.deactivate(id, kind, targetId))
  handle('pim:cancel', (id: string, kind: PimKind, requestId: string) => pim.cancel(id, kind, requestId))
  handle('exo:signIn', (id: string) => exoSignIn(id))
  handle('exo:consentReports', (id: string) => consentScopes(id, exo.REPORT_SCOPES))
  handle('exo:mailboxes', (id: string) => exo.listMailboxes(id))
  handle('exo:mailbox', (id: string, mbx: string) => exo.getMailbox(id, mbx))
  handle('exo:sizes', (id: string, upns: string[], live?: boolean) => exo.mailboxSizes(id, upns, live))
  handle('exo:setMailbox', (id: string, mbx: string, p: Parameters<typeof exo.setMailbox>[2]) => exo.setMailbox(id, mbx, p))
  handle('exo:enableArchive', (id: string, mbx: string) => exo.enableArchive(id, mbx))
  handle('exo:convertToShared', (id: string, mbx: string, userId: string, removeLic: boolean) => exo.convertToShared(id, mbx, userId, removeLic))
  handle('exo:newMailbox', (id: string, m: NewMailbox) => exo.newMailbox(id, m))
  handle('exo:permissions', (id: string, mbx: string) => exo.permissions(id, mbx))
  handle('exo:permissionsForUser', (id: string, user: string) => exo.permissionsForUser(id, user))
  handle('exo:grant', (id: string, mbx: string, user: string, right: MailRight, autoMap?: boolean) => exo.grant(id, mbx, user, right, autoMap))
  handle('exo:revoke', (id: string, mbx: string, user: string, right: MailRight) => exo.revoke(id, mbx, user, right))
  handle('exo:autoReply', (id: string, mbx: string) => exo.autoReply(id, mbx))
  handle('exo:setAutoReply', (id: string, mbx: string, a: AutoReply) => exo.setAutoReply(id, mbx, a))
  handle('exo:calendar', (id: string, mbx: string) => exo.calendar(id, mbx))
  handle('exo:setCalendar', (id: string, mbx: string, c: CalendarSettings) => exo.setCalendar(id, mbx, c))
  handle('exo:distGroups', (id: string) => exo.listDistGroups(id))
  handle('exo:distGroup', (id: string, g: string, kind: DistKind) => exo.distGroupDetail(id, g, kind))
  handle('exo:newDistGroup', (id: string, g: NewDistGroup) => exo.newDistGroup(id, g))
  handle('exo:setDistGroup', (id: string, g: string, kind: DistKind, p: Parameters<typeof exo.setDistGroup>[3]) => exo.setDistGroup(id, g, kind, p))
  handle('exo:addMember', (id: string, g: string, m: string) => exo.addMember(id, g, m))
  handle('exo:removeMember', (id: string, g: string, m: string) => exo.removeMember(id, g, m))
  handle('exo:removeDistGroup', (id: string, g: string, kind: DistKind) => exo.removeDistGroup(id, g, kind))
  handle('exo:recipients', (id: string) => exo.recipients(id))
  handle('exo:setAddresses', (id: string, r: Recipient, c: Parameters<typeof exo.setAddresses>[2]) => exo.setAddresses(id, r, c))
  handle('exo:inboxRules', (id: string) => exo.inboxRules(id))
  handle('exo:setInboxRule', (id: string, mbx: string, rule: string, action: 'disable' | 'enable' | 'remove') => exo.setInboxRule(id, mbx, rule, action))
  handle('exo:mailFlow', (id: string) => exo.mailFlow(id))
  handle('teams:consent', (id: string) => consentScopes(id, teams.TEAMS_SCOPES))
  handle('teams:consentReports', (id: string) => consentScopes(id, exo.REPORT_SCOPES))
  handle('teams:signIn', (id: string) => teamsSignIn(id))
  handle('teams:list', (id: string) => teams.listTeams(id))
  handle('teams:get', (id: string, teamId: string) => teams.team(id, teamId))
  handle('teams:activity', (id: string, teamIds: string[]) => teams.teamActivity(id, teamIds))
  handle('teams:create', (id: string, t: NewTeam) => teams.createTeam(id, t))
  handle('teams:update', (id: string, teamId: string, p: Parameters<typeof teams.updateTeam>[2]) => teams.updateTeam(id, teamId, p))
  handle('teams:archive', (id: string, teamId: string, archived: boolean, siteReadOnly?: boolean) => teams.archive(id, teamId, archived, siteReadOnly))
  handle('teams:members', (id: string, teamId: string) => teams.members(id, teamId))
  handle('teams:addMember', (id: string, teamId: string, user: string, owner: boolean) => teams.addMember(id, teamId, user, owner))
  handle('teams:setMemberRole', (id: string, teamId: string, mid: string, owner: boolean) => teams.setMemberRole(id, teamId, mid, owner))
  handle('teams:removeMember', (id: string, teamId: string, mid: string) => teams.removeMember(id, teamId, mid))
  handle('teams:channels', (id: string, teamId: string) => teams.channels(id, teamId))
  handle('teams:addChannel', (id: string, teamId: string, c: TemplateChannel, owners: string[]) => teams.addChannel(id, teamId, c, owners))
  handle('teams:removeChannel', (id: string, teamId: string, cid: string) => teams.removeChannel(id, teamId, cid))
  handle('teams:channelMembers', (id: string, teamId: string, cid: string) => teams.channelMembers(id, teamId, cid))
  handle('teams:addChannelMember', (id: string, teamId: string, cid: string, user: string, owner: boolean) =>
    teams.addChannelMember(id, teamId, cid, user, owner)
  )
  handle('teams:setChannelMemberRole', (id: string, teamId: string, cid: string, mid: string, owner: boolean) =>
    teams.setChannelMemberRole(id, teamId, cid, mid, owner)
  )
  handle('teams:removeChannelMember', (id: string, teamId: string, cid: string, mid: string) => teams.removeChannelMember(id, teamId, cid, mid))
  handle('teams:guests', (id: string, list: { id: string; displayName: string }[]) => teams.guests(id, list))
  handle('teams:removeGuest', (id: string, teamId: string, userId: string) => teams.removeGuest(id, teamId, userId))
  handle('teams:deleteGuest', (id: string, userId: string) => teams.deleteGuest(id, userId))
  handle('teams:templates', () => teamTemplates.list())
  handle('teams:saveTemplates', (items: TeamTemplate[]) => teamTemplates.save(items))
  handle('teams:policies', (id: string) => teams.policies(id))
  handle('teams:grantUsers', (id: string, type: PolicyType, policy: string | null, users: string[]) => teams.grantUsers(id, type, policy, users))
  handle('teams:grantGroup', (id: string, a: GroupPolicyAssignment) => teams.grantGroup(id, a))
  handle('teams:removeGroupPolicy', (id: string, type: PolicyType, groupId: string) => teams.removeGroupPolicy(id, type, groupId))
  handle('teams:access', (id: string) => teams.access(id))
  handle('teams:setFederation', (id: string, f: TeamsFederation) => teams.setFederation(id, f))
  handle('teams:setGuest', (id: string, g: Partial<TeamsGuestSettings>) => teams.setGuest(id, g))
  handle('teams:phone', (id: string) => teams.phone(id))
  handle('teams:assignNumber', (id: string, user: string, num: string, type: NumberType, loc?: string) => teams.assignNumber(id, user, num, type, loc))
  handle('teams:unassignNumber', (id: string, user: string, num: string, type: NumberType) => teams.unassignNumber(id, user, num, type))
  handle('teams:enterpriseVoice', (id: string, user: string, enabled: boolean) => teams.enterpriseVoice(id, user, enabled))
  handle('entra:consent', (id: string) => consentScopes(id, entra.ENTRA_SCOPES))
  handle('entra:apps', (id: string) => entra.listApps(id))
  handle('entra:app', (id: string, appId: string) => entra.getApp(id, appId))
  handle('entra:createApp', (id: string, n: NewAppReg) => entra.createApp(id, n))
  handle('entra:updateApp', (id: string, appId: string, p: Parameters<typeof entra.updateApp>[2]) => entra.updateApp(id, appId, p))
  handle('entra:deleteApp', (id: string, appId: string) => entra.deleteApp(id, appId))
  handle('entra:setAppRoles', (id: string, appId: string, r: AppRole[]) => entra.setAppRoles(id, appId, r))
  handle('entra:addSecret', (id: string, kind: 'app' | 'sp', obj: string, name: string, months: number) => entra.addSecret(id, kind, obj, name, months))
  handle('entra:removeCredential', (id: string, kind: 'app' | 'sp', obj: string, c: Pick<AppCredential, 'keyId' | 'kind'>) => entra.removeCredential(id, kind, obj, c))
  handle('entra:pickCert', () => entra.pickCert())
  handle('entra:addCert', (id: string, appId: string, c: CertFile, name: string) => entra.addCert(id, appId, c, name))
  handle('entra:resource', (id: string, appId: string) => entra.apiResource(id, appId))
  handle('entra:permissions', (id: string, appId: string) => entra.appPermissions(id, appId))
  handle('entra:grantConsent', (id: string, appId: string) => entra.grantConsent(id, appId))
  handle('entra:removePermission', (id: string, appId: string, res: string, perm: string) => entra.removePermission(id, appId, res, perm))
  handle('entra:owners', (id: string, kind: 'applications' | 'servicePrincipals', obj: string) => entra.owners(id, kind, obj))
  handle('entra:addOwner', (id: string, kind: 'applications' | 'servicePrincipals', obj: string, user: string) => entra.addOwner(id, kind, obj, user))
  handle('entra:removeOwner', (id: string, kind: 'applications' | 'servicePrincipals', obj: string, user: string) => entra.removeOwner(id, kind, obj, user))
  handle('entra:credentials', (id: string) => entra.credentials(id))
  handle('entra:enterpriseApps', (id: string) => entra.enterpriseApps(id))
  handle('entra:updateEnterpriseApp', (id: string, sp: string, p: Parameters<typeof entra.updateEnterpriseApp>[2]) => entra.updateEnterpriseApp(id, sp, p))
  handle('entra:assignments', (id: string, sp: string) => entra.assignments(id, sp))
  handle('entra:assign', (id: string, sp: string, principal: string, role: string) => entra.assign(id, sp, principal, role))
  handle('entra:unassign', (id: string, sp: string, a: string) => entra.unassign(id, sp, a))
  handle('entra:grants', (id: string) => entra.grants(id))
  handle('entra:revokeGrant', (id: string, sp: string, p: Pick<GrantedPermission, 'grantId' | 'kind' | 'value'>) => entra.revokeGrant(id, sp, p))
  handle('entra:caPolicies', (id: string) => entra.caPolicies(id))
  handle('entra:createCa', (id: string, d: CaDraft) => entra.createCa(id, d))
  handle('entra:setCaState', (id: string, p: string, s: CaState) => entra.setCaState(id, p, s))
  handle('entra:deleteCa', (id: string, p: string) => entra.deleteCa(id, p))
  handle('entra:resolve', (id: string, ids: string[]) => entra.resolveIds(id, ids))
  handle('entra:locations', (id: string) => entra.namedLocations(id))
  handle('entra:saveLocation', (id: string, l: Parameters<typeof entra.saveLocation>[1]) => entra.saveLocation(id, l))
  handle('entra:deleteLocation', (id: string, l: string) => entra.deleteLocation(id, l))
  handle('entra:roles', (id: string) => entra.roles(id))
  handle('entra:assignRole', (id: string, a: RoleAssign) => entra.assignRole(id, a))
  handle('entra:removeRole', (id: string, m: Parameters<typeof entra.removeRole>[1]) => entra.removeRole(id, m))
  handle('entra:gdap', (id: string) => entra.gdap(id))
  handle('entra:approvals', (id: string) => entra.approvals(id))
  handle('entra:review', (id: string, a: Pick<PimApproval, 'approvalId' | 'kind'>, ok: boolean, why: string) => entra.review(id, a, ok, why))
  handle('entra:roleSettings', (id: string, role: string) => entra.roleSettings(id, role))
  handle('entra:setRoleSettings', (id: string, role: string, s: Omit<PimRoleSettings, 'policyId'>) => entra.setRoleSettings(id, role, s))
  handle('entra:pimGroup', (id: string, g: string) => entra.pimGroup(id, g))
  handle('entra:pimGroupAssign', (id: string, g: string, a: Parameters<typeof entra.pimGroupAssign>[2]) => entra.pimGroupAssign(id, g, a))
  handle('entra:pimGroupRemove', (id: string, g: string, e: Parameters<typeof entra.pimGroupRemove>[2]) => entra.pimGroupRemove(id, g, e))
  handle('entra:pimAudit', (id: string, days: number) => entra.pimAudit(id, days))
  handle('entra:mfa', (id: string) => entra.mfaReport(id))
  handle('entra:methods', (id: string, user: string) => entra.methods(id, user))
  handle('entra:deleteMethod', (id: string, user: string, m: Pick<AuthMethod, 'id' | 'type'>) => entra.deleteMethod(id, user, m))
  handle('entra:resetMfa', (id: string, user: string, revoke: boolean) => entra.resetMfa(id, user, revoke))
  handle('entra:guests', (id: string) => entra.guestUsers(id))
  handle('entra:invite', (id: string, i: Invite) => entra.invite(id, i))
  handle('entra:deleteGuests', (id: string, ids: string[]) => entra.deleteGuests(id, ids))
  handle('sec:consent', (id: string, set: keyof typeof sec.SCOPE_SETS) => consentScopes(id, sec.SCOPE_SETS[set]))
  handle('sec:quarantine', (id: string, f: QuarantineFilter) => sec.quarantine(id, f))
  handle('sec:quarantinePreview', (id: string, q: string) => sec.quarantinePreview(id, q))
  handle('sec:quarantineRelease', (id: string, ids: string[], allow: boolean, report: boolean) => sec.quarantineRelease(id, ids, allow, report))
  handle('sec:quarantineDelete', (id: string, ids: string[]) => sec.quarantineDelete(id, ids))
  handle('sec:allowBlock', (id: string) => sec.allowBlock(id))
  handle('sec:newAllowBlock', (id: string, n: NewAllowBlock) => sec.newAllowBlock(id, n))
  handle('sec:removeAllowBlock', (id: string, lt: AllowBlockList, ids: string[]) => sec.removeAllowBlock(id, lt, ids))
  handle('sec:policies', (id: string) => sec.policies(id))
  handle('sec:setPolicy', (id: string, kind: SecPolicyKind, name: string, p: Record<string, unknown>) => sec.setPolicy(id, kind, name, p))
  handle('sec:setPolicyRule', (id: string, kind: SecPolicyKind, rule: string, enabled: boolean) => sec.setPolicyRule(id, kind, rule, enabled))
  handle('sec:restricted', (id: string) => sec.restricted(id))
  handle('sec:releaseSender', (id: string, s: RestrictedSender) => sec.releaseSender(id, s))
  handle('sec:trace', (id: string, f: TraceFilter) => sec.trace(id, f))
  handle('sec:traceDetail', (id: string, traceId: string, rcpt: string) => sec.traceDetail(id, traceId, rcpt))
  handle('sec:auditStatus', (id: string) => sec.auditStatus(id))
  handle('sec:enableAudit', (id: string) => sec.enableAudit(id))
  handle('sec:audit', (id: string, f: AuditFilter) => sec.audit(id, f))
  handle('sec:secureMailbox', (id: string, mbx: string, rules: boolean, fwd: boolean) => sec.secureMailbox(id, mbx, rules, fwd))
  handle('sec:signIns', (id: string, userId: string) => sec.signIns(id, userId))
  handle('sec:riskyUsers', (id: string) => sec.riskyUsers(id))
  handle('sec:riskDetections', (id: string, days: number) => sec.riskDetections(id, days))
  handle('sec:setRisk', (id: string, ids: string[], action: 'dismiss' | 'confirmCompromised') => sec.setRisk(id, ids, action))
  handle('sec:reported', (id: string, days: number) => sec.reported(id, days))
  handle('sec:retractJobs', (id: string) => sec.retractJobs(id))
  handle('sec:startRetract', (id: string, s: RetractSpec) => sec.startRetract(id, s))
  handle('sec:purgeRetract', (jobId: string, type: PurgeType) => sec.purgeRetract(jobId, type))
  handle('sec:discardRetract', (jobId: string) => sec.discardRetract(jobId))
  handle('sec:analyzed', (id: string, s: RetractSpec) => sec.analyzedEmails(id, s))
  handle('sec:remediate', (id: string, e: Pick<AnalyzedEmail, 'networkMessageId' | 'recipient'>[], a: DefenderAction, name: string) => sec.remediate(id, e, a, name))
  handle('update:state', () => updateState())
  handle('update:check', () => checkForUpdates())
  handle('update:download', () => downloadUpdate())
  handle('update:install', () => installUpdate())
}

// Development only: a separate profile, so a test run doesn't touch (or collide with) the installed app.
if (!app.isPackaged && process.env.TOOLBOX_USER_DATA) app.setPath('userData', process.env.TOOLBOX_USER_DATA)

if (!app.requestSingleInstanceLock()) app.quit()
else {
  app.on('second-instance', () => {
    if (win) {
      if (win.isMinimized()) win.restore()
      win.focus()
    }
  })
  app.setAppUserModelId('nl.rhis.m365toolbox')
  app.whenReady().then(() => {
    registerIpc()
    createWindow()
    // Check or install the Exchange module in the background so the first Exchange screen opens faster.
    exoPs.ensureModule().catch(() => undefined)
  })
  app.on('window-all-closed', () => app.quit())
  app.on('will-quit', () => (exoPs.closeAll(), teams.closeAll(), sec.closeAll()))
}
