import { X509Certificate, randomUUID } from 'crypto'
import { readFileSync } from 'fs'
import { basename } from 'path'
import { dialog } from 'electron'
import type {
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
  PermissionDef,
  PimApproval,
  PimAudit,
  PimGroupEntry,
  PimRoleSettings,
  ResourceAccess,
  RoleAssign,
  RoleDef,
  RoleMember,
  RolesData,
  UserMfa
} from '@shared/types'
import { tokenFor, tokenForHome, tokenForScopes } from './auth'
import { all, call, GraphError } from './graph'
import { tenants } from './store'
import { batch } from './teams'

/*
 * Entra ID runs entirely on Microsoft Graph. The write scopes for apps, Conditional Access, roles and PIM are
 * outside the base sign-in, so they are asked for once per tenant (consent banner on NEEDS_CONSENT), like Teams.
 */
export const ENTRA_SCOPES = [
  'Application.ReadWrite.All',
  'AppRoleAssignment.ReadWrite.All',
  'DelegatedPermissionGrant.ReadWrite.All',
  'Policy.Read.All',
  'Policy.ReadWrite.ConditionalAccess',
  'RoleManagement.ReadWrite.Directory',
  'RoleManagementPolicy.ReadWrite.Directory',
  'RoleEligibilitySchedule.ReadWrite.Directory',
  'RoleAssignmentSchedule.ReadWrite.Directory',
  'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup',
  'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup',
  'RoleManagementPolicy.ReadWrite.AzureADGroup',
  'User.Invite.All',
  'UserAuthenticationMethod.ReadWrite.All',
  'AuditLog.Read.All'
]

const GRAPH = 'https://graph.microsoft.com/v1.0'
const BETA = 'https://graph.microsoft.com/beta'
const token = (tid: string) => tokenForScopes(tid, ENTRA_SCOPES)
const json = (method: string, body?: unknown): RequestInit => ({ method, body: body === undefined ? undefined : JSON.stringify(body) })

export const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000'
/** Tenants Microsoft publishes its own apps from. */
const MICROSOFT_TENANTS = new Set(['f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47', '33e01921-4d64-4f8c-a055-5bdaffd5e33d'])

/** Role templates Microsoft counts as privileged; the same list its Conditional Access admin template uses. */
export const PRIVILEGED_ROLES = [
  '62e90394-69f5-4237-9190-012177145e10',
  '194ae4cb-b126-40b2-bd5b-6091b380977d',
  'f28a1f50-f6e7-4571-818b-6a12f2af6b6c',
  '29232cdf-9323-42fd-ade2-1d097af3e4de',
  'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9',
  '729827e3-9c14-49f7-bb1b-9608f156bbb8',
  'b0f54661-2d74-4c50-afa3-1ec803f12efe',
  'fe930be7-5e62-47db-91af-98c3a49a38b1',
  'c4e39bd9-1100-46d3-8c65-fb160da0071f',
  '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3',
  '158c047a-c907-4556-b7ef-446551a6b5f7',
  '966707d0-3269-4727-9be2-8c3a10f19b9d',
  '7be44c8a-adaf-4e2a-84d6-ab2649e08a13',
  'e8611ab8-c189-46e8-94e1-60213ab1f814'
]

type RawRef = { id: string; displayName?: string | null; userPrincipalName?: string | null; mail?: string | null; appId?: string | null; '@odata.type'?: string }
const ref = (p: RawRef | null | undefined, fallbackId = ''): DirectoryRef => ({
  id: p?.id || fallbackId,
  displayName: p?.displayName || p?.userPrincipalName || p?.mail || p?.id || fallbackId,
  detail: p?.userPrincipalName || p?.mail || p?.appId || null,
  type: (p?.['@odata.type'] || '').replace('#microsoft.graph.', '') || 'user'
})

// ---------- app registrations ----------

type RawPassword = { keyId: string; displayName: string | null; hint: string | null; startDateTime: string | null; endDateTime: string | null }
type RawKey = { keyId: string; displayName: string | null; customKeyIdentifier: string | null; startDateTime: string | null; endDateTime: string | null; type?: string; usage?: string; key?: string | null }
type RawApp = {
  id: string
  appId: string
  displayName: string
  signInAudience: string
  createdDateTime: string | null
  web?: { redirectUris: string[] }
  spa?: { redirectUris: string[] }
  publicClient?: { redirectUris: string[] }
  requiredResourceAccess: { resourceAppId: string; resourceAccess: { id: string; type: 'Scope' | 'Role' }[] }[]
  appRoles: AppRole[]
  passwordCredentials: RawPassword[]
  keyCredentials: RawKey[]
}

const APP_FIELDS = 'id,appId,displayName,signInAudience,createdDateTime,web,spa,publicClient,requiredResourceAccess,appRoles,passwordCredentials,keyCredentials'

const creds = (pw: RawPassword[] = [], keys: RawKey[] = []): AppCredential[] => [
  ...pw.map((p) => ({ keyId: p.keyId, kind: 'secret' as const, displayName: p.displayName, hint: p.hint, start: p.startDateTime, end: p.endDateTime })),
  // An app with a SAML signing certificate has both a Sign and a Verify key for one certificate; show it once.
  ...keys
    .filter((k) => k.usage !== 'Sign')
    .map((k) => ({ keyId: k.keyId, kind: 'cert' as const, displayName: k.displayName, hint: thumb(k.customKeyIdentifier), start: k.startDateTime, end: k.endDateTime }))
]

/** customKeyIdentifier is the thumbprint as base64; the portal shows it as hex. */
function thumb(v: string | null) {
  if (!v) return null
  try {
    const hex = Buffer.from(v, 'base64').toString('hex').toUpperCase()
    return hex.length === 40 ? hex : v
  } catch {
    return v
  }
}

const appReg = (a: RawApp): AppReg => ({
  id: a.id,
  appId: a.appId,
  displayName: a.displayName,
  audience: a.signInAudience,
  created: a.createdDateTime,
  web: a.web?.redirectUris || [],
  spa: a.spa?.redirectUris || [],
  publicClient: a.publicClient?.redirectUris || [],
  permissions: (a.requiredResourceAccess || []).map((r) => ({ resourceAppId: r.resourceAppId, access: r.resourceAccess })),
  appRoles: a.appRoles || [],
  credentials: creds(a.passwordCredentials, a.keyCredentials)
})

export async function listApps(tid: string): Promise<AppReg[]> {
  const tok = await token(tid)
  const raw = await all<RawApp>(tok, `/applications?$select=${APP_FIELDS}&$top=999`)
  return raw.map(appReg).sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function getApp(tid: string, id: string): Promise<AppReg> {
  const tok = await token(tid)
  return appReg(await call<RawApp>(tok, `/applications/${id}?$select=${APP_FIELDS}`))
}

const platformBody = (platform: NewAppReg['platform'], uris: string[]) =>
  platform === 'none' || !uris.length ? {} : { [platform]: { redirectUris: uris } }

const rraBody = (perms: ResourceAccess[]) =>
  perms.filter((p) => p.access.length).map((p) => ({ resourceAppId: p.resourceAppId, resourceAccess: p.access }))

export async function createApp(tid: string, n: NewAppReg): Promise<{ app: AppReg; consentError?: string }> {
  const tok = await token(tid)
  const a = await call<RawApp>(tok, '/applications', json('POST', {
    displayName: n.displayName,
    signInAudience: n.audience,
    ...platformBody(n.platform, n.redirectUris),
    requiredResourceAccess: rraBody(n.permissions)
  }))
  // The service principal makes the app usable in this tenant (sign-in, consent, enterprise apps list).
  await ensureSp(tok, a.appId)
  let consentError: string | undefined
  if (n.adminConsent && n.permissions.length) {
    try {
      await grantConsent(tid, a.id)
    } catch (e) {
      consentError = (e as Error).message
    }
  }
  return { app: appReg(a), consentError }
}

export async function updateApp(tid: string, id: string, p: { displayName?: string; web?: string[]; spa?: string[]; publicClient?: string[]; permissions?: ResourceAccess[] }) {
  const tok = await token(tid)
  const body: Record<string, unknown> = {}
  if (p.displayName !== undefined) body.displayName = p.displayName
  for (const k of ['web', 'spa', 'publicClient'] as const) if (p[k]) body[k] = { redirectUris: p[k] }
  if (p.permissions) body.requiredResourceAccess = rraBody(p.permissions)
  await call(tok, `/applications/${id}`, json('PATCH', body))
  return getApp(tid, id)
}

export async function deleteApp(tid: string, id: string) {
  await call(await token(tid), `/applications/${id}`, { method: 'DELETE' })
}

/** App roles can only be removed once disabled, so a removal is two updates. */
export async function setAppRoles(tid: string, id: string, roles: AppRole[]) {
  const tok = await token(tid)
  const cur = (await call<RawApp>(tok, `/applications/${id}?$select=appRoles`)).appRoles || []
  const keep = new Set(roles.map((r) => r.id))
  const removed = cur.filter((r) => !keep.has(r.id))
  if (removed.some((r) => r.isEnabled))
    await call(tok, `/applications/${id}`, json('PATCH', { appRoles: cur.map((r) => (keep.has(r.id) ? r : { ...r, isEnabled: false })) }))
  await call(tok, `/applications/${id}`, json('PATCH', { appRoles: roles.map((r) => ({ ...r, id: r.id || randomUUID() })) }))
  return getApp(tid, id)
}

export async function addSecret(tid: string, kind: 'app' | 'sp', objectId: string, displayName: string, months: number) {
  const tok = await token(tid)
  const end = new Date()
  end.setMonth(end.getMonth() + months)
  const path = kind === 'app' ? `/applications/${objectId}` : `/servicePrincipals/${objectId}`
  const r = await call<RawPassword & { secretText: string }>(tok, `${path}/addPassword`, json('POST', {
    passwordCredential: { displayName, endDateTime: end.toISOString() }
  }))
  return { secret: r.secretText, keyId: r.keyId, end: r.endDateTime }
}

export async function removeCredential(tid: string, kind: 'app' | 'sp', objectId: string, cred: Pick<AppCredential, 'keyId' | 'kind'>) {
  const tok = await token(tid)
  const path = kind === 'app' ? `/applications/${objectId}` : `/servicePrincipals/${objectId}`
  if (cred.kind === 'secret') return void (await call(tok, `${path}/removePassword`, json('POST', { keyId: cred.keyId })))
  // Certificates have no remove action: write back the key list without it. Keys are only returned on a single-object $select.
  const cur = (await call<{ keyCredentials: RawKey[] }>(tok, `${path}?$select=keyCredentials`)).keyCredentials
  const drop = cur.find((k) => k.keyId === cred.keyId)
  const left = cur.filter((k) => k.keyId !== cred.keyId && !(drop?.customKeyIdentifier && k.customKeyIdentifier === drop.customKeyIdentifier))
  await call(tok, path, json('PATCH', { keyCredentials: left.map(keyBody) }))
}

const keyBody = (k: RawKey) => ({ keyId: k.keyId, type: k.type, usage: k.usage, key: k.key, displayName: k.displayName, customKeyIdentifier: k.customKeyIdentifier })

/** Lets the user pick a .cer/.crt/.pem file and reads the public certificate from it. */
export async function pickCert(): Promise<CertFile | null> {
  const r = await dialog.showOpenDialog({ properties: ['openFile'], filters: [{ name: 'Certificate', extensions: ['cer', 'crt', 'pem'] }] })
  if (r.canceled || !r.filePaths[0]) return null
  const file = r.filePaths[0]
  let c: X509Certificate
  try {
    c = new X509Certificate(readFileSync(file))
  } catch {
    throw new Error('Not a valid public certificate (.cer, .crt or .pem)')
  }
  return {
    key: c.raw.toString('base64'),
    thumbprint: c.fingerprint.replace(/:/g, ''),
    subject: c.subject.replace(/\n/g, ', '),
    start: new Date(c.validFrom).toISOString(),
    end: new Date(c.validTo).toISOString(),
    fileName: basename(file)
  }
}

export async function addCert(tid: string, appObjectId: string, cert: CertFile, displayName: string) {
  const tok = await token(tid)
  const cur = (await call<{ keyCredentials: RawKey[] }>(tok, `/applications/${appObjectId}?$select=keyCredentials`)).keyCredentials
  await call(tok, `/applications/${appObjectId}`, json('PATCH', {
    keyCredentials: [...cur.map(keyBody), { type: 'AsymmetricX509Cert', usage: 'Verify', key: cert.key, displayName: displayName || cert.subject }]
  }))
  return getApp(tid, appObjectId)
}

// ---------- API permissions and admin consent ----------

type RawSpPerms = {
  id: string
  appId: string
  displayName: string
  appRoles: { id: string; value: string | null; displayName: string; description: string; allowedMemberTypes: string[]; isEnabled: boolean }[]
  oauth2PermissionScopes: { id: string; value: string; type: 'Admin' | 'User'; adminConsentDisplayName: string; adminConsentDescription: string; isEnabled: boolean }[]
}

const resourceCache = new Map<string, ApiResource>()

function toResource(sp: RawSpPerms): ApiResource {
  const perms: PermissionDef[] = [
    ...sp.oauth2PermissionScopes
      .filter((s) => s.isEnabled)
      .map((s) => ({ id: s.id, type: 'Scope' as const, value: s.value, name: s.adminConsentDisplayName || s.value, description: s.adminConsentDescription || '', admin: s.type === 'Admin' })),
    ...sp.appRoles
      .filter((r) => r.isEnabled && r.allowedMemberTypes.includes('Application') && r.value)
      .map((r) => ({ id: r.id, type: 'Role' as const, value: r.value!, name: r.displayName, description: r.description, admin: true }))
  ]
  return { appId: sp.appId, spId: sp.id, displayName: sp.displayName, permissions: perms.sort((a, b) => a.value.localeCompare(b.value)) }
}

/** An API (Microsoft Graph, Exchange, SharePoint or a custom one) with its delegated and application permissions. */
export async function apiResource(tid: string, appId: string): Promise<ApiResource> {
  const k = `${tid}/${appId}`
  const hit = resourceCache.get(k)
  if (hit) return hit
  const tok = await token(tid)
  const sp = await call<RawSpPerms>(tok, `/servicePrincipals(appId='${appId}')?$select=id,appId,displayName,appRoles,oauth2PermissionScopes`)
  const r = toResource(sp)
  resourceCache.set(k, r)
  return r
}

async function spIdOf(tok: string, appId: string): Promise<string | null> {
  const r = await call<{ value: { id: string }[] }>(tok, `/servicePrincipals?$filter=appId eq '${appId}'&$select=id`)
  return r.value[0]?.id || null
}

async function ensureSp(tok: string, appId: string): Promise<string> {
  const id = await spIdOf(tok, appId)
  if (id) return id
  // A new application can take a few seconds to replicate before its service principal can be created.
  for (let i = 0; ; i++) {
    try {
      return (await call<{ id: string }>(tok, '/servicePrincipals', json('POST', { appId }))).id
    } catch (e) {
      if (i >= 4) throw e
      await new Promise((r) => setTimeout(r, 2000 * (i + 1)))
    }
  }
}

type RawGrant = { id: string; clientId: string; consentType: 'AllPrincipals' | 'Principal'; principalId: string | null; resourceId: string; scope: string }
type RawAssignment = { id: string; appRoleId: string; resourceId: string; resourceDisplayName?: string; principalId: string; principalDisplayName?: string; principalType?: string; createdDateTime?: string | null }

export async function appPermissions(tid: string, appObjectId: string): Promise<AppPermissionState> {
  const tok = await token(tid)
  const app = await getApp(tid, appObjectId)
  const resources = await Promise.all(app.permissions.map((p) => apiResource(tid, p.resourceAppId).catch(() => null)))
  const spId = await spIdOf(tok, app.appId)
  const granted: AppPermissionState['granted'] = {}
  if (spId) {
    const [grants, roles] = await Promise.all([
      all<RawGrant>(tok, `/oauth2PermissionGrants?$filter=clientId eq '${spId}'`),
      all<RawAssignment>(tok, `/servicePrincipals/${spId}/appRoleAssignments`)
    ])
    for (const r of resources.filter(Boolean) as ApiResource[]) {
      granted[r.appId] = {
        scopes: grants.filter((g) => g.resourceId === r.spId && g.consentType === 'AllPrincipals').flatMap((g) => g.scope.split(' ').filter(Boolean)),
        roles: roles.filter((x) => x.resourceId === r.spId).map((x) => x.appRoleId)
      }
    }
  }
  return { spId, resources: resources.filter(Boolean) as ApiResource[], granted }
}

/** "Grant admin consent": delegated scopes for all users and every application permission the app requests. */
export async function grantConsent(tid: string, appObjectId: string) {
  const tok = await token(tid)
  const app = await getApp(tid, appObjectId)
  const spId = await ensureSp(tok, app.appId)
  const grants = await all<RawGrant>(tok, `/oauth2PermissionGrants?$filter=clientId eq '${spId}'`)
  const assigned = await all<RawAssignment>(tok, `/servicePrincipals/${spId}/appRoleAssignments`)
  for (const p of app.permissions) {
    const res = await apiResource(tid, p.resourceAppId)
    const scopes = p.access.filter((a) => a.type === 'Scope').map((a) => res.permissions.find((x) => x.id === a.id)?.value).filter(Boolean) as string[]
    if (scopes.length) {
      const g = grants.find((x) => x.resourceId === res.spId && x.consentType === 'AllPrincipals')
      if (g) {
        const merged = [...new Set([...g.scope.split(' ').filter(Boolean), ...scopes])].join(' ')
        if (merged !== g.scope.trim()) await call(tok, `/oauth2PermissionGrants/${g.id}`, json('PATCH', { scope: merged }))
      } else await call(tok, '/oauth2PermissionGrants', json('POST', { clientId: spId, consentType: 'AllPrincipals', resourceId: res.spId, scope: scopes.join(' ') }))
    }
    for (const a of p.access.filter((x) => x.type === 'Role')) {
      if (assigned.some((x) => x.resourceId === res.spId && x.appRoleId === a.id)) continue
      await call(tok, `/servicePrincipals/${res.spId}/appRoleAssignedTo`, json('POST', { principalId: spId, resourceId: res.spId, appRoleId: a.id }))
    }
  }
  return appPermissions(tid, appObjectId)
}

/** Removes one permission from the app and withdraws the consent for it. */
export async function removePermission(tid: string, appObjectId: string, resourceAppId: string, permId: string) {
  const tok = await token(tid)
  const app = await getApp(tid, appObjectId)
  const perms = app.permissions.map((p) => (p.resourceAppId === resourceAppId ? { ...p, access: p.access.filter((a) => a.id !== permId) } : p))
  await call(tok, `/applications/${appObjectId}`, json('PATCH', { requiredResourceAccess: rraBody(perms) }))
  const spId = await spIdOf(tok, app.appId)
  if (spId) {
    const res = await apiResource(tid, resourceAppId)
    const def = res.permissions.find((x) => x.id === permId)
    if (def?.type === 'Scope') {
      for (const g of await all<RawGrant>(tok, `/oauth2PermissionGrants?$filter=clientId eq '${spId}'`)) {
        if (g.resourceId !== res.spId || !g.scope.split(' ').includes(def.value)) continue
        const left = g.scope.split(' ').filter((s) => s && s !== def.value)
        await call(tok, `/oauth2PermissionGrants/${g.id}`, left.length ? json('PATCH', { scope: left.join(' ') }) : { method: 'DELETE' })
      }
    } else if (def) {
      const a = (await all<RawAssignment>(tok, `/servicePrincipals/${spId}/appRoleAssignments`)).find((x) => x.resourceId === res.spId && x.appRoleId === permId)
      if (a) await call(tok, `/servicePrincipals/${spId}/appRoleAssignments/${a.id}`, { method: 'DELETE' })
    }
  }
  return appPermissions(tid, appObjectId)
}

// ---------- owners (applications and service principals) ----------

type OwnerKind = 'applications' | 'servicePrincipals'

export async function owners(tid: string, kind: OwnerKind, id: string): Promise<DirectoryRef[]> {
  const tok = await token(tid)
  return (await all<RawRef>(tok, `/${kind}/${id}/owners?$select=id,displayName,userPrincipalName,mail,appId`)).map((o) => ref(o))
}

export async function addOwner(tid: string, kind: OwnerKind, id: string, userId: string) {
  const tok = await token(tid)
  await call(tok, `/${kind}/${id}/owners/$ref`, json('POST', { '@odata.id': `${GRAPH}/directoryObjects/${userId}` }))
  return owners(tid, kind, id)
}

export async function removeOwner(tid: string, kind: OwnerKind, id: string, userId: string) {
  const tok = await token(tid)
  await call(tok, `/${kind}/${id}/owners/${userId}/$ref`, { method: 'DELETE' })
  return owners(tid, kind, id)
}

// ---------- expiring secrets and certificates ----------

/** Every secret and certificate of the tenant's app registrations and service principals (SAML signing certificates). */
export async function credentials(tid: string): Promise<ExpiringCred[]> {
  const tok = await token(tid)
  const [apps, sps] = await Promise.all([
    all<RawApp>(tok, '/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials&$top=999'),
    all<{ id: string; appId: string; displayName: string; appOwnerOrganizationId: string | null; passwordCredentials: RawPassword[]; keyCredentials: RawKey[] }>(
      tok,
      '/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId,passwordCredentials,keyCredentials&$top=999'
    )
  ])
  const out: ExpiringCred[] = []
  for (const a of apps)
    for (const c of creds(a.passwordCredentials, a.keyCredentials)) out.push({ ...c, objectId: a.id, appId: a.appId, appName: a.displayName, source: 'app' })
  for (const s of sps) {
    if (s.appOwnerOrganizationId && MICROSOFT_TENANTS.has(s.appOwnerOrganizationId)) continue
    for (const c of creds(s.passwordCredentials, s.keyCredentials)) out.push({ ...c, objectId: s.id, appId: s.appId, appName: s.displayName, source: 'sp' })
  }
  return out
}

// ---------- enterprise apps ----------

type RawSp = {
  id: string
  appId: string
  displayName: string
  accountEnabled: boolean
  appRoleAssignmentRequired: boolean
  tags: string[]
  publisherName: string | null
  appOwnerOrganizationId: string | null
  servicePrincipalType: string
  appRoles: AppRole[]
  homepage: string | null
  verifiedPublisher?: { displayName: string | null }
}

const SP_FIELDS = 'id,appId,displayName,accountEnabled,appRoleAssignmentRequired,tags,publisherName,appOwnerOrganizationId,servicePrincipalType,appRoles,homepage,verifiedPublisher'

const enterprise = (tid: string, s: RawSp): EnterpriseApp => ({
  id: s.id,
  appId: s.appId,
  displayName: s.displayName,
  enabled: s.accountEnabled,
  assignmentRequired: s.appRoleAssignmentRequired,
  hidden: s.tags.includes('HideApp'),
  publisher: s.verifiedPublisher?.displayName || s.publisherName,
  microsoft: !!s.appOwnerOrganizationId && MICROSOFT_TENANTS.has(s.appOwnerOrganizationId),
  ownTenant: s.appOwnerOrganizationId === tid,
  type: s.servicePrincipalType,
  appRoles: (s.appRoles || []).filter((r) => r.isEnabled),
  homepage: s.homepage
})

/** The Entra portal's "Enterprise applications" list: service principals tagged as integrated apps. */
export async function enterpriseApps(tid: string): Promise<EnterpriseApp[]> {
  const tok = await token(tid)
  const raw = await all<RawSp>(
    tok,
    `/servicePrincipals?$filter=tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')&$select=${SP_FIELDS}&$top=999&$count=true`,
    { ConsistencyLevel: 'eventual' }
  )
  return raw.map((s) => enterprise(tid, s)).sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function updateEnterpriseApp(tid: string, id: string, p: { enabled?: boolean; assignmentRequired?: boolean; hidden?: boolean }) {
  const tok = await token(tid)
  const body: Record<string, unknown> = {}
  if (p.enabled !== undefined) body.accountEnabled = p.enabled
  if (p.assignmentRequired !== undefined) body.appRoleAssignmentRequired = p.assignmentRequired
  if (p.hidden !== undefined) {
    const tags = (await call<{ tags: string[] }>(tok, `/servicePrincipals/${id}?$select=tags`)).tags.filter((x) => x !== 'HideApp')
    body.tags = p.hidden ? [...tags, 'HideApp'] : tags
  }
  await call(tok, `/servicePrincipals/${id}`, json('PATCH', body))
  return enterprise(tid, await call<RawSp>(tok, `/servicePrincipals/${id}?$select=${SP_FIELDS}`))
}

export async function assignments(tid: string, spId: string): Promise<AppAssignment[]> {
  const tok = await token(tid)
  return (await all<RawAssignment>(tok, `/servicePrincipals/${spId}/appRoleAssignedTo`))
    .map((a) => ({
      id: a.id,
      principalId: a.principalId,
      principalName: a.principalDisplayName || a.principalId,
      principalType: a.principalType || 'User',
      appRoleId: a.appRoleId,
      created: a.createdDateTime || null
    }))
    .sort((a, b) => a.principalName.localeCompare(b.principalName))
}

/** appRoleId 00000000-… is the "Default Access" role for apps that define no roles. */
export async function assign(tid: string, spId: string, principalId: string, appRoleId: string) {
  const tok = await token(tid)
  await call(tok, `/servicePrincipals/${spId}/appRoleAssignedTo`, json('POST', { principalId, resourceId: spId, appRoleId }))
  return assignments(tid, spId)
}

export async function unassign(tid: string, spId: string, assignmentId: string) {
  const tok = await token(tid)
  await call(tok, `/servicePrincipals/${spId}/appRoleAssignedTo/${assignmentId}`, { method: 'DELETE' })
  return assignments(tid, spId)
}

// ---------- granted permissions and risky apps ----------

/** Permissions that let an app read or change mail, files or the directory of everyone. */
const RISKY = new Set([
  'Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'Mail.ReadBasic.All', 'MailboxSettings.ReadWrite', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared',
  'Calendars.ReadWrite', 'Contacts.ReadWrite', 'Files.Read.All', 'Files.ReadWrite.All', 'Sites.Read.All', 'Sites.ReadWrite.All', 'Sites.FullControl.All', 'Sites.Manage.All',
  'Notes.ReadWrite.All', 'Directory.ReadWrite.All', 'Directory.AccessAsUser.All', 'User.ReadWrite.All', 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All',
  'RoleManagement.ReadWrite.Directory', 'AppRoleAssignment.ReadWrite.All', 'Application.ReadWrite.All', 'Policy.ReadWrite.ConditionalAccess',
  'UserAuthenticationMethod.ReadWrite.All', 'Domain.ReadWrite.All', 'Chat.ReadWrite.All', 'ChannelMessage.Read.All', 'Chat.Read.All',
  'full_access_as_app', 'full_access_as_user', 'EWS.AccessAsUser.All', 'Exchange.ManageAsApp', 'Mail.ReadWrite.All', 'MailboxItem.Read.All',
  'AllSites.FullControl', 'AllSites.Manage', 'AllSites.Write', 'MyFiles.Write', 'Sites.Selected', 'offline_access_all'
])
const isRisky = (v: string) => RISKY.has(v) && v !== 'Sites.Selected' && v !== 'offline_access_all'

/** Every app that has been granted permissions, delegated (admin or user consent) or application. */
export async function grants(tid: string): Promise<AppGrants[]> {
  const tok = await token(tid)
  const [sps, delegated] = await Promise.all([
    all<RawSp>(tok, `/servicePrincipals?$select=${SP_FIELDS}&$top=999`),
    all<RawGrant>(tok, '/oauth2PermissionGrants')
  ])
  const byId = new Map(sps.map((s) => [s.id, s]))
  // Application permissions: assignments on the APIs that apps call, read per API that has any delegated grant or is a known Microsoft API.
  const apiIds = new Set<string>(delegated.map((g) => g.resourceId))
  for (const s of sps) if (['00000003-0000-0000-c000-000000000000', '00000002-0000-0ff1-ce00-000000000000', '00000003-0000-0ff1-ce00-000000000000'].includes(s.appId)) apiIds.add(s.id)
  const res = await batch(tok, [...apiIds].map((id) => ({ id, url: `/servicePrincipals/${id}/appRoleAssignedTo` })))
  const out = new Map<string, AppGrants>()
  const entry = (spId: string) => {
    let e = out.get(spId)
    if (!e) {
      const s = byId.get(spId)
      e = {
        spId,
        appId: s?.appId || '',
        displayName: s?.displayName || spId,
        publisher: s?.verifiedPublisher?.displayName || s?.publisherName || null,
        microsoft: !!s?.appOwnerOrganizationId && MICROSOFT_TENANTS.has(s.appOwnerOrganizationId),
        enabled: s?.accountEnabled ?? true,
        userConsents: 0,
        permissions: []
      }
      out.set(spId, e)
    }
    return e
  }
  for (const g of delegated) {
    const api = byId.get(g.resourceId)
    const e = entry(g.clientId)
    if (g.consentType === 'Principal') e.userConsents++
    for (const v of g.scope.split(' ').filter(Boolean)) {
      if (e.permissions.some((p) => p.kind === 'delegated' && p.value === v && p.resource === (api?.displayName || g.resourceId) && p.consent === (g.consentType === 'AllPrincipals' ? 'admin' : 'user'))) continue
      e.permissions.push({ grantId: g.id, kind: 'delegated', resource: api?.displayName || g.resourceId, value: v, consent: g.consentType === 'AllPrincipals' ? 'admin' : 'user', risky: isRisky(v) })
    }
  }
  for (const apiId of apiIds) {
    const r = res.get(apiId)
    if (r?.status !== 200) continue
    const api = byId.get(apiId)
    const next = r.body['@odata.nextLink'] as string | undefined
    const list = [...((r.body.value || []) as RawAssignment[]), ...(next ? await all<RawAssignment>(tok, next) : [])]
    for (const a of list) {
      if (a.principalType !== 'ServicePrincipal') continue
      const v = api?.appRoles.find((x) => x.id === a.appRoleId)?.value || a.appRoleId
      entry(a.principalId).permissions.push({ grantId: a.id, kind: 'application', resource: api?.displayName || apiId, value: v, consent: null, risky: isRisky(v) || (/\.All$/.test(v) && /Write/.test(v)) })
    }
  }
  return [...out.values()].sort((a, b) => a.displayName.localeCompare(b.displayName))
}

/** Withdraws one permission: a scope from a delegated grant (all grants of that app with it) or an application permission. */
export async function revokeGrant(tid: string, spId: string, p: Pick<GrantedPermission, 'grantId' | 'kind' | 'value'>) {
  const tok = await token(tid)
  if (p.kind === 'application') {
    await call(tok, `/servicePrincipals/${spId}/appRoleAssignments/${p.grantId}`, { method: 'DELETE' })
    return
  }
  const g = await call<RawGrant>(tok, `/oauth2PermissionGrants/${p.grantId}`)
  const left = g.scope.split(' ').filter((s) => s && s !== p.value)
  await call(tok, `/oauth2PermissionGrants/${g.id}`, left.length ? json('PATCH', { scope: left.join(' ') }) : { method: 'DELETE' })
}

// ---------- Conditional Access ----------

type RawCa = { id: string; displayName: string; state: CaState; createdDateTime: string | null; modifiedDateTime: string | null; [k: string]: unknown }

export async function caPolicies(tid: string): Promise<CaPolicy[]> {
  const tok = await token(tid)
  return (await all<RawCa>(tok, '/identity/conditionalAccess/policies'))
    .map((p) => ({ id: p.id, displayName: p.displayName, state: p.state, created: p.createdDateTime, modified: p.modifiedDateTime, raw: p }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

const OFFICE_365 = 'Office365'
const AZURE_MANAGEMENT = '797f4846-ba00-4fd7-ba43-dac1f8f63013'

/** Turns the toolbox's simple policy form into a Graph conditionalAccessPolicy. */
export function caBody(d: CaDraft) {
  const users: Record<string, unknown> = { includeUsers: [], excludeUsers: [], includeGroups: [], excludeGroups: [], includeRoles: [] }
  if (d.users === 'all') users.includeUsers = ['All']
  else if (d.users === 'guests') users.includeUsers = ['GuestsOrExternalUsers']
  else if (d.users === 'admins') users.includeRoles = PRIVILEGED_ROLES
  else {
    users.includeUsers = d.include.filter((x) => x.startsWith('u:')).map((x) => x.slice(2))
    users.includeGroups = d.include.filter((x) => x.startsWith('g:')).map((x) => x.slice(2))
  }
  users.excludeUsers = d.exclude.filter((x) => x.startsWith('u:')).map((x) => x.slice(2))
  users.excludeGroups = d.exclude.filter((x) => x.startsWith('g:')).map((x) => x.slice(2))
  const conditions: Record<string, unknown> = {
    users,
    applications: { includeApplications: [d.apps === 'all' ? 'All' : d.apps === 'office365' ? OFFICE_365 : AZURE_MANAGEMENT] },
    clientAppTypes: d.clientApps === 'legacy' ? ['exchangeActiveSync', 'other'] : ['all']
  }
  if (d.locations === 'excludeTrusted') conditions.locations = { includeLocations: ['All'], excludeLocations: ['AllTrusted'] }
  else if (d.locations === 'selected' && d.locationIds.length) conditions.locations = { includeLocations: d.locationIds }
  const grantControls =
    d.grant === 'block'
      ? { operator: 'OR', builtInControls: ['block'] }
      : d.grant === 'mfa'
        ? { operator: 'OR', builtInControls: ['mfa'] }
        : d.grant === 'compliant'
          ? { operator: 'OR', builtInControls: ['compliantDevice'] }
          : { operator: 'OR', builtInControls: ['mfa', 'compliantDevice'] }
  return { displayName: d.displayName, state: d.state, conditions, grantControls }
}

export async function createCa(tid: string, d: CaDraft) {
  const tok = await token(tid)
  await call(tok, '/identity/conditionalAccess/policies', json('POST', caBody(d)))
  return caPolicies(tid)
}

export async function setCaState(tid: string, id: string, state: CaState) {
  await call(await token(tid), `/identity/conditionalAccess/policies/${id}`, json('PATCH', { state }))
  return caPolicies(tid)
}

export async function deleteCa(tid: string, id: string) {
  await call(await token(tid), `/identity/conditionalAccess/policies/${id}`, { method: 'DELETE' })
  return caPolicies(tid)
}

/** Names for the ids in a policy (users, groups, roles, apps, locations), for the details view. */
export async function resolveIds(tid: string, ids: string[]): Promise<Record<string, string>> {
  const tok = await token(tid)
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
  const list = [...new Set(ids.filter((x) => guid.test(x)))]
  const out: Record<string, string> = {}
  for (let i = 0; i < list.length; i += 1000) {
    const r = await call<{ value: RawRef[] }>(tok, '/directoryObjects/getByIds', json('POST', { ids: list.slice(i, i + 1000) }))
    for (const o of r.value) out[o.id] = o.displayName || o.userPrincipalName || o.id
  }
  const missing = list.filter((x) => !out[x])
  if (missing.length) {
    const [defs, locs] = await Promise.all([
      all<{ id: string; templateId: string; displayName: string }>(tok, '/roleManagement/directory/roleDefinitions?$select=id,templateId,displayName').catch(() => []),
      all<{ id: string; displayName: string }>(tok, '/identity/conditionalAccess/namedLocations').catch(() => [])
    ])
    for (const d of defs) for (const k of [d.id, d.templateId]) if (missing.includes(k)) out[k] = d.displayName
    for (const l of locs) if (missing.includes(l.id)) out[l.id] = l.displayName
    const apps = missing.filter((x) => !out[x])
    if (apps.length) {
      const res = await batch(tok, apps.map((a) => ({ id: a, url: `/servicePrincipals(appId='${a}')?$select=displayName` })))
      for (const a of apps) {
        const r = res.get(a)
        if (r?.status === 200) out[a] = String(r.body.displayName)
      }
    }
  }
  return out
}

// ---------- named locations ----------

type RawLocation = {
  id: string
  displayName: string
  '@odata.type': string
  isTrusted?: boolean
  ipRanges?: { cidrAddress: string }[]
  countriesAndRegions?: string[]
  includeUnknownCountriesAndRegions?: boolean
  modifiedDateTime: string | null
}

export async function namedLocations(tid: string): Promise<NamedLocation[]> {
  const tok = await token(tid)
  return (await all<RawLocation>(tok, '/identity/conditionalAccess/namedLocations'))
    .map((l) => ({
      id: l.id,
      displayName: l.displayName,
      kind: l['@odata.type'].includes('country') ? ('country' as const) : ('ip' as const),
      trusted: !!l.isTrusted,
      ranges: (l.ipRanges || []).map((r) => r.cidrAddress),
      countries: l.countriesAndRegions || [],
      includeUnknown: !!l.includeUnknownCountriesAndRegions,
      modified: l.modifiedDateTime
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

function locationBody(l: Omit<NamedLocation, 'id' | 'modified'>) {
  if (l.kind === 'country')
    return { '@odata.type': '#microsoft.graph.countryNamedLocation', displayName: l.displayName, countriesAndRegions: l.countries, includeUnknownCountriesAndRegions: l.includeUnknown }
  return {
    '@odata.type': '#microsoft.graph.ipNamedLocation',
    displayName: l.displayName,
    isTrusted: l.trusted,
    ipRanges: l.ranges.map((c) => ({ '@odata.type': c.includes(':') ? '#microsoft.graph.iPv6CidrRange' : '#microsoft.graph.iPv4CidrRange', cidrAddress: c }))
  }
}

export async function saveLocation(tid: string, l: Omit<NamedLocation, 'modified'> & { id?: string }) {
  const tok = await token(tid)
  if (l.id) await call(tok, `/identity/conditionalAccess/namedLocations/${l.id}`, json('PATCH', locationBody(l)))
  else await call(tok, '/identity/conditionalAccess/namedLocations', json('POST', locationBody(l)))
  return namedLocations(tid)
}

export async function deleteLocation(tid: string, id: string) {
  await call(await token(tid), `/identity/conditionalAccess/namedLocations/${id}`, { method: 'DELETE' })
  return namedLocations(tid)
}

// ---------- admin roles ----------

const ROLE = '/roleManagement/directory'

export async function roles(tid: string): Promise<RolesData> {
  const tok = await token(tid)
  const [defs, perm] = await Promise.all([
    all<{ id: string; templateId: string; displayName: string; description: string; isBuiltIn: boolean }>(tok, `${ROLE}/roleDefinitions?$select=id,templateId,displayName,description,isBuiltIn`),
    all<{ id: string; roleDefinitionId: string; principalId: string; directoryScopeId: string; principal?: RawRef }>(tok, `${ROLE}/roleAssignments?$expand=principal&$top=999`)
  ])
  const members: RoleMember[] = perm
    .filter((a) => a.directoryScopeId === '/')
    .map((a) => ({ id: a.id, roleId: a.roleDefinitionId, principal: ref(a.principal, a.principalId), kind: 'permanent', start: null, end: null }))
  let pim = true
  try {
    type Inst = { id: string; roleDefinitionId: string; principalId: string; directoryScopeId: string; assignmentType?: string; startDateTime: string | null; endDateTime: string | null; principal?: RawRef }
    const [elig, act] = await Promise.all([
      all<Inst>(tok, `${ROLE}/roleEligibilityScheduleInstances?$expand=principal`),
      all<Inst>(tok, `${ROLE}/roleAssignmentScheduleInstances?$expand=principal`)
    ])
    // Time-bound and activated assignments also appear in roleAssignments; mark them with their end date.
    for (const a of act.filter((x) => x.directoryScopeId === '/' && (x.endDateTime || x.assignmentType === 'Activated'))) {
      const m = members.find((x) => x.roleId === a.roleDefinitionId && x.principal.id === a.principalId)
      if (m) Object.assign(m, { kind: 'active', start: a.startDateTime, end: a.endDateTime })
    }
    for (const e of elig.filter((x) => x.directoryScopeId === '/'))
      members.push({ id: e.id, roleId: e.roleDefinitionId, principal: ref(e.principal, e.principalId), kind: 'eligible', start: e.startDateTime, end: e.endDateTime })
  } catch (e) {
    if (!(e instanceof GraphError)) throw e
    pim = false
  }
  const privileged = new Set(PRIVILEGED_ROLES)
  return {
    roles: defs
      .map((d) => ({ id: d.id, templateId: d.templateId, displayName: d.displayName, description: d.description, builtIn: d.isBuiltIn, privileged: privileged.has(d.templateId) }))
      .sort((a, b) => a.displayName.localeCompare(b.displayName)),
    members,
    pim
  }
}

const schedule = (start?: string, end?: string | null) => ({
  startDateTime: start || new Date().toISOString(),
  expiration: end ? { type: 'afterDateTime', endDateTime: end } : { type: 'noExpiration' }
})

export async function assignRole(tid: string, a: RoleAssign) {
  const tok = await token(tid)
  const base = { principalId: a.principalId, roleDefinitionId: a.roleId, directoryScopeId: '/' }
  if (a.kind === 'permanent' && !a.end && !a.start) await call(tok, `${ROLE}/roleAssignments`, json('POST', base))
  else
    await call(tok, `${ROLE}/${a.kind === 'eligible' ? 'roleEligibilityScheduleRequests' : 'roleAssignmentScheduleRequests'}`, json('POST', {
      ...base,
      action: 'adminAssign',
      justification: a.justification || 'Assigned with M365 Admin Toolbox',
      scheduleInfo: schedule(a.start, a.end)
    }))
  return roles(tid)
}

export async function removeRole(tid: string, m: Pick<RoleMember, 'id' | 'kind' | 'roleId'> & { principalId: string }) {
  const tok = await token(tid)
  const req = { principalId: m.principalId, roleDefinitionId: m.roleId, directoryScopeId: '/', action: 'adminRemove', justification: 'Removed with M365 Admin Toolbox' }
  if (m.kind === 'eligible') await call(tok, `${ROLE}/roleEligibilityScheduleRequests`, json('POST', req))
  else if (m.kind === 'active') await call(tok, `${ROLE}/roleAssignmentScheduleRequests`, json('POST', req))
  else
    try {
      await call(tok, `${ROLE}/roleAssignments/${m.id}`, { method: 'DELETE' })
    } catch (e) {
      // Assignments made through PIM can only be removed with a schedule request.
      if (!(e instanceof GraphError) || e.status === 403) throw e
      await call(tok, `${ROLE}/roleAssignmentScheduleRequests`, json('POST', req))
    }
  return roles(tid)
}

/** GDAP relationships with this customer, read in the partner tenant. */
export async function gdap(tid: string): Promise<GdapStatus> {
  const t = tenants.get(tid)!
  if (t.mode !== 'gdap') return { gdap: false, relationships: [] }
  try {
    const tok = await tokenForHome(tid)
    type Rel = { id: string; displayName: string; status: string; activatedDateTime: string | null; endDateTime: string | null; autoExtendDuration: string | null; accessDetails: { unifiedRoles: { roleDefinitionId: string }[] } }
    const rels = await all<Rel>(tok, `/tenantRelationships/delegatedAdminRelationships?$filter=customer/tenantId eq '${tid}'`)
    const out = []
    for (const r of rels) {
      const aa = await all<{ status: string; accessContainer: { accessContainerId: string }; accessDetails: { unifiedRoles: { roleDefinitionId: string }[] } }>(
        tok,
        `/tenantRelationships/delegatedAdminRelationships/${r.id}/accessAssignments`
      ).catch(() => [])
      const names = new Map<string, string>()
      const ids = aa.map((a) => a.accessContainer.accessContainerId)
      if (ids.length) {
        const g = await call<{ value: RawRef[] }>(tok, '/directoryObjects/getByIds', json('POST', { ids, types: ['group'] })).catch(() => ({ value: [] as RawRef[] }))
        g.value.forEach((x) => names.set(x.id, x.displayName || x.id))
      }
      out.push({
        id: r.id,
        displayName: r.displayName,
        status: r.status,
        start: r.activatedDateTime,
        end: r.endDateTime,
        autoExtend: r.autoExtendDuration,
        roles: r.accessDetails.unifiedRoles.map((x) => x.roleDefinitionId),
        groups: aa
          .filter((a) => a.status === 'active')
          .map((a) => ({ id: a.accessContainer.accessContainerId, name: names.get(a.accessContainer.accessContainerId) || a.accessContainer.accessContainerId, roles: a.accessDetails.unifiedRoles.map((x) => x.roleDefinitionId) }))
      })
    }
    return { gdap: true, relationships: out }
  } catch (e) {
    const err = e as Error & { code?: string }
    if (err.code === 'NEEDS_HOME_SIGNIN') throw e
    return { gdap: true, relationships: [], error: err.message }
  }
}

// ---------- PIM management ----------

const GROUP = '/identityGovernance/privilegedAccess/group'

type RawRequest = {
  id: string
  approvalId: string | null
  status: string
  createdDateTime: string
  justification: string | null
  principalId: string
  principal?: RawRef
  roleDefinitionId?: string
  roleDefinition?: { displayName: string }
  groupId?: string
  group?: { displayName: string }
  accessId?: string
  scheduleInfo?: { startDateTime: string | null; expiration?: { endDateTime?: string | null; duration?: string | null } }
}

function endOf(s: RawRequest['scheduleInfo']) {
  if (!s) return null
  if (s.expiration?.endDateTime) return s.expiration.endDateTime
  const m = s.expiration?.duration?.match(/PT(\d+)H/)
  return m && s.startDateTime ? new Date(new Date(s.startDateTime).getTime() + +m[1] * 3600_000).toISOString() : null
}

/** Activation requests waiting for the signed-in user's approval: directory roles here, PIM for groups here too. */
export async function approvals(tid: string): Promise<PimApproval[]> {
  const tok = await token(tid)
  const pending = "filterByCurrentUser(on='approver')?$filter=status eq 'PendingApproval'"
  const [r, g] = await Promise.all([
    all<RawRequest>(tok, `${ROLE}/roleAssignmentScheduleRequests/${pending}&$expand=principal,roleDefinition($select=displayName)`).catch(() => [] as RawRequest[]),
    all<RawRequest>(tok, `${GROUP}/assignmentScheduleRequests/${pending}&$expand=principal,group($select=displayName)`).catch(() => [] as RawRequest[])
  ])
  const map = (x: RawRequest, kind: 'role' | 'group'): PimApproval => ({
    approvalId: x.approvalId || '',
    requestId: x.id,
    kind,
    targetId: (kind === 'role' ? x.roleDefinitionId : x.groupId) || '',
    targetName: (kind === 'role' ? x.roleDefinition?.displayName : x.group?.displayName) || '',
    principal: ref(x.principal, x.principalId),
    justification: x.justification,
    created: x.createdDateTime,
    accessId: x.accessId,
    start: x.scheduleInfo?.startDateTime || null,
    end: endOf(x.scheduleInfo)
  })
  return [...r.map((x) => map(x, 'role')), ...g.map((x) => map(x, 'group'))].filter((x) => x.approvalId).sort((a, b) => b.created.localeCompare(a.created))
}

export async function review(tid: string, a: Pick<PimApproval, 'approvalId' | 'kind'>, approve: boolean, justification: string) {
  const tok = await token(tid)
  // Directory role approvals only exist in beta; PIM for groups approvals are in v1.0.
  const base = a.kind === 'role' ? `${BETA}${ROLE}/roleAssignmentApprovals/${a.approvalId}` : `${GRAPH}${GROUP}/assignmentApprovals/${a.approvalId}`
  const steps = await call<{ value: { id: string; status: string; assignedToMe?: boolean }[] }>(tok, `${base}/steps`)
  const step = steps.value.find((s) => s.status === 'InProgress' && s.assignedToMe !== false) || steps.value.find((s) => s.status === 'InProgress')
  if (!step) throw new Error('No open approval step for you')
  await call(tok, `${base}/steps/${step.id}`, json('PATCH', { reviewResult: approve ? 'Approve' : 'Deny', justification }))
}

type RawRule = { id: string; '@odata.type': string; [k: string]: unknown }

async function policyFor(tok: string, scope: { kind: 'role'; roleId: string } | { kind: 'group'; groupId: string; accessId: string }) {
  const filter =
    scope.kind === 'role'
      ? `scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '${scope.roleId}'`
      : `scopeId eq '${scope.groupId}' and scopeType eq 'Group' and roleDefinitionId eq '${scope.accessId}'`
  const r = await call<{ value: { policyId: string; policy?: { rules: RawRule[] } }[] }>(tok, `/policies/roleManagementPolicyAssignments?$filter=${filter}&$expand=policy($expand=rules)`)
  const a = r.value[0]
  if (!a) throw new Error('No PIM policy found')
  return { policyId: a.policyId, rules: a.policy?.rules || [] }
}

const hoursOf = (iso: unknown) => {
  const m = String(iso || '').match(/P(?:(\d+)D)?T?(?:(\d+)H)?/)
  return m ? +(m[1] || 0) * 24 + +(m[2] || 0) : 8
}

export async function roleSettings(tid: string, roleId: string): Promise<PimRoleSettings> {
  const tok = await token(tid)
  const { policyId, rules } = await policyFor(tok, { kind: 'role', roleId })
  const exp = rules.find((r) => r.id === 'Expiration_EndUser_Assignment')
  const en = rules.find((r) => r.id === 'Enablement_EndUser_Assignment')
  const ap = rules.find((r) => r.id === 'Approval_EndUser_Assignment') as RawRule & {
    setting?: { isApprovalRequired: boolean; approvalStages?: { primaryApprovers?: { '@odata.type': string; userId?: string; groupId?: string; description?: string }[] }[] }
  }
  const enabled = (en?.enabledRules as string[]) || []
  const raw = ap?.setting?.approvalStages?.[0]?.primaryApprovers || []
  const ids = raw.map((x) => x.userId || x.groupId).filter(Boolean) as string[]
  const names = ids.length ? await resolveIds(tid, ids).catch(() => ({}) as Record<string, string>) : {}
  return {
    policyId,
    maxHours: hoursOf(exp?.maximumDuration),
    mfa: enabled.includes('MultiFactorAuthentication'),
    justification: enabled.includes('Justification'),
    ticket: enabled.includes('Ticketing'),
    approval: !!ap?.setting?.isApprovalRequired,
    approvers: raw.map((x) => {
      const id = (x.userId || x.groupId)!
      return { id, displayName: names[id] || x.description || id, detail: null, type: x.groupId ? 'group' : 'user' }
    })
  }
}

export async function setRoleSettings(tid: string, roleId: string, s: Omit<PimRoleSettings, 'policyId'>) {
  const tok = await token(tid)
  const { policyId, rules } = await policyFor(tok, { kind: 'role', roleId })
  const patch = (rule: RawRule | undefined, change: Record<string, unknown>) =>
    rule && call(tok, `/policies/roleManagementPolicies/${policyId}/rules/${rule.id}`, json('PATCH', { ...rule, ...change }))
  await patch(rules.find((r) => r.id === 'Expiration_EndUser_Assignment'), { maximumDuration: `PT${s.maxHours}H`, isExpirationRequired: true })
  await patch(rules.find((r) => r.id === 'Enablement_EndUser_Assignment'), {
    enabledRules: [s.mfa && 'MultiFactorAuthentication', s.justification && 'Justification', s.ticket && 'Ticketing'].filter(Boolean)
  })
  const ap = rules.find((r) => r.id === 'Approval_EndUser_Assignment') as (RawRule & { setting: Record<string, unknown> }) | undefined
  if (ap) {
    if (s.approval && !s.approvers.length) throw new Error('Approval needs at least one approver')
    await patch(ap, {
      setting: {
        ...ap.setting,
        isApprovalRequired: s.approval,
        approvalStages: [
          {
            ...((ap.setting.approvalStages as Record<string, unknown>[] | undefined)?.[0] || {}),
            approvalStageTimeOutInDays: 1,
            isApproverJustificationRequired: true,
            escalationTimeInMinutes: 0,
            isEscalationEnabled: false,
            primaryApprovers: s.approvers.map((a) =>
              a.type === 'group'
                ? { '@odata.type': '#microsoft.graph.groupMembers', groupId: a.id, description: a.displayName }
                : { '@odata.type': '#microsoft.graph.singleUser', userId: a.id, description: a.displayName }
            ),
            escalationApprovers: []
          }
        ]
      }
    })
  }
  return roleSettings(tid, roleId)
}

type GroupInst = { id: string; principalId: string; accessId: 'member' | 'owner'; assignmentType?: string; memberType?: string; startDateTime: string | null; endDateTime: string | null; principal?: RawRef }

/** Eligible and active members and owners of one group under PIM for groups. */
export async function pimGroup(tid: string, groupId: string): Promise<PimGroupEntry[]> {
  const tok = await token(tid)
  const [elig, act] = await Promise.all([
    all<GroupInst>(tok, `${GROUP}/eligibilityScheduleInstances?$filter=groupId eq '${groupId}'&$expand=principal`),
    all<GroupInst>(tok, `${GROUP}/assignmentScheduleInstances?$filter=groupId eq '${groupId}'&$expand=principal`)
  ])
  return [
    ...elig.map((e) => ({ id: e.id, principal: ref(e.principal, e.principalId), accessId: e.accessId, kind: 'eligible' as const, start: e.startDateTime, end: e.endDateTime })),
    ...act
      .filter((a) => a.memberType !== 'Inherited')
      .map((a) => ({ id: a.id, principal: ref(a.principal, a.principalId), accessId: a.accessId, kind: a.assignmentType === 'Activated' || a.endDateTime ? ('active' as const) : ('permanent' as const), start: a.startDateTime, end: a.endDateTime }))
  ].sort((a, b) => a.principal.displayName.localeCompare(b.principal.displayName))
}

export async function pimGroupAssign(tid: string, groupId: string, a: { principalId: string; accessId: 'member' | 'owner'; kind: 'eligible' | 'active'; start?: string; end?: string | null; justification?: string }) {
  const tok = await token(tid)
  await call(tok, `${GROUP}/${a.kind === 'eligible' ? 'eligibilityScheduleRequests' : 'assignmentScheduleRequests'}`, json('POST', {
    groupId,
    principalId: a.principalId,
    accessId: a.accessId,
    action: 'adminAssign',
    justification: a.justification || 'Assigned with M365 Admin Toolbox',
    scheduleInfo: schedule(a.start, a.end)
  }))
  return pimGroup(tid, groupId)
}

export async function pimGroupRemove(tid: string, groupId: string, e: Pick<PimGroupEntry, 'accessId' | 'kind'> & { principalId: string }) {
  const tok = await token(tid)
  await call(tok, `${GROUP}/${e.kind === 'eligible' ? 'eligibilityScheduleRequests' : 'assignmentScheduleRequests'}`, json('POST', {
    groupId,
    principalId: e.principalId,
    accessId: e.accessId,
    action: 'adminRemove',
    justification: 'Removed with M365 Admin Toolbox'
  }))
  return pimGroup(tid, groupId)
}

/** PIM activity from the directory audit log (who activated or assigned which role, and why). */
export async function pimAudit(tid: string, days: number): Promise<PimAudit[]> {
  const tok = await tokenFor(tid)
  const since = new Date(Date.now() - days * 86400000).toISOString()
  type Raw = {
    id: string
    activityDateTime: string
    activityDisplayName: string
    result: string
    resultReason: string | null
    initiatedBy: { user?: { userPrincipalName: string | null; displayName: string | null } | null; app?: { displayName: string | null } | null }
    targetResources: { displayName: string | null; type: string; userPrincipalName?: string | null }[]
    additionalDetails: { key: string; value: string }[]
  }
  const r = await call<{ value: Raw[] }>(
    tok,
    `/auditLogs/directoryAudits?$filter=loggedByService eq 'PIM' and activityDateTime ge ${since}&$orderby=activityDateTime desc&$top=500`
  )
  return r.value.map((x) => ({
    id: x.id,
    date: x.activityDateTime,
    activity: x.activityDisplayName,
    by: x.initiatedBy.user?.userPrincipalName || x.initiatedBy.user?.displayName || x.initiatedBy.app?.displayName || null,
    target: x.targetResources.map((t) => t.displayName || t.userPrincipalName).filter(Boolean).join(' · ') || null,
    result: x.result,
    reason: x.additionalDetails.find((d) => /justification/i.test(d.key))?.value || x.resultReason || null
  }))
}

// ---------- MFA methods ----------

const METHOD_PATH: Record<string, string> = {
  microsoftAuthenticatorAuthenticationMethod: 'microsoftAuthenticatorMethods',
  phoneAuthenticationMethod: 'phoneMethods',
  fido2AuthenticationMethod: 'fido2Methods',
  windowsHelloForBusinessAuthenticationMethod: 'windowsHelloForBusinessMethods',
  emailAuthenticationMethod: 'emailMethods',
  softwareOathAuthenticationMethod: 'softwareOathMethods',
  temporaryAccessPassAuthenticationMethod: 'temporaryAccessPassMethods',
  platformCredentialAuthenticationMethod: 'platformCredentialMethods'
}

/** Registered methods per user: the registration report when the tenant has it (Entra ID P1), otherwise per user. */
export async function mfaReport(tid: string): Promise<MfaReport> {
  const tok = await token(tid)
  try {
    type Row = { id: string; userPrincipalName: string; userDisplayName: string; isAdmin: boolean; isMfaCapable: boolean; methodsRegistered: string[]; defaultMfaMethod: string | null; userType: string }
    const rows = await all<Row>(tok, '/reports/authenticationMethods/userRegistrationDetails?$top=999')
    return {
      source: 'report',
      users: rows
        .filter((r) => r.userType !== 'guest')
        .map((r) => ({
          userId: r.id,
          displayName: r.userDisplayName,
          upn: r.userPrincipalName,
          methods: r.methodsRegistered,
          mfaCapable: r.isMfaCapable,
          defaultMethod: r.defaultMfaMethod && r.defaultMfaMethod !== 'none' ? r.defaultMfaMethod : null,
          isAdmin: r.isAdmin
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName))
    }
  } catch (e) {
    if (!(e instanceof GraphError) || (e.status !== 403 && e.status !== 400)) throw e
  }
  const users = await all<{ id: string; displayName: string; userPrincipalName: string }>(tok, "/users?$filter=userType eq 'Member'&$select=id,displayName,userPrincipalName&$top=999")
  const res = await batch(tok, users.map((u) => ({ id: u.id, url: `/users/${u.id}/authentication/methods` })))
  const out: UserMfa[] = users.map((u) => {
    const r = res.get(u.id)
    const types = r?.status === 200 ? ((r.body.value || []) as { '@odata.type': string }[]).map((m) => m['@odata.type'].replace('#microsoft.graph.', '').replace('AuthenticationMethod', '')) : []
    const strong = types.filter((x) => x !== 'password' && x !== 'email')
    return { userId: u.id, displayName: u.displayName, upn: u.userPrincipalName, methods: types, mfaCapable: strong.length > 0, defaultMethod: null, isAdmin: false }
  })
  return { source: 'methods', users: out.sort((a, b) => a.displayName.localeCompare(b.displayName)) }
}

export async function methods(tid: string, userId: string): Promise<AuthMethod[]> {
  const tok = await token(tid)
  type Raw = { id: string; '@odata.type': string; displayName?: string | null; phoneNumber?: string | null; emailAddress?: string | null; deviceTag?: string | null; model?: string | null; createdDateTime?: string | null }
  return (await all<Raw>(tok, `/users/${userId}/authentication/methods`)).map((m) => ({
    id: m.id,
    type: m['@odata.type'].replace('#microsoft.graph.', ''),
    label: m.displayName || m.phoneNumber || m.emailAddress || m.model || m.deviceTag || null,
    created: m.createdDateTime || null
  }))
}

export async function deleteMethod(tid: string, userId: string, m: Pick<AuthMethod, 'id' | 'type'>) {
  const path = METHOD_PATH[m.type]
  if (!path) throw new Error(`This method (${m.type}) cannot be removed here`)
  await call(await token(tid), `/users/${userId}/authentication/${path}/${m.id}`, { method: 'DELETE' })
}

/**
 * "Reset MFA": removes every method except the password, so the user registers again at the next sign-in.
 * The default method can only go once the others are gone, so failures get a second pass.
 */
export async function resetMfa(tid: string, userId: string, revoke: boolean) {
  let list = (await methods(tid, userId)).filter((m) => METHOD_PATH[m.type])
  const failed: { method: string; error: string }[] = []
  for (let pass = 0; pass < 2 && list.length; pass++) {
    const retry: AuthMethod[] = []
    for (const m of list) {
      try {
        await deleteMethod(tid, userId, m)
      } catch (e) {
        if (pass === 1) failed.push({ method: m.type, error: (e as Error).message })
        else retry.push(m)
      }
    }
    list = retry
  }
  if (revoke) await call(await tokenFor(tid), `/users/${userId}/revokeSignInSessions`, { method: 'POST' })
  return { failed, methods: await methods(tid, userId) }
}

// ---------- guests ----------

export async function guestUsers(tid: string): Promise<EntraGuest[]> {
  const tok = await tokenFor(tid)
  const sel = 'id,displayName,mail,userPrincipalName,externalUserState,externalUserStateChangeDateTime,createdDateTime,accountEnabled,companyName'
  type Raw = { id: string; displayName: string | null; mail: string | null; userPrincipalName: string; externalUserState: string | null; externalUserStateChangeDateTime: string | null; createdDateTime: string | null; accountEnabled: boolean; companyName: string | null; signInActivity?: { lastSignInDateTime: string | null; lastNonInteractiveSignInDateTime?: string | null } }
  let raw: Raw[]
  try {
    raw = await all<Raw>(tok, `/users?$filter=userType eq 'Guest'&$select=${sel},signInActivity&$top=999`)
  } catch (e) {
    if (!(e instanceof GraphError) || (e.status !== 403 && e.status !== 400)) throw e
    raw = await all<Raw>(tok, `/users?$filter=userType eq 'Guest'&$select=${sel}&$top=999`)
  }
  return raw
    .map((u) => {
      const a = u.signInActivity?.lastSignInDateTime || null
      const b = u.signInActivity?.lastNonInteractiveSignInDateTime || null
      return {
        id: u.id,
        displayName: u.displayName || u.mail || u.userPrincipalName,
        mail: u.mail,
        upn: u.userPrincipalName,
        state: u.externalUserState,
        stateChanged: u.externalUserStateChangeDateTime,
        created: u.createdDateTime,
        lastSignIn: a && b ? (a > b ? a : b) : a || b,
        enabled: u.accountEnabled,
        company: u.companyName
      }
    })
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function invite(tid: string, i: Invite) {
  const tok = await token(tid)
  const r = await call<{ inviteRedeemUrl: string; invitedUser: { id: string } }>(tok, '/invitations', json('POST', {
    invitedUserEmailAddress: i.email,
    invitedUserDisplayName: i.displayName || undefined,
    inviteRedirectUrl: 'https://myapps.microsoft.com',
    sendInvitationMessage: i.send,
    invitedUserMessageInfo: i.message ? { customizedMessageBody: i.message } : undefined
  }))
  return { userId: r.invitedUser.id, redeemUrl: r.inviteRedeemUrl }
}

/** Deletes guest accounts; they stay restorable for 30 days under Deleted users. */
export async function deleteGuests(tid: string, ids: string[]) {
  const tok = await tokenFor(tid)
  const res = await batch(tok, ids.map((id) => ({ id, url: `/users/${id}`, method: 'DELETE' })))
  return ids.filter((id) => res.get(id)?.status !== 204).map((id) => ({ id, error: String(res.get(id)?.body?.error?.message || 'failed') }))
}
