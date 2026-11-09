import type { Group, LicenseChange, NewGroup, NewUser, Sku, User } from '@shared/types'
import { tokenFor } from './auth'
import { skuName } from './skus'

const BASE = 'https://graph.microsoft.com/v1.0'

export class GraphError extends Error {
  constructor(
    message: string,
    public status: number,
    public code?: string
  ) {
    super(message)
  }
}

export async function call<T>(token: string, path: string, init: RequestInit = {}): Promise<T> {
  const res = await fetch(path.startsWith('http') ? path : BASE + path, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(init.headers || {}) }
  })
  if (res.status === 204) return undefined as T
  const body = await res.json().catch(() => ({}))
  if (!res.ok) {
    const err = body?.error || {}
    throw new GraphError(err.message || `Graph ${res.status}`, res.status, err.code)
  }
  return body as T
}

export async function all<T>(token: string, path: string, headers?: Record<string, string>): Promise<T[]> {
  const out: T[] = []
  let next: string | undefined = path
  while (next) {
    const page: { value: T[]; '@odata.nextLink'?: string } = await call(token, next, { headers })
    out.push(...page.value)
    next = page['@odata.nextLink']
  }
  return out
}

export async function organization(token: string) {
  const r = await call<{ value: { id: string; displayName: string; verifiedDomains: { name: string; isDefault: boolean }[] }[] }>(
    token,
    '/organization?$select=id,displayName,verifiedDomains'
  )
  const org = r.value[0]
  const domains = org.verifiedDomains.map((d) => d.name)
  const def = org.verifiedDomains.find((d) => d.isDefault)?.name || domains[0]
  return { name: org.displayName, defaultDomain: def, domains: [def, ...domains.filter((d) => d !== def)] }
}

type RawUser = {
  id: string
  displayName: string
  userPrincipalName: string
  department: string | null
  accountEnabled: boolean
  assignedLicenses: { skuId: string }[]
  createdDateTime: string | null
  signInActivity?: { lastSignInDateTime: string | null }
}

const USER_FIELDS = 'id,displayName,userPrincipalName,department,accountEnabled,assignedLicenses,createdDateTime'

export async function listUsers(tenantId: string): Promise<User[]> {
  const token = await tokenFor(tenantId)
  let raw: RawUser[]
  try {
    // signInActivity needs Entra ID P1 and AuditLog.Read.All; tenants without it get the list without that column.
    raw = await all<RawUser>(token, `/users?$top=999&$select=${USER_FIELDS},signInActivity`)
  } catch (e) {
    if (!(e instanceof GraphError) || (e.status !== 403 && e.status !== 400)) throw e
    raw = await all<RawUser>(token, `/users?$top=999&$select=${USER_FIELDS}`)
  }
  return raw
    .map((u) => ({
      id: u.id,
      displayName: u.displayName,
      userPrincipalName: u.userPrincipalName,
      department: u.department,
      accountEnabled: u.accountEnabled,
      skuIds: u.assignedLicenses.map((l) => l.skuId),
      lastSignIn: u.signInActivity?.lastSignInDateTime ?? null,
      createdDateTime: u.createdDateTime
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

type RawSku = {
  skuId: string
  skuPartNumber: string
  appliesTo: string
  capabilityStatus: string
  consumedUnits: number
  prepaidUnits: { enabled: number; suspended: number; warning: number }
}

export async function listSkus(tenantId: string): Promise<Sku[]> {
  const token = await tokenFor(tenantId)
  const raw = await all<RawSku>(token, '/subscribedSkus')
  return raw
    .filter((s) => s.appliesTo === 'User')
    .map((s) => {
      const enabled = s.prepaidUnits.enabled
      return {
        skuId: s.skuId,
        skuPartNumber: s.skuPartNumber,
        name: skuName(s.skuPartNumber),
        enabled,
        suspended: s.prepaidUnits.suspended,
        warning: s.prepaidUnits.warning,
        consumed: s.consumedUnits,
        available: Math.max(0, enabled - s.consumedUnits),
        shortage: Math.max(0, s.consumedUnits - enabled),
        capabilityStatus: s.capabilityStatus
      }
    })
    // Hide free/trial SKUs nobody uses and that have no paid seats, they only clutter the overview.
    .filter((s) => s.enabled > 0 || s.consumed > 0)
    .sort((a, b) => a.name.localeCompare(b.name))
}

export async function createUser(tenantId: string, u: NewUser): Promise<User> {
  const token = await tokenFor(tenantId)
  const created = await call<RawUser>(token, '/users', {
    method: 'POST',
    body: JSON.stringify({
      accountEnabled: true,
      displayName: u.displayName,
      givenName: u.givenName || undefined,
      surname: u.surname || undefined,
      mailNickname: u.mailNickname,
      userPrincipalName: u.userPrincipalName,
      usageLocation: u.usageLocation,
      department: u.department || undefined,
      passwordProfile: { password: u.password, forceChangePasswordNextSignIn: u.forceChangePasswordNextSignIn }
    })
  })
  if (u.skuIds.length) {
    await call(token, `/users/${created.id}/assignLicense`, {
      method: 'POST',
      body: JSON.stringify({ addLicenses: u.skuIds.map((skuId) => ({ skuId, disabledPlans: [] })), removeLicenses: [] })
    })
  }
  return {
    id: created.id,
    displayName: u.displayName,
    userPrincipalName: u.userPrincipalName,
    department: u.department || null,
    accountEnabled: true,
    skuIds: u.skuIds,
    lastSignIn: null,
    createdDateTime: new Date().toISOString()
  }
}

export async function setSignIn(tenantId: string, userId: string, enabled: boolean) {
  const token = await tokenFor(tenantId)
  await call(token, `/users/${userId}`, { method: 'PATCH', body: JSON.stringify({ accountEnabled: enabled }) })
}

export async function resetPassword(tenantId: string, userId: string, password: string, forceChange: boolean) {
  const token = await tokenFor(tenantId)
  await call(token, `/users/${userId}`, {
    method: 'PATCH',
    body: JSON.stringify({ passwordProfile: { password, forceChangePasswordNextSignIn: forceChange } })
  })
}

/** Requires the Temporary Access Pass method to be enabled in the tenant's authentication methods policy. */
export async function createTap(tenantId: string, userId: string, lifetimeInMinutes: number, isUsableOnce: boolean) {
  const token = await tokenFor(tenantId)
  const r = await call<{ temporaryAccessPass: string; lifetimeInMinutes: number; startDateTime: string }>(
    token,
    `/users/${userId}/authentication/temporaryAccessPassMethods`,
    { method: 'POST', body: JSON.stringify({ lifetimeInMinutes, isUsableOnce }) }
  )
  return { pass: r.temporaryAccessPass, lifetimeInMinutes: r.lifetimeInMinutes, start: r.startDateTime }
}

export async function revokeSessions(tenantId: string, userId: string) {
  const token = await tokenFor(tenantId)
  await call(token, `/users/${userId}/revokeSignInSessions`, { method: 'POST' })
}

/** Applies license changes one user at a time and reports per-user failures instead of stopping. */
export async function changeLicenses(tenantId: string, changes: LicenseChange[]) {
  const token = await tokenFor(tenantId)
  const failed: { userId: string; error: string }[] = []
  for (const c of changes) {
    if (!c.add.length && !c.remove.length) continue
    // A user needs a usage location before a license can be assigned; default to NL when missing.
    try {
      await call(token, `/users/${c.userId}/assignLicense`, {
        method: 'POST',
        body: JSON.stringify({ addLicenses: c.add.map((skuId) => ({ skuId, disabledPlans: [] })), removeLicenses: c.remove })
      })
    } catch (e) {
      if (e instanceof GraphError && /usage location/i.test(e.message)) {
        try {
          await call(token, `/users/${c.userId}`, { method: 'PATCH', body: JSON.stringify({ usageLocation: 'NL' }) })
          await call(token, `/users/${c.userId}/assignLicense`, {
            method: 'POST',
            body: JSON.stringify({ addLicenses: c.add.map((skuId) => ({ skuId, disabledPlans: [] })), removeLicenses: c.remove })
          })
          continue
        } catch (e2) {
          failed.push({ userId: c.userId, error: (e2 as Error).message })
          continue
        }
      }
      failed.push({ userId: c.userId, error: (e as Error).message })
    }
  }
  return { failed }
}

type RawGroup = {
  id: string
  displayName: string
  description: string | null
  groupTypes: string[]
  mailEnabled: boolean
  securityEnabled: boolean
  mail: string | null
  createdDateTime: string | null
}

const kindOf = (g: RawGroup): Group['kind'] =>
  g.groupTypes.includes('Unified') ? 'm365' : g.mailEnabled && g.securityEnabled ? 'mailSecurity' : g.mailEnabled ? 'distribution' : 'security'

const GROUP_FIELDS = 'id,displayName,description,groupTypes,mailEnabled,securityEnabled,mail,createdDateTime'

export async function listGroups(tenantId: string): Promise<Group[]> {
  const token = await tokenFor(tenantId)
  const raw = await all<RawGroup>(token, `/groups?$top=999&$select=${GROUP_FIELDS}`)
  return raw
    .map((g) => ({ id: g.id, displayName: g.displayName, description: g.description, kind: kindOf(g), mail: g.mail, createdDateTime: g.createdDateTime }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function createGroup(tenantId: string, g: NewGroup): Promise<Group> {
  const token = await tokenFor(tenantId)
  const m365 = g.kind === 'm365'
  const body: Record<string, unknown> = {
    displayName: g.displayName,
    description: g.description || undefined,
    mailNickname: g.mailNickname,
    groupTypes: m365 ? ['Unified'] : [],
    mailEnabled: m365,
    securityEnabled: !m365
  }
  if (g.ownerId) body['owners@odata.bind'] = [`${BASE}/users/${g.ownerId}`]
  const r = await call<RawGroup>(token, '/groups', { method: 'POST', body: JSON.stringify(body) })
  return { id: r.id, displayName: r.displayName, description: r.description, kind: kindOf(r), mail: r.mail, createdDateTime: r.createdDateTime }
}
