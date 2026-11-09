import type { Group, MemberGroup, SearchHit } from '@shared/types'
import { tokenFor } from './auth'
import * as exo from './exchange'
import { all, call } from './graph'

const BASE = 'https://graph.microsoft.com/v1.0'
const EVENTUAL = { ConsistencyLevel: 'eventual' }
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i

type RawUser = {
  id: string
  displayName: string
  userPrincipalName: string
  mail: string | null
  department: string | null
  accountEnabled: boolean
  userType: string | null
  proxyAddresses: string[]
}
type RawGroup = {
  id: string
  displayName: string
  groupTypes: string[]
  mailEnabled: boolean
  securityEnabled: boolean
  mail: string | null
  description: string | null
  createdDateTime: string | null
  proxyAddresses?: string[]
  onPremisesSyncEnabled?: boolean | null
}
type RawSp = { id: string; appId: string; displayName: string; accountEnabled: boolean; publisherName: string | null }
type RawApp = { id: string; appId: string; displayName: string; publisherDomain: string | null }

const kindOf = (g: RawGroup): Group['kind'] =>
  g.groupTypes.includes('Unified') ? 'm365' : g.mailEnabled && g.securityEnabled ? 'mailSecurity' : g.mailEnabled ? 'distribution' : 'security'

/** The proxy address that matched the query, when it isn't the primary one. */
function alias(q: string, primary: string | null, proxies: string[] = []) {
  const l = q.toLowerCase()
  if (primary?.toLowerCase().startsWith(l)) return null
  const hit = proxies.map((p) => p.replace(/^smtp:/i, '')).find((p) => p.toLowerCase().startsWith(l))
  return hit || null
}

/** Finds users, guests, groups and apps whose name, UPN, mail address, alias or app id matches. */
export async function search(tid: string, query: string): Promise<SearchHit[]> {
  const token = await tokenFor(tid)
  const q = query.replace(/["\\]/g, '').trim()
  if (q.length < 2) return []
  // Encoded, so a plus or ampersand in an address survives the query string.
  const esc = encodeURIComponent(q.replace(/'/g, "''"))
  const sq = encodeURIComponent(q)
  const mail = q.includes('@')
  const guid = GUID.test(q)
  const uSel = 'id,displayName,userPrincipalName,mail,department,accountEnabled,userType,proxyAddresses'
  const gSel = 'id,displayName,groupTypes,mailEnabled,securityEnabled,mail,description,createdDateTime,proxyAddresses'

  const userPath = guid
    ? `/users?$filter=id eq '${esc}'&$select=${uSel}`
    : mail
      ? `/users?$count=true&$top=25&$select=${uSel}&$filter=startswith(mail,'${esc}') or startswith(userPrincipalName,'${esc}') or proxyAddresses/any(p:startswith(p,'smtp:${esc}'))`
      : `/users?$count=true&$top=25&$select=${uSel}&$search="displayName:${sq}" OR "userPrincipalName:${sq}" OR "mail:${sq}"`
  const groupPath = guid
    ? `/groups?$filter=id eq '${esc}'&$select=${gSel}`
    : mail
      ? `/groups?$count=true&$top=25&$select=${gSel}&$filter=startswith(mail,'${esc}') or proxyAddresses/any(p:startswith(p,'smtp:${esc}'))`
      : `/groups?$count=true&$top=25&$select=${gSel}&$search="displayName:${sq}" OR "mail:${sq}"`
  const spSel = 'id,appId,displayName,accountEnabled,publisherName'
  const appSel = 'id,appId,displayName,publisherDomain'
  const spPath = guid
    ? `/servicePrincipals?$filter=appId eq '${esc}' or id eq '${esc}'&$select=${spSel}`
    : `/servicePrincipals?$count=true&$top=15&$select=${spSel}&$search="displayName:${sq}"`
  const appPath = guid
    ? `/applications?$filter=appId eq '${esc}' or id eq '${esc}'&$select=${appSel}`
    : `/applications?$count=true&$top=15&$select=${appSel}&$search="displayName:${sq}"`

  const get = <T>(p: string) => call<{ value: T[] }>(token, p, { headers: EVENTUAL }).then((r) => r.value)
  const [users, groups, sps, apps] = await Promise.allSettled([
    get<RawUser>(userPath),
    get<RawGroup>(groupPath),
    mail ? Promise.resolve([] as RawSp[]) : get<RawSp>(spPath),
    mail ? Promise.resolve([] as RawApp[]) : get<RawApp>(appPath)
  ])
  // Only fail the tenant when the directory itself can't be read; apps are a bonus.
  if (users.status === 'rejected' && groups.status === 'rejected') throw users.reason

  const hits: SearchHit[] = []
  if (users.status === 'fulfilled')
    for (const u of users.value)
      hits.push({
        kind: u.userType === 'Guest' ? 'guest' : 'user',
        id: u.id,
        name: u.displayName,
        sub: u.userPrincipalName,
        detail: u.department,
        enabled: u.accountEnabled,
        matched: mail ? alias(q, u.mail, u.proxyAddresses) : null
      })
  if (groups.status === 'fulfilled')
    for (const g of groups.value)
      hits.push({ kind: 'group', id: g.id, name: g.displayName, sub: g.mail, detail: kindOf(g), enabled: null, matched: mail ? alias(q, g.mail, g.proxyAddresses) : null })
  if (sps.status === 'fulfilled')
    for (const s of sps.value)
      hits.push({ kind: 'app', id: s.id, name: s.displayName, sub: s.appId, detail: s.publisherName, enabled: s.accountEnabled, matched: null })
  if (apps.status === 'fulfilled')
    for (const a of apps.value) hits.push({ kind: 'appReg', id: a.id, name: a.displayName, sub: a.appId, detail: a.publisherDomain, enabled: null, matched: null })
  return hits
}

/** Groups the user is a direct member of. */
export async function userGroups(tid: string, user: string): Promise<MemberGroup[]> {
  const token = await tokenFor(tid)
  const raw = await all<RawGroup & { membershipRule?: string | null }>(
    token,
    `/users/${encodeURIComponent(user)}/memberOf/microsoft.graph.group?$top=999&$select=id,displayName,description,groupTypes,mailEnabled,securityEnabled,mail,createdDateTime,onPremisesSyncEnabled`
  )
  return raw
    .map((g) => ({
      id: g.id,
      displayName: g.displayName,
      description: g.description,
      kind: kindOf(g),
      mail: g.mail,
      createdDateTime: g.createdDateTime,
      readOnly: g.groupTypes.includes('DynamicMembership') || !!g.onPremisesSyncEnabled
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

type GroupRef = Pick<Group, 'id' | 'displayName' | 'kind' | 'mail'>

/**
 * Adds or removes a user in several groups. Graph can't change distribution lists or mail-enabled
 * security groups, so those go through Exchange.
 */
export async function setGroups(tid: string, userId: string, upn: string, groups: GroupRef[], add: boolean) {
  const token = await tokenFor(tid)
  const failed: { id: string; error: string }[] = []
  for (const g of groups) {
    try {
      if (g.kind === 'distribution' || g.kind === 'mailSecurity') {
        if (add) await exo.addMember(tid, g.mail || g.id, upn)
        else await exo.removeMember(tid, g.mail || g.id, upn)
      } else if (add) {
        await call(token, `/groups/${g.id}/members/$ref`, { method: 'POST', body: JSON.stringify({ '@odata.id': `${BASE}/directoryObjects/${userId}` }) })
      } else {
        await call(token, `/groups/${g.id}/members/${userId}/$ref`, { method: 'DELETE' })
      }
    } catch (e) {
      failed.push({ id: g.displayName, error: (e as Error).message })
    }
  }
  return { failed }
}

/** Removes all licenses from a user. */
export async function removeLicenses(tid: string, user: string) {
  const token = await tokenFor(tid)
  const u = await call<{ assignedLicenses: { skuId: string }[] }>(token, `/users/${encodeURIComponent(user)}?$select=assignedLicenses`)
  if (u.assignedLicenses.length)
    await call(token, `/users/${encodeURIComponent(user)}/assignLicense`, {
      method: 'POST',
      body: JSON.stringify({ addLicenses: [], removeLicenses: u.assignedLicenses.map((l) => l.skuId) })
    })
  return { removed: u.assignedLicenses.length }
}

/** Looks a user up by UPN or object id, for the bulk license and wizard flows. */
export async function findUser(tid: string, user: string) {
  const token = await tokenFor(tid)
  const u = await call<{ id: string; displayName: string; userPrincipalName: string; assignedLicenses: { skuId: string }[] }>(
    token,
    `/users/${encodeURIComponent(user)}?$select=id,displayName,userPrincipalName,assignedLicenses`
  )
  return { id: u.id, displayName: u.displayName, userPrincipalName: u.userPrincipalName, skuIds: u.assignedLicenses.map((l) => l.skuId) }
}
