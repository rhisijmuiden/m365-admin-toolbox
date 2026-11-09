import type { PimData, PimItem, PimRequest } from '@shared/types'
import { homeTenantOf, stepUp, tokenFor, tokenForHome } from './auth'
import { all, call, GraphError } from './graph'
import { tenants } from './store'
import { track } from './pimWatch'

/*
 * Directory tenants (own admin login): the signed-in user activates Entra roles in that tenant.
 * GDAP tenants: the partner user has no roles in the customer tenant. Access comes from security
 * groups in the partner tenant, so activation is PIM for groups membership there.
 */

const ROLE = '/roleManagement/directory'
const GROUP = '/identityGovernance/privilegedAccess/group'

type RoleInstance = {
  roleDefinitionId: string
  directoryScopeId: string
  assignmentType?: string
  startDateTime: string | null
  endDateTime: string | null
  roleDefinition?: { displayName: string }
}
type GroupInstance = {
  groupId: string
  accessId: string
  assignmentType?: string
  startDateTime: string | null
  endDateTime: string | null
  group?: { displayName: string }
}
type RoleRequest = { id: string; roleDefinitionId: string; status: string; createdDateTime: string; roleDefinition?: { displayName: string } }
type GroupRequest = { id: string; groupId: string; status: string; createdDateTime: string; group?: { displayName: string } }

const isDirectory = (tenantId: string) => tenants.get(tenantId)!.mode === 'direct'

async function me(token: string) {
  return (await call<{ id: string }>(token, '/me?$select=id')).id
}

/** Group ids linked to this customer through active GDAP relationships, or null when that can't be read. */
async function gdapGroups(homeToken: string, customerTenantId: string): Promise<Set<string> | null> {
  try {
    const rels = await all<{ id: string; status: string; customer: { tenantId: string } }>(
      homeToken,
      `/tenantRelationships/delegatedAdminRelationships?$filter=customer/tenantId eq '${customerTenantId}'`
    )
    const ids = new Set<string>()
    for (const r of rels.filter((x) => x.status === 'active')) {
      const aa = await all<{ status: string; accessContainer: { accessContainerId: string } }>(
        homeToken,
        `/tenantRelationships/delegatedAdminRelationships/${r.id}/accessAssignments`
      )
      aa.filter((a) => a.status === 'active').forEach((a) => ids.add(a.accessContainer.accessContainerId))
    }
    return ids.size ? ids : null
  } catch {
    return null
  }
}

export async function load(tenantId: string): Promise<PimData> {
  if (isDirectory(tenantId)) {
    const token = await tokenFor(tenantId)
    const sel = "filterByCurrentUser(on='principal')?$expand=roleDefinition($select=displayName)"
    const [elig, act, reqs] = await Promise.all([
      all<RoleInstance>(token, `${ROLE}/roleEligibilityScheduleInstances/${sel}`),
      all<RoleInstance>(token, `${ROLE}/roleAssignmentScheduleInstances/${sel}`),
      all<RoleRequest>(token, `${ROLE}/roleAssignmentScheduleRequests/${sel}`)
    ])
    const active: PimItem[] = act.map((a) => ({
      kind: 'role',
      targetId: a.roleDefinitionId,
      name: a.roleDefinition?.displayName || a.roleDefinitionId,
      state: a.assignmentType === 'Activated' ? 'active' : 'permanent',
      start: a.startDateTime,
      end: a.endDateTime
    }))
    const activeIds = new Set(active.map((a) => a.targetId))
    const data: PimData = {
      mode: 'directory',
      homeTenantId: tenantId,
      filtered: false,
      active,
      eligible: dedupe(
        elig.map((e) => ({
          kind: 'role' as const,
          targetId: e.roleDefinitionId,
          name: e.roleDefinition?.displayName || e.roleDefinitionId,
          state: 'eligible' as const,
          end: e.endDateTime
        }))
      ).filter((e) => !activeIds.has(e.targetId)),
      pending: reqs
        .filter((r) => r.status === 'PendingApproval' || r.status === 'PendingProvisioning')
        .map((r) => ({
          kind: 'role',
          targetId: r.roleDefinitionId,
          name: r.roleDefinition?.displayName || r.roleDefinitionId,
          state: 'pending',
          start: r.createdDateTime,
          requestId: r.id
        }))
    }
    track(tenantId, data.active)
    return data
  }

  const home = homeTenantOf(tenantId)
  const token = await tokenForHome(tenantId)
  const sel = "filterByCurrentUser(on='principal')?$expand=group($select=displayName)"
  const [elig, act, reqs, only] = await Promise.all([
    all<GroupInstance>(token, `${GROUP}/eligibilityScheduleInstances/${sel}`),
    all<GroupInstance>(token, `${GROUP}/assignmentScheduleInstances/${sel}`),
    all<GroupRequest>(token, `${GROUP}/assignmentScheduleRequests/${sel}`),
    gdapGroups(token, tenantId)
  ])
  const keep = (id: string) => !only || only.has(id)
  const active: PimItem[] = act
    .filter((a) => a.accessId === 'member' && keep(a.groupId))
    .map((a) => ({
      kind: 'group',
      targetId: a.groupId,
      name: a.group?.displayName || a.groupId,
      state: a.assignmentType === 'Activated' ? 'active' : 'permanent',
      start: a.startDateTime,
      end: a.endDateTime
    }))
  const activeIds = new Set(active.map((a) => a.targetId))
  const data: PimData = {
    mode: 'groups',
    homeTenantId: home,
    filtered: !!only,
    active,
    eligible: dedupe(
      elig
        .filter((e) => e.accessId === 'member' && keep(e.groupId))
        .map((e) => ({ kind: 'group' as const, targetId: e.groupId, name: e.group?.displayName || e.groupId, state: 'eligible' as const, end: e.endDateTime }))
    ).filter((e) => !activeIds.has(e.targetId)),
    pending: reqs
      .filter((r) => (r.status === 'PendingApproval' || r.status === 'PendingProvisioning') && keep(r.groupId))
      .map((r) => ({ kind: 'group', targetId: r.groupId, name: r.group?.displayName || r.groupId, state: 'pending', start: r.createdDateTime, requestId: r.id }))
  }
  track(tenantId, data.active)
  return data
}

const dedupe = (items: PimItem[]) => [...new Map(items.map((i) => [i.targetId, i])).values()].sort((a, b) => a.name.localeCompare(b.name))

type Action = 'selfActivate' | 'selfDeactivate' | 'selfExtend'

function body(action: Action, principalId: string, r: Pick<PimRequest, 'kind' | 'targetId'> & Partial<PimRequest>) {
  const withSchedule = action !== 'selfDeactivate'
  const base: Record<string, unknown> = { action, principalId }
  if (r.kind === 'role') Object.assign(base, { roleDefinitionId: r.targetId, directoryScopeId: '/' })
  else Object.assign(base, { groupId: r.targetId, accessId: 'member' })
  if (withSchedule) {
    base.justification = r.justification
    base.scheduleInfo = { startDateTime: new Date().toISOString(), expiration: { type: 'afterDuration', duration: `PT${r.hours}H` } }
    if (r.ticket) base.ticketInfo = { ticketNumber: r.ticket, ticketSystem: 'RHIS' }
  }
  return base
}

/** Pulls the claims challenge out of a PIM authentication-context error, or null for a plain MFA requirement. */
function claimsFrom(e: GraphError): string | null {
  const m = e.message.match(/claims=([^&\s"']+)/)
  if (m) return decodeURIComponent(m[1])
  if (/Acrs/i.test(e.code || '') || /authentication context/i.test(e.message)) {
    const c = e.message.match(/\bc([1-9]|[1-9][0-9])\b/)
    return JSON.stringify({ access_token: { acrs: { essential: true, value: c ? c[0] : 'c1' } } })
  }
  return null
}

const needsStepUp = (e: unknown): e is GraphError =>
  e instanceof GraphError && (/Acrs|MfaRule|AuthenticationContext/i.test(e.code || '') || /claims=|multi-factor|MFA|authentication context/i.test(e.message))

async function send(tenantId: string, action: Action, r: Pick<PimRequest, 'kind' | 'targetId'> & Partial<PimRequest>) {
  const t = tenants.get(tenantId)!
  const authorityTenant = r.kind === 'role' ? tenantId : homeTenantOf(tenantId)
  let token = r.kind === 'role' ? await tokenFor(tenantId) : await tokenForHome(tenantId)
  const path = r.kind === 'role' ? `${ROLE}/roleAssignmentScheduleRequests` : `${GROUP}/assignmentScheduleRequests`
  const principalId = await me(token)
  try {
    return await call<{ id: string; status: string }>(token, path, { method: 'POST', body: JSON.stringify(body(action, principalId, r)) })
  } catch (e) {
    if (!needsStepUp(e)) throw e
    token = await stepUp(authorityTenant, t.username, claimsFrom(e) || undefined)
    return await call<{ id: string; status: string }>(token, path, { method: 'POST', body: JSON.stringify(body(action, principalId, r)) })
  }
}

export async function activate(tenantId: string, r: PimRequest) {
  const res = await send(tenantId, 'selfActivate', r)
  if (res.status !== 'PendingApproval') {
    const end = new Date(Date.now() + r.hours * 3600_000).toISOString()
    track(tenantId, [{ kind: r.kind, targetId: r.targetId, name: r.name, state: 'active', end }], true)
  }
  return { status: res.status }
}

export async function deactivate(tenantId: string, kind: PimRequest['kind'], targetId: string) {
  await send(tenantId, 'selfDeactivate', { kind, targetId })
}

export async function extend(tenantId: string, r: PimRequest) {
  const res = await send(tenantId, 'selfExtend', r)
  return { status: res.status }
}

export async function cancel(tenantId: string, kind: PimRequest['kind'], requestId: string) {
  const token = kind === 'role' ? await tokenFor(tenantId) : await tokenForHome(tenantId)
  const path = kind === 'role' ? `${ROLE}/roleAssignmentScheduleRequests` : `${GROUP}/assignmentScheduleRequests`
  await call(token, `${path}/${requestId}/cancel`, { method: 'POST' })
}
