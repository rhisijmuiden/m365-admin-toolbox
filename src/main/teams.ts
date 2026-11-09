import type {
  ChannelType,
  GroupPolicyAssignment,
  NewTeam,
  NewTeamResult,
  NumberType,
  PolicyType,
  Team,
  TeamActivity,
  TeamActivityReport,
  TeamChannel,
  TeamGuest,
  TeamMember,
  TeamPerson,
  TeamsAccess,
  TeamsFederation,
  TeamsGuestSettings,
  TeamsPhone,
  TeamsPolicies,
  CsUser,
  TemplateChannel
} from '@shared/types'
import helperSource from './teams/helper.ps1?raw'
import { NeedsConsent, teamsTokens, tokenFor, tokenForScopes } from './auth'
import { all, call, GraphError } from './graph'
import { psHelper } from './ps'
import { REPORT_SCOPES } from './exchange'

/*
 * Teams, channels and members go through Graph with the app's own sign-in. The Teams endpoints need a few
 * scopes outside the base set, which existing tenants have not consented to yet, so they are asked for
 * separately (the screen offers a consent button on NEEDS_CONSENT) instead of breaking every other screen.
 */
export const TEAMS_SCOPES = [
  'Team.Create',
  'TeamMember.ReadWrite.All',
  'TeamSettings.ReadWrite.All',
  'Channel.Create',
  'Channel.Delete.All',
  'ChannelMember.ReadWrite.All'
]

const GRAPH = 'https://graph.microsoft.com/v1.0'
const token = (tid: string) => tokenForScopes(tid, TEAMS_SCOPES)
const userRef = (u: string) => `${GRAPH}/users('${encodeURIComponent(u)}')`
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

type BatchReq = { id: string; url: string; method?: string; body?: unknown }
type BatchRes = { id: string; status: number; body: { value?: unknown[]; error?: { message: string }; [k: string]: unknown }; headers?: Record<string, string> }

/** Runs Graph requests through $batch, 20 per call and a few calls at a time, retrying throttled ones. */
export async function batch(tok: string, reqs: BatchReq[]): Promise<Map<string, BatchRes>> {
  const out = new Map<string, BatchRes>()
  const chunks: BatchReq[][] = []
  for (let i = 0; i < reqs.length; i += 20) chunks.push(reqs.slice(i, i + 20))
  async function send(chunk: BatchReq[], attempt = 0): Promise<void> {
    const r = await call<{ responses: BatchRes[] }>(tok, '/$batch', {
      method: 'POST',
      body: JSON.stringify({
        requests: chunk.map((q) => ({
          id: q.id,
          method: q.method || 'GET',
          url: q.url,
          ...(q.body ? { body: q.body, headers: { 'Content-Type': 'application/json' } } : {})
        }))
      })
    })
    const retry: BatchReq[] = []
    let wait = 0
    for (const res of r.responses) {
      if (res.status === 429 && attempt < 4) {
        retry.push(chunk.find((q) => q.id === res.id)!)
        wait = Math.max(wait, Number(res.headers?.['Retry-After'] || 2) * 1000)
      } else out.set(res.id, res)
    }
    if (retry.length) {
      await sleep(wait)
      await send(retry, attempt + 1)
    }
  }
  for (let i = 0; i < chunks.length; i += 4) await Promise.all(chunks.slice(i, i + 4).map((c) => send(c)))
  return out
}

type RawPerson = { id: string; displayName: string | null; userPrincipalName?: string | null; mail?: string | null }
const person = (p: RawPerson): TeamPerson => ({ id: p.id, displayName: p.displayName || p.mail || p.id, upn: p.userPrincipalName || p.mail || '' })

type RawTeamGroup = { id: string; displayName: string; description: string | null; visibility: string | null; mail: string | null; createdDateTime: string | null }

const TEAM_FILTER = "resourceProvisioningOptions/Any(x:x eq 'Team')"
const TEAM_FIELDS = 'id,displayName,description,visibility,mail,createdDateTime'

/** Every team with its owners and archive state: one list call, then owners and settings via $batch. */
export async function listTeams(tid: string): Promise<Team[]> {
  const tok = await token(tid)
  const groups = await all<RawTeamGroup>(tok, `/groups?$filter=${TEAM_FILTER}&$select=${TEAM_FIELDS}&$top=999`)
  const res = await batch(
    tok,
    groups.flatMap((g) => [
      { id: `o${g.id}`, url: `/groups/${g.id}/owners?$select=id,displayName,userPrincipalName,mail&$top=100` },
      { id: `t${g.id}`, url: `/teams/${g.id}?$select=id,isArchived` }
    ])
  )
  return groups
    .map((g) => {
      const o = res.get(`o${g.id}`)
      const s = res.get(`t${g.id}`)
      return {
        id: g.id,
        displayName: g.displayName,
        description: g.description,
        visibility: (g.visibility || 'private').toLowerCase(),
        mail: g.mail,
        created: g.createdDateTime,
        owners: o?.status === 200 ? ((o.body.value || []) as RawPerson[]).map(person) : [],
        archived: s?.status === 200 ? Boolean(s.body.isArchived) : null
      }
    })
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
}

async function getTeam(tok: string, id: string): Promise<Team> {
  const [g, owners, s] = await Promise.all([
    call<RawTeamGroup>(tok, `/groups/${id}?$select=${TEAM_FIELDS}`),
    all<RawPerson>(tok, `/groups/${id}/owners?$select=id,displayName,userPrincipalName,mail`),
    call<{ isArchived: boolean }>(tok, `/teams/${id}?$select=id,isArchived`).catch(() => null)
  ])
  return {
    id,
    displayName: g.displayName,
    description: g.description,
    visibility: (g.visibility || 'private').toLowerCase(),
    mail: g.mail,
    created: g.createdDateTime,
    owners: owners.map(person),
    archived: s ? s.isArchived : null
  }
}

export const team = async (tid: string, id: string) => getTeam(await token(tid), id)

type ReportRow = {
  teamId: string
  lastActivityDate: string | null
  reportRefreshDate?: string
  isDeleted?: boolean
  details?: { reportPeriod?: string; activeUsers?: number; guests?: number; channelMessages?: number; postMessages?: number; replyMessages?: number }[]
}

/** Last activity per team from the Teams usage report (90 days, a day or two behind). */
export async function teamActivity(tid: string, teamIds: string[]): Promise<TeamActivityReport> {
  try {
    const tok = await tokenForScopes(tid, REPORT_SCOPES)
    const rows = (await all<ReportRow>(tok, "https://graph.microsoft.com/beta/reports/getTeamsTeamActivityDetail(period='D90')?$format=application/json")).filter(
      (r) => !r.isDeleted
    )
    const known = new Set(teamIds)
    // With "conceal user, group and site names in reports" on, ids come back obfuscated and match nothing.
    if (rows.length && !rows.some((r) => known.has(r.teamId))) return { rows: [], reportDate: null, issue: 'concealed' }
    return {
      reportDate: rows[0]?.reportRefreshDate || null,
      rows: rows.map((r): TeamActivity => {
        const d = r.details?.[0]
        return {
          teamId: r.teamId,
          lastActivity: r.lastActivityDate,
          activeUsers: d?.activeUsers ?? null,
          guests: d?.guests ?? null,
          messages: d ? (d.channelMessages ?? 0) + (d.postMessages ?? 0) + (d.replyMessages ?? 0) : null
        }
      })
    }
  } catch (e) {
    const issue = e instanceof NeedsConsent ? 'consent' : e instanceof GraphError && e.status === 403 ? 'denied' : 'error'
    return { rows: [], reportDate: null, issue }
  }
}

// ---------- members ----------

type RawMember = { id: string; userId: string; displayName: string | null; email: string | null; roles: string[] }
const member = (m: RawMember): TeamMember => ({ id: m.id, userId: m.userId, displayName: m.displayName || m.email || m.userId, email: m.email, roles: m.roles || [] })
const memberBody = (user: string, owner: boolean) => ({
  '@odata.type': '#microsoft.graph.aadUserConversationMember',
  roles: owner ? ['owner'] : [],
  'user@odata.bind': userRef(user)
})

export async function members(tid: string, teamId: string): Promise<TeamMember[]> {
  const tok = await token(tid)
  return (await all<RawMember>(tok, `/teams/${teamId}/members`)).map(member).sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function addMember(tid: string, teamId: string, user: string, owner: boolean) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/members`, { method: 'POST', body: JSON.stringify(memberBody(user, owner)) })
  return members(tid, teamId)
}

export async function setMemberRole(tid: string, teamId: string, membershipId: string, owner: boolean) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/members/${membershipId}`, {
    method: 'PATCH',
    body: JSON.stringify({ '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: owner ? ['owner'] : [] })
  })
  return members(tid, teamId)
}

export async function removeMember(tid: string, teamId: string, membershipId: string) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/members/${membershipId}`, { method: 'DELETE' })
  return members(tid, teamId)
}

/** Adds several users in one call; returns the ones Teams refused with the reason. */
async function addMany(tok: string, teamId: string, users: { user: string; owner: boolean }[]) {
  const failed: { what: string; error: string }[] = []
  for (let i = 0; i < users.length; i += 200) {
    const part = users.slice(i, i + 200)
    try {
      const r = await call<{ value: { error: { message: string } | null; userId: string }[] }>(tok, `/teams/${teamId}/members/add`, {
        method: 'POST',
        body: JSON.stringify({ values: part.map((u) => memberBody(u.user, u.owner)) })
      })
      r.value.forEach((x, n) => x.error && failed.push({ what: part[n]?.user || x.userId, error: x.error.message }))
    } catch (e) {
      part.forEach((u) => failed.push({ what: u.user, error: (e as Error).message }))
    }
  }
  return failed
}

// ---------- channels ----------

type RawChannel = { id: string; displayName: string; description: string | null; membershipType: string | null; createdDateTime: string | null }
const channel = (c: RawChannel): TeamChannel => ({
  id: c.id,
  displayName: c.displayName,
  description: c.description,
  type: (c.membershipType === 'private' || c.membershipType === 'shared' ? c.membershipType : 'standard') as ChannelType,
  created: c.createdDateTime
})

export async function channels(tid: string, teamId: string): Promise<TeamChannel[]> {
  const tok = await token(tid)
  const list = (await all<RawChannel>(tok, `/teams/${teamId}/channels`)).map(channel)
  // General first, the rest by name, the way Teams itself lists them.
  return list.sort((a, b) => Number(b.displayName === 'General') - Number(a.displayName === 'General') || a.displayName.localeCompare(b.displayName))
}

/**
 * Private and shared channels need an owner when they are created. The first owner goes in the create call;
 * the other owners are added right after, so a failure there leaves a usable channel.
 */
async function createChannel(tok: string, teamId: string, c: TemplateChannel, owners: string[]) {
  const special = c.type !== 'standard'
  if (special && !owners.length) throw new Error('A private or shared channel needs an owner')
  const body: Record<string, unknown> = { displayName: c.name, description: c.description || undefined, membershipType: c.type }
  if (special) body.members = [memberBody(owners[0], true)]
  let created: RawChannel | undefined
  // Right after a team is created, its channels endpoint can answer 404 for a short while.
  for (let attempt = 0; !created; attempt++) {
    try {
      created = await call<RawChannel>(tok, `/teams/${teamId}/channels`, { method: 'POST', body: JSON.stringify(body) })
    } catch (e) {
      if (!(e instanceof GraphError) || e.status !== 404 || attempt >= 5) throw e
      await sleep(5000)
    }
  }
  const failed: { what: string; error: string }[] = []
  if (special)
    for (const o of owners.slice(1))
      try {
        await call(tok, `/teams/${teamId}/channels/${created.id}/members`, { method: 'POST', body: JSON.stringify(memberBody(o, true)) })
      } catch (e) {
        failed.push({ what: `${c.name}: ${o}`, error: (e as Error).message })
      }
  return { channel: channel(created), failed }
}

export async function addChannel(tid: string, teamId: string, c: TemplateChannel, owners: string[]) {
  return createChannel(await token(tid), teamId, c, owners)
}

export async function removeChannel(tid: string, teamId: string, channelId: string) {
  await call(await token(tid), `/teams/${teamId}/channels/${channelId}`, { method: 'DELETE' })
}

export async function channelMembers(tid: string, teamId: string, channelId: string): Promise<TeamMember[]> {
  const tok = await token(tid)
  return (await all<RawMember>(tok, `/teams/${teamId}/channels/${channelId}/members`)).map(member).sort((a, b) => a.displayName.localeCompare(b.displayName))
}

export async function addChannelMember(tid: string, teamId: string, channelId: string, user: string, owner: boolean) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/channels/${channelId}/members`, { method: 'POST', body: JSON.stringify(memberBody(user, owner)) })
  return channelMembers(tid, teamId, channelId)
}

export async function setChannelMemberRole(tid: string, teamId: string, channelId: string, membershipId: string, owner: boolean) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/channels/${channelId}/members/${membershipId}`, {
    method: 'PATCH',
    body: JSON.stringify({ '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: owner ? ['owner'] : [] })
  })
  return channelMembers(tid, teamId, channelId)
}

export async function removeChannelMember(tid: string, teamId: string, channelId: string, membershipId: string) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/channels/${channelId}/members/${membershipId}`, { method: 'DELETE' })
  return channelMembers(tid, teamId, channelId)
}

// ---------- create, archive ----------

/** POST /teams answers 202 with the new team's id in Content-Location and an operation to poll. */
export async function createTeam(tid: string, t: NewTeam): Promise<NewTeamResult> {
  if (!t.owners.length) throw new Error('A team needs at least one owner')
  const tok = await token(tid)
  const res = await fetch(`${GRAPH}/teams`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${tok}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      'template@odata.bind': `${GRAPH}/teamsTemplates('standard')`,
      displayName: t.displayName,
      description: t.description || undefined,
      visibility: t.visibility,
      members: [memberBody(t.owners[0], true)],
      channels: t.channels
        .filter((c) => c.type === 'standard' && c.name.trim().toLowerCase() !== 'general')
        .map((c) => ({ displayName: c.name, description: c.description || undefined }))
    })
  })
  if (res.status !== 202 && res.status !== 201) {
    const body = await res.json().catch(() => ({}))
    throw new GraphError(body?.error?.message || `Graph ${res.status}`, res.status, body?.error?.code)
  }
  const loc = res.headers.get('Content-Location') || res.headers.get('Location') || ''
  const teamId = /teams\('([^']+)'\)/.exec(loc)?.[1]
  if (!teamId) throw new Error('Teams did not return the new team id')
  const op = res.headers.get('Location')
  if (op && /operations/.test(op)) {
    for (let i = 0; i < 60; i++) {
      await sleep(i < 5 ? 2000 : 5000)
      const s = await call<{ status: string; error?: { message: string } | null }>(tok, op.startsWith('http') ? op : GRAPH + op).catch(() => null)
      if (s?.status === 'succeeded') break
      if (s?.status === 'failed') throw new Error(s.error?.message || 'Creating the team failed')
    }
  }
  const failed = await addMany(tok, teamId, [...t.owners.slice(1).map((user) => ({ user, owner: true })), ...t.members.map((user) => ({ user, owner: false }))])
  for (const c of t.channels.filter((c) => c.type !== 'standard'))
    try {
      failed.push(...(await createChannel(tok, teamId, c, t.owners)).failed)
    } catch (e) {
      failed.push({ what: c.name, error: (e as Error).message })
    }
  return { team: await getTeam(tok, teamId), failed }
}

/** Archiving makes the team read-only; optionally the SharePoint site becomes read-only for members too. */
export async function archive(tid: string, teamId: string, archived: boolean, siteReadOnly = false) {
  const tok = await token(tid)
  await call(tok, `/teams/${teamId}/${archived ? 'archive' : 'unarchive'}`, {
    method: 'POST',
    body: archived ? JSON.stringify({ shouldSetSpoSiteReadOnlyForMembers: siteReadOnly }) : undefined
  })
  // The archive operation runs asynchronously; report the new state without waiting for it.
  return { ...(await getTeam(tok, teamId)), archived }
}

export async function updateTeam(tid: string, teamId: string, p: { displayName?: string; description?: string; visibility?: 'private' | 'public' }) {
  const tok = await token(tid)
  await call(tok, `/groups/${teamId}`, { method: 'PATCH', body: JSON.stringify(p) })
  return getTeam(tok, teamId)
}

// ---------- guests ----------

type RawGuest = RawPerson & { userType?: string; createdDateTime?: string | null; externalUserState?: string | null; '@odata.type'?: string }

/** All guests in all teams, from the group members of each team, plus their last sign-in where available. */
export async function guests(tid: string, teams: { id: string; displayName: string }[]): Promise<TeamGuest[]> {
  const tok = await tokenFor(tid)
  const sel = 'id,displayName,mail,userPrincipalName,userType,createdDateTime,externalUserState'
  const res = await batch(
    tok,
    teams.map((t) => ({ id: t.id, url: `/groups/${t.id}/members/microsoft.graph.user?$select=${sel}&$top=999` }))
  )
  const out: TeamGuest[] = []
  for (const t of teams) {
    const r = res.get(t.id)
    if (r?.status !== 200) continue
    const next = r.body['@odata.nextLink'] as string | undefined
    const list = [...((r.body.value || []) as RawGuest[]), ...(next ? await all<RawGuest>(tok, next) : [])]
    for (const u of list)
      if (u.userType === 'Guest')
        out.push({
          teamId: t.id,
          teamName: t.displayName,
          userId: u.id,
          displayName: u.displayName || u.mail || u.id,
          mail: u.mail || null,
          upn: u.userPrincipalName || '',
          created: u.createdDateTime || null,
          state: u.externalUserState || null,
          lastSignIn: null
        })
  }
  // signInActivity needs Entra ID P1 and AuditLog.Read.All; without it the column stays empty.
  try {
    const seen = await all<{ id: string; signInActivity?: { lastSignInDateTime: string | null } }>(
      tok,
      "/users?$filter=userType eq 'Guest'&$select=id,signInActivity&$top=999"
    )
    const last = new Map(seen.map((u) => [u.id, u.signInActivity?.lastSignInDateTime ?? null]))
    for (const g of out) g.lastSignIn = last.get(g.userId) ?? null
  } catch {
    /* no sign-in data */
  }
  return out
}

export async function removeGuest(tid: string, teamId: string, userId: string) {
  await call(await tokenFor(tid), `/groups/${teamId}/members/${userId}/$ref`, { method: 'DELETE' })
}

/** Deletes the guest account itself (it goes to the recycle bin for 30 days). */
export async function deleteGuest(tid: string, userId: string) {
  const tok = await tokenFor(tid)
  const u = await call<{ userType: string }>(tok, `/users/${userId}?$select=userType`)
  if (u.userType !== 'Guest') throw new Error('Only guest accounts can be deleted here')
  await call(tok, `/users/${userId}`, { method: 'DELETE' })
}

// ---------- PowerShell: policies, access, phone ----------

const helper = psHelper({
  name: 'teams',
  source: helperSource,
  async connect(tenantId) {
    const { graph, admin } = await teamsTokens(tenantId)
    return { key: graph + admin, args: { GraphToken: graph, AdminToken: admin } }
  }
})

export const { ensureModule, closeSession, closeAll } = helper
const run = helper.run

export const policies = (tid: string) => run<TeamsPolicies>(tid, 'Tb-Policies')
export const grantUsers = (tid: string, type: PolicyType, policy: string | null, users: string[]) =>
  run<{ users: CsUser[]; failed: { user: string; error: string }[] }>(tid, 'Tb-GrantUsers', { Type: type, Policy: policy || '', Users: users })
export const grantGroup = (tid: string, a: GroupPolicyAssignment) => run<void>(tid, 'Tb-GrantGroup', { Type: a.type, Policy: a.policy, GroupId: a.groupId, Rank: a.rank })
export const removeGroupPolicy = (tid: string, type: PolicyType, groupId: string) => run<void>(tid, 'Tb-RemoveGroupPolicy', { Type: type, GroupId: groupId })

export const access = (tid: string) => run<TeamsAccess>(tid, 'Tb-Access')
export const setFederation = (tid: string, f: TeamsFederation) =>
  run<TeamsAccess>(tid, 'Tb-SetFederation', {
    AllowFederated: f.allowFederated,
    AllowAll: f.allowAll,
    Allowed: f.allowed,
    Blocked: f.blocked,
    Consumer: f.consumer,
    ConsumerInbound: f.consumerInbound
  })
export const setGuest = (tid: string, g: Partial<TeamsGuestSettings>) => run<TeamsAccess>(tid, 'Tb-SetGuest', { Params: g })

export const phone = (tid: string) => run<TeamsPhone>(tid, 'Tb-Phone')
export const assignNumber = (tid: string, user: string, number: string, type: NumberType, locationId?: string) =>
  run<CsUser>(tid, 'Tb-AssignNumber', { User: user, Number: number, Type: type, LocationId: locationId || '' })
export const unassignNumber = (tid: string, user: string, number: string, type: NumberType) =>
  run<CsUser>(tid, 'Tb-UnassignNumber', { User: user, Number: number, Type: type })
export const enterpriseVoice = (tid: string, user: string, enabled: boolean) => run<CsUser>(tid, 'Tb-EnterpriseVoice', { User: user, Enabled: enabled })
