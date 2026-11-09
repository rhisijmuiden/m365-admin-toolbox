import type {
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
} from '@shared/types'
import { tokenFor, tokenForScopes } from './auth'
import { all, call } from './graph'
import { run } from './exo'
import * as scc from './scc'

/*
 * Phase 5. Quarantine, the Tenant Allow/Block List, the protection policies, restricted senders, message
 * trace and the audit log run in the Exchange Online helper. Pulling mail out of mailboxes runs a compliance
 * search in the Security & Compliance helper, or uses the Defender remediation API with Plan 2. Risk data
 * and user-reported messages come from Graph, each with its own scopes so one missing licence or consent
 * doesn't block the other screens.
 */
export const RISK_SCOPES = ['IdentityRiskyUser.ReadWrite.All', 'IdentityRiskEvent.Read.All']
export const SUBMISSION_SCOPES = ['ThreatSubmission.Read.All']
export const DEFENDER_SCOPES = ['SecurityAnalyzedMessage.ReadWrite.All']
export const SCOPE_SETS = { risk: RISK_SCOPES, submissions: SUBMISSION_SCOPES, defender: DEFENDER_SCOPES }

const GRAPH = 'https://graph.microsoft.com/v1.0'
const BETA = 'https://graph.microsoft.com/beta'
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

// ---------- quarantine ----------

export const quarantine = (tid: string, f: QuarantineFilter) =>
  run<QuarantineMsg[]>(tid, 'Tb-Quarantine', {
    Sender: f.sender || '',
    Recipient: f.recipient || '',
    Type: f.type || '',
    Subject: f.subject || '',
    Start: f.start || null,
    End: f.end || null,
    Page: f.page || 1
  })
export const quarantinePreview = (tid: string, id: string) => run<QuarantinePreview>(tid, 'Tb-QuarantinePreview', { Identity: id })
export const quarantineRelease = (tid: string, ids: string[], allowSender: boolean, reportFalsePositive: boolean) =>
  run<{ failed: { id: string; error: string }[] }>(tid, 'Tb-QuarantineRelease', { Ids: ids, AllowSender: allowSender, ReportFalsePositive: reportFalsePositive })
export const quarantineDelete = (tid: string, ids: string[]) => run<{ failed: { id: string; error: string }[] }>(tid, 'Tb-QuarantineDelete', { Ids: ids })

// ---------- tenant allow/block list ----------

export const allowBlock = (tid: string) => run<AllowBlockItem[]>(tid, 'Tb-AllowBlock')
export const newAllowBlock = (tid: string, n: NewAllowBlock) =>
  run<AllowBlockItem[]>(tid, 'Tb-NewAllowBlock', { ListType: n.listType, Entries: n.entries, Action: n.action, Days: n.days || null, Notes: n.notes || '' })
export const removeAllowBlock = (tid: string, listType: AllowBlockList, ids: string[]) => run<AllowBlockItem[]>(tid, 'Tb-RemoveAllowBlock', { ListType: listType, Ids: ids })

// ---------- policies ----------

export const policies = (tid: string) => run<SecPolicies>(tid, 'Tb-SecPolicies')
export const setPolicy = (tid: string, kind: SecPolicyKind, name: string, params: Record<string, unknown>) =>
  run<SecPolicies>(tid, 'Tb-SetSecPolicy', { Kind: kind, Identity: name, Params: params })
export const setPolicyRule = (tid: string, kind: SecPolicyKind, rule: string, enabled: boolean) => run<SecPolicies>(tid, 'Tb-SetSecRule', { Kind: kind, Rule: rule, Enabled: enabled })

// ---------- restricted senders, message trace, audit ----------

export const restricted = (tid: string) => run<RestrictedSender[]>(tid, 'Tb-RestrictedSenders')
export const releaseSender = (tid: string, s: RestrictedSender) => run<RestrictedSender[]>(tid, 'Tb-ReleaseSender', { Sender: s.sender, Reason: s.reason })

export const trace = (tid: string, f: TraceFilter) =>
  run<TraceRow[]>(tid, 'Tb-MessageTrace', { Sender: f.sender || '', Recipient: f.recipient || '', Subject: f.subject || '', Status: f.status || '', Start: f.start, End: f.end })
export const traceDetail = (tid: string, traceId: string, recipient: string) => run<TraceEvent[]>(tid, 'Tb-MessageTraceDetail', { TraceId: traceId, Recipient: recipient })

export const auditStatus = (tid: string) => run<{ enabled: boolean }>(tid, 'Tb-AuditStatus')
export const enableAudit = (tid: string) => run<{ enabled: boolean }>(tid, 'Tb-EnableAudit')
export const audit = (tid: string, f: AuditFilter) =>
  run<AuditRecord[]>(tid, 'Tb-Audit', {
    Start: f.start,
    End: f.end,
    Users: f.users || [],
    Operations: f.operations || [],
    RecordType: f.recordType || '',
    FreeText: f.freeText || ''
  })

// ---------- compromised account ----------

export const secureMailbox = (tid: string, mailbox: string, rules: boolean, forwarding: boolean) =>
  run<SecureMailboxResult>(tid, 'Tb-SecureMailbox', { Identity: mailbox, Rules: rules, Forwarding: forwarding })

type GraphSignIn = {
  id: string
  createdDateTime: string
  appDisplayName: string
  ipAddress: string
  clientAppUsed: string
  isInteractive: boolean
  riskLevelDuringSignIn: string
  conditionalAccessStatus: string
  location?: { city?: string; countryOrRegion?: string }
  deviceDetail?: { browser?: string; operatingSystem?: string }
  status?: { errorCode: number; failureReason?: string }
}

const place = (l?: { city?: string | null; countryOrRegion?: string | null }) => [l?.city, l?.countryOrRegion].filter(Boolean).join(', ')

/** The latest sign-ins of one user. Needs Entra ID P1 for the sign-in log. */
export async function signIns(tid: string, userId: string): Promise<SignIn[]> {
  const token = await tokenFor(tid)
  const r = await call<{ value: GraphSignIn[] }>(token, `/auditLogs/signIns?$filter=userId eq '${userId}'&$top=30`)
  return r.value.map((s) => ({
    id: s.id,
    date: s.createdDateTime,
    app: s.appDisplayName,
    ip: s.ipAddress,
    location: place(s.location),
    client: s.clientAppUsed,
    device: [s.deviceDetail?.operatingSystem, s.deviceDetail?.browser].filter(Boolean).join(' · '),
    ok: (s.status?.errorCode ?? 0) === 0,
    error: s.status?.failureReason || '',
    risk: s.riskLevelDuringSignIn,
    ca: s.conditionalAccessStatus,
    interactive: s.isInteractive
  }))
}

// ---------- risky users (Entra ID P2) ----------

export async function riskyUsers(tid: string): Promise<RiskyUser[]> {
  const token = await tokenForScopes(tid, RISK_SCOPES)
  const rows = await all<{ id: string; userDisplayName: string; userPrincipalName: string; riskLevel: string; riskState: string; riskDetail: string; riskLastUpdatedDateTime: string }>(
    token,
    `${GRAPH}/identityProtection/riskyUsers?$top=500`
  )
  return rows.map((u) => ({
    id: u.id,
    displayName: u.userDisplayName,
    upn: u.userPrincipalName,
    riskLevel: u.riskLevel,
    riskState: u.riskState,
    riskDetail: u.riskDetail,
    updated: u.riskLastUpdatedDateTime
  }))
}

export async function riskDetections(tid: string, days: number): Promise<RiskDetection[]> {
  const token = await tokenForScopes(tid, RISK_SCOPES)
  const since = new Date(Date.now() - days * 86400000).toISOString()
  const rows = await all<{
    id: string
    riskEventType: string
    riskLevel: string
    riskState: string
    detectedDateTime: string
    userDisplayName: string
    userPrincipalName: string
    userId: string
    ipAddress: string
    activity: string
    location?: { city?: string; countryOrRegion?: string }
  }>(token, `${GRAPH}/identityProtection/riskDetections?$filter=detectedDateTime ge ${since}&$top=500`)
  return rows
    .map((d) => ({
      id: d.id,
      type: d.riskEventType,
      level: d.riskLevel,
      state: d.riskState,
      date: d.detectedDateTime,
      user: d.userDisplayName,
      upn: d.userPrincipalName,
      userId: d.userId,
      ip: d.ipAddress,
      location: place(d.location),
      activity: d.activity
    }))
    .sort((a, b) => (b.date || '').localeCompare(a.date || ''))
}

export async function setRisk(tid: string, userIds: string[], action: 'dismiss' | 'confirmCompromised') {
  const token = await tokenForScopes(tid, RISK_SCOPES)
  await call(token, `${GRAPH}/identityProtection/riskyUsers/${action}`, { method: 'POST', body: JSON.stringify({ userIds }) })
}

// ---------- user-reported messages ----------

export async function reported(tid: string, days: number): Promise<ReportedMsg[]> {
  const token = await tokenForScopes(tid, SUBMISSION_SCOPES)
  const since = new Date(Date.now() - days * 86400000).toISOString()
  const rows = await all<{
    id: string
    createdDateTime: string
    category: string
    source: string
    status: string
    sender?: string
    subject?: string
    recipientEmailAddress: string
    receivedDateTime?: string
    internetMessageId?: string
    createdBy?: { user?: { email?: string; displayName?: string } }
    result?: { category?: string; detail?: string }
  }>(token, `${BETA}/security/threatSubmission/emailThreats?$filter=createdDateTime ge ${since}`)
  return rows
    .map((r) => ({
      id: r.id,
      created: r.createdDateTime,
      category: r.category,
      source: r.source,
      status: r.status,
      sender: r.sender || '',
      subject: r.subject || '',
      recipient: r.recipientEmailAddress,
      received: r.receivedDateTime || null,
      reportedBy: r.createdBy?.user?.email || r.createdBy?.user?.displayName || '',
      result: [r.result?.category, r.result?.detail].filter(Boolean).join(' · '),
      internetMessageId: r.internetMessageId || ''
    }))
    .sort((a, b) => (b.created || '').localeCompare(a.created || ''))
}

// ---------- pull mail out of mailboxes: compliance search ----------

const jobs = new Map<string, RetractJob>()
/** A purge action removes at most 10 items per mailbox, so it is repeated; this caps it at 100 per mailbox. */
const MAX_ROUNDS = 10
const SEARCH_TIMEOUT = 60 * 60000

const kql = (s: string) => s.replace(/["\\]/g, '').trim()

/** KQL for the compliance search. A sender, subject or query is required so a search can never match everything. */
export function retractQuery(s: RetractSpec) {
  const parts: string[] = []
  if (s.sender?.trim()) parts.push(`from:"${kql(s.sender)}"`)
  if (s.subject?.trim()) parts.push(`subject:"${kql(s.subject)}"`)
  if (s.query?.trim()) parts.push(`(${s.query.trim()})`)
  if (!parts.length) throw new Error('Enter at least a sender, subject or query')
  if (s.start) parts.push(`received>=${s.start.slice(0, 10)}`)
  if (s.end) parts.push(`received<=${s.end.slice(0, 10)}`)
  return parts.join(' AND ')
}

type SearchStatus = { status: string; items: number; mailboxes: RetractJob['mailboxes']; errors: string }

async function waitSearch(job: RetractJob) {
  const until = Date.now() + SEARCH_TIMEOUT
  for (;;) {
    await sleep(8000)
    const s = await scc.run<SearchStatus>(job.tenantId, 'Tb-SearchStatus', { Name: job.searchName })
    if (s.status === 'Completed') {
      job.items = s.items || 0
      job.mailboxes = (s.mailboxes || []).sort((a, b) => b.items - a.items)
      return
    }
    if (/Failed|Stopped/i.test(s.status)) throw new Error(s.errors || `Search ${s.status}`)
    if (Date.now() > until) throw new Error('The search did not finish within an hour')
  }
}

async function waitAction(job: RetractJob, action: string) {
  const until = Date.now() + SEARCH_TIMEOUT
  for (;;) {
    await sleep(10000)
    const a = await scc.run<{ status: string; errors: string }>(job.tenantId, 'Tb-ActionStatus', { Action: action })
    if (a.status === 'Completed') {
      if (a.errors) throw new Error(a.errors)
      return
    }
    if (/Failed|Stopped/i.test(a.status)) throw new Error(a.errors || `Purge ${a.status}`)
    if (Date.now() > until) throw new Error('The purge did not finish within an hour')
  }
}

const fail = (job: RetractJob, e: unknown) => {
  job.phase = 'failed'
  job.error = (e as Error).message
  job.finished = new Date().toISOString()
}

export const retractJobs = (tid: string) => [...jobs.values()].filter((j) => j.tenantId === tid).sort((a, b) => b.started.localeCompare(a.started))

export async function startRetract(tid: string, spec: RetractSpec): Promise<RetractJob> {
  const query = retractQuery(spec)
  const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14)
  const job: RetractJob = {
    id: `${stamp}-${Math.random().toString(36).slice(2, 6)}`,
    tenantId: tid,
    searchName: '',
    spec,
    query,
    phase: 'searching',
    items: 0,
    mailboxes: [],
    purgeType: null,
    round: 0,
    rounds: 0,
    error: null,
    started: new Date().toISOString(),
    finished: null
  }
  job.searchName = `Toolbox-${job.id}`
  await scc.run(tid, 'Tb-StartSearch', { Name: job.searchName, Query: query, Locations: spec.mailboxes || [] })
  jobs.set(job.id, job)
  waitSearch(job).then(
    () => (job.phase = 'found'),
    (e) => fail(job, e)
  )
  return job
}

export function purgeRetract(jobId: string, type: PurgeType): RetractJob {
  const job = jobs.get(jobId)
  if (!job || job.phase !== 'found') throw new Error('This search is not ready to delete')
  if (!job.items) throw new Error('Nothing to delete')
  const most = Math.max(1, ...job.mailboxes.map((m) => m.items))
  job.purgeType = type
  job.rounds = Math.min(MAX_ROUNDS, Math.ceil(most / 10))
  job.round = 0
  job.phase = 'purging'
  ;(async () => {
    try {
      for (let r = 1; r <= job.rounds; r++) {
        job.round = r
        // Each round re-runs the search, so it acts on what is still there.
        if (r > 1) {
          await scc.run(job.tenantId, 'Tb-RestartSearch', { Name: job.searchName })
          await waitSearch(job)
        }
        const { action } = await scc.run<{ action: string }>(job.tenantId, 'Tb-Purge', { Name: job.searchName, Type: type })
        await waitAction(job, action)
      }
      job.phase = 'done'
      job.finished = new Date().toISOString()
      await scc.run(job.tenantId, 'Tb-RemoveSearch', { Name: job.searchName }).catch(() => undefined)
    } catch (e) {
      fail(job, e)
    }
  })()
  return job
}

/** Forgets a search and removes it from the tenant (purges keep running on Microsoft's side once started). */
export async function discardRetract(jobId: string) {
  const job = jobs.get(jobId)
  if (!job) return
  if (job.phase === 'searching' || job.phase === 'purging') throw new Error('Wait until this search or purge has finished')
  jobs.delete(jobId)
  if (job.phase !== 'done') await scc.run(job.tenantId, 'Tb-RemoveSearch', { Name: job.searchName }).catch(() => undefined)
}

// ---------- pull mail out of mailboxes: Defender for Office 365 Plan 2 ----------

type Analyzed = {
  id: string
  networkMessageId: string
  loggedDateTime: string
  subject: string
  recipientEmailAddress: string
  directionality: string
  threatTypes?: string[]
  latestDeliveryLocation?: string
  senderDetail?: { fromAddress?: string; mailFromAddress?: string }
}

/** Messages Defender analyzed in the period, filtered on sender and subject here. Plan 2 only. */
export async function analyzedEmails(tid: string, spec: RetractSpec): Promise<AnalyzedEmail[]> {
  const token = await tokenForScopes(tid, DEFENDER_SCOPES)
  const end = spec.end ? new Date(spec.end) : new Date()
  const start = spec.start ? new Date(spec.start) : new Date(end.getTime() - 7 * 86400000)
  const sender = spec.sender?.trim().toLowerCase()
  const subject = spec.subject?.trim().toLowerCase()
  if (!sender && !subject) throw new Error('Enter at least a sender or subject')
  const mailboxes = new Set((spec.mailboxes || []).map((m) => m.toLowerCase()))
  const out: AnalyzedEmail[] = []
  let next: string | undefined = `${BETA}/security/collaboration/analyzedEmails?startTime=${start.toISOString()}&endTime=${end.toISOString()}`
  for (let page = 0; next && page < 40; page++) {
    const r: { value: Analyzed[]; '@odata.nextLink'?: string } = await call(token, next)
    for (const m of r.value) {
      const from = (m.senderDetail?.fromAddress || m.senderDetail?.mailFromAddress || '').toLowerCase()
      if (sender && !from.includes(sender)) continue
      if (subject && !(m.subject || '').toLowerCase().includes(subject)) continue
      if (mailboxes.size && !mailboxes.has((m.recipientEmailAddress || '').toLowerCase())) continue
      out.push({
        id: m.id,
        networkMessageId: m.networkMessageId,
        received: m.loggedDateTime,
        sender: from,
        subject: m.subject,
        recipient: m.recipientEmailAddress,
        direction: m.directionality,
        threats: m.threatTypes || [],
        location: m.latestDeliveryLocation || ''
      })
    }
    next = r['@odata.nextLink']
  }
  return out
}

export async function remediate(tid: string, emails: Pick<AnalyzedEmail, 'networkMessageId' | 'recipient'>[], action: DefenderAction, name: string) {
  const token = await tokenForScopes(tid, DEFENDER_SCOPES)
  await call(token, `${BETA}/security/collaboration/analyzedEmails/remediate`, {
    method: 'POST',
    body: JSON.stringify({
      displayName: name,
      description: 'Started from M365 Admin Toolbox',
      severity: 'medium',
      action,
      remediateSendersCopy: false,
      analyzedEmails: emails.map((e) => ({ networkMessageId: e.networkMessageId, recipientEmailAddress: e.recipient }))
    })
  })
}

export const closeAll = () => scc.closeAll()
export const closeSession = (tid: string) => scc.closeSession(tid)
