import { useState } from 'react'
import type { MemberGroup } from '@shared/types'
import { genPassword, unwrap, useApp } from '../ctx'
import { useUsers } from '../components/teams'
import { StepList, TenantScope, TenantSelect, useSteps, type Step } from '../components/cross'

/** Leaving employee: block, sign out, hand over the mailbox, convert it to shared and free the license. */
export function OffboardView() {
  const { t, tenant, tenants } = useApp()
  const [tid, setTid] = useState(tenant!.tenantId)
  const x = tenants.find((y) => y.tenantId === tid) || tenant!
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('offboard')}</h1>
        </div>
        <div className="top-r">
          <label className="hint" htmlFor="ofT">
            {t('scTenant')}
          </label>
          <TenantSelect id="ofT" value={tid} onChange={setTid} />
        </div>
      </header>
      <TenantScope tenant={x}>
        <OffboardForm key={tid} />
      </TenantScope>
    </>
  )
}

type Who = { id: string; displayName: string; userPrincipalName: string; skuIds: string[] }
type Opt = 'block' | 'revoke' | 'password' | 'autoReply' | 'forward' | 'access' | 'hide' | 'groups' | 'shared' | 'licenses'
const DEFAULTS: Record<Opt, boolean> = {
  block: true,
  revoke: true,
  password: true,
  autoReply: false,
  forward: false,
  access: false,
  hide: false,
  groups: true,
  shared: true,
  licenses: true
}

function OffboardForm() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const users = useUsers(tid)
  const [pick, setPick] = useState('')
  const [who, setWho] = useState<Who | null>(null)
  const [groups, setGroups] = useState<MemberGroup[]>([])
  const [finding, setFinding] = useState(false)
  const [opt, setOpt] = useState(DEFAULTS)
  const [keep, setKeep] = useState<Set<string>>(new Set())
  const [reply, setReply] = useState('')
  const [fwd, setFwd] = useState('')
  const [delegate, setDelegate] = useState('')
  const steps = useSteps()

  async function find() {
    setFinding(true)
    try {
      const u = await unwrap(window.api.cross.findUser(tid, pick.trim()))
      setWho(u)
      setReply(t('ofReplyDefault', u.displayName))
      steps.reset()
      setGroups(await unwrap(window.api.cross.userGroups(tid, u.id)).catch(() => []))
      setKeep(new Set())
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setFinding(false)
    }
  }

  const on = (k: Opt) => opt[k]
  const set = (k: Opt, v: boolean) => setOpt((o) => ({ ...o, [k]: v }))
  const removeGroups = groups.filter((g) => !g.readOnly && !keep.has(g.id))

  const plan: Step[] = []
  if (who) {
    const upn = who.userPrincipalName
    if (on('block')) plan.push({ key: 'block', label: t('ofStBlock'), run: () => unwrap(window.api.setSignIn(tid, who.id, false)) })
    if (on('revoke')) plan.push({ key: 'revoke', label: t('ofStRevoke'), run: () => unwrap(window.api.revokeSessions(tid, who.id)) })
    if (on('password'))
      plan.push({ key: 'password', label: t('ofStPassword'), run: () => unwrap(window.api.resetPassword(tid, who.id, genPassword(20), false)) })
    if (on('autoReply') && reply.trim())
      plan.push({
        key: 'autoReply',
        label: t('ofStReply'),
        run: async () =>
          void (await unwrap(
            window.api.exo.setAutoReply(tid, upn, { state: 'Enabled', start: null, end: null, internalMessage: reply, externalMessage: reply, externalAudience: 'All' })
          ))
      })
    if (on('forward') && fwd.trim())
      plan.push({
        key: 'forward',
        label: t('ofStForward', fwd.trim()),
        run: async () => void (await unwrap(window.api.exo.setMailbox(tid, upn, { ForwardingAddress: fwd.trim(), DeliverToMailboxAndForward: true })))
      })
    if (on('access') && delegate.trim())
      plan.push({
        key: 'access',
        label: t('ofStAccess', delegate.trim()),
        run: async () => void (await unwrap(window.api.exo.grant(tid, upn, delegate.trim(), 'FullAccess', true)))
      })
    if (on('hide'))
      plan.push({ key: 'hide', label: t('ofStHide'), run: async () => void (await unwrap(window.api.exo.setMailbox(tid, upn, { HiddenFromAddressListsEnabled: true }))) })
    if (on('groups') && removeGroups.length)
      plan.push({
        key: 'groups',
        label: t('ofStGroups', removeGroups.length),
        run: async () => {
          const r = await unwrap(window.api.cross.removeFromGroups(tid, who.id, upn, removeGroups))
          if (r.failed.length) throw new Error(r.failed.map((f) => `${f.id}: ${f.error}`).join(' · '))
        }
      })
    // Converting first keeps the mailbox: a user mailbox without a license is deleted after 30 days.
    if (on('shared'))
      plan.push({
        key: 'shared',
        label: on('licenses') ? t('ofStSharedLic') : t('ofStShared'),
        run: async () => {
          const r = await unwrap(window.api.exo.convertToShared(tid, upn, who.id, on('licenses')))
          return on('licenses') ? t('ofLicRemoved', r.licensesRemoved) : undefined
        }
      })
    else if (on('licenses'))
      plan.push({
        key: 'licenses',
        label: t('ofStLicenses'),
        run: async () => t('ofLicRemoved', (await unwrap(window.api.cross.removeLicenses(tid, who.id))).removed)
      })
  }

  const started = Object.keys(steps.status).length > 0
  const failed = plan.filter((s) => steps.status[s.key]?.state === 'fail')

  function start() {
    if (!who || !confirm(t('ofConfirm', who.displayName, tenant!.name))) return
    steps.run(plan)
  }

  const box = (k: Opt, label: string, extra?: React.ReactNode) => (
    <div>
      <label className={`opt ${opt[k] ? 'on' : ''}`} style={{ padding: '8px 10px' }}>
        <input type="checkbox" className="chk" checked={opt[k]} disabled={steps.running} onChange={(e) => set(k, e.target.checked)} />
        <div>{label}</div>
      </label>
      {opt[k] && extra && <div style={{ padding: '8px 0 0 30px' }}>{extra}</div>}
    </div>
  )

  return (
    <div className="content stack">
      <div className="panel" style={{ padding: 14, display: 'grid', gap: 8 }}>
        <b>{t('ofWho')}</b>
        <div className="pw">
          <input
            list="ofUsers"
            className="mono"
            placeholder={t('ofWhoPh')}
            value={pick}
            onChange={(e) => setPick(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && pick.trim() && find()}
          />
          <datalist id="ofUsers">
            {users.slice(0, 1000).map((u) => (
              <option key={u.id} value={u.userPrincipalName}>
                {u.displayName}
              </option>
            ))}
          </datalist>
          <button className="btn sm" disabled={!pick.trim() || finding || steps.running} onClick={find}>
            {finding && <span className="spin" />}
            {t('ofFind')}
          </button>
        </div>
      </div>

      {who && (
        <div className="row2" style={{ alignItems: 'start' }}>
          <div className="panel" style={{ padding: 16, display: 'grid', gap: 8 }}>
            <div>
              <div className="u-n">{who.displayName}</div>
              <div className="hint mono">{who.userPrincipalName}</div>
            </div>
            <b style={{ marginTop: 6 }}>{t('ofAccount')}</b>
            {box('block', t('ofBlock'))}
            {box('revoke', t('ofRevoke'))}
            {box('password', t('ofPassword'))}
            <b style={{ marginTop: 6 }}>{t('ofMailbox')}</b>
            {box('autoReply', t('ofReply'), <textarea rows={3} style={{ width: '100%' }} value={reply} onChange={(e) => setReply(e.target.value)} />)}
            {box('forward', t('ofForward'), <AddressInput value={fwd} onChange={setFwd} placeholder={t('ofColleague')} list="ofUsers" />)}
            {box('access', t('ofAccess'), <AddressInput value={delegate} onChange={setDelegate} placeholder={t('ofColleague')} list="ofUsers" />)}
            {box('hide', t('ofHide'))}
            {box('shared', t('ofShared'))}
            {box('licenses', t('ofLicenses', who.skuIds.length))}
            <b style={{ marginTop: 6 }}>{t('groups')}</b>
            {box(
              'groups',
              t('ofGroups', groups.filter((g) => !g.readOnly).length),
              <div style={{ display: 'grid', gap: 4, maxHeight: 220, overflow: 'auto' }}>
                {groups.map((g) => (
                  <label key={g.id} className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <input
                      type="checkbox"
                      className="chk"
                      disabled={g.readOnly}
                      checked={!g.readOnly && !keep.has(g.id)}
                      onChange={() => setKeep((s) => (s.has(g.id) ? new Set([...s].filter((y) => y !== g.id)) : new Set([...s, g.id])))}
                    />
                    {g.displayName}
                    {g.readOnly && <span className="pill off">{t('ofGroupRo')}</span>}
                  </label>
                ))}
                {!groups.length && <span className="hint">{t('ofNoGroups')}</span>}
              </div>
            )}
          </div>

          <div className="panel" style={{ padding: 16, display: 'grid', gap: 12 }}>
            <b>{t('ofPlan')}</b>
            <StepList steps={plan} status={steps.status} />
            {!started ? (
              <button className="btn pri" style={{ justifySelf: 'start' }} disabled={!plan.length} onClick={start}>
                {t('ofRun', plan.length)}
              </button>
            ) : (
              !steps.running &&
              failed.length > 0 && (
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button className="btn" onClick={() => steps.run(plan, (s) => steps.status[s.key]?.state === 'fail')}>
                    {t('obRetry', failed.length)}
                  </button>
                </div>
              )
            )}
            {started && !steps.running && !failed.length && <span className="pill ok" style={{ justifySelf: 'start' }}>{t('ofDone')}</span>}
            <p className="hint">{t('ofHint')}</p>
          </div>
        </div>
      )}
    </div>
  )
}

function AddressInput({ value, onChange, placeholder, list }: { value: string; onChange(v: string): void; placeholder: string; list: string }) {
  return <input list={list} className="mono" style={{ width: '100%' }} placeholder={placeholder} value={value} onChange={(e) => onChange(e.target.value)} />
}
