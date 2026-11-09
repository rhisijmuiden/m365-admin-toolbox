import { useMemo, useState } from 'react'
import type { CsUser, NumberType, PhoneNumber, PolicyType } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { Tabs } from '../components/exo'
import { TeamsLoad } from '../components/teams'

const TYPES: NumberType[] = ['CallingPlan', 'DirectRouting', 'OperatorConnect']
const numberType = (n: PhoneNumber): NumberType => (TYPES.find((x) => x.toLowerCase() === n.type.toLowerCase()) || 'DirectRouting') as NumberType
const lineNumber = (uri: string | null) => (uri || '').replace(/^tel:/i, '')

/** Teams Phone: phone numbers and who has them, Enterprise Voice and the calling and voice routing policies. */
export function TeamPhoneView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.teams.phone(tid), [tid])
  const [tab, setTab] = useState<'numbers' | 'users'>('numbers')
  const [q, setQ] = useState('')
  const [assign, setAssign] = useState<PhoneNumber | null>(null)
  const [busy, setBusy] = useState<string | null>(null)

  const users = data.data?.users || []
  const byId = useMemo(() => new Map(users.map((u) => [u.id, u])), [users])
  const numbers = data.data?.numbers || []
  const free = numbers.filter((n) => !n.assignedTo)
  const without = users.filter((u) => u.phoneSystem && !u.lineUri)
  const match = (s: string) => s.toLowerCase().includes(q.toLowerCase())

  const putUser = (u: CsUser) => data.set((d) => ({ ...d, users: d.users.some((x) => x.id === u.id) ? d.users.map((x) => (x.id === u.id ? u : x)) : [...d.users, u] }))

  async function step(key: string, fn: () => Promise<CsUser | void>, after?: () => void) {
    setBusy(key)
    try {
      const u = await fn()
      if (u) putUser(u)
      after?.()
      toast(t('dlSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  const unassign = (n: PhoneNumber) => {
    const u = byId.get(n.assignedTo!)
    if (!confirm(t('tphUnassignConfirm', n.number, u?.displayName || n.assignedTo!))) return
    step(n.number, () => unwrap(window.api.teams.unassignNumber(tid, u?.upn || n.assignedTo!, n.number, numberType(n))), () =>
      data.set((d) => ({ ...d, numbers: d.numbers.map((x) => (x.number === n.number ? { ...x, assignedTo: null } : x)) }))
    )
  }

  const grant = (u: CsUser, type: PolicyType, policy: string) =>
    step(u.id + type, async () => {
      const r = await unwrap(window.api.teams.grantUsers(tid, type, policy || null, [u.upn]))
      if (r.failed.length) throw new Error(r.failed[0].error)
      return r.users[0]
    })

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('teams')}
          </div>
          <h1>{t('teamPhone')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('tphSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={data.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <TeamsLoad loading={data.loading && !data.data} error={data.error} onRetry={data.reload} />
        {data.data && !numbers.length && !users.length && (
          <div className="empty">
            <h3>{t('tphNoneT')}</h3>
            <p>{t('tphNoneD')}</p>
          </div>
        )}
        {data.data && (numbers.length > 0 || users.length > 0) && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('tphNumbers')}</div>
                <div className="v">{numbers.length}</div>
                <div className="s">{t('tphAssignedN', numbers.length - free.length)}</div>
              </div>
              <div className="tile">
                <div className="k">{t('tphFree')}</div>
                <div className="v">{free.length}</div>
                <div className="s">{free.slice(0, 3).map((n) => n.number).join(', ') || '—'}</div>
              </div>
              <div className="tile">
                <div className="k">{t('tphUsers')}</div>
                <div className="v">{users.filter((u) => u.phoneSystem).length}</div>
                <div className="s">{t('tphVoiceN', users.filter((u) => u.enterpriseVoice).length)}</div>
              </div>
              <div className={`tile ${without.length ? 'warn' : ''}`}>
                <div className="k">{t('tphWithout')}</div>
                <div className="v">{without.length}</div>
                <div className="s">{without.map((u) => u.displayName).join(', ') || t('ok')}</div>
              </div>
            </div>
            <Tabs<typeof tab>
              tabs={[
                ['numbers', t('tphNumbers')],
                ['users', t('tphUsers')]
              ]}
              value={tab}
              onChange={setTab}
            />
            {tab === 'numbers' ? (
              <div className="panel">
                <div className="tbl-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>{t('tphNumber')}</th>
                        <th>{t('typeL')}</th>
                        <th>{t('tphCity')}</th>
                        <th>{t('tphAssignedTo')}</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {numbers
                        .filter((n) => match(n.number + (byId.get(n.assignedTo || '')?.displayName || '') + (n.city || '')))
                        .map((n) => {
                          const u = n.assignedTo ? byId.get(n.assignedTo) : null
                          return (
                            <tr key={n.number}>
                              <td className="mono">{n.number}</td>
                              <td>
                                <span className="lic-tag">{t(`nt_${numberType(n)}`)}</span>
                                {n.activation && n.activation !== 'Activated' && <span className="pill warn">{n.activation}</span>}
                              </td>
                              <td>{n.city || '—'}</td>
                              <td>{n.assignedTo ? u ? <>{u.displayName}<div className="mono">{u.upn}</div></> : <span className="mono">{n.assignedTo}</span> : <span className="pill ok">{t('tphFreePill')}</span>}</td>
                              <td style={{ width: 120, textAlign: 'right' }}>
                                {n.assignedTo ? (
                                  <button className="btn sm" disabled={!!busy} onClick={() => unassign(n)}>
                                    {busy === n.number && <span className="spin" />}
                                    {t('tphUnassign')}
                                  </button>
                                ) : (
                                  <button className="btn sm" disabled={!!busy} onClick={() => setAssign(n)}>
                                    {t('tphAssign')}
                                  </button>
                                )}
                              </td>
                            </tr>
                          )
                        })}
                    </tbody>
                  </table>
                  {!numbers.length && (
                    <p className="hint" style={{ padding: 16 }}>
                      {t('tphNoNumbers')}
                    </p>
                  )}
                </div>
              </div>
            ) : (
              <div className="panel">
                <div className="tbl-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>{t('name')}</th>
                        <th>{t('tphNumber')}</th>
                        <th>{t('tphVoice')}</th>
                        <th>{t('pt_TeamsCallingPolicy')}</th>
                        <th>{t('pt_OnlineVoiceRoutingPolicy')}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {users
                        .filter((u) => match(u.displayName + u.upn + (u.lineUri || '')))
                        .sort((a, b) => a.displayName.localeCompare(b.displayName))
                        .map((u) => (
                          <tr key={u.id}>
                            <td>
                              <div className="u-n">{u.displayName}</div>
                              <div className="mono">{u.upn}</div>
                            </td>
                            <td className="mono">{lineNumber(u.lineUri) || '—'}</td>
                            <td>
                              <input
                                type="checkbox"
                                className="chk"
                                checked={u.enterpriseVoice}
                                disabled={!!busy}
                                onChange={(e) => step(u.id, () => unwrap(window.api.teams.enterpriseVoice(tid, u.upn, e.target.checked)))}
                              />
                            </td>
                            <td>
                              <NameSelect names={data.data!.calling} value={u.policies.TeamsCallingPolicy || ''} disabled={!!busy} onChange={(v) => grant(u, 'TeamsCallingPolicy', v)} />
                            </td>
                            <td>
                              <NameSelect names={data.data!.routing} value={u.policies.OnlineVoiceRoutingPolicy || ''} disabled={!!busy} onChange={(v) => grant(u, 'OnlineVoiceRoutingPolicy', v)} />
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>
      {assign && (
        <AssignNumberModal
          number={assign}
          users={users}
          onClose={() => setAssign(null)}
          onDone={(u) => {
            putUser(u)
            data.set((d) => ({ ...d, numbers: d.numbers.map((x) => (x.number === assign.number ? { ...x, assignedTo: u.id } : x)) }))
            setAssign(null)
            toast(t('tphAssigned', assign.number, u.displayName))
          }}
        />
      )}
    </>
  )
}

function NameSelect({ names, value, onChange, disabled }: { names: string[]; value: string; onChange(v: string): void; disabled?: boolean }) {
  const { t } = useApp()
  return (
    <select value={value} disabled={disabled} onChange={(e) => onChange(e.target.value)}>
      <option value="">{t('tpGlobal')}</option>
      {names
        .filter((n) => n !== 'Global')
        .map((n) => (
          <option key={n} value={n}>
            {n}
          </option>
        ))}
    </select>
  )
}

function AssignNumberModal({ number, users, onClose, onDone }: { number: PhoneNumber; users: CsUser[]; onClose(): void; onDone(u: CsUser): void }) {
  const { t, tenant } = useApp()
  const [user, setUser] = useState('')
  const [loc, setLoc] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const candidates = users.filter((u) => !u.lineUri).sort((a, b) => a.displayName.localeCompare(b.displayName))
  const type = numberType(number)

  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      onDone(await unwrap(window.api.teams.assignNumber(tenant!.tenantId, user, number.number, type, loc.trim() || undefined)))
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('tphAssignT', number.number)}
      sub={t(`nt_${type}`)}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !user} onClick={submit}>
            {busy && <span className="spin" />}
            {t('tphAssign')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="ph-u">{t('tphUser')}</label>
        <select id="ph-u" value={user} onChange={(e) => setUser(e.target.value)}>
          <option value="">—</option>
          {candidates.map((u) => (
            <option key={u.id} value={u.upn}>
              {u.displayName} · {u.upn}
            </option>
          ))}
        </select>
        <span className="hint">{t('tphUserHint')}</span>
      </div>
      {type !== 'DirectRouting' && (
        <div className="fld">
          <label htmlFor="ph-l">
            {t('tphLocation')} <span className="hint">· {t('optional')}</span>
          </label>
          <input id="ph-l" className="mono" value={loc} onChange={(e) => setLoc(e.target.value)} />
          <span className="hint">{t('tphLocationHint')}</span>
        </div>
      )}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
