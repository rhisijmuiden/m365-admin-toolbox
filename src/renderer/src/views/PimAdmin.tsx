import { useState } from 'react'
import type { DirectoryRef, Group, PimApproval, PimGroupEntry, PimRoleSettings, RoleDef } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { EntraLoad, localInput, PrincipalInput, PrincipalList, Tabs, toIso } from '../components/entra'
import { useGroups } from '../components/teams'

type Tab = 'approvals' | 'settings' | 'groups' | 'audit'

/** PIM management: approve requests, role settings, PIM for groups and the PIM audit trail. */
export function PimAdminView() {
  const { t, tenant } = useApp()
  const [tab, setTab] = useState<Tab>('approvals')
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('pimAdmin')}</h1>
        </div>
        <div className="top-r">
          <Tabs
            tabs={[
              ['approvals', t('paApprovals')],
              ['settings', t('paSettings')],
              ['groups', t('paGroups')],
              ['audit', t('paAudit')]
            ]}
            value={tab}
            onChange={setTab}
          />
        </div>
      </header>
      <div className="content">
        {tab === 'approvals' && <Approvals />}
        {tab === 'settings' && <Settings />}
        {tab === 'groups' && <Groups />}
        {tab === 'audit' && <Audit />}
        <p className="hint">{t('paLicense')}</p>
      </div>
    </>
  )
}

function Approvals() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.approvals(tid), [tid])
  const [review, setReview] = useState<{ a: PimApproval; ok: boolean } | null>(null)
  const when = (iso: string | null) => (iso ? new Date(iso).toLocaleString(lang, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—')
  return (
    <>
      <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
      {list.data && (
        <div className="panel">
          <div className="panel-h">
            <h2>{t('paPending', list.data.length)}</h2>
            <div className="r">
              <button className="btn sm" onClick={list.reload} aria-label={t('refresh')}>
                {Icon.refresh}
              </button>
            </div>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('paRequester')}</th>
                  <th>{t('roleOrGroup')}</th>
                  <th>{t('reason')}</th>
                  <th>{t('paWindow')}</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {list.data.map((a) => (
                  <tr key={a.requestId}>
                    <td>
                      <div className="u-n">{a.principal.displayName}</div>
                      <div className="mono">{a.principal.detail}</div>
                    </td>
                    <td>
                      {a.targetName} {a.kind === 'group' && <span className="lic-tag">{a.accessId}</span>}
                    </td>
                    <td style={{ whiteSpace: 'normal', maxWidth: 280 }}>{a.justification || '—'}</td>
                    <td className="mono">
                      {when(a.start)} → {when(a.end)}
                    </td>
                    <td style={{ textAlign: 'right' }}>
                      <div style={{ display: 'inline-flex', gap: 6 }}>
                        <button className="btn sm pri" onClick={() => setReview({ a, ok: true })}>
                          {t('paApprove')}
                        </button>
                        <button className="btn sm danger" onClick={() => setReview({ a, ok: false })}>
                          {t('paDeny')}
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!list.data.length && (
              <p className="hint" style={{ padding: 16 }}>
                {t('paNoPending')}
              </p>
            )}
          </div>
        </div>
      )}
      {review && (
        <ReviewModal
          {...review}
          onClose={() => setReview(null)}
          onDone={() => {
            list.set((d) => d.filter((x) => x.requestId !== review.a.requestId))
            toast(review.ok ? t('paApproved') : t('paDenied'))
            setReview(null)
          }}
        />
      )}
    </>
  )
}

function ReviewModal({ a, ok, onClose, onDone }: { a: PimApproval; ok: boolean; onClose(): void; onDone(): void }) {
  const { t, tenant } = useApp()
  const [why, setWhy] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={ok ? t('paApproveT', a.targetName) : t('paDenyT', a.targetName)}
      sub={a.principal.displayName}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className={`btn ${ok ? 'pri' : 'danger'}`}
            disabled={busy || !why.trim()}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                await unwrap(window.api.entra.review(tenant!.tenantId, a, ok, why.trim()))
                onDone()
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {ok ? t('paApprove') : t('paDeny')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="pw">{t('reason')}</label>
        <input id="pw" autoFocus value={why} onChange={(e) => setWhy(e.target.value)} />
      </div>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function Settings() {
  const { t, tenant } = useApp()
  const tid = tenant!.tenantId
  const roles = useLoad(() => window.api.entra.roles(tid), [tid])
  const [q, setQ] = useState('')
  const [open, setOpen] = useState<RoleDef | null>(null)
  const list = (roles.data?.roles || [])
    .filter((r) => r.builtIn && r.displayName.toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => Number(b.privileged) - Number(a.privileged) || a.displayName.localeCompare(b.displayName))
  const count = (r: RoleDef) => roles.data?.members.filter((m) => m.roleId === r.id && m.kind === 'eligible').length || 0
  return (
    <>
      <EntraLoad loading={roles.loading && !roles.data} error={roles.error} onRetry={roles.reload} />
      {roles.data && (
        <div className="panel">
          <div className="panel-h">
            <label className="search">
              {Icon.search}
              <input placeholder={t('rlSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
            </label>
            <span className="hint">{t('paSettingsHint')}</span>
          </div>
          <div className="tbl-wrap" style={{ maxHeight: 520 }}>
            <table>
              <thead>
                <tr>
                  <th>{t('rlRole')}</th>
                  <th>{t('rlEligible')}</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {list.map((r) => (
                  <tr key={r.id}>
                    <td className="u-n">
                      {r.displayName} {r.privileged && <span className="flag">{t('rlPrivTag')}</span>}
                    </td>
                    <td>{count(r) || '—'}</td>
                    <td style={{ textAlign: 'right' }}>
                      <button className="btn sm" onClick={() => setOpen(r)}>
                        {t('paEdit')}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
      {open && <SettingsModal role={open} onClose={() => setOpen(null)} />}
    </>
  )
}

function SettingsModal({ role, onClose }: { role: RoleDef; onClose(): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const cur = useLoad(() => window.api.entra.roleSettings(tid, role.id), [role.id])
  const [s, setS] = useState<PimRoleSettings | null>(null)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const v = s || cur.data
  const set = (p: Partial<PimRoleSettings>) => setS({ ...(v as PimRoleSettings), ...p })
  const check = (k: 'mfa' | 'justification' | 'ticket' | 'approval', label: string) => (
    <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
      <input type="checkbox" className="chk" checked={!!v?.[k]} onChange={(e) => set({ [k]: e.target.checked })} />
      {label}
    </label>
  )
  return (
    <Modal
      title={t('paSettingsT', role.displayName)}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !s}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const { policyId: _p, ...rest } = s!
                await unwrap(window.api.entra.setRoleSettings(tid, role.id, rest))
                toast(t('saved'))
                onClose()
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('save')}
          </button>
        </>
      }
    >
      <EntraLoad loading={cur.loading} error={cur.error} onRetry={cur.reload} />
      {v && (
        <>
          <div className="fld">
            <label htmlFor="mh">{t('paMaxHours')}</label>
            <select id="mh" value={v.maxHours} onChange={(e) => set({ maxHours: +e.target.value })}>
              {[1, 2, 4, 6, 8, 10, 12, 24].map((h) => (
                <option key={h} value={h}>
                  {t('hours', h)}
                </option>
              ))}
            </select>
          </div>
          {check('mfa', t('paReqMfa'))}
          {check('justification', t('paReqReason'))}
          {check('ticket', t('paReqTicket'))}
          {check('approval', t('paReqApproval'))}
          {v.approval && (
            <div className="fld">
              <label>{t('paApprovers')}</label>
              <PrincipalList groups items={v.approvers} onChange={(approvers: DirectoryRef[]) => set({ approvers })} />
            </div>
          )}
        </>
      )}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function Groups() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const groups = useGroups(tid)
  const [g, setG] = useState<Group | null>(null)
  const [q, setQ] = useState('')
  const list = useLoad(() => (g ? window.api.entra.pimGroup(tid, g.id) : null), [g?.id])
  const [add, setAdd] = useState(false)
  const [busy, setBusy] = useState<string | null>(null)
  const when = (iso: string | null) => (iso ? new Date(iso).toLocaleDateString(lang, { day: 'numeric', month: 'short', year: 'numeric' }) : t('rlNoEnd'))
  const pick = groups.filter((x) => x.kind === 'security' || x.kind === 'm365').filter((x) => x.displayName.toLowerCase().includes(q.toLowerCase())).slice(0, 50)
  async function remove(e: PimGroupEntry) {
    if (!g || !confirm(t('rlRemoveConfirm', e.principal.displayName, g.displayName))) return
    setBusy(e.id)
    try {
      const r = await unwrap(window.api.entra.pimGroupRemove(tid, g.id, { accessId: e.accessId, kind: e.kind === 'eligible' ? 'eligible' : 'active', principalId: e.principal.id }))
      list.set(() => r)
      toast(t('rlRemoved'))
    } catch (err) {
      toast((err as Error).message)
    } finally {
      setBusy(null)
    }
  }
  return (
    <>
      <p className="hint" style={{ margin: '0 0 12px', maxWidth: '80ch' }}>
        {t('paGroupsIntro')}
      </p>
      <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 16, alignItems: 'start' }}>
        <div className="panel">
          <div className="panel-h">
            <input placeholder={t('paGroupSearch')} value={q} onChange={(e) => setQ(e.target.value)} style={{ width: '100%', padding: '6px 9px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface-2)' }} />
          </div>
          <div style={{ maxHeight: 480, overflow: 'auto', padding: 6 }}>
            {pick.map((x) => (
              <button key={x.id} className="tm-item" style={{ width: '100%', borderRadius: 7, background: g?.id === x.id ? 'var(--surface-2)' : undefined }} onClick={() => setG(x)}>
                <span>{x.displayName}</span>
              </button>
            ))}
          </div>
        </div>
        <div>
          {!g && <div className="empty">{t('paPickGroup')}</div>}
          {g && (
            <div className="panel">
              <div className="panel-h">
                <h2>{g.displayName}</h2>
                <div className="r">
                  <button className="btn sm pri" onClick={() => setAdd(true)}>
                    {Icon.plus}
                    {t('rlAssign')}
                  </button>
                </div>
              </div>
              <div style={{ padding: list.data ? 0 : 14 }}>
                <EntraLoad loading={list.loading && !list.data?.length} error={list.error} onRetry={list.reload} />
              </div>
              {list.data && (
                <table>
                  <thead>
                    <tr>
                      <th>{t('eaPrincipal')}</th>
                      <th>{t('paAccess')}</th>
                      <th>{t('rlKind')}</th>
                      <th>{t('rlUntil')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {list.data.map((e) => (
                      <tr key={e.id}>
                        <td>
                          <div className="u-n">{e.principal.displayName}</div>
                          <div className="mono">{e.principal.detail}</div>
                        </td>
                        <td>{e.accessId === 'owner' ? t('paOwner') : t('paMember')}</td>
                        <td>
                          <span className={`pill ${e.kind === 'eligible' ? 'off' : e.kind === 'active' ? 'warn' : 'ok'}`}>{t(`rk_${e.kind}`)}</span>
                        </td>
                        <td className="mono">{when(e.end)}</td>
                        <td style={{ textAlign: 'right' }}>
                          <button className="btn sm danger" disabled={!!busy} onClick={() => remove(e)}>
                            {busy === e.id && <span className="spin" />}
                            {t('remove')}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {list.data && !list.data.length && !list.loading && (
                <p className="hint" style={{ padding: 14 }}>
                  {t('paGroupEmpty')}
                </p>
              )}
            </div>
          )}
        </div>
      </div>
      {add && g && (
        <GroupAssignModal
          group={g}
          onClose={() => setAdd(false)}
          onDone={(r) => {
            list.set(() => r)
            setAdd(false)
          }}
        />
      )}
    </>
  )
}

function GroupAssignModal({ group, onClose, onDone }: { group: Group; onClose(): void; onDone(r: PimGroupEntry[]): void }) {
  const { t, tenant, toast } = useApp()
  const [who, setWho] = useState<DirectoryRef | null>(null)
  const [access, setAccess] = useState<'member' | 'owner'>('member')
  const [kind, setKind] = useState<'eligible' | 'active'>('eligible')
  const [start, setStart] = useState(localInput(new Date()))
  const [end, setEnd] = useState(localInput(new Date(Date.now() + 180 * 86400000)))
  const [noEnd, setNoEnd] = useState(false)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('rlAssignT', group.displayName)}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !who}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const r = await unwrap(window.api.entra.pimGroupAssign(tenant!.tenantId, group.id, { principalId: who!.id, accessId: access, kind, start: toIso(start), end: noEnd ? null : toIso(end) }))
                toast(t('rlAssigned', who!.displayName))
                onDone(r)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('rlAssign')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label>{t('eaPrincipal')}</label>
        {who ? (
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <b style={{ flex: 1 }}>{who.displayName}</b>
            <button className="btn sm" onClick={() => setWho(null)}>
              {t('change')}
            </button>
          </div>
        ) : (
          <PrincipalInput groups label={t('select')} onPick={setWho} />
        )}
      </div>
      <div className="row2">
        <div className="fld">
          <label htmlFor="ga">{t('paAccess')}</label>
          <select id="ga" value={access} onChange={(e) => setAccess(e.target.value as 'member' | 'owner')}>
            <option value="member">{t('paMember')}</option>
            <option value="owner">{t('paOwner')}</option>
          </select>
        </div>
        <div className="fld">
          <label htmlFor="gk">{t('rlKind')}</label>
          <select id="gk" value={kind} onChange={(e) => setKind(e.target.value as 'eligible' | 'active')}>
            <option value="eligible">{t('rk_eligible')}</option>
            <option value="active">{t('rk_active')}</option>
          </select>
        </div>
      </div>
      <div className="row2">
        <div className="fld">
          <label htmlFor="gs">{t('rlStart')}</label>
          <input id="gs" type="datetime-local" value={start} onChange={(e) => setStart(e.target.value)} />
        </div>
        <div className="fld">
          <label htmlFor="ge">{t('rlEnd')}</label>
          <input id="ge" type="datetime-local" value={end} disabled={noEnd} onChange={(e) => setEnd(e.target.value)} />
        </div>
      </div>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={noEnd} onChange={(e) => setNoEnd(e.target.checked)} />
        {t('rlPermanentOpt')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function Audit() {
  const { t, lang, tenant } = useApp()
  const [days, setDays] = useState(7)
  const [q, setQ] = useState('')
  const list = useLoad(() => window.api.entra.pimAudit(tenant!.tenantId, days), [tenant!.tenantId, days])
  const rows = (list.data || []).filter((r) => [r.activity, r.by, r.target, r.reason].join(' ').toLowerCase().includes(q.toLowerCase()))
  return (
    <div className="panel">
      <div className="panel-h">
        <div className="seg">
          {[1, 7, 30].map((d) => (
            <button key={d} className={days === d ? 'on' : ''} onClick={() => setDays(d)}>
              {t('paDays', d)}
            </button>
          ))}
        </div>
        <label className="search">
          {Icon.search}
          <input placeholder={t('paAuditSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
        </label>
      </div>
      <div style={{ padding: list.data ? 0 : 14 }}>
        <EntraLoad loading={list.loading} error={list.error} onRetry={list.reload} />
      </div>
      {list.data && (
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>{t('paWhen')}</th>
                <th>{t('paActivity')}</th>
                <th>{t('paBy')}</th>
                <th>{t('paTarget')}</th>
                <th>{t('reason')}</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id}>
                  <td className="mono">{new Date(r.date).toLocaleString(lang, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</td>
                  <td>
                    {r.activity} {r.result !== 'success' && <span className="pill bad">{r.result}</span>}
                  </td>
                  <td>{r.by || '—'}</td>
                  <td style={{ whiteSpace: 'normal', maxWidth: 300 }}>{r.target || '—'}</td>
                  <td style={{ whiteSpace: 'normal', maxWidth: 260 }}>{r.reason || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {!rows.length && (
            <p className="hint" style={{ padding: 16 }}>
              {t('paNoAudit')}
            </p>
          )}
        </div>
      )}
    </div>
  )
}
