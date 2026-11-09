import { useState } from 'react'
import type { DirectoryRef, GdapRelationship, RoleAssign, RoleDef, RoleMember, RolesData } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { daysLeft, EntraLoad, localInput, PrincipalInput, toIso } from '../components/entra'

const GLOBAL_ADMIN = '62e90394-69f5-4237-9190-012177145e10'

/** Roles each phase of the toolbox needs, by template id; checked against the GDAP relationship. */
export const NEEDED: [string, string, number][] = [
  ['fe930be7-5e62-47db-91af-98c3a49a38b1', 'User Administrator', 1],
  ['4d6ac14f-3453-41d0-bef9-a3e0c569773a', 'License Administrator', 1],
  ['fdd7a751-b60b-444a-984c-02652fe8fa1c', 'Groups Administrator', 1],
  ['29232cdf-9323-42fd-ade2-1d097af3e4de', 'Exchange Administrator', 2],
  ['69091246-20e8-4a56-aa4d-066075b2a7a8', 'Teams Administrator', 3],
  ['9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3', 'Application Administrator', 4],
  ['b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', 'Conditional Access Administrator', 4],
  ['e8611ab8-c189-46e8-94e1-60213ab1f814', 'Privileged Role Administrator', 4],
  ['c4e39bd9-1100-46d3-8c65-fb160da0071f', 'Authentication Administrator', 4],
  ['194ae4cb-b126-40b2-bd5b-6091b380977d', 'Security Administrator', 5]
]

export function RolesView() {
  const { t, lang, tenant } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.entra.roles(tid), [tid])
  const gdap = useLoad(() => (tenant!.mode === 'gdap' ? window.api.entra.gdap(tid) : null), [tid])
  const [q, setQ] = useState('')
  const [onlyUsed, setOnlyUsed] = useState(true)
  const [open, setOpen] = useState<RoleDef | null>(null)
  const d = data.data
  const of = (r: RoleDef) => d?.members.filter((m) => m.roleId === r.id || m.roleId === r.templateId) || []
  const shown = (d?.roles || [])
    .filter((r) => !onlyUsed || of(r).length)
    .filter((r) => (r.displayName + of(r).map((m) => m.principal.displayName + (m.principal.detail || '')).join(' ')).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => Number(b.privileged) - Number(a.privileged) || a.displayName.localeCompare(b.displayName))
  const ga = d ? of(d.roles.find((r) => r.templateId === GLOBAL_ADMIN) || ({ id: GLOBAL_ADMIN, templateId: GLOBAL_ADMIN } as RoleDef)).filter((m) => m.kind !== 'eligible') : []
  const privileged = d ? new Set(d.members.filter((m) => d.roles.find((r) => (r.id === m.roleId || r.templateId === m.roleId) && r.privileged) && m.kind !== 'eligible').map((m) => m.principal.id)) : new Set()

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('roles')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('rlSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={() => (data.reload(), gdap.reload())} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={data.loading && !d} error={data.error} onRetry={data.reload} />
        {d && (
          <>
            <div className="tiles">
              <div className={`tile ${ga.length > 4 || ga.length < 2 ? 'warn' : ''}`}>
                <div className="k">{t('rlGlobal')}</div>
                <div className="v">{ga.length}</div>
                <div className="s">{t('rlGlobalD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('rlPrivileged')}</div>
                <div className="v">{privileged.size}</div>
                <div className="s">{t('rlPrivilegedD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('rlEligible')}</div>
                <div className="v">{d.pim ? d.members.filter((m) => m.kind === 'eligible').length : '—'}</div>
                <div className="s">{d.pim ? t('rlEligibleD') : t('rlNoPim')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('rlInUse')}</div>
                <div className="v">{d.roles.filter((r) => of(r).length).length}</div>
                <div className="s">{t('rlOfRoles', d.roles.length)}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={onlyUsed} onChange={(e) => setOnlyUsed(e.target.checked)} />
                  {t('rlOnlyUsed')}
                </label>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('rlRole')}</th>
                      <th>{t('rlMembers')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((r) => {
                      const ms = of(r)
                      return (
                        <tr key={r.id}>
                          <td>
                            <button className="linkish" onClick={() => setOpen(r)}>
                              <div className="u-n">
                                {r.displayName} {r.privileged && <span className="flag">{t('rlPrivTag')}</span>}
                              </div>
                            </button>
                          </td>
                          <td style={{ whiteSpace: 'normal' }}>
                            {ms.slice(0, 5).map((m) => (
                              <span key={m.id} className="lic-tag" title={m.principal.detail || ''}>
                                {m.principal.displayName}
                                {m.kind !== 'permanent' && <span className="hint"> · {t(`rk_${m.kind}`)}</span>}
                              </span>
                            ))}
                            {ms.length > 5 && <span className="hint">+{ms.length - 5}</span>}
                            {!ms.length && <span className="hint">—</span>}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            <button className="btn sm" onClick={() => setOpen(r)}>
                              {t('rlManage')}
                            </button>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
        {tenant!.mode === 'gdap' && (
          <div className="panel" style={{ marginTop: 16 }}>
            <div className="panel-h">
              <h2>{t('rlGdap')}</h2>
            </div>
            <div style={{ padding: 14, display: 'grid', gap: 12 }}>
              <EntraLoad loading={gdap.loading && !gdap.data} error={gdap.error} onRetry={gdap.reload} />
              {gdap.data?.error && <div className="banner bad">{gdap.data.error}</div>}
              {gdap.data && !gdap.data.relationships.length && !gdap.data.error && <p className="hint">{t('rlGdapNone')}</p>}
              {gdap.data?.relationships.map((r) => <Gdap key={r.id} r={r} names={d?.roles || []} lang={lang} />)}
            </div>
          </div>
        )}
      </div>
      {open && d && <RoleModal role={open} data={d} onClose={() => setOpen(null)} onChange={(n) => data.set(() => n)} />}
    </>
  )
}

function Gdap({ r, names, lang }: { r: GdapRelationship; names: RoleDef[]; lang: string }) {
  const { t } = useApp()
  const name = (id: string) => names.find((x) => x.templateId === id)?.displayName || NEEDED.find((x) => x[0] === id)?.[1] || id
  const left = daysLeft(r.end)
  const active = r.status === 'active'
  return (
    <div style={{ border: '1px solid var(--line)', borderRadius: 8, padding: 12, display: 'grid', gap: 8 }}>
      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        <b>{r.displayName}</b>
        <span className={`pill ${active ? (left !== null && left <= 30 ? 'warn' : 'ok') : 'bad'}`}>{r.status}</span>
        <span className="hint">
          {r.end ? t('rlGdapEnds', new Date(r.end).toLocaleDateString(lang)) : ''} {r.autoExtend && r.autoExtend !== 'PT0S' ? `· ${t('rlAutoExtend')}` : `· ${t('rlNoAutoExtend')}`}
        </span>
      </div>
      <div className="hint">{t('rlGdapRoles', r.roles.length)}</div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
        {NEEDED.map(([id, n, phase]) => (
          <span key={id} className={`flag ${r.roles.includes(id) ? '' : 'external'}`} style={r.roles.includes(id) ? { background: 'var(--ok-soft)', color: 'var(--ok)' } : undefined} title={`${t('phase')} ${phase}`}>
            {r.roles.includes(id) ? '✓' : '✗'} {n}
          </span>
        ))}
      </div>
      {r.groups.length > 0 && (
        <table>
          <tbody>
            {r.groups.map((g) => (
              <tr key={g.id}>
                <td className="u-n">{g.name}</td>
                <td className="hint" style={{ whiteSpace: 'normal' }}>{g.roles.map(name).join(', ')}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  )
}

function RoleModal({ role, data, onClose, onChange }: { role: RoleDef; data: RolesData; onClose(): void; onChange(d: RolesData): void }) {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [assign, setAssign] = useState(false)
  const [busy, setBusy] = useState<string | null>(null)
  const ms = data.members.filter((m) => m.roleId === role.id || m.roleId === role.templateId)
  const when = (iso: string | null) => (iso ? new Date(iso).toLocaleString(lang, { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '')
  async function remove(m: RoleMember) {
    if (!confirm(t('rlRemoveConfirm', m.principal.displayName, role.displayName))) return
    setBusy(m.id)
    try {
      onChange(await unwrap(window.api.entra.removeRole(tid, { id: m.id, kind: m.kind, roleId: role.id, principalId: m.principal.id })))
      toast(t('rlRemoved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }
  if (assign)
    return (
      <AssignModal
        role={role}
        pim={data.pim}
        onClose={() => setAssign(false)}
        onDone={(n) => {
          onChange(n)
          setAssign(false)
        }}
      />
    )
  return (
    <Modal wide title={role.displayName} sub={role.description} onClose={onClose}>
      <div>
        <button className="btn sm pri" onClick={() => setAssign(true)}>
          {Icon.plus}
          {t('rlAssign')}
        </button>
      </div>
      <div className="panel">
        <table>
          <thead>
            <tr>
              <th>{t('eaPrincipal')}</th>
              <th>{t('rlKind')}</th>
              <th>{t('rlUntil')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {ms.map((m) => (
              <tr key={m.id}>
                <td>
                  <div className="u-n">{m.principal.displayName}</div>
                  <div className="mono">{m.principal.detail}</div>
                </td>
                <td>
                  <span className={`pill ${m.kind === 'eligible' ? 'off' : m.kind === 'active' ? 'warn' : 'ok'}`}>{t(`rk_${m.kind}`)}</span>
                </td>
                <td className="mono">{m.end ? when(m.end) : t('rlNoEnd')}</td>
                <td style={{ textAlign: 'right' }}>
                  <button className="btn sm danger" disabled={!!busy} onClick={() => remove(m)}>
                    {busy === m.id && <span className="spin" />}
                    {t('remove')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!ms.length && (
          <p className="hint" style={{ padding: 14 }}>
            {t('rlNoMembers')}
          </p>
        )}
      </div>
    </Modal>
  )
}

function AssignModal({ role, pim, onClose, onDone }: { role: RoleDef; pim: boolean; onClose(): void; onDone(d: RolesData): void }) {
  const { t, tenant, toast } = useApp()
  const [who, setWho] = useState<DirectoryRef | null>(null)
  const [kind, setKind] = useState<RoleAssign['kind']>(pim ? 'eligible' : 'permanent')
  const [start, setStart] = useState(localInput(new Date()))
  const [end, setEnd] = useState(localInput(new Date(Date.now() + 365 * 86400000)))
  const [noEnd, setNoEnd] = useState(kind === 'permanent')
  const [why, setWhy] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('rlAssignT', role.displayName)}
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
                const timed = kind !== 'permanent'
                const r = await unwrap(
                  window.api.entra.assignRole(tenant!.tenantId, {
                    roleId: role.id,
                    principalId: who!.id,
                    kind,
                    start: timed ? toIso(start) : undefined,
                    end: timed && !noEnd ? toIso(end) : null,
                    justification: why || undefined
                  })
                )
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
            <b style={{ flex: 1 }}>
              {who.displayName} <span className="mono">{who.detail}</span>
            </b>
            <button className="btn sm" onClick={() => setWho(null)}>
              {t('change')}
            </button>
          </div>
        ) : (
          <PrincipalInput groups label={t('select')} onPick={setWho} />
        )}
      </div>
      <div className="fld">
        <label>{t('rlKind')}</label>
        {(['eligible', 'active', 'permanent'] as const).map((k) => (
          <label key={k} className={`opt ${kind === k ? 'on' : ''}`} style={!pim && k !== 'permanent' ? { opacity: 0.5 } : undefined}>
            <input type="radio" name="rk" checked={kind === k} disabled={!pim && k !== 'permanent'} onChange={() => (setKind(k), setNoEnd(k === 'permanent'))} />
            <span>
              {t(`rk_${k}`)}
              <small>{t(`rkD_${k}`)}</small>
            </span>
          </label>
        ))}
        {!pim && <p className="hint">{t('rlNoPimHint')}</p>}
      </div>
      {kind !== 'permanent' && (
        <>
          <div className="row2">
            <div className="fld">
              <label htmlFor="rs">{t('rlStart')}</label>
              <input id="rs" type="datetime-local" value={start} onChange={(e) => setStart(e.target.value)} />
            </div>
            <div className="fld">
              <label htmlFor="re">{t('rlEnd')}</label>
              <input id="re" type="datetime-local" value={end} disabled={noEnd} onChange={(e) => setEnd(e.target.value)} />
            </div>
          </div>
          <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <input type="checkbox" className="chk" checked={noEnd} onChange={(e) => setNoEnd(e.target.checked)} />
            {t('rlPermanentOpt')}
          </label>
          <div className="fld">
            <label htmlFor="rw">{t('reason')}</label>
            <input id="rw" value={why} onChange={(e) => setWhy(e.target.value)} />
          </div>
        </>
      )}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
