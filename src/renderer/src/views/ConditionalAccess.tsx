import { useState } from 'react'
import type { CaDraft, CaPolicy, CaState, DirectoryRef } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate, type Key, type T } from '../i18n'
import { Icon, Modal, RowMenu } from '../components/ui'
import { EntraLoad, PrincipalList } from '../components/entra'

const STATES: CaState[] = ['enabledForReportingButNotEnforced', 'enabled', 'disabled']

const blank = (name = ''): CaDraft => ({
  displayName: name,
  state: 'enabledForReportingButNotEnforced',
  users: 'all',
  include: [],
  exclude: [],
  apps: 'all',
  clientApps: 'all',
  locations: 'any',
  locationIds: [],
  grant: 'mfa'
})

/** Microsoft's baseline policies, always deployed in report-only first. */
export const TEMPLATES: { id: string; draft: (name: string) => CaDraft }[] = [
  { id: 'mfaAll', draft: (n) => ({ ...blank(n), users: 'all', grant: 'mfa' }) },
  { id: 'legacy', draft: (n) => ({ ...blank(n), users: 'all', clientApps: 'legacy', grant: 'block' }) },
  { id: 'mfaAdmins', draft: (n) => ({ ...blank(n), users: 'admins', grant: 'mfa' }) },
  { id: 'mfaAzure', draft: (n) => ({ ...blank(n), users: 'all', apps: 'azureManagement', grant: 'mfa' }) },
  { id: 'mfaGuests', draft: (n) => ({ ...blank(n), users: 'guests', grant: 'mfa' }) },
  { id: 'mfaOutside', draft: (n) => ({ ...blank(n), users: 'all', locations: 'excludeTrusted', grant: 'mfa' }) }
]

const statePill = (s: CaState, t: T) =>
  s === 'enabled' ? <span className="pill ok">{t('caOn')}</span> : s === 'disabled' ? <span className="pill off">{t('caOff')}</span> : <span className="pill warn">{t('caReport')}</span>

export function CaView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.caPolicies(tid), [tid])
  const [modal, setModal] = useState<{ k: 'new'; draft: CaDraft; template?: string } | { k: 'pick' } | { k: 'view'; p: CaPolicy } | null>(null)
  const rows = list.data || []

  async function run(fn: () => Promise<CaPolicy[]>, msg: string) {
    try {
      const r = await fn()
      list.set(() => r)
      toast(msg)
    } catch (e) {
      toast((e as Error).message)
    }
  }
  const setState = (p: CaPolicy, s: CaState) => {
    if (s === 'enabled' && !confirm(t('caEnableConfirm', p.displayName))) return
    run(() => unwrap(window.api.entra.setCaState(tid, p.id, s)), t('caStateSet', p.displayName))
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('ca')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn" onClick={() => setModal({ k: 'pick' })}>
            {t('caFromTemplate')}
          </button>
          <button className="btn pri" onClick={() => setModal({ k: 'new', draft: blank() })}>
            {Icon.plus}
            {t('caNew')}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('caPolicies')}</div>
                <div className="v">{rows.length}</div>
                <div className="s">{t('caPoliciesD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('caOn')}</div>
                <div className="v">{rows.filter((p) => p.state === 'enabled').length}</div>
                <div className="s">{t('caOnD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('caReport')}</div>
                <div className="v">{rows.filter((p) => p.state === 'enabledForReportingButNotEnforced').length}</div>
                <div className="s">{t('caReportD')}</div>
              </div>
              <div className={`tile ${rows.some((p) => p.state === 'enabled' && hasMfaAll(p)) ? '' : 'warn'}`}>
                <div className="k">{t('caMfaAll')}</div>
                <div className="v">{rows.some((p) => p.state === 'enabled' && hasMfaAll(p)) ? t('yes') : t('no')}</div>
                <div className="s">{t('caMfaAllD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('arName')}</th>
                      <th>{t('status')}</th>
                      <th>{t('caGrant')}</th>
                      <th>{t('caModified')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((p) => (
                      <tr key={p.id}>
                        <td>
                          <button className="linkish" onClick={() => setModal({ k: 'view', p })}>
                            <div className="u-n">{p.displayName}</div>
                          </button>
                        </td>
                        <td>{statePill(p.state, t)}</td>
                        <td className="mono">{grantText(p)}</td>
                        <td className="mono">{fmtDate(p.modified || p.created, lang, '—')}</td>
                        <td style={{ width: 40 }}>
                          <RowMenu
                            items={[
                              [t('caDetails'), () => setModal({ k: 'view', p })],
                              ...STATES.filter((s) => s !== p.state).map((s) => [t(`caSet_${s}`), () => setState(p, s)] as [string, () => void]),
                              [t('caDelete'), () => confirm(t('caDeleteConfirm', p.displayName)) && run(() => unwrap(window.api.entra.deleteCa(tid, p.id)), t('caDeleted'))]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!rows.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('caNone')}
                  </p>
                )}
              </div>
            </div>
            <p className="hint">{t('caHint')}</p>
          </>
        )}
      </div>
      {modal?.k === 'pick' && (
        <Modal title={t('caFromTemplate')} sub={t('caTemplatesSub')} onClose={() => setModal(null)}>
          {TEMPLATES.map((x) => (
            <button key={x.id} className="opt" style={{ textAlign: 'left', background: 'none' }} onClick={() => setModal({ k: 'new', draft: x.draft(t(`caT_${x.id}` as Key)), template: x.id })}>
              <span>
                <b>{t(`caT_${x.id}` as Key)}</b>
                <small>{t(`caTD_${x.id}` as Key)}</small>
              </span>
            </button>
          ))}
        </Modal>
      )}
      {modal?.k === 'new' && (
        <CaModal
          initial={modal.draft}
          onClose={() => setModal(null)}
          onSave={async (d) => {
            const r = await unwrap(window.api.entra.createCa(tid, d))
            list.set(() => r)
            setModal(null)
            toast(t('caCreated', d.displayName))
          }}
        />
      )}
      {modal?.k === 'view' && <PolicyDetail p={modal.p} onClose={() => setModal(null)} />}
    </>
  )
}

type Raw = {
  conditions?: { users?: { includeUsers?: string[] }; applications?: { includeApplications?: string[] }; clientAppTypes?: string[] }
  grantControls?: { operator?: string; builtInControls?: string[] } | null
  sessionControls?: unknown
}
const hasMfaAll = (p: CaPolicy) => {
  const r = p.raw as Raw
  return !!r.conditions?.users?.includeUsers?.includes('All') && !!r.conditions?.applications?.includeApplications?.includes('All') && !!r.grantControls?.builtInControls?.includes('mfa')
}
const grantText = (p: CaPolicy) => {
  const g = (p.raw as Raw).grantControls
  return g?.builtInControls?.length ? g.builtInControls.join(g.operator === 'AND' ? ' + ' : ' / ') : (p.raw as Raw).sessionControls ? 'session' : '—'
}

function CaModal({ initial, onClose, onSave }: { initial: CaDraft; onClose(): void; onSave(d: CaDraft): Promise<void> }) {
  const { t, tenant } = useApp()
  const [d, setD] = useState(initial)
  const [inc, setInc] = useState<DirectoryRef[]>([])
  const [exc, setExc] = useState<DirectoryRef[]>([])
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const locs = useLoad(() => (d.locations === 'selected' ? window.api.entra.locations(tenant!.tenantId) : null), [d.locations])
  const set = (p: Partial<CaDraft>) => setD((x) => ({ ...x, ...p }))
  const key = (p: DirectoryRef) => `${p.type === 'group' ? 'g' : 'u'}:${p.id}`
  const sel = (id: string, label: Key, value: string, options: [string, Key][], on: (v: string) => void) => (
    <div className="fld">
      <label htmlFor={id}>{t(label)}</label>
      <select id={id} value={value} onChange={(e) => on(e.target.value)}>
        {options.map(([v, k]) => (
          <option key={v} value={v}>
            {t(k)}
          </option>
        ))}
      </select>
    </div>
  )
  return (
    <Modal
      wide
      title={t('caNew')}
      sub={t('caNewSub')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !d.displayName.trim() || (d.users === 'selected' && !inc.length)}
            onClick={async () => {
              if (d.state === 'enabled' && !confirm(t('caEnableConfirm', d.displayName))) return
              setBusy(true)
              setErr(null)
              try {
                await onSave({ ...d, displayName: d.displayName.trim(), include: inc.map(key), exclude: exc.map(key) })
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('caCreate')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="cn">{t('arName')}</label>
        <input id="cn" value={d.displayName} onChange={(e) => set({ displayName: e.target.value })} />
      </div>
      <div className="row2">
        {sel('cu', 'caUsers', d.users, [['all', 'caU_all'], ['guests', 'caU_guests'], ['admins', 'caU_admins'], ['selected', 'caU_selected']], (v) => set({ users: v as CaDraft['users'] }))}
        {sel('ca', 'caApps', d.apps, [['all', 'caA_all'], ['office365', 'caA_office365'], ['azureManagement', 'caA_azureManagement']], (v) => set({ apps: v as CaDraft['apps'] }))}
      </div>
      {d.users === 'selected' && (
        <div className="fld">
          <label>{t('caInclude')}</label>
          <PrincipalList groups items={inc} onChange={setInc} />
        </div>
      )}
      <div className="fld">
        <label>
          {t('caExclude')} <span className="hint">{t('caExcludeHint')}</span>
        </label>
        <PrincipalList groups items={exc} onChange={setExc} />
      </div>
      <div className="row2">
        {sel('cc', 'caClientApps', d.clientApps, [['all', 'caC_all'], ['legacy', 'caC_legacy']], (v) => set({ clientApps: v as CaDraft['clientApps'] }))}
        {sel('cl', 'caLocations', d.locations, [['any', 'caL_any'], ['excludeTrusted', 'caL_excludeTrusted'], ['selected', 'caL_selected']], (v) => set({ locations: v as CaDraft['locations'] }))}
      </div>
      {d.locations === 'selected' && (
        <div className="fld">
          <label>{t('locations')}</label>
          {locs.data?.map((l) => (
            <label key={l.id} className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input
                type="checkbox"
                className="chk"
                checked={d.locationIds.includes(l.id)}
                onChange={(e) => set({ locationIds: e.target.checked ? [...d.locationIds, l.id] : d.locationIds.filter((x) => x !== l.id) })}
              />
              {l.displayName}
            </label>
          ))}
          {locs.data && !locs.data.length && <p className="hint">{t('nlNone')}</p>}
        </div>
      )}
      <div className="row2">
        {sel('cg', 'caGrant', d.grant, [['mfa', 'caG_mfa'], ['block', 'caG_block'], ['compliant', 'caG_compliant'], ['mfaOrCompliant', 'caG_mfaOrCompliant']], (v) => set({ grant: v as CaDraft['grant'] }))}
        {sel('cs', 'status', d.state, STATES.map((s) => [s, `caS_${s}` as Key]), (v) => set({ state: v as CaState }))}
      </div>
      {d.state === 'enabled' && <div className="banner">{t('caEnabledWarn')}</div>}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

/** A readable summary of a policy, with ids replaced by names, plus the raw JSON. */
function PolicyDetail({ p, onClose }: { p: CaPolicy; onClose(): void }) {
  const { t, tenant } = useApp()
  const ids = JSON.stringify(p.raw).match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) || []
  const names = useLoad(() => window.api.entra.resolve(tenant!.tenantId, ids.filter((x) => x !== p.id)), [p.id])
  const c = (p.raw.conditions || {}) as Record<string, Record<string, string[]> | string[] | null>
  const nm = (v: string) => names.data?.[v] || v
  const list = (v?: string[] | null) => (v?.length ? v.map(nm).join(', ') : '—')
  const users = (c.users || {}) as Record<string, string[]>
  const apps = (c.applications || {}) as Record<string, string[]>
  const locs = (c.locations || {}) as Record<string, string[]>
  const rows: [string, string][] = [
    [t('caInclude'), [list(users.includeUsers), list(users.includeGroups), list(users.includeRoles)].filter((x) => x !== '—').join(' · ') || '—'],
    [t('caExclude'), [list(users.excludeUsers), list(users.excludeGroups), list(users.excludeRoles)].filter((x) => x !== '—').join(' · ') || '—'],
    [t('caApps'), list(apps.includeApplications) + (apps.excludeApplications?.length ? ` (−${list(apps.excludeApplications)})` : '')],
    [t('caClientApps'), list(c.clientAppTypes as string[])],
    [t('caLocations'), locs.includeLocations ? `${list(locs.includeLocations)}${locs.excludeLocations?.length ? ` (−${list(locs.excludeLocations)})` : ''}` : '—'],
    [t('caGrant'), grantText(p)]
  ]
  return (
    <Modal wide title={p.displayName} sub={statePill(p.state, t)} onClose={onClose}>
      <div className="creds" style={{ fontFamily: 'var(--f-body)' }}>
        {rows.map(([k, v]) => (
          <div key={k} style={{ display: 'flex', gap: 10 }}>
            <span className="hint" style={{ width: 130, flex: 'none' }}>{k}</span>
            <span style={{ whiteSpace: 'normal' }}>{v}</span>
          </div>
        ))}
      </div>
      <details>
        <summary>JSON</summary>
        <pre className="mono" style={{ whiteSpace: 'pre-wrap', maxHeight: 320, overflow: 'auto', background: 'var(--surface-2)', padding: 10, borderRadius: 8 }}>
          {JSON.stringify(p.raw, null, 2)}
        </pre>
      </details>
    </Modal>
  )
}
