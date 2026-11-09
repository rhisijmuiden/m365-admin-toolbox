import { useEffect, useMemo, useState } from 'react'
import type { ApiResource, AppPermissionState, AppReg, AppRole, NewAppReg, PermissionDef, ResourceAccess, SignInAudience } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { CopyButton, Icon, Modal, RowMenu } from '../components/ui'
import { daysLeft, EntraLoad, ExpiryPill, OwnersPanel, pending, SecretModal, Tabs } from '../components/entra'

export const API_APPS: [string, string][] = [
  ['00000003-0000-0000-c000-000000000000', 'Microsoft Graph'],
  ['00000002-0000-0ff1-ce00-000000000000', 'Office 365 Exchange Online'],
  ['00000003-0000-0ff1-ce00-000000000000', 'SharePoint']
]
const AUDIENCES: SignInAudience[] = ['AzureADMyOrg', 'AzureADMultipleOrgs', 'AzureADandPersonalMicrosoftAccount', 'PersonalMicrosoftAccount']

/** Earliest end date over an app's secrets and certificates. */
export const firstEnd = (a: Pick<AppReg, 'credentials'>) =>
  a.credentials.map((c) => c.end).filter(Boolean).sort()[0] || null

export function AppRegsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.apps(tid), [tid])
  const [q, setQ] = useState('')
  const [modal, setModal] = useState<{ k: 'new' } | { k: 'app'; id: string } | null>(null)

  // Opened from the expiry overview.
  useEffect(() => {
    if (pending.appId && list.data) {
      const a = list.data.find((x) => x.id === pending.appId)
      pending.appId = null
      if (a) setModal({ k: 'app', id: a.id })
    }
  }, [list.data])

  const rows = list.data || []
  const shown = rows.filter((a) => (a.displayName + a.appId).toLowerCase().includes(q.toLowerCase()))
  const allCreds = rows.flatMap((a) => a.credentials)
  const expired = allCreds.filter((c) => (daysLeft(c.end) ?? 1) < 0).length
  const soon = allCreds.filter((c) => {
    const d = daysLeft(c.end)
    return d !== null && d >= 0 && d <= 30
  }).length

  async function remove(a: AppReg) {
    if (!confirm(t('arDeleteConfirm', a.displayName))) return
    try {
      await unwrap(window.api.entra.deleteApp(tid, a.id))
      list.set((d) => d.filter((x) => x.id !== a.id))
      toast(t('arDeleted', a.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('appRegs')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('arSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn pri" onClick={() => setModal({ k: 'new' })}>
            {Icon.plus}
            {t('arNew')}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('appRegs')}</div>
                <div className="v">{rows.length}</div>
                <div className="s">{t('arWithCreds', rows.filter((a) => a.credentials.length).length)}</div>
              </div>
              <div className={`tile ${expired ? 'bad' : ''}`}>
                <div className="k">{t('enExpiredN')}</div>
                <div className="v">{expired}</div>
                <div className="s">{t('enSecretsCerts')}</div>
              </div>
              <div className={`tile ${soon ? 'warn' : ''}`}>
                <div className="k">{t('enWithin', 30)}</div>
                <div className="v">{soon}</div>
                <div className="s">{t('enSecretsCerts')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('arMultiTenant')}</div>
                <div className="v">{rows.filter((a) => a.audience !== 'AzureADMyOrg').length}</div>
                <div className="s">{t('arMultiTenantD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('arName')}</th>
                      <th>{t('arAppId')}</th>
                      <th>{t('arAudience')}</th>
                      <th>{t('enSecretsCerts')}</th>
                      <th>{t('arCreated')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((a) => (
                      <tr key={a.id}>
                        <td>
                          <button className="linkish" onClick={() => setModal({ k: 'app', id: a.id })}>
                            <div className="u-n">{a.displayName}</div>
                          </button>
                        </td>
                        <td className="mono">{a.appId}</td>
                        <td>{t(`aud_${a.audience}` as 'aud_AzureADMyOrg') || a.audience}</td>
                        <td>
                          {a.credentials.length ? (
                            <span style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
                              <ExpiryPill end={firstEnd(a)} />
                              <span className="hint">{t('arCredCount', a.credentials.length)}</span>
                            </span>
                          ) : (
                            <span className="hint">{t('arNoCreds')}</span>
                          )}
                        </td>
                        <td className="mono">{fmtDate(a.created, lang, '—')}</td>
                        <td style={{ width: 40 }}>
                          <RowMenu
                            items={[
                              [t('arOpen'), () => setModal({ k: 'app', id: a.id })],
                              [t('arDelete'), () => remove(a)]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('arNone')}
                  </p>
                )}
              </div>
            </div>
          </>
        )}
      </div>
      {modal?.k === 'new' && (
        <CreateAppModal
          onClose={() => setModal(null)}
          onCreated={(a) => {
            list.set((d) => [...d, a].sort((x, y) => x.displayName.localeCompare(y.displayName)))
            setModal({ k: 'app', id: a.id })
          }}
        />
      )}
      {modal?.k === 'app' && (
        <AppDetail id={modal.id} onClose={() => setModal(null)} onChange={(a) => list.set((d) => d.map((x) => (x.id === a.id ? a : x)))} />
      )}
    </>
  )
}

// ---------- create ----------

function CreateAppModal({ onClose, onCreated }: { onClose(): void; onCreated(a: AppReg): void }) {
  const { t, tenant, toast } = useApp()
  const [n, setN] = useState<NewAppReg>({ displayName: '', audience: 'AzureADMyOrg', platform: 'none', redirectUris: [], permissions: [], adminConsent: true })
  const [uris, setUris] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const set = (p: Partial<NewAppReg>) => setN((x) => ({ ...x, ...p }))
  const count = n.permissions.reduce((s, p) => s + p.access.length, 0)
  return (
    <Modal
      wide
      title={t('arNew')}
      sub={t('arNewSub')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !n.displayName.trim()}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const redirectUris = uris.split(/\s+/).map((x) => x.trim()).filter(Boolean)
                const r = await unwrap(window.api.entra.createApp(tenant!.tenantId, { ...n, displayName: n.displayName.trim(), redirectUris }))
                toast(r.consentError ? t('arCreatedNoConsent', r.consentError) : t('arRegistered', r.app.displayName))
                onCreated(r.app)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('arCreate')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="an">{t('arName')}</label>
        <input id="an" autoFocus value={n.displayName} onChange={(e) => set({ displayName: e.target.value })} />
      </div>
      <div className="fld">
        <label htmlFor="aa">{t('arAudience')}</label>
        <select id="aa" value={n.audience} onChange={(e) => set({ audience: e.target.value as SignInAudience })}>
          {AUDIENCES.map((a) => (
            <option key={a} value={a}>
              {t(`aud_${a}`)}
            </option>
          ))}
        </select>
      </div>
      <div className="row2">
        <div className="fld">
          <label htmlFor="ap">{t('arPlatform')}</label>
          <select id="ap" value={n.platform} onChange={(e) => set({ platform: e.target.value as NewAppReg['platform'] })}>
            {(['none', 'web', 'spa', 'publicClient'] as const).map((p) => (
              <option key={p} value={p}>
                {t(`plat_${p}`)}
              </option>
            ))}
          </select>
        </div>
        <div className="fld">
          <label htmlFor="au">{t('arRedirects')}</label>
          <textarea id="au" className="mono" disabled={n.platform === 'none'} placeholder="https://localhost/callback" value={uris} onChange={(e) => setUris(e.target.value)} style={{ minHeight: 38 }} />
        </div>
      </div>
      <div className="fld">
        <label>
          {t('arPermissions')} <span className="hint">{t('arPermCount', count)}</span>
        </label>
        <PermissionPicker value={n.permissions} onChange={(permissions) => set({ permissions })} />
      </div>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={n.adminConsent} onChange={(e) => set({ adminConsent: e.target.checked })} />
        {t('arConsentNow')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

/** Picks delegated and application permissions of Microsoft Graph, Exchange Online or SharePoint. */
function PermissionPicker({ value, onChange }: { value: ResourceAccess[]; onChange(v: ResourceAccess[]): void }) {
  const { t, tenant } = useApp()
  const [api, setApi] = useState(API_APPS[0][0])
  const [type, setType] = useState<'Scope' | 'Role'>('Scope')
  const [q, setQ] = useState('')
  const res = useLoad(() => window.api.entra.resource(tenant!.tenantId, api), [api])
  const chosen = new Set(value.find((v) => v.resourceAppId === api)?.access.map((a) => a.id) || [])
  const perms = (res.data?.permissions || []).filter((p) => p.type === type && (p.value + p.name).toLowerCase().includes(q.toLowerCase()))
  const toggle = (p: PermissionDef) => {
    const cur = value.find((v) => v.resourceAppId === api) || { resourceAppId: api, access: [] }
    const access = chosen.has(p.id) ? cur.access.filter((a) => a.id !== p.id) : [...cur.access, { id: p.id, type: p.type }]
    onChange([...value.filter((v) => v.resourceAppId !== api), { resourceAppId: api, access }].filter((v) => v.access.length))
  }
  return (
    <div style={{ border: '1px solid var(--line)', borderRadius: 8 }}>
      <div style={{ display: 'flex', gap: 8, padding: 8, flexWrap: 'wrap', borderBottom: '1px solid var(--line)' }}>
        <select value={api} onChange={(e) => setApi(e.target.value)} style={{ padding: '5px 8px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }}>
          {API_APPS.map(([id, name]) => (
            <option key={id} value={id}>
              {name}
            </option>
          ))}
        </select>
        <Tabs tabs={[['Scope', t('arDelegated')], ['Role', t('arApplication')]]} value={type} onChange={setType} />
        <input placeholder={t('arPermSearch')} value={q} onChange={(e) => setQ(e.target.value)} style={{ flex: 1, minWidth: 140, padding: '5px 8px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }} />
      </div>
      <div style={{ maxHeight: 230, overflow: 'auto', padding: 6 }}>
        <EntraLoad loading={res.loading} error={res.error} onRetry={res.reload} />
        {perms.slice(0, 300).map((p) => (
          <label key={p.id} style={{ display: 'flex', gap: 8, padding: '5px 6px', alignItems: 'flex-start', cursor: 'pointer' }}>
            <input type="checkbox" className="chk" checked={chosen.has(p.id)} onChange={() => toggle(p)} />
            <span style={{ minWidth: 0 }}>
              <span className="mono" style={{ color: 'var(--fg)' }}>{p.value}</span> {p.admin && <span className="flag">{t('arAdminOnly')}</span>}
              <div className="hint" style={{ whiteSpace: 'normal' }}>{p.name}</div>
            </span>
          </label>
        ))}
      </div>
    </div>
  )
}

// ---------- detail ----------

type Tab = 'overview' | 'auth' | 'perms' | 'creds' | 'roles' | 'owners'

function AppDetail({ id, onClose, onChange }: { id: string; onClose(): void; onChange(a: AppReg): void }) {
  const { t, tenant } = useApp()
  const tid = tenant!.tenantId
  const app = useLoad(() => window.api.entra.app(tid, id), [tid, id])
  const [tab, setTab] = useState<Tab>('overview')
  const a = app.data
  const update = (n: AppReg) => {
    app.set(() => n)
    onChange(n)
  }
  return (
    <Modal wide title={a?.displayName || t('loading')} sub={a && <span className="mono">{a.appId}</span>} onClose={onClose}>
      <Tabs
        tabs={[
          ['overview', t('arOverview')],
          ['auth', t('arAuth')],
          ['perms', t('arPermissions')],
          ['creds', t('enSecretsCerts')],
          ['roles', t('arRoles')],
          ['owners', t('arOwners')]
        ]}
        value={tab}
        onChange={setTab}
      />
      <EntraLoad loading={app.loading && !a} error={app.error} onRetry={app.reload} />
      {a && tab === 'overview' && <Overview a={a} />}
      {a && tab === 'auth' && <AuthTab a={a} onSaved={update} />}
      {a && tab === 'perms' && <PermsTab a={a} onSaved={update} />}
      {a && tab === 'creds' && <CredsTab a={a} reload={app.reload} onSaved={update} />}
      {a && tab === 'roles' && <RolesTab a={a} onSaved={update} />}
      {a && tab === 'owners' && <OwnersPanel kind="applications" id={a.id} />}
    </Modal>
  )
}

function Overview({ a }: { a: AppReg }) {
  const { t, tenant } = useApp()
  const rows: [string, string][] = [
    [t('arAppId'), a.appId],
    [t('arObjectId'), a.id],
    [t('arTenantId'), tenant!.tenantId],
    [t('arAuthority'), `https://login.microsoftonline.com/${tenant!.tenantId}/v2.0`]
  ]
  return (
    <div className="creds">
      {rows.map(([k, v]) => (
        <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span className="hint" style={{ width: 120 }}>{k}</span>
          <span style={{ flex: 1, wordBreak: 'break-all' }}>{v}</span>
          <CopyButton text={v} />
        </div>
      ))}
      <div style={{ display: 'flex', gap: 10 }}>
        <span className="hint" style={{ width: 120 }}>{t('arAudience')}</span>
        <span>{t(`aud_${a.audience}` as 'aud_AzureADMyOrg') || a.audience}</span>
      </div>
    </div>
  )
}

function AuthTab({ a, onSaved }: { a: AppReg; onSaved(a: AppReg): void }) {
  const { t, tenant, toast } = useApp()
  const [v, setV] = useState({ web: a.web.join('\n'), spa: a.spa.join('\n'), publicClient: a.publicClient.join('\n') })
  const [busy, setBusy] = useState(false)
  const lines = (s: string) => s.split(/\s+/).map((x) => x.trim()).filter(Boolean)
  return (
    <>
      {(['web', 'spa', 'publicClient'] as const).map((k) => (
        <div className="fld" key={k}>
          <label htmlFor={`r-${k}`}>{t(`plat_${k}`)}</label>
          <textarea id={`r-${k}`} className="mono" value={v[k]} onChange={(e) => setV({ ...v, [k]: e.target.value })} placeholder={k === 'publicClient' ? 'http://localhost' : 'https://'} />
        </div>
      ))}
      <p className="hint">{t('arRedirectHint')}</p>
      <div>
        <button
          className="btn pri"
          disabled={busy}
          onClick={async () => {
            setBusy(true)
            try {
              onSaved(await unwrap(window.api.entra.updateApp(tenant!.tenantId, a.id, { web: lines(v.web), spa: lines(v.spa), publicClient: lines(v.publicClient) })))
              toast(t('saved'))
            } catch (e) {
              toast((e as Error).message)
            } finally {
              setBusy(false)
            }
          }}
        >
          {busy && <span className="spin" />}
          {t('save')}
        </button>
      </div>
    </>
  )
}

function PermsTab({ a, onSaved }: { a: AppReg; onSaved(a: AppReg): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const st = useLoad(() => window.api.entra.permissions(tid, a.id), [tid, a.id, a.permissions.length])
  const [adding, setAdding] = useState(false)
  const [draft, setDraft] = useState<ResourceAccess[]>([])
  const [busy, setBusy] = useState<string | null>(null)
  const s = st.data
  const run = async (key: string, fn: () => Promise<AppPermissionState | void>, msg: string) => {
    setBusy(key)
    try {
      const r = await fn()
      if (r) st.set(() => r)
      toast(msg)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }
  const def = (res: ApiResource | undefined, id: string) => res?.permissions.find((p) => p.id === id)
  const missing = s
    ? a.permissions.some((p) =>
        p.access.some((x) => {
          const res = s.resources.find((r) => r.appId === p.resourceAppId)
          const g = s.granted[p.resourceAppId]
          const d = def(res, x.id)
          return x.type === 'Role' ? !g?.roles.includes(x.id) : !!d && !g?.scopes.includes(d.value)
        })
      )
    : false
  return (
    <>
      <EntraLoad loading={st.loading && !s} error={st.error} onRetry={st.reload} />
      {s && (
        <>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {missing ? <span className="pill warn">{t('arConsentMissing')}</span> : a.permissions.length ? <span className="pill ok">{t('arConsentOk')}</span> : null}
            <span style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
              <button className="btn sm" onClick={() => (setDraft(a.permissions), setAdding(!adding))}>
                {Icon.plus}
                {t('arAddPerm')}
              </button>
              <button className="btn sm pri" disabled={!!busy || !a.permissions.length} onClick={() => confirm(t('arConsentConfirm', tenant!.name)) && run('consent', () => unwrap(window.api.entra.grantConsent(tid, a.id)), t('arConsentDone'))}>
                {busy === 'consent' && <span className="spin" />}
                {t('arGrantConsent')}
              </button>
            </span>
          </div>
          {adding && (
            <>
              <PermissionPicker value={draft} onChange={setDraft} />
              <div>
                <button
                  className="btn pri sm"
                  disabled={busy === 'save'}
                  onClick={async () => {
                    setBusy('save')
                    try {
                      onSaved(await unwrap(window.api.entra.updateApp(tid, a.id, { permissions: draft })))
                      setAdding(false)
                      toast(t('saved'))
                    } catch (e) {
                      toast((e as Error).message)
                    } finally {
                      setBusy(null)
                    }
                  }}
                >
                  {t('save')}
                </button>
              </div>
            </>
          )}
          <div className="panel">
            <table>
              <thead>
                <tr>
                  <th>{t('arPermission')}</th>
                  <th>{t('arType')}</th>
                  <th>{t('arStatus')}</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {a.permissions.flatMap((p) => {
                  const res = s.resources.find((r) => r.appId === p.resourceAppId)
                  const g = s.granted[p.resourceAppId]
                  return p.access.map((x) => {
                    const d = def(res, x.id)
                    const ok = x.type === 'Role' ? g?.roles.includes(x.id) : !!d && g?.scopes.includes(d.value)
                    return (
                      <tr key={p.resourceAppId + x.id}>
                        <td>
                          <div className="mono" style={{ color: 'var(--fg)' }}>{d?.value || x.id}</div>
                          <div className="hint">{res?.displayName || p.resourceAppId}</div>
                        </td>
                        <td>{x.type === 'Role' ? t('arApplication') : t('arDelegated')}</td>
                        <td>{ok ? <span className="pill ok">{t('arGranted')}</span> : <span className="pill warn">{t('arNotGranted')}</span>}</td>
                        <td style={{ textAlign: 'right' }}>
                          <button
                            className="btn sm danger"
                            disabled={!!busy}
                            onClick={() =>
                              confirm(t('arPermRemoveConfirm', d?.value || x.id)) &&
                              run(x.id, async () => {
                                const r = await unwrap(window.api.entra.removePermission(tid, a.id, p.resourceAppId, x.id))
                                onSaved(await unwrap(window.api.entra.app(tid, a.id)))
                                return r
                              }, t('arPermRemoved'))
                            }
                          >
                            {busy === x.id && <span className="spin" />}
                            {t('remove')}
                          </button>
                        </td>
                      </tr>
                    )
                  })
                })}
              </tbody>
            </table>
            {!a.permissions.length && (
              <p className="hint" style={{ padding: 14 }}>
                {t('arNoPerms')}
              </p>
            )}
          </div>
        </>
      )}
    </>
  )
}

function CredsTab({ a, reload, onSaved }: { a: AppReg; reload(): void; onSaved(a: AppReg): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [secret, setSecret] = useState(false)
  const [busy, setBusy] = useState<string | null>(null)
  const certs = a.credentials.filter((c) => c.kind === 'cert')
  const secrets = a.credentials.filter((c) => c.kind === 'secret')

  async function uploadCert() {
    try {
      const c = await unwrap(window.api.entra.pickCert())
      if (!c) return
      if (!confirm(t('arCertConfirm', c.subject, new Date(c.end).toLocaleDateString()))) return
      setBusy('cert')
      onSaved(await unwrap(window.api.entra.addCert(tid, a.id, c, c.fileName)))
      toast(t('arCertAdded'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  const table = (items: typeof a.credentials, hint: string) => (
    <div className="panel">
      <table>
        <thead>
          <tr>
            <th>{t('enDescription')}</th>
            <th>{hint}</th>
            <th>{t('enExpires')}</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {items.map((c) => (
            <tr key={c.keyId}>
              <td>{c.displayName || '—'}</td>
              <td className="mono">{c.kind === 'secret' ? (c.hint ? `${c.hint}…` : '—') : c.hint}</td>
              <td>
                <ExpiryPill end={c.end} />
              </td>
              <td style={{ textAlign: 'right' }}>
                <button
                  className="btn sm danger"
                  disabled={!!busy}
                  onClick={async () => {
                    if (!confirm(t('arCredRemoveConfirm', c.displayName || c.keyId))) return
                    setBusy(c.keyId)
                    try {
                      await unwrap(window.api.entra.removeCredential(tid, 'app', a.id, c))
                      reload()
                      toast(t('arCredRemoved'))
                    } catch (e) {
                      toast((e as Error).message)
                    } finally {
                      setBusy(null)
                    }
                  }}
                >
                  {busy === c.keyId && <span className="spin" />}
                  {t('remove')}
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {!items.length && (
        <p className="hint" style={{ padding: 14 }}>
          {t('arNone')}
        </p>
      )}
    </div>
  )

  return (
    <>
      <div className="panel-h" style={{ padding: 0, border: 0 }}>
        <h2>{t('arSecrets')}</h2>
        <div className="r">
          <button className="btn sm" onClick={() => setSecret(true)}>
            {Icon.plus}
            {t('enNewSecret')}
          </button>
        </div>
      </div>
      {table(secrets, t('arHint'))}
      <div className="panel-h" style={{ padding: 0, border: 0 }}>
        <h2>{t('arCerts')}</h2>
        <div className="r">
          <button className="btn sm" disabled={busy === 'cert'} onClick={uploadCert}>
            {busy === 'cert' ? <span className="spin" /> : Icon.plus}
            {t('arUploadCert')}
          </button>
        </div>
      </div>
      {table(certs, t('arThumbprint'))}
      <p className="hint">{t('arCertHint')}</p>
      {secret && <SecretModal kind="app" objectId={a.id} appName={a.displayName} onClose={() => setSecret(false)} onDone={reload} />}
    </>
  )
}

const newRole = (): AppRole => ({ id: crypto.randomUUID(), displayName: '', description: '', value: '', allowedMemberTypes: ['User'], isEnabled: true })

function RolesTab({ a, onSaved }: { a: AppReg; onSaved(a: AppReg): void }) {
  const { t, tenant, toast } = useApp()
  const [roles, setRoles] = useState<AppRole[]>(a.appRoles)
  const [busy, setBusy] = useState(false)
  const dirty = useMemo(() => JSON.stringify(roles) !== JSON.stringify(a.appRoles), [roles, a.appRoles])
  const set = (i: number, p: Partial<AppRole>) => setRoles(roles.map((r, j) => (j === i ? { ...r, ...p } : r)))
  const members = (r: AppRole) => (r.allowedMemberTypes.length === 2 ? 'both' : r.allowedMemberTypes[0] === 'Application' ? 'apps' : 'users')
  return (
    <>
      <p className="hint">{t('arRolesHint')}</p>
      {roles.map((r, i) => (
        <div key={r.id} style={{ border: '1px solid var(--line)', borderRadius: 8, padding: 10, display: 'grid', gap: 8 }}>
          <div className="row2">
            <input placeholder={t('arRoleName')} value={r.displayName} onChange={(e) => set(i, { displayName: e.target.value })} style={{ padding: '7px 9px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }} />
            <input className="mono" placeholder={t('arRoleValue')} value={r.value || ''} onChange={(e) => set(i, { value: e.target.value.replace(/\s/g, '') })} style={{ padding: '7px 9px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }} />
          </div>
          <input placeholder={t('enDescription')} value={r.description} onChange={(e) => set(i, { description: e.target.value })} style={{ padding: '7px 9px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }} />
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <select
              value={members(r)}
              onChange={(e) => set(i, { allowedMemberTypes: e.target.value === 'both' ? ['User', 'Application'] : e.target.value === 'apps' ? ['Application'] : ['User'] })}
              style={{ padding: '5px 8px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }}
            >
              <option value="users">{t('arRoleUsers')}</option>
              <option value="apps">{t('arRoleApps')}</option>
              <option value="both">{t('arRoleBoth')}</option>
            </select>
            <label className="hint" style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <input type="checkbox" className="chk" checked={r.isEnabled} onChange={(e) => set(i, { isEnabled: e.target.checked })} />
              {t('arRoleEnabled')}
            </label>
            <button className="btn sm danger" style={{ marginLeft: 'auto' }} onClick={() => setRoles(roles.filter((_, j) => j !== i))}>
              {t('remove')}
            </button>
          </div>
        </div>
      ))}
      <div style={{ display: 'flex', gap: 8 }}>
        <button className="btn sm" onClick={() => setRoles([...roles, newRole()])}>
          {Icon.plus}
          {t('arAddRole')}
        </button>
        <button
          className="btn sm pri"
          disabled={busy || !dirty || roles.some((r) => !r.displayName.trim() || !r.value)}
          onClick={async () => {
            setBusy(true)
            try {
              const n = await unwrap(window.api.entra.setAppRoles(tenant!.tenantId, a.id, roles))
              setRoles(n.appRoles)
              onSaved(n)
              toast(t('saved'))
            } catch (e) {
              toast((e as Error).message)
            } finally {
              setBusy(false)
            }
          }}
        >
          {busy && <span className="spin" />}
          {t('save')}
        </button>
      </div>
    </>
  )
}
