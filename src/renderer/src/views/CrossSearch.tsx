import { useState } from 'react'
import type { SearchHit, SearchKind, Tenant } from '@shared/types'
import { setNextSearch, unwrap, useApp, type ApiError } from '../ctx'
import { Icon } from '../components/ui'
import { pending } from '../components/entra'
import { eachTenant, errText, fixable, fixTenant, TenantTag } from '../components/cross'

type Row = SearchHit & { tenant: Tenant }
type Filter = 'all' | 'people' | 'group' | 'apps'
const FILTERS: Filter[] = ['all', 'people', 'group', 'apps']
const inFilter = (k: SearchKind, f: Filter) =>
  f === 'all' || (f === 'people' ? k === 'user' || k === 'guest' : f === 'apps' ? k === 'app' || k === 'appReg' : k === f)

/** Finds a user, mail address, group or app in every connected tenant at once. */
export function CrossSearchView() {
  const { t, tenants, toast, go } = useApp()
  const [q, setQ] = useState('')
  const [done, setDone] = useState('')
  const [rows, setRows] = useState<Row[]>([])
  const [errors, setErrors] = useState<Record<string, ApiError>>({})
  const [busy, setBusy] = useState(0)
  const [filter, setFilter] = useState<Filter>('all')

  async function searchOne(x: Tenant, term: string) {
    setBusy((n) => n + 1)
    try {
      const hits = await unwrap(window.api.cross.search(x.tenantId, term))
      setRows((r) => [...r.filter((y) => y.tenant.tenantId !== x.tenantId), ...hits.map((h) => ({ ...h, tenant: x }))])
      setErrors(({ [x.tenantId]: _, ...rest }) => rest)
    } catch (e) {
      setErrors((s) => ({ ...s, [x.tenantId]: e as ApiError }))
    } finally {
      setBusy((n) => n - 1)
    }
  }

  function search(e?: React.FormEvent) {
    e?.preventDefault()
    const term = q.trim()
    if (term.length < 2) return
    setRows([])
    setErrors({})
    setDone(term)
    eachTenant(tenants, 6, (x) => searchOne(x, term))
  }

  function open(r: Row) {
    if (r.kind === 'user' || r.kind === 'guest') {
      setNextSearch(r.sub || r.name)
      go('users', r.tenant.tenantId)
    } else if (r.kind === 'group') {
      setNextSearch(r.name)
      go('groups', r.tenant.tenantId)
    } else if (r.kind === 'app') {
      pending.spId = r.id
      go('apps', r.tenant.tenantId)
    } else {
      pending.appId = r.id
      go('appRegs', r.tenant.tenantId)
    }
  }

  const shown = rows
    .filter((r) => inFilter(r.kind, filter))
    .sort((a, b) => a.tenant.name.localeCompare(b.tenant.name) || a.kind.localeCompare(b.kind) || a.name.localeCompare(b.name))
  const count = (f: Filter) => rows.filter((r) => inFilter(r.kind, f)).length
  const hitTenants = new Set(rows.map((r) => r.tenant.tenantId)).size

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('search')}</h1>
        </div>
      </header>
      <div className="content stack">
        <form className="panel" onSubmit={search} style={{ padding: 14, display: 'flex', gap: 10, alignItems: 'center' }}>
          <label className="search" style={{ flex: 1 }}>
            {Icon.search}
            <input autoFocus placeholder={t('xsPlaceholder')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn pri" type="submit" disabled={q.trim().length < 2 || busy > 0}>
            {busy > 0 && <span className="spin" />}
            {t('xsGo')}
          </button>
        </form>
        {Object.entries(errors).map(([id, e]) => {
          const x = tenants.find((y) => y.tenantId === id)!
          return (
            <div className={`banner ${fixable(e) ? '' : 'bad'}`} key={id}>
              <span>
                <b>{x.name}:</b> {errText(t, e)}
              </span>
              {fixable(e) && (
                <button
                  className="btn sm"
                  onClick={() =>
                    fixTenant(x, e)
                      .then(() => searchOne(x, done))
                      .catch((err) => toast((err as Error).message))
                  }
                >
                  {t('signInAgain')}
                </button>
              )}
            </div>
          )
        })}
        {done && (
          <div className="panel">
            <div className="panel-h">
              <div className="seg">
                {FILTERS.map((f) => (
                  <button key={f} className={filter === f ? 'on' : ''} onClick={() => setFilter(f)}>
                    {t(`xsF_${f}`)} ({count(f)})
                  </button>
                ))}
              </div>
              <span className="hint">{busy ? t('xsBusy', tenants.length - busy, tenants.length) : t('xsCount', rows.length, hitTenants)}</span>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('scTenant')}</th>
                    <th>{t('arType')}</th>
                    <th>{t('name')}</th>
                    <th>{t('xsAddress')}</th>
                    <th>{t('status')}</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {shown.map((r) => (
                    <tr key={r.tenant.tenantId + r.kind + r.id}>
                      <td>
                        <TenantTag tenant={r.tenant} />
                      </td>
                      <td>{t(`xsK_${r.kind}`)}</td>
                      <td>
                        <div className="u-n">{r.name}</div>
                        {r.detail && <div className="hint">{r.kind === 'group' ? t(`kind_${r.detail}` as 'kind_m365') : r.detail}</div>}
                      </td>
                      <td>
                        <span className="mono">{r.sub}</span>
                        {r.matched && <div className="hint">{t('xsAlias', r.matched)}</div>}
                      </td>
                      <td>{r.enabled == null ? '' : <span className={`pill ${r.enabled ? 'ok' : 'off'}`}>{r.enabled ? t('active') : t('blocked')}</span>}</td>
                      <td style={{ width: 90 }}>
                        <button className="btn sm" onClick={() => open(r)}>
                          {t('arOpen')}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!shown.length && !busy && <p className="hint" style={{ padding: 16 }}>{t('xsNone', done)}</p>}
            </div>
          </div>
        )}
        <p className="hint">{t('xsHint')}</p>
      </div>
    </>
  )
}
