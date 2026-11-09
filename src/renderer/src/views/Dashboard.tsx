import { useCallback, useEffect, useState } from 'react'
import type { ExpiringCred, Sku, Team, Tenant } from '@shared/types'
import { unwrap, useApp, type ApiError } from '../ctx'
import { Icon } from '../components/ui'
import { daysLeft, ExpiryPill, pending } from '../components/entra'
import { eachTenant, errText, fixable, fixTenant, TenantTag } from '../components/cross'

type Part = 'skus' | 'creds' | 'quarantine' | 'teams'
type Cell<T> = { loading: boolean; data?: T; error?: ApiError }
type Row = { skus: Cell<Sku[]>; creds: Cell<ExpiringCred[]>; quarantine: Cell<number>; teams: Cell<Team[]> }

const QUARANTINE_KEY = 'dash.quarantine'
/** Exchange returns one page of 500; more than that shows as 500+. */
const PAGE = 500
const SOON = 30

function readFlag() {
  try {
    return localStorage.getItem(QUARANTINE_KEY) !== '0'
  } catch {
    return true
  }
}

/** The morning overview: license shortages, expiring secrets, quarantine and teams without owner, for all tenants. */
export function DashboardView() {
  const { t, tenants, toast, go } = useApp()
  const [rows, setRows] = useState<Record<string, Row>>({})
  const [withQ, setWithQ] = useState(readFlag)

  const put = useCallback(<K extends Part>(id: string, k: K, c: Row[K]) => setRows((r) => ({ ...r, [id]: { ...(r[id] || ({} as Row)), [k]: c } })), [])

  const loadPart = useCallback(
    async (x: Tenant, k: Part) => {
      const id = x.tenantId
      put(id, k, { loading: true })
      try {
        if (k === 'skus') put(id, k, { loading: false, data: await unwrap(window.api.listSkus(id)) })
        else if (k === 'creds') put(id, k, { loading: false, data: await unwrap(window.api.entra.credentials(id)) })
        else if (k === 'teams') put(id, k, { loading: false, data: await unwrap(window.api.teams.list(id)) })
        else {
          const since = new Date(Date.now() - 86400000).toISOString()
          put(id, k, { loading: false, data: (await unwrap(window.api.sec.quarantine(id, { start: since }))).length })
        }
      } catch (e) {
        put(id, k, { loading: false, error: e as ApiError })
      }
    },
    [put]
  )

  const loadAll = useCallback(() => {
    // Graph parts run side by side; the Exchange part starts one PowerShell session per tenant, so it goes two at a time.
    eachTenant(tenants, 4, async (x) => {
      await Promise.all((['skus', 'creds', 'teams'] as Part[]).map((k) => loadPart(x, k)))
    })
    if (withQ) eachTenant(tenants, 2, (x) => loadPart(x, 'quarantine'))
  }, [tenants, withQ, loadPart])
  useEffect(loadAll, [loadAll])

  function toggleQ(v: boolean) {
    setWithQ(v)
    try {
      localStorage.setItem(QUARANTINE_KEY, v ? '1' : '0')
    } catch {
      // Not remembered, that's fine.
    }
  }

  async function fix(x: Tenant, k: Part, e: ApiError) {
    try {
      await fixTenant(x, e, () => unwrap(k === 'teams' ? window.api.teams.consent(x.tenantId) : window.api.entra.consent(x.tenantId)))
      loadPart(x, k)
    } catch (err) {
      toast((err as Error).message)
    }
  }

  const all = tenants.map((x) => ({ x, r: rows[x.tenantId] }))
  const short = all.flatMap(({ x, r }) => (r?.skus?.data || []).filter((s) => s.shortage > 0).map((s) => ({ x, s })))
  const soon = all
    .flatMap(({ x, r }) => (r?.creds?.data || []).map((c) => ({ x, c, left: daysLeft(c.end) ?? Infinity })))
    .filter((c) => c.left <= SOON)
    .sort((a, b) => a.left - b.left)
  const ownerless = all.flatMap(({ x, r }) => (r?.teams?.data || []).filter((tm) => !tm.owners.length && tm.archived !== true).map((tm) => ({ x, tm })))
  const quarantined = all.reduce((n, { r }) => n + (r?.quarantine?.data || 0), 0)
  const busy = all.some(({ r }) => r && Object.values(r).some((c) => c?.loading))

  const cell = <K extends Part>(x: Tenant, k: K, value: (d: NonNullable<Row[K]['data']>) => number, view: string, label?: (n: number) => string) => {
    const c = rows[x.tenantId]?.[k]
    if (!c || c.loading) return <span className="spin" />
    if (c.error)
      return fixable(c.error) ? (
        <button className="btn sm" onClick={() => fix(x, k, c.error!)} title={errText(t, c.error)}>
          {c.error.code === 'NEEDS_CONSENT' ? t('grantAccess') : c.error.code === 'NEEDS_EXO_SIGNIN' ? t('exoSignInBtn') : t('signInAgain')}
        </button>
      ) : (
        <span className="pill off" title={c.error.message}>
          {t('dbNoAccess')}
        </span>
      )
    const n = value(c.data as NonNullable<Row[K]['data']>)
    return (
      <button className="linkish" onClick={() => go(view, x.tenantId)}>
        <span className={`pill ${n ? (k === 'quarantine' ? 'warn' : 'bad') : 'ok'}`}>{label ? label(n) : n}</span>
      </button>
    )
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('dashboard')}</h1>
        </div>
        <div className="top-r">
          <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <input type="checkbox" className="chk" checked={withQ} onChange={(e) => toggleQ(e.target.checked)} />
            {t('dbWithQ')}
          </label>
          <button className="btn" onClick={() => (setRows({}), loadAll())} title={t('refresh')} aria-label={t('refresh')}>
            {busy ? <span className="spin" /> : Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content stack">
        <div className="tiles">
          <div className={`tile ${short.length ? 'bad' : ''}`}>
            <div className="k">{t('short')}</div>
            <div className="v">{short.reduce((n, { s }) => n + s.shortage, 0)}</div>
            <div className="s">{t('dbShortD', new Set(short.map(({ x }) => x.tenantId)).size)}</div>
          </div>
          <div className={`tile ${soon.length ? 'bad' : ''}`}>
            <div className="k">{t('dbSecrets', SOON)}</div>
            <div className="v">{soon.length}</div>
            <div className="s">{t('dbSecretsD', soon.filter((c) => c.left < 0).length)}</div>
          </div>
          <div className={`tile ${quarantined ? 'warn' : ''}`}>
            <div className="k">{t('dbQuarantine')}</div>
            <div className="v">{withQ ? quarantined : '—'}</div>
            <div className="s">{withQ ? t('dbQuarantineD') : t('dbQuarantineOff')}</div>
          </div>
          <div className={`tile ${ownerless.length ? 'warn' : ''}`}>
            <div className="k">{t('dbOwnerless')}</div>
            <div className="v">{ownerless.length}</div>
            <div className="s">{t('dbOwnerlessD')}</div>
          </div>
        </div>

        <div className="panel">
          <div className="panel-h">
            <b>{t('dbPerTenant')}</b>
            <span className="hint">{t('dbClick')}</span>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('scTenant')}</th>
                  <th>{t('short')}</th>
                  <th>{t('dbSecrets', SOON)}</th>
                  {withQ && <th>{t('dbQuarantine')}</th>}
                  <th>{t('dbOwnerless')}</th>
                </tr>
              </thead>
              <tbody>
                {tenants.map((x) => (
                  <tr key={x.tenantId}>
                    <td>
                      <TenantTag tenant={x} />
                    </td>
                    <td>{cell(x, 'skus', (d) => d.reduce((n, s) => n + s.shortage, 0), 'licenses')}</td>
                    <td>{cell(x, 'creds', (d) => d.filter((c) => (daysLeft(c.end) ?? Infinity) <= SOON).length, 'secrets')}</td>
                    {withQ && <td>{cell(x, 'quarantine', (d) => d, 'quarantine', (n) => (n >= PAGE ? `${PAGE}+` : String(n)))}</td>}
                    <td>{cell(x, 'teams', (d) => d.filter((tm) => !tm.owners.length && tm.archived !== true).length, 'teamsList')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="row2" style={{ alignItems: 'start' }}>
          <div className="panel">
            <div className="panel-h">
              <b>{t('short')}</b>
            </div>
            <div className="tbl-wrap">
              <table>
                <tbody>
                  {short.map(({ x, s }) => (
                    <tr key={x.tenantId + s.skuId}>
                      <td>
                        <TenantTag tenant={x} />
                      </td>
                      <td>{s.name}</td>
                      <td style={{ textAlign: 'right' }}>
                        <button className="linkish" onClick={() => go('licenses', x.tenantId)}>
                          <span className="pill bad">{t('dbShortN', s.shortage)}</span>
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!short.length && <p className="hint" style={{ padding: 16 }}>{t('dbNoShort')}</p>}
            </div>
          </div>
          <div className="panel">
            <div className="panel-h">
              <b>{t('dbSecrets', SOON)}</b>
            </div>
            <div className="tbl-wrap">
              <table>
                <tbody>
                  {soon.slice(0, 12).map(({ x, c }) => (
                    <tr key={x.tenantId + c.objectId + c.keyId}>
                      <td>
                        <TenantTag tenant={x} />
                      </td>
                      <td>
                        <button
                          className="linkish"
                          onClick={() => {
                            if (c.source === 'app') pending.appId = c.objectId
                            else pending.spId = c.objectId
                            go(c.source === 'app' ? 'appRegs' : 'apps', x.tenantId)
                          }}
                        >
                          {c.appName}
                        </button>
                      </td>
                      <td style={{ textAlign: 'right' }}>
                        <ExpiryPill end={c.end} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!soon.length && <p className="hint" style={{ padding: 16 }}>{t('dbNoSecrets')}</p>}
              {soon.length > 12 && (
                <p className="hint" style={{ padding: '8px 16px' }}>
                  <button className="linkish" onClick={() => go('secrets')}>
                    {t('dbMore', soon.length - 12)}
                  </button>
                </p>
              )}
            </div>
          </div>
        </div>

        {ownerless.length > 0 && (
          <div className="panel">
            <div className="panel-h">
              <b>{t('dbOwnerless')}</b>
            </div>
            <div className="tbl-wrap">
              <table>
                <tbody>
                  {ownerless.map(({ x, tm }) => (
                    <tr key={x.tenantId + tm.id}>
                      <td style={{ width: 260 }}>
                        <TenantTag tenant={x} />
                      </td>
                      <td>
                        <button className="linkish" onClick={() => go('teamsList', x.tenantId)}>
                          {tm.displayName}
                        </button>
                      </td>
                      <td className="hint">{tm.mail}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </>
  )
}
