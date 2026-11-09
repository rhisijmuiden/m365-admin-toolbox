import { useCallback, useEffect, useState, type ReactNode } from 'react'
import type { ExpiringCred, Tenant } from '@shared/types'
import { AppCtx, initials, unwrap, useApp, type ApiError } from '../ctx'
import { Icon, RowMenu } from '../components/ui'
import { daysLeft, ExpiryPill, pending, SecretModal } from '../components/entra'

type Row = ExpiringCred & { tenant: Tenant }
type TenantState = { loading: boolean; error: ApiError | null; count: number }
const WINDOWS = [30, 60, 90, 0] as const

/** Secrets and certificates of all tenants, sorted by expiry, colored for 30, 60 and 90 days. */
export function SecretsView() {
  const { t, tenants, toast, go } = useApp()
  const [rows, setRows] = useState<Row[]>([])
  const [state, setState] = useState<Record<string, TenantState>>({})
  const [win, setWin] = useState<(typeof WINDOWS)[number]>(90)
  const [q, setQ] = useState('')
  const [secret, setSecret] = useState<Row | null>(null)

  const loadOne = useCallback(async (x: Tenant) => {
    setState((s) => ({ ...s, [x.tenantId]: { loading: true, error: null, count: s[x.tenantId]?.count || 0 } }))
    try {
      const list = await unwrap(window.api.entra.credentials(x.tenantId))
      setRows((r) => [...r.filter((y) => y.tenant.tenantId !== x.tenantId), ...list.map((c) => ({ ...c, tenant: x }))])
      setState((s) => ({ ...s, [x.tenantId]: { loading: false, error: null, count: list.length } }))
    } catch (e) {
      setState((s) => ({ ...s, [x.tenantId]: { loading: false, error: e as ApiError, count: 0 } }))
    }
  }, [])
  const loadAll = useCallback(() => tenants.forEach(loadOne), [tenants, loadOne])
  useEffect(loadAll, [loadAll])

  const left = (r: Row) => daysLeft(r.end) ?? Infinity
  const shown = rows
    .filter((r) => !win || left(r) <= win)
    .filter((r) => (r.appName + r.tenant.name + (r.displayName || '')).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => left(a) - left(b))
  const n = (lo: number, hi: number) => rows.filter((r) => left(r) >= lo && left(r) <= hi).length
  const expired = rows.filter((r) => left(r) < 0).length
  const busy = Object.values(state).some((s) => s.loading)
  const failed = tenants.filter((x) => state[x.tenantId]?.error)

  async function fix(x: Tenant, e: ApiError) {
    try {
      if (e.code === 'NEEDS_CONSENT') await unwrap(window.api.entra.consent(x.tenantId))
      else await unwrap(window.api.reauthTenant(x.tenantId))
      loadOne(x)
    } catch (err) {
      toast((err as Error).message)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)} · {t('entra')}
          </div>
          <h1>{t('secrets')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('scSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={() => (setRows([]), loadAll())} title={t('refresh')} aria-label={t('refresh')}>
            {busy ? <span className="spin" /> : Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        {failed.map((x) => {
          const e = state[x.tenantId].error!
          const fixable = e.code === 'NEEDS_CONSENT' || e.code === 'NEEDS_SIGNIN'
          return (
            <div className={`banner ${fixable ? '' : 'bad'}`} key={x.tenantId}>
              <span>
                <b>{x.name}:</b> {e.code === 'NEEDS_CONSENT' ? t('enConsentT') : e.code === 'NEEDS_SIGNIN' ? t('needsSignIn') : e.message}
              </span>
              {fixable && (
                <button className="btn sm" onClick={() => fix(x, e)}>
                  {e.code === 'NEEDS_CONSENT' ? t('grantAccess') : t('signInAgain')}
                </button>
              )}
            </div>
          )
        })}
        <div className="tiles">
          <div className={`tile ${expired ? 'bad' : ''}`}>
            <div className="k">{t('enExpiredN')}</div>
            <div className="v">{expired}</div>
            <div className="s">{t('scExpiredD')}</div>
          </div>
          <div className={`tile ${n(0, 30) ? 'bad' : ''}`}>
            <div className="k">{t('enWithin', 30)}</div>
            <div className="v">{n(0, 30)}</div>
            <div className="s">{t('scRenewNow')}</div>
          </div>
          <div className={`tile ${n(31, 60) ? 'warn' : ''}`}>
            <div className="k">{t('scBetween', 31, 60)}</div>
            <div className="v">{n(31, 60)}</div>
            <div className="s">{t('scPlan')}</div>
          </div>
          <div className={`tile ${n(61, 90) ? 'warn' : ''}`}>
            <div className="k">{t('scBetween', 61, 90)}</div>
            <div className="v">{n(61, 90)}</div>
            <div className="s">{t('scTenants', tenants.length - failed.length, tenants.length)}</div>
          </div>
        </div>
        <div className="panel">
          <div className="panel-h">
            <div className="seg">
              {WINDOWS.map((w) => (
                <button key={w} className={win === w ? 'on' : ''} onClick={() => setWin(w)}>
                  {w ? t('enWithin', w) : t('scAll')}
                </button>
              ))}
            </div>
            <span className="hint">{t('scCount', shown.length)}</span>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('scTenant')}</th>
                  <th>{t('scApp')}</th>
                  <th>{t('arType')}</th>
                  <th>{t('enDescription')}</th>
                  <th>{t('enExpires')}</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {shown.map((r) => (
                  <tr key={r.tenant.tenantId + r.objectId + r.keyId}>
                    <td>
                      <span style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
                        <span className="t-av" style={{ background: r.tenant.color, width: 22, height: 22, fontSize: 10 }}>{initials(r.tenant.name)}</span>
                        {r.tenant.name}
                      </span>
                    </td>
                    <td>
                      <div className="u-n">{r.appName}</div>
                      <div className="hint">{r.source === 'app' ? t('appRegs') : t('apps')}</div>
                    </td>
                    <td>{r.kind === 'secret' ? t('scSecret') : t('scCert')}</td>
                    <td>{r.displayName || <span className="mono">{r.hint}</span>}</td>
                    <td>
                      <ExpiryPill end={r.end} />
                    </td>
                    <td style={{ width: 40 }}>
                      <RowMenu
                        items={[
                          [t('scNewSecret'), () => setSecret(r)],
                          [
                            t('arOpen'),
                            () => {
                              if (r.source === 'app') pending.appId = r.objectId
                              else pending.spId = r.objectId
                              go(r.source === 'app' ? 'appRegs' : 'apps', r.tenant.tenantId)
                            }
                          ]
                        ]}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!shown.length && !busy && (
              <p className="hint" style={{ padding: 16 }}>
                {t('scNone')}
              </p>
            )}
          </div>
        </div>
        <p className="hint">{t('scHint')}</p>
      </div>
      {secret && (
        <TenantScope tenant={secret.tenant}>
          <SecretModal kind={secret.source} objectId={secret.objectId} appName={`${secret.appName} · ${secret.tenant.name}`} onClose={() => setSecret(null)} onDone={() => loadOne(secret.tenant)} />
        </TenantScope>
      )}
    </>
  )
}

/** Runs a tenant-bound component (the secret dialog) for another tenant than the selected one. */
function TenantScope({ tenant, children }: { tenant: Tenant; children: ReactNode }) {
  const ctx = useApp()
  return <AppCtx.Provider value={{ ...ctx, tenant }}>{children}</AppCtx.Provider>
}
