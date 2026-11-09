import { useState } from 'react'
import type { Sku } from '@shared/types'
import { initials, useApp, useLoad } from '../ctx'
import { Icon, LoadState } from '../components/ui'

type Row = { tenantId: string; skus?: Sku[]; error?: string; needsSignIn?: boolean }

export function LicensesView() {
  const { t, tenant, tenants } = useApp()
  const tid = tenant!.tenantId
  const skus = useLoad(() => window.api.listSkus(tid), [tid])
  const [allRows, setAllRows] = useState<Row[] | null>(null)
  const [busyAll, setBusyAll] = useState(false)

  async function checkAll() {
    setBusyAll(true)
    const rows = await Promise.all(
      tenants.map(async (x): Promise<Row> => {
        const r = await window.api.listSkus(x.tenantId)
        return r.ok ? { tenantId: x.tenantId, skus: r.data } : { tenantId: x.tenantId, error: r.error, needsSignIn: r.code === 'NEEDS_SIGNIN' }
      })
    )
    setAllRows(rows)
    setBusyAll(false)
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('m365')}
          </div>
          <h1>{t('licenses')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={skus.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ margin: '0 0 14px', maxWidth: '75ch' }}>
          {t('licIntro')}
        </p>
        <LoadState loading={skus.loading && !skus.data} error={skus.error} />
        {skus.data && (
          <div className="panel">
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('sku')}</th>
                    <th>{t('total')}</th>
                    <th>{t('assigned')}</th>
                    <th>{t('avail')}</th>
                    <th>{t('usage')}</th>
                    <th>{t('status')}</th>
                  </tr>
                </thead>
                <tbody>
                  {skus.data.map((s) => (
                    <SkuRow key={s.skuId} s={s} />
                  ))}
                  {!skus.data.length && (
                    <tr>
                      <td colSpan={6} className="hint">
                        {t('noSkus')}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="panel" style={{ marginTop: 16 }}>
          <div className="panel-h">
            <h2>{t('allTenants')}</h2>
            <span className="hint">{t('allTenantsHint')}</span>
            <div className="r">
              <button className="btn sm" disabled={busyAll} onClick={checkAll}>
                {busyAll && <span className="spin" />}
                {t('loadAll')}
              </button>
            </div>
          </div>
          {allRows && (
            <div className="tbl-wrap">
              <table>
                <tbody>
                  {allRows.map((r) => {
                    const x = tenants.find((y) => y.tenantId === r.tenantId)!
                    const short = r.skus?.filter((s) => s.shortage > 0) || []
                    const n = short.reduce((a, s) => a + s.shortage, 0)
                    return (
                      <tr key={r.tenantId}>
                        <td>
                          <div className="u">
                            <div className="t-av" style={{ background: x.color }}>
                              {initials(x.name)}
                            </div>
                            <div>
                              <div className="u-n">{x.name}</div>
                              <div className="mono">{x.defaultDomain}</div>
                            </div>
                          </div>
                        </td>
                        <td>
                          <span className={`chip ${x.mode}`}>{t(x.mode)}</span>
                        </td>
                        <td style={{ whiteSpace: 'normal' }}>
                          {r.error ? (
                            <span className="pill warn" title={r.error}>
                              {r.needsSignIn ? t('needsSignIn') : t('error')}
                            </span>
                          ) : n ? (
                            <>
                              <span className="pill bad">{t('shortTxt', n)}</span>{' '}
                              <span className="hint">{short.map((s) => s.name).join(', ')}</span>
                            </>
                          ) : (
                            <span className="pill ok">{t('ok')}</span>
                          )}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </>
  )
}

function SkuRow({ s }: { s: Sku }) {
  const { t } = useApp()
  const pct = s.enabled ? Math.min(100, Math.round((s.consumed / s.enabled) * 100)) : 100
  const cls = s.shortage ? 'full' : pct >= 85 ? 'hi' : ''
  return (
    <tr>
      <td>
        <div className="u-n">{s.name}</div>
        <div className="mono">{s.skuPartNumber}</div>
      </td>
      <td>
        {s.enabled}
        {s.warning > 0 && <span className="hint"> (+{s.warning} ⚠)</span>}
      </td>
      <td>{s.consumed}</td>
      <td>{s.available}</td>
      <td>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <div className={`bar ${cls}`}>
            <i style={{ width: `${pct}%` }} />
          </div>
          <span className="mono">{pct}%</span>
        </div>
      </td>
      <td>
        {s.shortage ? (
          <span className="pill bad">{t('shortTxt', s.shortage)}</span>
        ) : pct >= 85 ? (
          <span className="pill warn">{t('nearly')}</span>
        ) : (
          <span className="pill ok">{t('ok')}</span>
        )}
      </td>
    </tr>
  )
}
