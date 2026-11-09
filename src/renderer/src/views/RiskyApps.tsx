import { useState } from 'react'
import type { AppGrants, GrantedPermission } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon } from '../components/ui'
import { EntraLoad } from '../components/entra'

/** Apps with permissions on mail, files or the whole directory, with the option to withdraw them or block the app. */
export function RiskyAppsView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.grants(tid), [tid])
  const [q, setQ] = useState('')
  const [all, setAll] = useState(false)
  const [ms, setMs] = useState(false)
  const [open, setOpen] = useState<string | null>(null)
  const [busy, setBusy] = useState<string | null>(null)

  const rows = (list.data || []).filter((a) => a.permissions.length)
  const risky = (a: AppGrants) => a.permissions.some((p) => p.risky)
  const shown = rows
    .filter((a) => ms || !a.microsoft)
    .filter((a) => all || risky(a))
    .filter((a) => (a.displayName + (a.publisher || '') + a.permissions.map((p) => p.value).join(' ')).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => Number(risky(b)) - Number(risky(a)) || b.permissions.filter((p) => p.risky).length - a.permissions.filter((p) => p.risky).length || a.displayName.localeCompare(b.displayName))
  const third = rows.filter((a) => !a.microsoft)

  async function revoke(a: AppGrants, p: GrantedPermission) {
    if (!confirm(t('raRevokeConfirm', p.value, a.displayName))) return
    setBusy(p.grantId + p.value)
    try {
      await unwrap(window.api.entra.revokeGrant(tid, a.spId, p))
      list.set((d) =>
        d.map((x) => (x.spId === a.spId ? { ...x, permissions: x.permissions.filter((y) => !(y.grantId === p.grantId && (p.kind === 'application' || y.value === p.value))) } : x))
      )
      toast(t('raRevoked', p.value))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  async function block(a: AppGrants) {
    if (!confirm(t('eaDisableConfirm', a.displayName))) return
    setBusy(a.spId)
    try {
      await unwrap(window.api.entra.updateEnterpriseApp(tid, a.spId, { enabled: false }))
      list.set((d) => d.map((x) => (x.spId === a.spId ? { ...x, enabled: false } : x)))
      toast(t('raBlocked', a.displayName))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('riskyApps')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('raSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} text={t('raLoading')} />
        {list.data && (
          <>
            <div className="tiles">
              <div className={`tile ${third.filter(risky).length ? 'bad' : ''}`}>
                <div className="k">{t('raRiskyApps')}</div>
                <div className="v">{third.filter(risky).length}</div>
                <div className="s">{t('raRiskyAppsD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('raAppPerms')}</div>
                <div className="v">{third.filter((a) => a.permissions.some((p) => p.kind === 'application')).length}</div>
                <div className="s">{t('raAppPermsD')}</div>
              </div>
              <div className={`tile ${third.some((a) => a.userConsents) ? 'warn' : ''}`}>
                <div className="k">{t('raUserConsent')}</div>
                <div className="v">{third.filter((a) => a.userConsents).length}</div>
                <div className="s">{t('raUserConsentD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('raWithGrants')}</div>
                <div className="v">{third.length}</div>
                <div className="s">{t('raWithGrantsD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={all} onChange={(e) => setAll(e.target.checked)} />
                  {t('raShowAll')}
                </label>
                <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={ms} onChange={(e) => setMs(e.target.checked)} />
                  {t('raShowMs')}
                </label>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('scApp')}</th>
                      <th>{t('raRiskyPerms')}</th>
                      <th>{t('raConsent')}</th>
                      <th>{t('eaSignIn')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((a) => {
                      const r = a.permissions.filter((p) => p.risky)
                      const isOpen = open === a.spId
                      return [
                        <tr key={a.spId}>
                          <td>
                            <button className="linkish" onClick={() => setOpen(isOpen ? null : a.spId)}>
                              <div className="u-n">{a.displayName}</div>
                              <div className="hint">{a.publisher || t('raUnknownPub')}</div>
                            </button>
                          </td>
                          <td style={{ whiteSpace: 'normal', maxWidth: 380 }}>
                            {r.slice(0, 6).map((p) => (
                              <span key={p.kind + p.value} className={`flag ${p.kind === 'application' ? 'external' : ''}`}>
                                {p.value}
                              </span>
                            ))}
                            {r.length > 6 && <span className="hint">+{r.length - 6}</span>}
                            {!r.length && <span className="hint">{t('raNoRisky', a.permissions.length)}</span>}
                          </td>
                          <td>
                            {a.permissions.some((p) => p.kind === 'application') && <span className="lic-tag">{t('raAppPerm')}</span>}
                            {a.permissions.some((p) => p.consent === 'admin') && <span className="lic-tag">{t('raAdmin')}</span>}
                            {a.userConsents > 0 && <span className="lic-tag">{t('raUsersN', a.userConsents)}</span>}
                          </td>
                          <td>{a.enabled ? <span className="pill ok">{t('eaOn')}</span> : <span className="pill bad">{t('eaOff')}</span>}</td>
                          <td style={{ textAlign: 'right' }}>
                            <div style={{ display: 'inline-flex', gap: 6 }}>
                              <button className="btn sm" onClick={() => setOpen(isOpen ? null : a.spId)}>
                                {isOpen ? t('raHide') : t('raDetails')}
                              </button>
                              {a.enabled && (
                                <button className="btn sm danger" disabled={busy === a.spId} onClick={() => block(a)}>
                                  {t('raBlock')}
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>,
                        isOpen && (
                          <tr key={a.spId + '-d'}>
                            <td colSpan={5} style={{ background: 'var(--surface-2)', whiteSpace: 'normal' }}>
                              <table>
                                <tbody>
                                  {a.permissions.map((p) => (
                                    <tr key={p.grantId + p.value + p.consent}>
                                      <td>
                                        <span className="mono" style={{ color: 'var(--fg)' }}>{p.value}</span> {p.risky && <span className="flag external">{t('raRisky')}</span>}
                                      </td>
                                      <td className="hint">{p.resource}</td>
                                      <td>{p.kind === 'application' ? t('raAppPerm') : p.consent === 'admin' ? t('raAdmin') : t('raUser')}</td>
                                      <td style={{ textAlign: 'right' }}>
                                        <button className="btn sm danger" disabled={busy === p.grantId + p.value} onClick={() => revoke(a, p)}>
                                          {busy === p.grantId + p.value && <span className="spin" />}
                                          {t('raRevoke')}
                                        </button>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </td>
                          </tr>
                        )
                      ]
                    })}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('raNone')}
                  </p>
                )}
              </div>
            </div>
            <p className="hint">{t('raHint')}</p>
          </>
        )}
      </div>
    </>
  )
}
