import { useApp, useLoad } from '../ctx'
import { fmtDate, type Key } from '../i18n'
import { Icon } from '../components/ui'
import { ExoLoad } from '../components/exo'

export function MailFlowView() {
  const { t, lang, tenant } = useApp()
  const tid = tenant!.tenantId
  const flow = useLoad(() => window.api.exo.mailFlow(tid), [tid])
  const d = flow.data

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('exchange')}
          </div>
          <h1>{t('mailflow')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={flow.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ marginTop: 0 }}>
          {t('mfReadOnly')}
        </p>
        <ExoLoad loading={flow.loading && !d} error={flow.error} onRetry={flow.reload} />
        {d && (
          <>
            <div className="panel" style={{ marginBottom: 16 }}>
              <div className="panel-h">
                <h2>{t('mfDomains')}</h2>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('domain')}</th>
                      <th>{t('typeL')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {d.domains.map((x) => (
                      <tr key={x.name}>
                        <td className="mono">{x.name}</td>
                        <td>{(t(`dt_${x.type}` as Key) as string | undefined) || x.type}</td>
                        <td>{x.isDefault && <span className="pill ok">{t('mfDefault')}</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <h2>{t('mfRules')}</h2>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style={{ width: 70 }}>{t('mfPriority')}</th>
                      <th>{t('name')}</th>
                      <th>{t('status')}</th>
                      <th>{t('mfMode')}</th>
                      <th>{t('mfModified')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...d.rules]
                      .sort((a, b) => a.priority - b.priority)
                      .map((r) => (
                        <tr key={r.name} className={r.state === 'Enabled' ? '' : 'muted'}>
                          <td className="mono">{r.priority}</td>
                          <td>
                            <div className="u-n">{r.name}</div>
                            <div className="hint" style={{ maxWidth: 560, whiteSpace: 'normal' }}>
                              {r.description}
                            </div>
                          </td>
                          <td>{r.state === 'Enabled' ? <span className="pill ok">{t('active')}</span> : <span className="pill off">{t('disabledL')}</span>}</td>
                          <td className="mono">{r.mode}</td>
                          <td className="mono">{fmtDate(r.modified, lang, '—')}</td>
                        </tr>
                      ))}
                  </tbody>
                </table>
                {!d.rules.length && <p className="hint" style={{ padding: 16 }}>{t('mfNoRules')}</p>}
              </div>
            </div>
          </>
        )}
      </div>
    </>
  )
}
