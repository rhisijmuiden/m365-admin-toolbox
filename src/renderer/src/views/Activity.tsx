import { useState } from 'react'
import { useApp, useLoad, unwrap } from '../ctx'
import { Icon, LoadState } from '../components/ui'
import { fmtWhen } from '../components/security'
import { TenantSelect } from '../components/cross'

type Show = 'all' | 'failed'

/** The local log of every change made through the app, filterable per tenant and exportable to CSV. */
export function ActivityView() {
  const { t, lang, tenants, toast } = useApp()
  const [tid, setTid] = useState('')
  const [q, setQ] = useState('')
  const [show, setShow] = useState<Show>('all')
  const log = useLoad(() => window.api.activity.list(tid || undefined), [tid])

  const shown = (log.data || [])
    .filter((e) => show === 'all' || !e.ok)
    .filter((e) => (e.tenantName + e.action + e.details + (e.error || '')).toLowerCase().includes(q.toLowerCase()))
    .slice(0, 1000)

  async function exportCsv() {
    try {
      const p = await unwrap(window.api.activity.export(tid || undefined))
      if (p) toast(t('acSaved', p))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('activity')}</h1>
        </div>
        <div className="top-r">
          <TenantSelect value={tid} onChange={setTid} all={t('acAllTenants')} />
          <label className="search">
            {Icon.search}
            <input placeholder={t('acSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={log.reload} title={t('refresh')} aria-label={t('refresh')}>
            {log.loading ? <span className="spin" /> : Icon.refresh}
          </button>
          <button className="btn" onClick={exportCsv}>
            {t('acExport')}
          </button>
        </div>
      </header>
      <div className="content stack">
        <LoadState loading={log.loading && !log.data} error={log.error} />
        <div className="panel">
          <div className="panel-h">
            <div className="seg">
              {(['all', 'failed'] as Show[]).map((s) => (
                <button key={s} className={show === s ? 'on' : ''} onClick={() => setShow(s)}>
                  {t(`acShow_${s}`)}
                </button>
              ))}
            </div>
            <span className="hint">{t('acCount', shown.length)}</span>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('acWhen')}</th>
                  <th>{t('scTenant')}</th>
                  <th>{t('acAction')}</th>
                  <th>{t('acDetails')}</th>
                  <th>{t('status')}</th>
                </tr>
              </thead>
              <tbody>
                {shown.map((e, i) => (
                  <tr key={e.at + i}>
                    <td style={{ whiteSpace: 'nowrap' }}>{fmtWhen(e.at, lang)}</td>
                    <td>{e.tenantName || <span className="hint">—</span>}</td>
                    <td className="mono">{e.action}</td>
                    <td className="mono" style={{ maxWidth: 420, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={e.details}>
                      {e.details}
                    </td>
                    <td>
                      <span className={`pill ${e.ok ? 'ok' : 'bad'}`} title={e.error}>
                        {e.ok ? t('crSt_ok') : t('crSt_fail')}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!shown.length && !log.loading && <p className="hint" style={{ padding: 16 }}>{t('acNone')}</p>}
          </div>
        </div>
        <p className="hint">{t('acHint')}</p>
      </div>
    </>
  )
}
