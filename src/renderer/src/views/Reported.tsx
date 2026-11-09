import { useState } from 'react'
import type { ReportedMsg } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, RowMenu } from '../components/ui'
import { fmtWhen, Head, retractPrefill, SecLoad } from '../components/security'

const catCls = (c: string) => (/phish|malware/i.test(c) ? 'bad' : /spam|junk/i.test(c) ? 'warn' : 'ok')

/** Messages users reported with the Report button in Outlook, plus admin submissions, and what Microsoft concluded. */
export function ReportedView() {
  const { t, lang, tenant, toast, go } = useApp()
  const tid = tenant!.tenantId
  const [days, setDays] = useState(30)
  const [src, setSrc] = useState<'user' | 'all'>('user')
  const list = useLoad(() => window.api.sec.reported(tid, days), [tid, days])
  const rows = (list.data || []).filter((r) => src === 'all' || r.source === 'user')

  async function block(r: ReportedMsg) {
    if (!r.sender || !confirm(t('seQBlockConfirm', r.sender))) return
    try {
      await unwrap(window.api.sec.newAllowBlock(tid, { listType: 'Sender', entries: [r.sender], action: 'Block', notes: 'M365 Admin Toolbox: reported by a user' }))
      toast(t('seQBlocked', 1))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  const retract = (r: ReportedMsg) => {
    const at = r.received || r.created
    retractPrefill.spec = { sender: r.sender, subject: r.subject, start: at ? new Date(new Date(at).getTime() - 86400000).toISOString() : undefined, end: new Date().toISOString() }
    go('retract')
  }

  return (
    <>
      <Head title={t('reported')}>
        <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
      </Head>
      <div className="content">
        <p className="hint" style={{ maxWidth: 760, marginTop: 0 }}>
          {t('seRpIntro')}
        </p>
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} consent="submissions" />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <div className="seg">
                {(['user', 'all'] as const).map((k) => (
                  <button key={k} className={src === k ? 'on' : ''} onClick={() => setSrc(k)}>
                    {t(`seRpSrc_${k}`)}
                  </button>
                ))}
              </div>
              <div className="r" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <span className="hint">{t('scCount', rows.length)}</span>
                <select value={days} onChange={(e) => setDays(+e.target.value)} aria-label={t('seRuPeriod')}>
                  {[7, 30, 90].map((d) => (
                    <option key={d} value={d}>
                      {t('seLastDays', d)}
                    </option>
                  ))}
                </select>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('seRpReported')}</th>
                    <th>{t('seRpBy')}</th>
                    <th>{t('seSender')}</th>
                    <th>{t('seSubject')}</th>
                    <th>{t('seRpAs')}</th>
                    <th>{t('seRpResult')}</th>
                    <th style={{ width: 40 }} />
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r) => (
                    <tr key={r.id}>
                      <td className="mono">{fmtWhen(r.created, lang)}</td>
                      <td className="mono">{r.reportedBy || r.recipient}</td>
                      <td className="mono">{r.sender || '—'}</td>
                      <td style={{ whiteSpace: 'normal', maxWidth: 340 }}>{r.subject || t('seNoSubject')}</td>
                      <td>
                        <span className={`pill ${catCls(r.category)}`}>{(t(`seCat_${r.category}` as 'seCat_phishing') as string | undefined) || r.category}</span>
                      </td>
                      <td className="hint" style={{ whiteSpace: 'normal', maxWidth: 260 }}>
                        {r.result || (t(`seSt_${r.status}` as 'seSt_pending') as string | undefined) || r.status}
                      </td>
                      <td>
                        <RowMenu
                          items={[
                            [t('seRtFromHere'), () => retract(r)],
                            [t('seBlockSender'), () => block(r)]
                          ]}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!rows.length && <p className="hint" style={{ padding: 16 }}>{t('seRpNone')}</p>}
            </div>
          </div>
        )}
      </div>
    </>
  )
}
