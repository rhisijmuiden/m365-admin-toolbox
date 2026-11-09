import { useState } from 'react'
import type { RestrictedSender } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon } from '../components/ui'
import { fmtWhen, Head, SecLoad } from '../components/security'
import { compromisedPrefill } from './Compromised'

/** Users Microsoft blocked from sending because they sent spam, usually after their account was taken over. */
export function RestrictedSendersView() {
  const { t, lang, tenant, toast, go } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.sec.restricted(tid), [tid])
  const [busy, setBusy] = useState<string | null>(null)

  async function release(s: RestrictedSender) {
    if (!confirm(t('seRsConfirm', s.sender))) return
    setBusy(s.sender)
    try {
      const r = await unwrap(window.api.sec.releaseSender(tid, s))
      list.set(() => r)
      toast(t('seRsReleased', s.sender))
    } catch (e) {
      toast((e as Error).message)
      list.reload()
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <Head title={t('restricted')}>
        <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
      </Head>
      <div className="content">
        <p className="hint" style={{ maxWidth: 760, marginTop: 0 }}>
          {t('seRsIntro')}
        </p>
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{list.data.length ? t('seRsCount', list.data.length) : t('seRsNone')}</h2>
            </div>
            {list.data.length > 0 && (
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('seSender')}</th>
                      <th>{t('seReason')}</th>
                      <th>{t('seRsSince')}</th>
                      <th style={{ width: 260 }} />
                    </tr>
                  </thead>
                  <tbody>
                    {list.data.map((s) => (
                      <tr key={s.sender + s.reason}>
                        <td className="mono">{s.sender}</td>
                        <td>
                          <span className="flag external">{s.reason}</span>
                        </td>
                        <td className="mono">{fmtWhen(s.created, lang)}</td>
                        <td>
                          <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                            <button
                              className="btn sm"
                              onClick={() => {
                                compromisedPrefill.upn = s.sender
                                go('compromised')
                              }}
                            >
                              {t('seRsSecure')}
                            </button>
                            <button className="btn sm pri" disabled={busy === s.sender} onClick={() => release(s)}>
                              {busy === s.sender && <span className="spin" />}
                              {t('seRelease')}
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </>
  )
}
