import { useState } from 'react'
import type { TraceFilter, TraceRow } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal, RowMenu } from '../components/ui'
import { fmtBytes } from '../components/exo'
import { dayEnd, dayInput, dayStart, fmtWhen, Head, retractPrefill, SecLoad } from '../components/security'

const STATUSES = ['Delivered', 'Failed', 'Pending', 'Quarantined', 'FilteredAsSpam', 'Expanded', 'GettingStatus']
const statusCls = (s: string) => (s === 'Delivered' || s === 'Expanded' ? 'ok' : s === 'Failed' ? 'bad' : s === 'Pending' || s === 'GettingStatus' ? 'off' : 'warn')

/** Message trace: where did a message go, and why. */
export function MessageTraceView() {
  const { t, lang, tenant, toast, go } = useApp()
  const tid = tenant!.tenantId
  const [draft, setDraft] = useState({ sender: '', recipient: '', subject: '', status: '', from: dayInput(2), to: dayInput(0) })
  const [filter, setFilter] = useState<TraceFilter | null>(null)
  const list = useLoad(() => (filter ? window.api.sec.trace(tid, filter) : null), [tid, filter])
  const [detail, setDetail] = useState<TraceRow | null>(null)

  const apply = () =>
    setFilter({
      sender: draft.sender.trim(),
      recipient: draft.recipient.trim(),
      subject: draft.subject.trim(),
      status: draft.status,
      start: dayStart(draft.from)!,
      end: draft.to === dayInput(0) ? new Date().toISOString() : dayEnd(draft.to)!
    })

  async function block(sender: string) {
    if (!confirm(t('seQBlockConfirm', sender))) return
    try {
      await unwrap(window.api.sec.newAllowBlock(tid, { listType: 'Sender', entries: [sender], action: 'Block', notes: 'M365 Admin Toolbox: blocked from message trace' }))
      toast(t('seQBlocked', 1))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  const retract = (r: TraceRow) => {
    retractPrefill.spec = { sender: r.sender, subject: r.subject, start: r.received || undefined, end: new Date().toISOString() }
    go('retract')
  }

  return (
    <>
      <Head title={t('messageTrace')} />
      <div className="content">
        <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
          <div className="row2" style={{ gridTemplateColumns: 'repeat(3, minmax(0,1fr))' }}>
            <div className="fld">
              <label htmlFor="ts">{t('seSender')}</label>
              <input id="ts" className="mono" value={draft.sender} onChange={(e) => setDraft({ ...draft, sender: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="tr">{t('seRecipient')}</label>
              <input id="tr" className="mono" value={draft.recipient} onChange={(e) => setDraft({ ...draft, recipient: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="tj">{t('seSubject')}</label>
              <input id="tj" value={draft.subject} onChange={(e) => setDraft({ ...draft, subject: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="tt">{t('status')}</label>
              <select id="tt" value={draft.status} onChange={(e) => setDraft({ ...draft, status: e.target.value })}>
                <option value="">{t('seAnyStatus')}</option>
                {STATUSES.map((s) => (
                  <option key={s} value={s}>
                    {t(`seMs_${s}` as 'seMs_Delivered')}
                  </option>
                ))}
              </select>
            </div>
            <div className="fld">
              <label htmlFor="tf">{t('seFrom')}</label>
              <input id="tf" type="date" min={dayInput(10)} value={draft.from} onChange={(e) => setDraft({ ...draft, from: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="te">{t('seTo')}</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input id="te" type="date" style={{ flex: 1 }} value={draft.to} onChange={(e) => setDraft({ ...draft, to: e.target.value })} />
                <button className="btn pri" onClick={apply}>
                  {Icon.search}
                  {t('seSearch')}
                </button>
              </div>
            </div>
          </div>
          <p className="hint" style={{ margin: 0 }}>
            {t('seMtHint')}
          </p>
        </div>
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{t('seMtCount', list.data.length)}</h2>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('seReceived')}</th>
                    <th>{t('seSender')}</th>
                    <th>{t('seRecipient')}</th>
                    <th>{t('seSubject')}</th>
                    <th>{t('status')}</th>
                    <th style={{ width: 40 }} />
                  </tr>
                </thead>
                <tbody>
                  {list.data.map((r) => (
                    <tr key={r.traceId + r.recipient}>
                      <td className="mono">{fmtWhen(r.received, lang)}</td>
                      <td className="mono">{r.sender}</td>
                      <td className="mono">{r.recipient}</td>
                      <td style={{ whiteSpace: 'normal', maxWidth: 380 }}>
                        <button className="linkish" onClick={() => setDetail(r)}>
                          <span className="u-n">{r.subject || t('seNoSubject')}</span>
                        </button>
                        <div className="hint">
                          {fmtBytes(r.size, lang)} {r.fromIp && `· ${r.fromIp}`}
                        </div>
                      </td>
                      <td>
                        <span className={`pill ${statusCls(r.status)}`}>{(t(`seMs_${r.status}` as 'seMs_Delivered') as string | undefined) || r.status}</span>
                      </td>
                      <td>
                        <RowMenu
                          items={[
                            [t('seMtDetails'), () => setDetail(r)],
                            [t('seRtFromHere'), () => retract(r)],
                            [t('seBlockSender'), () => block(r.sender)]
                          ]}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!list.data.length && <p className="hint" style={{ padding: 16 }}>{t('seMtNone')}</p>}
              {list.data.length >= 5000 && <p className="hint" style={{ padding: 12 }}>{t('seMtCapped')}</p>}
            </div>
          </div>
        )}
      </div>
      {detail && <DetailModal r={detail} onClose={() => setDetail(null)} />}
    </>
  )
}

function DetailModal({ r, onClose }: { r: TraceRow; onClose(): void }) {
  const { t, lang, tenant } = useApp()
  const ev = useLoad(() => window.api.sec.traceDetail(tenant!.tenantId, r.traceId, r.recipient), [r.traceId, r.recipient])
  return (
    <Modal wide title={r.subject || t('seNoSubject')} sub={`${r.sender} → ${r.recipient}`} onClose={onClose}>
      <div className="creds">
        {[
          [t('seReceived'), fmtWhen(r.received, lang)],
          [t('status'), r.status],
          ['Message-ID', r.messageId],
          [t('seMtFromIp'), r.fromIp || '—'],
          [t('seMtToIp'), r.toIp || '—']
        ].map(([k, v]) => (
          <div key={k} style={{ display: 'flex', gap: 10 }}>
            <span className="hint" style={{ width: 110 }}>
              {k}
            </span>
            <span className="mono" style={{ flex: 1, wordBreak: 'break-all' }}>
              {v}
            </span>
          </div>
        ))}
      </div>
      <SecLoad loading={ev.loading} error={ev.error} onRetry={ev.reload} />
      {ev.data && (
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>{t('seWhen')}</th>
                <th>{t('seMtEvent')}</th>
                <th>{t('seMtDetail')}</th>
              </tr>
            </thead>
            <tbody>
              {ev.data.map((e, i) => (
                <tr key={i}>
                  <td className="mono">{fmtWhen(e.date, lang)}</td>
                  <td>
                    <div className="u-n">{e.event}</div>
                    {e.action && <div className="hint">{e.action}</div>}
                  </td>
                  <td className="hint" style={{ whiteSpace: 'normal' }}>
                    {e.detail}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Modal>
  )
}
