import { useState } from 'react'
import type { QuarantineFilter, QuarantineMsg, QuarantineType } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { fmtBytes } from '../components/exo'
import { dayEnd, dayInput, dayStart, fmtWhen, Head, MailBody, SecLoad } from '../components/security'

const TYPES: QuarantineType[] = ['Spam', 'HighConfSpam', 'Phish', 'HighConfPhish', 'Bulk', 'Malware', 'TransportRule']
const PAGE = 500

/** The quarantine queue: filter, preview, release (optionally allowing the sender), delete, or block the sender. */
export function QuarantineView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [draft, setDraft] = useState({ sender: '', recipient: '', subject: '', type: '' as QuarantineType | '', from: dayInput(14), to: dayInput(0) })
  const [filter, setFilter] = useState<QuarantineFilter>({ start: dayStart(dayInput(14)), end: dayEnd(dayInput(0)) })
  const [page, setPage] = useState(1)
  const [more, setMore] = useState<QuarantineMsg[]>([])
  const list = useLoad(() => window.api.sec.quarantine(tid, { ...filter, page: 1 }), [tid, filter])
  const [sel, setSel] = useState<Set<string>>(new Set())
  const [modal, setModal] = useState<{ k: 'preview'; m: QuarantineMsg } | { k: 'release'; ids: string[] } | null>(null)
  const [busy, setBusy] = useState(false)
  const rows = [...(list.data || []), ...more].filter((m) => m.releaseStatus !== 'RELEASED')
  const released = [...(list.data || []), ...more].length - rows.length
  const chosen = rows.filter((m) => sel.has(m.id))

  const apply = () => {
    setMore([])
    setPage(1)
    setSel(new Set())
    setFilter({ sender: draft.sender.trim(), recipient: draft.recipient.trim(), subject: draft.subject.trim(), type: draft.type, start: dayStart(draft.from), end: dayEnd(draft.to) })
  }

  async function loadMore() {
    setBusy(true)
    try {
      const next = await unwrap(window.api.sec.quarantine(tid, { ...filter, page: page + 1 }))
      setMore((m) => [...m, ...next])
      setPage((p) => p + 1)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  const drop = (ids: string[]) => {
    const s = new Set(ids)
    list.set((d) => d.filter((m) => !s.has(m.id)))
    setMore((d) => d.filter((m) => !s.has(m.id)))
    setSel(new Set())
  }

  async function remove(ids: string[]) {
    if (!confirm(t('seQDeleteConfirm', ids.length))) return
    setBusy(true)
    try {
      const r = await unwrap(window.api.sec.quarantineDelete(tid, ids))
      drop(ids.filter((id) => !r.failed.some((f) => f.id === id)))
      toast(r.failed.length ? t('seSomeFailed', r.failed.length, r.failed[0].error) : t('seQDeleted', ids.length))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function blockSenders(msgs: QuarantineMsg[]) {
    const senders = [...new Set(msgs.map((m) => m.sender.toLowerCase()).filter(Boolean))]
    if (!senders.length || !confirm(t('seQBlockConfirm', senders.join(', ')))) return
    setBusy(true)
    try {
      await unwrap(window.api.sec.newAllowBlock(tid, { listType: 'Sender', entries: senders, action: 'Block', notes: 'M365 Admin Toolbox: blocked from quarantine' }))
      toast(t('seQBlocked', senders.length))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <Head title={t('quarantine')}>
        <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
      </Head>
      <div className="content">
        <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
          <div className="row2" style={{ gridTemplateColumns: 'repeat(3, minmax(0,1fr))' }}>
            <div className="fld">
              <label htmlFor="qs">{t('seSender')}</label>
              <input id="qs" className="mono" value={draft.sender} onChange={(e) => setDraft({ ...draft, sender: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="qr">{t('seRecipient')}</label>
              <input id="qr" className="mono" value={draft.recipient} onChange={(e) => setDraft({ ...draft, recipient: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="qj">{t('seSubject')}</label>
              <input id="qj" value={draft.subject} onChange={(e) => setDraft({ ...draft, subject: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
            </div>
            <div className="fld">
              <label htmlFor="qt">{t('seReason')}</label>
              <select id="qt" value={draft.type} onChange={(e) => setDraft({ ...draft, type: e.target.value as QuarantineType })}>
                <option value="">{t('seAnyReason')}</option>
                {TYPES.map((x) => (
                  <option key={x} value={x}>
                    {t(`seQt_${x}`)}
                  </option>
                ))}
              </select>
            </div>
            <div className="fld">
              <label htmlFor="qf">{t('seFrom')}</label>
              <input id="qf" type="date" value={draft.from} onChange={(e) => setDraft({ ...draft, from: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="qe">{t('seTo')}</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input id="qe" type="date" style={{ flex: 1 }} value={draft.to} onChange={(e) => setDraft({ ...draft, to: e.target.value })} />
                <button className="btn pri" onClick={apply}>
                  {Icon.search}
                  {t('seSearch')}
                </button>
              </div>
            </div>
          </div>
        </div>
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{t('seQCount', rows.length)}</h2>
              <div className="r" style={{ display: 'flex', gap: 6 }}>
                {released > 0 && <span className="hint">{t('seQReleasedHidden', released)}</span>}
                <button className="btn sm pri" disabled={!chosen.length || busy} onClick={() => setModal({ k: 'release', ids: chosen.map((m) => m.id) })}>
                  {t('seRelease')} {chosen.length > 0 && `(${chosen.length})`}
                </button>
                <button className="btn sm" disabled={!chosen.length || busy} onClick={() => blockSenders(chosen)}>
                  {t('seBlockSender')}
                </button>
                <button className="btn sm danger" disabled={!chosen.length || busy} onClick={() => remove(chosen.map((m) => m.id))}>
                  {t('remove')}
                </button>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th style={{ width: 30 }}>
                      <input
                        type="checkbox"
                        className="chk"
                        aria-label={t('seSelectAll')}
                        checked={rows.length > 0 && chosen.length === rows.length}
                        onChange={(e) => setSel(e.target.checked ? new Set(rows.map((m) => m.id)) : new Set())}
                      />
                    </th>
                    <th>{t('seReceived')}</th>
                    <th>{t('seSender')}</th>
                    <th>{t('seSubject')}</th>
                    <th>{t('seRecipient')}</th>
                    <th>{t('seReason')}</th>
                    <th>{t('seExpires')}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((m) => (
                    <tr key={m.id}>
                      <td>
                        <input
                          type="checkbox"
                          className="chk"
                          checked={sel.has(m.id)}
                          onChange={(e) => {
                            const s = new Set(sel)
                            e.target.checked ? s.add(m.id) : s.delete(m.id)
                            setSel(s)
                          }}
                        />
                      </td>
                      <td className="mono">{fmtWhen(m.received, lang)}</td>
                      <td className="mono">{m.sender}</td>
                      <td style={{ whiteSpace: 'normal', maxWidth: 360 }}>
                        <button className="linkish" onClick={() => setModal({ k: 'preview', m })}>
                          <span className="u-n">{m.subject || t('seNoSubject')}</span>
                        </button>
                        <div className="hint">
                          {fmtBytes(m.size, lang)} · {m.direction}
                        </div>
                      </td>
                      <td className="mono" style={{ whiteSpace: 'normal' }}>
                        {m.recipients.join(', ')}
                      </td>
                      <td>
                        <span className={`flag ${/Phish|Malware/i.test(m.type) ? 'external' : ''}`}>{(t(`seQt_${m.type}` as 'seQt_Spam') as string | undefined) || m.type}</span>
                        <div className="hint">{m.policy}</div>
                      </td>
                      <td className="mono">{m.expires ? new Date(m.expires).toLocaleDateString(lang) : '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!rows.length && <p className="hint" style={{ padding: 16 }}>{t('seQEmpty')}</p>}
              {(list.data.length + more.length) % PAGE === 0 && list.data.length > 0 && (
                <div style={{ padding: 12 }}>
                  <button className="btn sm" disabled={busy} onClick={loadMore}>
                    {busy && <span className="spin" />}
                    {t('seLoadMore')}
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
      {modal?.k === 'preview' && (
        <PreviewModal
          m={modal.m}
          onClose={() => setModal(null)}
          onRelease={() => setModal({ k: 'release', ids: [modal.m.id] })}
          onDelete={() => (setModal(null), remove([modal.m.id]))}
          onBlock={() => blockSenders([modal.m])}
        />
      )}
      {modal?.k === 'release' && (
        <ReleaseModal
          ids={modal.ids}
          onClose={() => setModal(null)}
          onDone={(ok) => {
            drop(ok)
            setModal(null)
          }}
        />
      )}
    </>
  )
}

function PreviewModal({ m, onClose, onRelease, onDelete, onBlock }: { m: QuarantineMsg; onClose(): void; onRelease(): void; onDelete(): void; onBlock(): void }) {
  const { t, lang, tenant } = useApp()
  const p = useLoad(() => window.api.sec.quarantinePreview(tenant!.tenantId, m.id), [m.id])
  const [headers, setHeaders] = useState(false)
  return (
    <Modal
      wide
      title={m.subject || t('seNoSubject')}
      sub={`${m.sender} → ${m.recipients.join(', ')} · ${fmtWhen(m.received, lang)}`}
      onClose={onClose}
      footer={
        <>
          <button className="btn danger" onClick={onDelete}>
            {t('remove')}
          </button>
          <button className="btn" onClick={onBlock}>
            {t('seBlockSender')}
          </button>
          <button className="btn pri" onClick={onRelease}>
            {t('seRelease')}
          </button>
        </>
      }
    >
      <SecLoad loading={p.loading} error={p.error} onRetry={p.reload} />
      {p.data && (
        <>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <span className={`flag ${/Phish|Malware/i.test(m.type) ? 'external' : ''}`}>{(t(`seQt_${m.type}` as 'seQt_Spam') as string | undefined) || m.type}</span>
            {p.data.attachments.map((a) => (
              <span key={a} className="lic-tag">
                📎 {a}
              </span>
            ))}
            {p.data.headers && (
              <button className="btn sm" style={{ marginLeft: 'auto' }} onClick={() => setHeaders(!headers)}>
                {headers ? t('seShowBody') : t('seShowHeaders')}
              </button>
            )}
          </div>
          {headers ? (
            <pre className="mono" style={{ whiteSpace: 'pre-wrap', maxHeight: 360, overflow: 'auto', fontSize: 11.5 }}>
              {p.data.headers}
            </pre>
          ) : (
            <MailBody body={p.data.body} isHtml={p.data.isHtml} />
          )}
          <p className="hint">{t('seQPreviewSafe')}</p>
        </>
      )}
    </Modal>
  )
}

function ReleaseModal({ ids, onClose, onDone }: { ids: string[]; onClose(): void; onDone(released: string[]): void }) {
  const { t, tenant, toast } = useApp()
  const [allow, setAllow] = useState(false)
  const [report, setReport] = useState(true)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('seReleaseT', ids.length)}
      sub={t('seReleaseD')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const r = await unwrap(window.api.sec.quarantineRelease(tenant!.tenantId, ids, allow, report))
                const ok = ids.filter((id) => !r.failed.some((f) => f.id === id))
                if (r.failed.length) setErr(t('seSomeFailed', r.failed.length, r.failed[0].error))
                if (ok.length) toast(t('seReleased', ok.length))
                if (!r.failed.length) onDone(ok)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('seRelease')}
          </button>
        </>
      }
    >
      <label className={`opt ${report ? 'on' : ''}`}>
        <input type="checkbox" className="chk" checked={report} onChange={(e) => setReport(e.target.checked)} />
        <span>
          {t('seReportFp')}
          <small>{t('seReportFpD')}</small>
        </span>
      </label>
      <label className={`opt ${allow ? 'on' : ''}`}>
        <input type="checkbox" className="chk" checked={allow} onChange={(e) => setAllow(e.target.checked)} />
        <span>
          {t('seAllowSender')}
          <small>{t('seAllowSenderD')}</small>
        </span>
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
