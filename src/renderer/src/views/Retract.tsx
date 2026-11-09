import { useEffect, useState } from 'react'
import type { AnalyzedEmail, DefenderAction, PurgeType, RetractJob, RetractSpec } from '@shared/types'
import { unwrap, useApp, useLoad, type ApiError } from '../ctx'
import { Icon } from '../components/ui'
import { fmtBytes } from '../components/exo'
import { dayEnd, dayInput, dayStart, entries, fmtWhen, Head, retractPrefill, SecLoad } from '../components/security'

type Method = 'search' | 'defender'

/**
 * Pulls a message out of every mailbox: find it with a compliance search, check which mailboxes have it,
 * then delete it (soft or hard). With Defender for Office 365 Plan 2 the remediation API can be used instead.
 */
export function RetractView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [method, setMethod] = useState<Method>('search')
  const [spec, setSpec] = useState(() => {
    const p = retractPrefill.spec
    retractPrefill.spec = null
    return {
      sender: p?.sender || '',
      subject: p?.subject || '',
      from: p?.start ? p.start.slice(0, 10) : dayInput(7),
      to: p?.end ? p.end.slice(0, 10) : dayInput(0),
      mailboxes: (p?.mailboxes || []).join(', '),
      query: p?.query || ''
    }
  })
  const jobs = useLoad(() => window.api.sec.retractJobs(tid), [tid])
  const setJobs = jobs.set
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<ApiError | null>(null)
  const [found, setFound] = useState<AnalyzedEmail[] | null>(null)
  const running = (jobs.data || []).some((j) => j.phase === 'searching' || j.phase === 'purging')

  // Jobs run in the background in the main process; refresh while one is busy.
  useEffect(() => {
    if (!running) return
    const i = setInterval(() => unwrap(window.api.sec.retractJobs(tid)).then((r) => setJobs(() => r), () => undefined), 5000)
    return () => clearInterval(i)
  }, [running, tid, setJobs])

  const asSpec = (): RetractSpec => ({
    sender: spec.sender.trim(),
    subject: spec.subject.trim(),
    start: dayStart(spec.from),
    end: dayEnd(spec.to),
    mailboxes: entries(spec.mailboxes),
    query: spec.query.trim()
  })

  async function start() {
    setBusy(true)
    setErr(null)
    try {
      if (method === 'search') {
        const j = await unwrap(window.api.sec.startRetract(tid, asSpec()))
        jobs.set((d) => [j, ...d])
        toast(t('seRtStarted'))
      } else setFound(await unwrap(window.api.sec.analyzed(tid, asSpec())))
    } catch (e) {
      setErr(e as ApiError)
    } finally {
      setBusy(false)
    }
  }

  const ready = !!(spec.sender.trim() || spec.subject.trim() || (method === 'search' && spec.query.trim()))

  return (
    <>
      <Head title={t('retract')}>
        <button className="btn" onClick={jobs.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
      </Head>
      <div className="content">
        <p className="hint" style={{ maxWidth: 780, marginTop: 0 }}>
          {t('seRtIntro')}
        </p>
        <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
          <div className="seg" style={{ marginBottom: 12 }}>
            {(['search', 'defender'] as Method[]).map((m) => (
              <button key={m} className={method === m ? 'on' : ''} onClick={() => (setMethod(m), setErr(null))}>
                {t(`seRtM_${m}`)}
              </button>
            ))}
          </div>
          <div className="row2" style={{ gridTemplateColumns: 'repeat(2, minmax(0,1fr))' }}>
            <div className="fld">
              <label htmlFor="rs">{t('seSenderOrDomain')}</label>
              <input id="rs" className="mono" value={spec.sender} onChange={(e) => setSpec({ ...spec, sender: e.target.value })} placeholder="phisher@example.com" />
            </div>
            <div className="fld">
              <label htmlFor="rj">{t('seSubject')}</label>
              <input id="rj" value={spec.subject} onChange={(e) => setSpec({ ...spec, subject: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="rf">{t('seReceivedFrom')}</label>
              <input id="rf" type="date" value={spec.from} onChange={(e) => setSpec({ ...spec, from: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="rt">{t('seTo')}</label>
              <input id="rt" type="date" value={spec.to} onChange={(e) => setSpec({ ...spec, to: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="rm">{t('seRtMailboxes')}</label>
              <input id="rm" className="mono" value={spec.mailboxes} onChange={(e) => setSpec({ ...spec, mailboxes: e.target.value })} placeholder={t('seRtAllMailboxes')} />
            </div>
            {method === 'search' && (
              <div className="fld">
                <label htmlFor="rq">{t('seRtQuery')}</label>
                <input id="rq" className="mono" value={spec.query} onChange={(e) => setSpec({ ...spec, query: e.target.value })} placeholder='hasattachment:true AND attachmentnames:"invoice.html"' />
              </div>
            )}
          </div>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginTop: 4 }}>
            <button className="btn pri" disabled={busy || !ready} onClick={start}>
              {busy ? <span className="spin" /> : Icon.search}
              {t('seRtFind')}
            </button>
            <span className="hint">{method === 'search' ? t('seRtSearchHint') : t('seRtDefenderHint')}</span>
          </div>
        </div>
        {err && <SecLoad loading={false} error={err} onRetry={start} scc={method === 'search'} consent={method === 'defender' ? 'defender' : undefined} />}
        {busy && method === 'search' && <SecLoad loading error={null} onRetry={start} scc />}
        {method === 'defender' && found && <DefenderResults rows={found} onDone={() => setFound(null)} />}
        {method === 'search' && (
          <>
            <SecLoad loading={jobs.loading && !jobs.data} error={jobs.error} onRetry={jobs.reload} />
            {(jobs.data || []).map((j) => (
              <JobPanel key={j.id} job={j} onChange={(nj) => jobs.set((d) => (nj ? d.map((x) => (x.id === nj.id ? nj : x)) : d.filter((x) => x.id !== j.id)))} />
            ))}
          </>
        )}
      </div>
    </>
  )
}

function JobPanel({ job, onChange }: { job: RetractJob; onChange(j: RetractJob | null): void }) {
  const { t, lang, toast } = useApp()
  const [busy, setBusy] = useState(false)
  const most = Math.max(0, ...job.mailboxes.map((m) => m.items))

  async function purge(type: PurgeType) {
    if (!confirm(t(type === 'HardDelete' ? 'seRtHardConfirm' : 'seRtSoftConfirm', job.items, job.mailboxes.length))) return
    setBusy(true)
    try {
      onChange(await unwrap(window.api.sec.purgeRetract(job.id, type)))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function discard() {
    setBusy(true)
    try {
      await unwrap(window.api.sec.discardRetract(job.id))
      onChange(null)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  const pill =
    job.phase === 'done' ? 'ok' : job.phase === 'failed' ? 'bad' : job.phase === 'found' ? (job.items ? 'warn' : 'off') : 'off'
  return (
    <div className="panel" style={{ marginBottom: 14 }}>
      <div className="panel-h">
        <div>
          <h2 style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {(job.phase === 'searching' || job.phase === 'purging') && <span className="spin" />}
            {[job.spec.sender, job.spec.subject && `“${job.spec.subject}”`].filter(Boolean).join(' · ') || job.query}
          </h2>
          <div className="mono hint">{job.query}</div>
        </div>
        <div className="r" style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
          <span className={`pill ${pill}`}>
            {job.phase === 'purging' ? t('seRtPurging', job.round, job.rounds) : job.phase === 'found' && !job.items ? t('seRtNothing') : t(`seRtP_${job.phase}`)}
          </span>
          {job.phase === 'found' && job.items > 0 && (
            <>
              <button className="btn sm" disabled={busy} onClick={() => purge('SoftDelete')}>
                {t('seRtSoft')}
              </button>
              <button className="btn sm danger" disabled={busy} onClick={() => purge('HardDelete')}>
                {t('seRtHard')}
              </button>
            </>
          )}
          {(job.phase === 'found' || job.phase === 'done' || job.phase === 'failed') && (
            <button className="btn sm" disabled={busy} onClick={discard}>
              {job.phase === 'found' ? t('cancel') : t('close')}
            </button>
          )}
        </div>
      </div>
      <div style={{ padding: '0 16px 14px' }}>
        <p className="hint" style={{ margin: '0 0 8px' }}>
          {t('seRtStartedAt', fmtWhen(job.started, lang))}
          {job.finished && ` · ${t('seRtFinishedAt', fmtWhen(job.finished, lang))}`}
          {job.phase !== 'searching' && ` · ${t('seRtFound', job.items, job.mailboxes.length)}`}
          {job.purgeType && ` · ${job.purgeType === 'HardDelete' ? t('seRtHard') : t('seRtSoft')}`}
        </p>
        {job.error && <div className="err">{job.error}</div>}
        {job.phase === 'found' && most > 10 && <p className="hint">{t('seRtRounds', Math.min(10, Math.ceil(most / 10)))}</p>}
        {job.phase === 'done' && <p className="hint">{t('seRtDoneD')}</p>}
        {job.mailboxes.length > 0 && (
          <div className="tbl-wrap" style={{ maxHeight: 260 }}>
            <table>
              <thead>
                <tr>
                  <th>{t('irMailbox')}</th>
                  <th>{t('seItems')}</th>
                  <th>{t('seSize')}</th>
                </tr>
              </thead>
              <tbody>
                {job.mailboxes.map((m) => (
                  <tr key={m.location}>
                    <td className="mono">{m.location}</td>
                    <td>{m.items}</td>
                    <td>{fmtBytes(m.size, lang)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}

function DefenderResults({ rows, onDone }: { rows: AnalyzedEmail[]; onDone(): void }) {
  const { t, lang, tenant, toast } = useApp()
  const [sel, setSel] = useState<Set<string>>(() => new Set(rows.map((r) => r.id)))
  const [action, setAction] = useState<DefenderAction>('softDelete')
  const [busy, setBusy] = useState(false)
  const chosen = rows.filter((r) => sel.has(r.id))
  return (
    <div className="panel">
      <div className="panel-h">
        <h2>{t('seRtDefFound', rows.length)}</h2>
        <div className="r" style={{ display: 'flex', gap: 6 }}>
          <select value={action} onChange={(e) => setAction(e.target.value as DefenderAction)}>
            {(['softDelete', 'hardDelete', 'moveToDeletedItems', 'moveToJunk'] as DefenderAction[]).map((a) => (
              <option key={a} value={a}>
                {t(`seRtA_${a}`)}
              </option>
            ))}
          </select>
          <button
            className="btn sm pri"
            disabled={busy || !chosen.length}
            onClick={async () => {
              if (!confirm(t('seRtDefConfirm', chosen.length))) return
              setBusy(true)
              try {
                await unwrap(window.api.sec.remediate(tenant!.tenantId, chosen, action, `Toolbox ${new Date().toISOString().slice(0, 16)}`))
                toast(t('seRtDefStarted'))
                onDone()
              } catch (e) {
                toast((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('seRtDefRun')} ({chosen.length})
          </button>
        </div>
      </div>
      <div className="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th style={{ width: 30 }}>
                <input type="checkbox" className="chk" aria-label={t('seSelectAll')} checked={chosen.length === rows.length && rows.length > 0} onChange={(e) => setSel(e.target.checked ? new Set(rows.map((r) => r.id)) : new Set())} />
              </th>
              <th>{t('seReceived')}</th>
              <th>{t('seSender')}</th>
              <th>{t('seSubject')}</th>
              <th>{t('seRecipient')}</th>
              <th>{t('seLocation')}</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id}>
                <td>
                  <input
                    type="checkbox"
                    className="chk"
                    checked={sel.has(r.id)}
                    onChange={(e) => {
                      const s = new Set(sel)
                      e.target.checked ? s.add(r.id) : s.delete(r.id)
                      setSel(s)
                    }}
                  />
                </td>
                <td className="mono">{fmtWhen(r.received, lang)}</td>
                <td className="mono">{r.sender}</td>
                <td style={{ whiteSpace: 'normal' }}>
                  {r.subject}
                  {r.threats.map((x) => (
                    <span key={x} className="flag external" style={{ marginLeft: 6 }}>
                      {x}
                    </span>
                  ))}
                </td>
                <td className="mono">{r.recipient}</td>
                <td>{r.location}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {!rows.length && <p className="hint" style={{ padding: 16 }}>{t('seRtDefNone')}</p>}
      </div>
    </div>
  )
}
