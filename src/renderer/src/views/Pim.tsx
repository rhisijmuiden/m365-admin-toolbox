import { useState } from 'react'
import type { PimItem } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, LoadState, Modal } from '../components/ui'

const HOURS = [1, 2, 4, 8]

export function PimView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.pimLoad(tid), [tid])
  const [modal, setModal] = useState<{ item: PimItem; mode: 'activate' | 'extend' } | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const d = data.data
  const when = (iso?: string | null) =>
    iso ? new Date(iso).toLocaleString(lang, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : ''

  async function run(key: string, fn: () => Promise<unknown>, msg: string) {
    setBusy(key)
    try {
      await fn()
      toast(msg)
      data.reload()
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
            {tenant!.name} · {t('access')}
          </div>
          <h1>{t('pim')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={data.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ margin: '0 0 14px', maxWidth: '80ch' }}>
          {tenant!.mode === 'gdap' ? t('pimIntroGdap') : t('pimIntroDir')} {d?.filtered && t('pimFiltered')}
        </p>
        {data.error?.code === 'NEEDS_HOME_SIGNIN' ? (
          <div className="banner">
            <span>{t('homeSignIn')}</span>
            <button
              className="btn sm"
              onClick={async () => {
                try {
                  await unwrap(window.api.reauthHome(tid))
                  data.reload()
                } catch (e) {
                  toast((e as Error).message)
                }
              }}
            >
              {t('homeSignInBtn')}
            </button>
          </div>
        ) : (
          <LoadState loading={data.loading && !d} error={data.error} />
        )}
        {d && (
          <>
            <Section title={t('pimActive')} empty={t('noActive')} items={d.active}>
              {(i) => (
                <>
                  <td>{i.state === 'permanent' ? <span className="pill off">{t('permanent')}</span> : <span className="pill ok">{t('until', when(i.end))}</span>}</td>
                  <td style={{ textAlign: 'right' }}>
                    {i.state === 'active' && (
                      <div style={{ display: 'inline-flex', gap: 6 }}>
                        <button className="btn sm" onClick={() => setModal({ item: i, mode: 'extend' })}>
                          {t('extend')}
                        </button>
                        <button
                          className="btn sm danger"
                          disabled={busy === i.targetId}
                          onClick={() =>
                            confirm(t('deactConfirm', i.name)) &&
                            run(i.targetId, () => unwrap(window.api.pimDeactivate(tid, i.kind, i.targetId)), t('deactivatedT', i.name))
                          }
                        >
                          {t('deactivate')}
                        </button>
                      </div>
                    )}
                  </td>
                </>
              )}
            </Section>
            {d.pending.length > 0 && (
              <Section title={t('pimPending')} items={d.pending}>
                {(i) => (
                  <>
                    <td>
                      <span className="pill warn">{when(i.start)}</span>
                    </td>
                    <td style={{ textAlign: 'right' }}>
                      <button
                        className="btn sm"
                        disabled={busy === i.requestId}
                        onClick={() => run(i.requestId!, () => unwrap(window.api.pimCancel(tid, i.kind, i.requestId!)), t('cancelledT'))}
                      >
                        {t('cancelReq')}
                      </button>
                    </td>
                  </>
                )}
              </Section>
            )}
            <Section title={t('pimEligible')} empty={t('noEligible')} items={d.eligible}>
              {(i) => (
                <>
                  <td className="hint">{i.end ? t('until', when(i.end)) : t('permanent')}</td>
                  <td style={{ textAlign: 'right' }}>
                    <button className="btn pri sm" onClick={() => setModal({ item: i, mode: 'activate' })}>
                      {t('activate')}
                    </button>
                  </td>
                </>
              )}
            </Section>
            <p className="hint">{t('pimNote')}</p>
          </>
        )}
      </div>
      {modal && (
        <ActivateModal
          item={modal.item}
          mode={modal.mode}
          onClose={() => setModal(null)}
          onDone={(msg) => {
            setModal(null)
            toast(msg)
            data.reload()
          }}
        />
      )}
    </>
  )
}

function Section({ title, empty, items, children }: { title: string; empty?: string; items: PimItem[]; children(i: PimItem): JSX.Element }) {
  return (
    <div className="panel" style={{ marginBottom: 16 }}>
      <div className="panel-h">
        <h2>{title}</h2>
        <span className="hint">{items.length}</span>
      </div>
      {items.length ? (
        <div className="tbl-wrap">
          <table>
            <tbody>
              {items.map((i) => (
                <tr key={i.kind + i.targetId + (i.requestId || '')}>
                  <td style={{ width: '50%' }}>
                    <div className="u">
                      <span className="ico" style={{ background: 'linear-gradient(135deg,#ff8a6b,#e0503c)' }}>
                        {Icon.access}
                      </span>
                      <div className="u-n">{i.name}</div>
                    </div>
                  </td>
                  {children(i)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <p className="hint" style={{ padding: '12px 14px', margin: 0 }}>
          {empty}
        </p>
      )}
    </div>
  )
}

function ActivateModal({ item, mode, onClose, onDone }: { item: PimItem; mode: 'activate' | 'extend'; onClose(): void; onDone(msg: string): void }) {
  const { t, tenant } = useApp()
  const [reason, setReason] = useState('')
  const [hours, setHours] = useState(4)
  const [ticket, setTicket] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  async function submit() {
    setBusy(true)
    setErr(null)
    const req = { kind: item.kind, targetId: item.targetId, name: item.name, justification: reason.trim(), hours, ticket: ticket.trim() || undefined }
    try {
      const fn = mode === 'activate' ? window.api.pimActivate : window.api.pimExtend
      const r = await unwrap(fn(tenant!.tenantId, req))
      onDone(mode === 'extend' ? t('extendedT', item.name) : r.status === 'PendingApproval' ? t('pendingT', item.name) : t('activatedT', item.name))
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={mode === 'activate' ? t('actTitle', item.name) : t('extTitle', item.name)}
      sub={t('stepUpHint')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !reason.trim()} onClick={submit}>
            {busy && <span className="spin" />}
            {mode === 'activate' ? t('activate') : t('extend')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="pr">{t('reason')}</label>
        <input id="pr" autoFocus placeholder={t('reasonPh')} value={reason} onChange={(e) => setReason(e.target.value)} />
      </div>
      <div className="row2">
        <div className="fld">
          <label htmlFor="ph">{t('duration')}</label>
          <select id="ph" value={hours} onChange={(e) => setHours(+e.target.value)}>
            {HOURS.map((h) => (
              <option key={h} value={h}>
                {t('hours', h)}
              </option>
            ))}
          </select>
        </div>
        <div className="fld">
          <label htmlFor="pt">
            {t('ticket')} <span className="hint">({t('optional')})</span>
          </label>
          <input id="pt" className="mono" value={ticket} onChange={(e) => setTicket(e.target.value)} />
        </div>
      </div>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
