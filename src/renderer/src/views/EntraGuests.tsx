import { useState } from 'react'
import type { EntraGuest, Invite } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { CopyButton, Icon, Modal, RowMenu } from '../components/ui'
import { EntraLoad } from '../components/entra'
import { daysSince } from '../components/teams'

const STALE_DAYS = 90
const PENDING_DAYS = 30

/** Tenant-wide guest accounts: invite, resend, and clean up guests who never accepted or stopped signing in. */
export function EntraGuestsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.guests(tid), [tid])
  const [q, setQ] = useState('')
  const [onlyStale, setOnlyStale] = useState(false)
  const [sel, setSel] = useState<Set<string>>(new Set())
  const [modal, setModal] = useState<{ k: 'invite'; email?: string; name?: string } | { k: 'link'; url: string; email: string } | null>(null)
  const [busy, setBusy] = useState(false)

  const rows = list.data || []
  const pendingOld = (g: EntraGuest) => g.state === 'PendingAcceptance' && (daysSince(g.stateChanged || g.created) ?? 0) >= PENDING_DAYS
  const inactive = (g: EntraGuest) => g.state !== 'PendingAcceptance' && (g.lastSignIn ? (daysSince(g.lastSignIn) ?? 0) >= STALE_DAYS : (daysSince(g.created) ?? 0) >= STALE_DAYS)
  const stale = (g: EntraGuest) => pendingOld(g) || inactive(g)
  const shown = rows.filter((g) => !onlyStale || stale(g)).filter((g) => (g.displayName + (g.mail || '') + (g.company || '')).toLowerCase().includes(q.toLowerCase()))
  const domains = new Set(rows.map((g) => (g.mail || '').split('@')[1]).filter(Boolean))

  async function remove(ids: string[]) {
    if (!ids.length || !confirm(t('egDeleteConfirm', ids.length))) return
    setBusy(true)
    try {
      const failed = await unwrap(window.api.entra.deleteGuests(tid, ids))
      const bad = new Set(failed.map((f) => f.id))
      list.set((d) => d.filter((g) => !ids.includes(g.id) || bad.has(g.id)))
      setSel(new Set())
      toast(failed.length ? t('egDeletedSome', ids.length - failed.length, failed.length) : t('egDeleted', ids.length))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function block(g: EntraGuest) {
    try {
      await unwrap(window.api.setSignIn(tid, g.id, !g.enabled))
      list.set((d) => d.map((x) => (x.id === g.id ? { ...x, enabled: !g.enabled } : x)))
      toast(g.enabled ? t('egBlocked', g.displayName) : t('egUnblocked', g.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  const toggle = (id: string) => setSel((s) => (s.has(id) ? new Set([...s].filter((x) => x !== id)) : new Set([...s, id])))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('entraGuests')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('tgSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn pri" onClick={() => setModal({ k: 'invite' })}>
            {Icon.plus}
            {t('egInvite')}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('tgGuests')}</div>
                <div className="v">{rows.length}</div>
                <div className="s">{t('egDomains', domains.size)}</div>
              </div>
              <div className={`tile ${rows.some(pendingOld) ? 'warn' : ''}`}>
                <div className="k">{t('egPendingOld', PENDING_DAYS)}</div>
                <div className="v">{rows.filter(pendingOld).length}</div>
                <div className="s">{t('egPendingOldD')}</div>
              </div>
              <div className={`tile ${rows.some(inactive) ? 'warn' : ''}`}>
                <div className="k">{t('tgStale', STALE_DAYS)}</div>
                <div className="v">{rows.filter(inactive).length}</div>
                <div className="s">{t('tgStaleD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('egBlockedN')}</div>
                <div className="v">{rows.filter((g) => !g.enabled).length}</div>
                <div className="s">{t('egBlockedD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={onlyStale} onChange={(e) => setOnlyStale(e.target.checked)} />
                  {t('egOnlyStale')}
                </label>
                <div className="r">
                  {onlyStale && shown.length > 0 && (
                    <button className="btn sm" onClick={() => setSel(new Set(shown.map((g) => g.id)))}>
                      {t('egSelectAll', shown.length)}
                    </button>
                  )}
                  <button className="btn sm danger" disabled={!sel.size || busy} onClick={() => remove([...sel])}>
                    {busy && <span className="spin" />}
                    {t('egDeleteSel', sel.size)}
                  </button>
                </div>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style={{ width: 30 }} />
                      <th>{t('tgGuest')}</th>
                      <th>{t('egCompany')}</th>
                      <th>{t('tgSince')}</th>
                      <th>{t('lastSign')}</th>
                      <th>{t('status')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((g) => (
                      <tr key={g.id} className={g.enabled ? '' : 'muted'}>
                        <td>
                          <input type="checkbox" className="chk" checked={sel.has(g.id)} onChange={() => toggle(g.id)} aria-label={g.displayName} />
                        </td>
                        <td>
                          <div className="u-n">{g.displayName}</div>
                          <div className="mono">{g.mail || g.upn}</div>
                        </td>
                        <td>{g.company || '—'}</td>
                        <td className="mono">{fmtDate(g.created, lang, '—')}</td>
                        <td className="mono">{fmtDate(g.lastSignIn, lang, t('never'))}</td>
                        <td>
                          {g.state === 'PendingAcceptance' ? (
                            <span className="pill warn">{t('tgPendingPill')}</span>
                          ) : !g.enabled ? (
                            <span className="pill bad">{t('egBlockedPill')}</span>
                          ) : inactive(g) ? (
                            <span className="pill warn">{t('tmInactivePill')}</span>
                          ) : (
                            <span className="pill ok">{t('active')}</span>
                          )}
                        </td>
                        <td style={{ width: 40 }}>
                          <RowMenu
                            items={[
                              ...(g.state === 'PendingAcceptance' && g.mail ? [[t('egResend'), () => setModal({ k: 'invite', email: g.mail!, name: g.displayName })] as [string, () => void]] : []),
                              [g.enabled ? t('egBlock') : t('egUnblock'), () => block(g)],
                              [t('tgDeleteAccount'), () => remove([g.id])]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('tgNone')}
                  </p>
                )}
              </div>
            </div>
            <p className="hint">{t('egHint')}</p>
          </>
        )}
      </div>
      {modal?.k === 'invite' && (
        <InviteModal
          email={modal.email}
          name={modal.name}
          onClose={() => setModal(null)}
          onDone={(r, email) => {
            list.reload()
            setModal({ k: 'link', url: r.redeemUrl, email })
          }}
        />
      )}
      {modal?.k === 'link' && (
        <Modal title={t('egInvited', modal.email)} sub={t('egLinkSub')} onClose={() => setModal(null)} footer={<button className="btn pri" onClick={() => setModal(null)}>{t('close')}</button>}>
          <div className="creds">
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <b style={{ flex: 1, wordBreak: 'break-all' }}>{modal.url}</b>
              <CopyButton text={modal.url} />
            </div>
          </div>
        </Modal>
      )}
    </>
  )
}

function InviteModal({ email, name, onClose, onDone }: { email?: string; name?: string; onClose(): void; onDone(r: { userId: string; redeemUrl: string }, email: string): void }) {
  const { t, tenant } = useApp()
  const [i, setI] = useState<Invite>({ email: email || '', displayName: name || '', message: '', send: true })
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const valid = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(i.email.trim())
  return (
    <Modal
      title={email ? t('egResend') : t('egInvite')}
      sub={t('egInviteSub')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !valid}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const r = await unwrap(window.api.entra.invite(tenant!.tenantId, { ...i, email: i.email.trim(), displayName: i.displayName.trim(), message: i.message?.trim() || undefined }))
                onDone(r, i.email.trim())
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('egSend')}
          </button>
        </>
      }
    >
      <div className="row2">
        <div className="fld">
          <label htmlFor="ie">{t('egEmail')}</label>
          <input id="ie" type="email" autoFocus value={i.email} disabled={!!email} onChange={(e) => setI({ ...i, email: e.target.value })} />
        </div>
        <div className="fld">
          <label htmlFor="in">{t('displayName')}</label>
          <input id="in" value={i.displayName} onChange={(e) => setI({ ...i, displayName: e.target.value })} />
        </div>
      </div>
      <div className="fld">
        <label htmlFor="im">
          {t('egMessage')} <span className="hint">{t('optional')}</span>
        </label>
        <textarea id="im" value={i.message} onChange={(e) => setI({ ...i, message: e.target.value })} />
      </div>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={i.send} onChange={(e) => setI({ ...i, send: e.target.checked })} />
        {t('egSendMail')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
