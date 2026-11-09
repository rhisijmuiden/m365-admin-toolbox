import { useState } from 'react'
import type { AuthMethod, User, UserMfa } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { CopyButton, Icon, Modal, RowMenu } from '../components/ui'
import { EntraLoad } from '../components/entra'
import { TapModal } from './Users'

type Filter = 'all' | 'none' | 'admins' | 'weak'
/** Methods that count as phishing-resistant or app-based, as opposed to SMS and voice. */
const STRONG = /authenticator|fido|windowsHello|passKey|softwareOath|hardwareOath|platformCredential/i
const PHONE = /mobilePhone|phone|sms|voice/i

const label = (m: string) =>
  m
    .replace(/^microsoft/, '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/^./, (c) => c.toUpperCase())

/** Registered MFA methods per user, with "reset MFA" and a Temporary Access Pass to register again. */
export function MfaView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.mfa(tid), [tid])
  const [q, setQ] = useState('')
  const [f, setF] = useState<Filter>('none')
  const [modal, setModal] = useState<{ k: 'methods' | 'reset' | 'tap'; u: UserMfa } | { k: 'pass'; u: UserMfa; pass: string; minutes: number } | null>(null)
  const rows = list.data?.users || []
  const shown = rows
    .filter((u) => (f === 'none' ? !u.mfaCapable : f === 'admins' ? u.isAdmin : f === 'weak' ? u.mfaCapable && !u.methods.some((m) => STRONG.test(m)) : true))
    .filter((u) => (u.displayName + u.upn).toLowerCase().includes(q.toLowerCase()))
  const asUser = (u: UserMfa): User => ({ id: u.userId, displayName: u.displayName, userPrincipalName: u.upn, department: null, accountEnabled: true, skuIds: [], lastSignIn: null, createdDateTime: null })

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('mfa')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('searchUsers')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} text={t('mfLoading')} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('totUsers')}</div>
                <div className="v">{rows.length}</div>
                <div className="s">{list.data.source === 'report' ? t('mfFromReport') : t('mfFromMethods')}</div>
              </div>
              <div className={`tile ${rows.some((u) => !u.mfaCapable) ? 'bad' : ''}`}>
                <div className="k">{t('mfNone')}</div>
                <div className="v">{rows.filter((u) => !u.mfaCapable).length}</div>
                <div className="s">{t('mfNoneD')}</div>
              </div>
              <div className={`tile ${rows.some((u) => u.mfaCapable && !u.methods.some((m) => STRONG.test(m))) ? 'warn' : ''}`}>
                <div className="k">{t('mfWeak')}</div>
                <div className="v">{rows.filter((u) => u.mfaCapable && !u.methods.some((m) => STRONG.test(m))).length}</div>
                <div className="s">{t('mfWeakD')}</div>
              </div>
              <div className={`tile ${rows.some((u) => u.isAdmin && !u.mfaCapable) ? 'bad' : ''}`}>
                <div className="k">{t('mfAdminsNone')}</div>
                <div className="v">{list.data.source === 'report' ? rows.filter((u) => u.isAdmin && !u.mfaCapable).length : '—'}</div>
                <div className="s">{t('mfAdminsD', rows.filter((u) => u.isAdmin).length)}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <div className="seg">
                  {(['none', 'weak', 'admins', 'all'] as Filter[]).map((k) => (
                    <button key={k} className={f === k ? 'on' : ''} disabled={k === 'admins' && list.data!.source !== 'report'} onClick={() => setF(k)}>
                      {t(`mfF_${k}`)}
                    </button>
                  ))}
                </div>
                <span className="hint">{t('scCount', shown.length)}</span>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('user')}</th>
                      <th>{t('mfMethods')}</th>
                      <th>{t('mfDefault')}</th>
                      <th>{t('status')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.slice(0, 500).map((u) => (
                      <tr key={u.userId}>
                        <td>
                          <div className="u-n">
                            {u.displayName} {u.isAdmin && <span className="flag">{t('mfAdmin')}</span>}
                          </div>
                          <div className="mono">{u.upn}</div>
                        </td>
                        <td style={{ whiteSpace: 'normal', maxWidth: 360 }}>
                          {u.methods
                            .filter((m) => m !== 'password')
                            .map((m) => (
                              <span key={m} className="lic-tag" style={PHONE.test(m) && !STRONG.test(m) ? { color: 'var(--warn)' } : undefined}>
                                {label(m)}
                              </span>
                            ))}
                          {!u.methods.filter((m) => m !== 'password').length && <span className="hint">—</span>}
                        </td>
                        <td>{u.defaultMethod ? label(u.defaultMethod) : '—'}</td>
                        <td>{u.mfaCapable ? <span className="pill ok">{t('mfCapable')}</span> : <span className="pill bad">{t('mfNotCapable')}</span>}</td>
                        <td style={{ width: 40 }}>
                          <RowMenu
                            items={[
                              [t('mfView'), () => setModal({ k: 'methods', u })],
                              [t('mfReset'), () => setModal({ k: 'reset', u })],
                              [t('tap'), () => setModal({ k: 'tap', u })]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {shown.length > 500 && <p className="hint" style={{ padding: 12 }}>{t('mfFirst', 500)}</p>}
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('mfAllGood')}
                  </p>
                )}
              </div>
            </div>
            {list.data.source === 'methods' && <p className="hint">{t('mfMethodsHint')}</p>}
          </>
        )}
      </div>
      {modal?.k === 'methods' && <MethodsModal u={modal.u} onClose={() => setModal(null)} />}
      {modal?.k === 'reset' && (
        <ResetModal
          u={modal.u}
          onClose={() => setModal(null)}
          onDone={(left) => {
            list.set((d) => ({ ...d, users: d.users.map((x) => (x.userId === modal.u.userId ? { ...x, methods: left.map((m) => m.type.replace('AuthenticationMethod', '')), mfaCapable: false, defaultMethod: null } : x)) }))
            toast(t('mfResetDone', modal.u.displayName))
            setModal({ k: 'tap', u: modal.u })
          }}
        />
      )}
      {modal?.k === 'tap' && <TapModal user={asUser(modal.u)} onClose={() => setModal(null)} onDone={(pass, minutes) => setModal({ k: 'pass', u: modal.u, pass, minutes })} />}
      {modal?.k === 'pass' && (
        <Modal title={t('tapDone')} sub={t('tapValid', modal.minutes)} onClose={() => setModal(null)} footer={<button className="btn pri" onClick={() => setModal(null)}>{t('close')}</button>}>
          <div className="creds">
            {[[t('uname'), modal.u.upn], ['TAP', modal.pass]].map(([k, v]) => (
              <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span className="hint" style={{ width: 110 }}>{k}</span>
                <b style={{ flex: 1, wordBreak: 'break-all' }}>{v}</b>
                <CopyButton text={v} />
              </div>
            ))}
          </div>
        </Modal>
      )}
    </>
  )
}

function MethodsModal({ u, onClose }: { u: UserMfa; onClose(): void }) {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.methods(tid, u.userId), [u.userId])
  const [busy, setBusy] = useState<string | null>(null)
  async function remove(m: AuthMethod) {
    if (!confirm(t('mfRemoveConfirm', label(m.type.replace('AuthenticationMethod', '')), u.displayName))) return
    setBusy(m.id)
    try {
      await unwrap(window.api.entra.deleteMethod(tid, u.userId, m))
      list.set((d) => d.filter((x) => x.id !== m.id))
      toast(t('mfRemoved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }
  return (
    <Modal wide title={t('mfMethodsOf', u.displayName)} sub={u.upn} onClose={onClose}>
      <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
      {list.data && (
        <div className="panel">
          <table>
            <tbody>
              {list.data.map((m) => (
                <tr key={m.id}>
                  <td className="u-n">{label(m.type.replace('AuthenticationMethod', ''))}</td>
                  <td className="mono">{m.label || '—'}</td>
                  <td className="mono">{fmtDate(m.created, lang, '—')}</td>
                  <td style={{ textAlign: 'right' }}>
                    {m.type !== 'passwordAuthenticationMethod' && (
                      <button className="btn sm danger" disabled={!!busy} onClick={() => remove(m)}>
                        {busy === m.id && <span className="spin" />}
                        {t('remove')}
                      </button>
                    )}
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

function ResetModal({ u, onClose, onDone }: { u: UserMfa; onClose(): void; onDone(left: AuthMethod[]): void }) {
  const { t, tenant } = useApp()
  const [revoke, setRevoke] = useState(true)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('mfReset')}
      sub={u.displayName}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn danger"
            disabled={busy}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const r = await unwrap(window.api.entra.resetMfa(tenant!.tenantId, u.userId, revoke))
                if (r.failed.length) setErr(t('mfResetPartial', r.failed.map((f) => `${f.method}: ${f.error}`).join('\n')))
                else onDone(r.methods)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('mfResetBtn')}
          </button>
        </>
      }
    >
      <p style={{ margin: 0 }}>{t('mfResetD')}</p>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={revoke} onChange={(e) => setRevoke(e.target.checked)} />
        {t('mfRevoke')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
