import { useState } from 'react'
import type { SecureMailboxResult, User } from '@shared/types'
import { genPassword, unwrap, useApp, useLoad } from '../ctx'
import { CopyButton } from '../components/ui'
import { useUsers } from '../components/teams'
import { fmtWhen, Head, SecLoad } from '../components/security'

/** Set by other screens (restricted senders, risky users) to open this one with a user picked. */
export const compromisedPrefill: { upn: string | null } = { upn: null }
/** Set here so the audit log opens filtered on the user. */
export const auditPrefill: { user: string | null } = { user: null }

type Step = 'block' | 'revoke' | 'password' | 'rules' | 'forwarding' | 'mfa'
const STEPS: Step[] = ['block', 'revoke', 'password', 'rules', 'forwarding', 'mfa']
type State = { s: 'wait' | 'run' | 'ok' | 'fail' | 'skip'; note?: string }

/**
 * "Account compromised": blocks sign-in, signs the user out everywhere, sets a new password, removes inbox
 * rules and forwarding that leak or hide mail, optionally clears MFA methods, and shows the latest sign-ins.
 */
export function CompromisedView() {
  const { t, lang, tenant, go } = useApp()
  const tid = tenant!.tenantId
  const users = useUsers(tid)
  const [q, setQ] = useState(() => {
    const p = compromisedPrefill.upn
    compromisedPrefill.upn = null
    return p || ''
  })
  const user: User | undefined = users.find((u) => u.userPrincipalName.toLowerCase() === q.trim().toLowerCase())
  const [on, setOn] = useState<Record<Step, boolean>>({ block: true, revoke: true, password: true, rules: true, forwarding: true, mfa: false })
  const [state, setState] = useState<Partial<Record<Step, State>>>({})
  const [running, setRunning] = useState(false)
  const [pw, setPw] = useState<string | null>(null)
  const [mbx, setMbx] = useState<SecureMailboxResult | null>(null)
  const signIns = useLoad(() => (user ? window.api.sec.signIns(tid, user.id) : null), [tid, user?.id])
  const done = STEPS.some((s) => state[s]?.s === 'ok' || state[s]?.s === 'fail')

  async function runAll() {
    if (!user || !confirm(t('seCpConfirm', user.displayName))) return
    setRunning(true)
    setPw(null)
    setMbx(null)
    const st: Partial<Record<Step, State>> = Object.fromEntries(STEPS.map((s) => [s, { s: on[s] ? 'wait' : 'skip' }]))
    setState({ ...st })
    const step = async (s: Step, fn: () => Promise<string | undefined | void>) => {
      if (!on[s]) return
      st[s] = { s: 'run' }
      setState({ ...st })
      try {
        st[s] = { s: 'ok', note: (await fn()) || undefined }
      } catch (e) {
        st[s] = { s: 'fail', note: (e as Error).message }
      }
      setState({ ...st })
    }
    await step('block', () => unwrap(window.api.setSignIn(tid, user.id, false)))
    await step('revoke', () => unwrap(window.api.revokeSessions(tid, user.id)))
    await step('password', async () => {
      const p = genPassword(16)
      await unwrap(window.api.resetPassword(tid, user.id, p, true))
      setPw(p)
    })
    // Rules and forwarding are one Exchange call; its result is reported on both steps.
    if (on.rules || on.forwarding) {
      for (const s of ['rules', 'forwarding'] as Step[]) if (on[s]) st[s] = { s: 'run' }
      setState({ ...st })
      try {
        const r = await unwrap(window.api.sec.secureMailbox(tid, user.userPrincipalName, on.rules, on.forwarding))
        setMbx(r)
        if (on.rules)
          st.rules = r.failed.length
            ? { s: 'fail', note: r.failed.map((f) => `${f.rule}: ${f.error}`).join('; ') }
            : { s: 'ok', note: r.rulesRemoved.length ? t('seCpRulesRemoved', r.rulesRemoved.join(', ')) : t('seCpNoRules') }
        if (on.forwarding) st.forwarding = { s: 'ok', note: r.forwardingRemoved ? t('seCpFwdRemoved', r.forwardingRemoved) : t('seCpNoFwd') }
      } catch (e) {
        for (const s of ['rules', 'forwarding'] as Step[]) if (on[s]) st[s] = { s: 'fail', note: (e as Error).message }
      }
      setState({ ...st })
    }
    await step('mfa', async () => {
      const r = await unwrap(window.api.entra.resetMfa(tid, user.id, false))
      if (r.failed.length) throw new Error(r.failed.map((f) => `${f.method}: ${f.error}`).join('; '))
    })
    setRunning(false)
    signIns.reload()
  }

  return (
    <>
      <Head title={t('compromised')} />
      <div className="content">
        <p className="hint" style={{ maxWidth: 780, marginTop: 0 }}>
          {t('seCpIntro')}
        </p>
        <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
          <div className="fld">
            <label htmlFor="cpu">{t('user')}</label>
            <input id="cpu" list="cp-users" className="mono" value={q} onChange={(e) => (setQ(e.target.value), setState({}), setPw(null), setMbx(null))} placeholder={t('enUserPh')} />
            <datalist id="cp-users">
              {users.slice(0, 1500).map((u) => (
                <option key={u.id} value={u.userPrincipalName}>
                  {u.displayName}
                </option>
              ))}
            </datalist>
            {q.trim() && !user && users.length > 0 && <small className="hint">{t('enNotFound', q.trim())}</small>}
          </div>
          <div style={{ display: 'grid', gap: 8, marginTop: 6 }}>
            {STEPS.map((s) => {
              const x = state[s]
              return (
                <label key={s} className={`opt ${on[s] ? 'on' : ''}`} style={{ alignItems: 'center' }}>
                  <input type="checkbox" className="chk" disabled={running} checked={on[s]} onChange={(e) => setOn({ ...on, [s]: e.target.checked })} />
                  <span style={{ flex: 1 }}>
                    {t(`seCp_${s}`)}
                    <small>{x?.note || t(`seCp_${s}D`)}</small>
                  </span>
                  {x?.s === 'run' && <span className="spin" />}
                  {x?.s === 'ok' && <span className="pill ok">{t('seDone')}</span>}
                  {x?.s === 'fail' && <span className="pill bad">{t('seFailed')}</span>}
                  {x?.s === 'wait' && <span className="pill off">{t('seWaiting')}</span>}
                </label>
              )
            })}
          </div>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginTop: 12 }}>
            <button className="btn pri" style={{ background: 'var(--bad)', borderColor: 'var(--bad)' }} disabled={!user || running || !STEPS.some((s) => on[s])} onClick={runAll}>
              {running && <span className="spin" />}
              {t('seCpRun')}
            </button>
            {user && <span className="hint">{user.displayName}</span>}
          </div>
        </div>

        {pw && user && (
          <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
            <h2 style={{ marginTop: 0 }}>{t('seCpNewPw')}</h2>
            <div className="creds">
              {[[t('uname'), user.userPrincipalName], [t('pw'), pw]].map(([k, v]) => (
                <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span className="hint" style={{ width: 110 }}>{k}</span>
                  <b style={{ flex: 1, wordBreak: 'break-all' }}>{v}</b>
                  <CopyButton text={v} />
                </div>
              ))}
            </div>
            <p className="hint">{t('seCpPwNote')}</p>
          </div>
        )}

        {done && user && (
          <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
            <h2 style={{ marginTop: 0 }}>{t('seCpNext')}</h2>
            {mbx?.restricted && (
              <div className="banner">
                <span>{t('seCpRestricted', mbx.restricted)}</span>
                <button className="btn sm" onClick={() => go('restricted')}>
                  {t('restricted')}
                </button>
              </div>
            )}
            <ul className="hint" style={{ margin: 0, paddingLeft: 18, lineHeight: 1.7 }}>
              <li>{t('seCpNext1')}</li>
              <li>{t('seCpNext2')}</li>
              <li>{t('seCpNext3')}</li>
            </ul>
            <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
              <button className="btn sm" onClick={() => ((auditPrefill.user = user.userPrincipalName), go('auditLog'))}>
                {t('seCpAudit')}
              </button>
              <button className="btn sm" onClick={() => go('inboxRules')}>
                {t('inboxRules')}
              </button>
              <button className="btn sm" onClick={() => go('mfa')}>
                {t('mfa')}
              </button>
            </div>
          </div>
        )}

        {user && (
          <div className="panel">
            <div className="panel-h">
              <h2>{t('seCpSignIns')}</h2>
            </div>
            <div style={{ padding: signIns.error || signIns.loading ? '0 16px' : 0 }}>
              <SecLoad loading={signIns.loading} error={signIns.error} onRetry={signIns.reload} />
            </div>
            {signIns.data && (
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('seWhen')}</th>
                      <th>{t('seCpApp')}</th>
                      <th>{t('seCpIp')}</th>
                      <th>{t('seLocation')}</th>
                      <th>{t('seCpDevice')}</th>
                      <th>{t('status')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {signIns.data.map((s) => (
                      <tr key={s.id}>
                        <td className="mono">{fmtWhen(s.date, lang)}</td>
                        <td>
                          {s.app}
                          {!s.interactive && <div className="hint">{t('seCpNonInteractive')}</div>}
                        </td>
                        <td className="mono">{s.ip}</td>
                        <td>{s.location || '—'}</td>
                        <td className="hint">{s.device || s.client}</td>
                        <td>
                          {s.ok ? <span className="pill ok">{t('seCpOk')}</span> : <span className="pill bad" title={s.error}>{t('seFailed')}</span>}
                          {s.risk && s.risk !== 'none' && s.risk !== 'hidden' && <span className="flag external" style={{ marginLeft: 6 }}>{t('seRisk')}: {s.risk}</span>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!signIns.data.length && <p className="hint" style={{ padding: 16 }}>{t('seCpNoSignIns')}</p>}
              </div>
            )}
          </div>
        )}
      </div>
    </>
  )
}
