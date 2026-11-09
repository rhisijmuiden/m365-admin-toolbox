import { useState } from 'react'
import type { InboxRule } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon } from '../components/ui'
import { ExoLoad, isExternal } from '../components/exo'

const SUSPICIOUS = ['external', 'delete', 'move', 'hide', 'forward']
const suspicious = (r: InboxRule) => r.flags.some((f) => SUSPICIOUS.includes(f))
const severity = (r: InboxRule) => (r.flags.includes('external') ? 3 : r.flags.includes('hide') || r.flags.includes('delete') ? 2 : suspicious(r) ? 1 : 0)

export function InboxRulesView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [run, setRun] = useState(0)
  const rules = useLoad(() => (run ? window.api.exo.inboxRules(tid) : null), [tid, run])
  // Mailbox-level forwarding is not an inbox rule but has the same effect, so it is shown here too.
  const mbx = useLoad(() => (run ? window.api.exo.mailboxes(tid) : null), [tid, run])
  const [only, setOnly] = useState(true)
  const [busy, setBusy] = useState<string | null>(null)

  const list = (rules.data || []).filter((r) => !only || suspicious(r)).sort((a, b) => severity(b) - severity(a) || a.mailbox.localeCompare(b.mailbox))
  const flagged = (rules.data || []).filter(suspicious).length
  const extFwd = (mbx.data?.mailboxes || []).filter((m) => isExternal(m.forwardingSmtp, mbx.data!.domains))

  async function act(r: InboxRule, action: 'disable' | 'enable' | 'remove') {
    if (action === 'remove' && !confirm(t('irRemoveConfirm', r.name))) return
    setBusy(r.mailboxId + r.ruleId)
    try {
      await unwrap(window.api.exo.setInboxRule(tid, r.mailboxId, r.ruleId, action))
      rules.set((d) =>
        action === 'remove' ? d.filter((x) => !(x.mailboxId === r.mailboxId && x.ruleId === r.ruleId)) : d.map((x) => (x === r ? { ...x, enabled: action === 'enable' } : x))
      )
      toast(t('irDone'))
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
            {tenant!.name} · {t('exchange')}
          </div>
          <h1>{t('inboxRules')}</h1>
        </div>
        <div className="top-r">
          <button className="btn pri" disabled={rules.loading} onClick={() => setRun((n) => n + 1)}>
            {rules.loading ? <span className="spin" /> : Icon.refresh}
            {run ? t('irRescan') : t('irScan')}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ maxWidth: 760, marginTop: 0 }}>
          {t('irIntro')}
        </p>
        <ExoLoad loading={rules.loading} error={rules.error} onRetry={() => setRun((n) => n + 1)} text={t('irScanning')} />
        {extFwd.length > 0 && (
          <div className="panel" style={{ marginBottom: 16 }}>
            <div className="panel-h">
              <h2>{t('irFwdMbx')}</h2>
            </div>
            <div className="tbl-wrap">
              <table>
                <tbody>
                  {extFwd.map((m) => (
                    <tr key={m.id}>
                      <td>
                        <div className="u-n">{m.displayName}</div>
                        <div className="mono">{m.primarySmtp}</div>
                      </td>
                      <td className="mono">→ {m.forwardingSmtp}</td>
                      <td>
                        <span className="flag external">{t('flag_external')}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {rules.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{flagged ? t('irFound', flagged) : t('irClean')}</h2>
              <div className="r">
                <label className="hint" style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={only} onChange={(e) => setOnly(e.target.checked)} />
                  {t('irOnlySusp')}
                </label>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('irMailbox')}</th>
                    <th>{t('irRule')}</th>
                    <th>{t('irWhat')}</th>
                    <th style={{ width: 210 }} aria-label={t('actions')} />
                  </tr>
                </thead>
                <tbody>
                  {list.map((r) => (
                    <tr key={r.mailboxId + r.ruleId} className={r.enabled ? '' : 'muted'}>
                      <td>
                        <div className="u-n">{r.mailbox}</div>
                        <div className="mono">{r.mailboxSmtp}</div>
                      </td>
                      <td>
                        <div className="u-n">{r.name}</div>
                        {!r.enabled && <span className="pill off">{t('disabledL')}</span>}
                      </td>
                      <td>
                        <div style={{ marginBottom: 4 }}>
                          {r.flags.map((f) => (
                            <span key={f} className={`flag ${f}`}>
                              {t(`flag_${f}` as 'flag_forward')}
                            </span>
                          ))}
                        </div>
                        {r.targets.length > 0 && <div className="mono">→ {r.targets.join(', ')}</div>}
                        <div className="hint" style={{ maxWidth: 520, whiteSpace: 'normal' }}>
                          {r.description}
                        </div>
                      </td>
                      <td>
                        <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                          <button className="btn sm" disabled={busy === r.mailboxId + r.ruleId} onClick={() => act(r, r.enabled ? 'disable' : 'enable')}>
                            {r.enabled ? t('irDisable') : t('irEnable')}
                          </button>
                          <button className="btn sm" style={{ color: 'var(--bad)' }} disabled={busy === r.mailboxId + r.ruleId} onClick={() => act(r, 'remove')}>
                            {t('remove')}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!list.length && <p className="hint" style={{ padding: 16 }}>{rules.data.length ? t('irClean') : t('irNone')}</p>}
            </div>
          </div>
        )}
      </div>
    </>
  )
}
