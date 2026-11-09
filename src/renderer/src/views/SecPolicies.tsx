import { useState } from 'react'
import type { SecPolicies, SecPolicy, SecPolicyKind } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import type { Key } from '../i18n'
import { Icon, Modal } from '../components/ui'
import { Tabs } from '../components/exo'
import { entries, Head, SecLoad } from '../components/security'

type Field = { k: string; type: 'bool' | 'num' | 'sel' | 'list'; opts?: string[]; min?: number; max?: number; mdo?: boolean; warn?: boolean }

const ACTIONS = ['MoveToJmf', 'Quarantine', 'Delete', 'AddXHeader', 'ModifySubject']
const IMPERSONATION = ['NoAction', 'MoveToJmf', 'Quarantine', 'Delete']

/** The settings the app edits, in the order shown; everything else stays as it is in Exchange. */
const FIELDS: Record<SecPolicyKind, Field[]> = {
  spam: [
    { k: 'BulkThreshold', type: 'num', min: 1, max: 9 },
    { k: 'SpamAction', type: 'sel', opts: ACTIONS },
    { k: 'HighConfidenceSpamAction', type: 'sel', opts: ACTIONS },
    { k: 'PhishSpamAction', type: 'sel', opts: ACTIONS },
    { k: 'BulkSpamAction', type: 'sel', opts: ACTIONS },
    { k: 'QuarantineRetentionPeriod', type: 'num', min: 1, max: 30 },
    { k: 'SpamZapEnabled', type: 'bool' },
    { k: 'PhishZapEnabled', type: 'bool' },
    { k: 'InlineSafetyTipsEnabled', type: 'bool' },
    { k: 'BlockedSenders', type: 'list' },
    { k: 'BlockedSenderDomains', type: 'list' },
    { k: 'AllowedSenders', type: 'list', warn: true },
    { k: 'AllowedSenderDomains', type: 'list', warn: true },
    { k: 'EnableLanguageBlockList', type: 'bool' },
    { k: 'LanguageBlockList', type: 'list' },
    { k: 'EnableRegionBlockList', type: 'bool' },
    { k: 'RegionBlockList', type: 'list' }
  ],
  phish: [
    { k: 'Enabled', type: 'bool' },
    { k: 'PhishThresholdLevel', type: 'sel', opts: ['1', '2', '3', '4'] },
    { k: 'EnableSpoofIntelligence', type: 'bool' },
    { k: 'AuthenticationFailAction', type: 'sel', opts: ['MoveToJmf', 'Quarantine'] },
    { k: 'EnableUnauthenticatedSender', type: 'bool' },
    { k: 'EnableViaTag', type: 'bool' },
    { k: 'EnableFirstContactSafetyTips', type: 'bool' },
    { k: 'EnableMailboxIntelligence', type: 'bool', mdo: true },
    { k: 'EnableMailboxIntelligenceProtection', type: 'bool', mdo: true },
    { k: 'MailboxIntelligenceProtectionAction', type: 'sel', opts: IMPERSONATION, mdo: true },
    { k: 'EnableTargetedUserProtection', type: 'bool', mdo: true },
    { k: 'TargetedUsersToProtect', type: 'list', mdo: true },
    { k: 'TargetedUserProtectionAction', type: 'sel', opts: IMPERSONATION, mdo: true },
    { k: 'EnableOrganizationDomainsProtection', type: 'bool', mdo: true },
    { k: 'EnableTargetedDomainsProtection', type: 'bool', mdo: true },
    { k: 'TargetedDomainsToProtect', type: 'list', mdo: true },
    { k: 'TargetedDomainProtectionAction', type: 'sel', opts: IMPERSONATION, mdo: true },
    { k: 'EnableSimilarUsersSafetyTips', type: 'bool', mdo: true },
    { k: 'EnableSimilarDomainsSafetyTips', type: 'bool', mdo: true },
    { k: 'EnableUnusualCharactersSafetyTips', type: 'bool', mdo: true },
    { k: 'ExcludedSenders', type: 'list', mdo: true },
    { k: 'ExcludedDomains', type: 'list', mdo: true }
  ],
  outbound: [
    { k: 'RecipientLimitExternalPerHour', type: 'num', min: 0, max: 10000 },
    { k: 'RecipientLimitInternalPerHour', type: 'num', min: 0, max: 10000 },
    { k: 'RecipientLimitPerDay', type: 'num', min: 0, max: 10000 },
    { k: 'ActionWhenThresholdReached', type: 'sel', opts: ['BlockUserForToday', 'BlockUser', 'Alert'] },
    { k: 'AutoForwardingMode', type: 'sel', opts: ['Automatic', 'Off', 'On'] },
    { k: 'NotifyOutboundSpam', type: 'bool' },
    { k: 'NotifyOutboundSpamRecipients', type: 'list' },
    { k: 'BccSuspiciousOutboundMail', type: 'bool' },
    { k: 'BccSuspiciousOutboundAdditionalRecipients', type: 'list' }
  ]
}

const SUMMARY: Record<SecPolicyKind, string[]> = {
  spam: ['SpamAction', 'HighConfidenceSpamAction', 'PhishSpamAction', 'BulkSpamAction'],
  phish: ['PhishThresholdLevel', 'EnableSpoofIntelligence', 'EnableTargetedUserProtection', 'EnableMailboxIntelligenceProtection'],
  outbound: ['AutoForwardingMode', 'RecipientLimitExternalPerHour', 'ActionWhenThresholdReached']
}

function useLabels() {
  const { t } = useApp()
  const f = (k: string) => (t(`seF_${k}` as Key) as string | undefined) || k
  const v = (x: unknown, k?: string) => {
    if (typeof x === 'number' && k !== 'PhishThresholdLevel') return String(x)
    if (x === true) return t('yes')
    if (x === false) return t('no')
    if (x == null || x === '') return '—'
    if (Array.isArray(x)) return x.length ? String(x.length) : '—'
    return (t(`seV_${x}` as Key) as string | undefined) || String(x)
  }
  return { f, v }
}

/** Anti-spam, anti-phishing and outbound spam policies, with the main settings editable. */
export function SecPoliciesView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.sec.policies(tid), [tid])
  const [tab, setTab] = useState<SecPolicyKind>('spam')
  const [edit, setEdit] = useState<SecPolicy | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const { f, v } = useLabels()
  const list = data.data?.[tab] || []

  async function toggle(p: SecPolicy) {
    if (!p.scope) return
    setBusy(p.name)
    try {
      const r = await unwrap(window.api.sec.setPolicyRule(tid, p.kind, p.scope.rule, !p.scope.enabled))
      data.set(() => r)
      toast(t('seSaved'))
    } catch (e) {
      toast((e as Error).message)
      data.reload()
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <Head title={t('secPolicies')}>
        <button className="btn" onClick={data.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
      </Head>
      <div className="content">
        <Tabs<SecPolicyKind>
          tabs={[
            ['spam', t('seP_spam')],
            ['phish', t('seP_phish')],
            ['outbound', t('seP_outbound')]
          ]}
          value={tab}
          onChange={setTab}
        />
        <p className="hint" style={{ maxWidth: 760, marginTop: 0 }}>
          {t(`seP_${tab}D`)}
        </p>
        <SecLoad loading={data.loading} error={data.error} onRetry={data.reload} />
        {data.data && tab === 'phish' && !data.data.mdo && <div className="banner">{t('seNoMdo')}</div>}
        {data.data && (
          <div className="panel">
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('sePolicy')}</th>
                    <th>{t('seAppliesTo')}</th>
                    {SUMMARY[tab].map((k) => (
                      <th key={k}>{f(k)}</th>
                    ))}
                    <th style={{ width: 160 }} />
                  </tr>
                </thead>
                <tbody>
                  {list.map((p) => (
                    <tr key={p.name} className={p.scope && !p.scope.enabled ? 'muted' : ''}>
                      <td>
                        <div className="u-n">{p.name}</div>
                        {p.isDefault && <span className="lic-tag">{t('seDefault')}</span>}
                        {(p.preset || p.builtIn) && <span className="lic-tag">{t('sePreset')}</span>}
                      </td>
                      <td style={{ whiteSpace: 'normal', maxWidth: 260 }}>
                        {p.isDefault ? (
                          <span className="hint">{t('seEveryone')}</span>
                        ) : p.scope ? (
                          <span className="hint">{[...p.scope.users, ...p.scope.groups, ...p.scope.domains].join(', ') || '—'}</span>
                        ) : (
                          '—'
                        )}
                      </td>
                      {SUMMARY[tab].map((k) => (
                        <td key={k}>{v(p.settings[k], k)}</td>
                      ))}
                      <td>
                        <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                          {p.scope && !p.preset && !p.builtIn && (
                            <button className="btn sm" disabled={busy === p.name} onClick={() => toggle(p)}>
                              {p.scope.enabled ? t('irDisable') : t('irEnable')}
                            </button>
                          )}
                          <button className="btn sm" onClick={() => setEdit(p)}>
                            {p.preset || p.builtIn ? t('details') : t('edit')}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
      {edit && (
        <EditModal
          p={edit}
          mdo={!!data.data?.mdo}
          onClose={() => setEdit(null)}
          onDone={(r) => {
            data.set(() => r)
            setEdit(null)
            toast(t('seSaved'))
          }}
        />
      )}
    </>
  )
}

type Val = string | number | boolean | string[] | null

function EditModal({ p, mdo, onClose, onDone }: { p: SecPolicy; mdo: boolean; onClose(): void; onDone(r: SecPolicies): void }) {
  const { t, tenant } = useApp()
  const { f, v } = useLabels()
  const readOnly = p.preset || p.builtIn
  const fields = FIELDS[p.kind].filter((x) => !x.mdo || mdo)
  const [vals, setVals] = useState<Record<string, Val>>(() => Object.fromEntries(fields.map((x) => [x.k, p.settings[x.k] ?? null])))
  const [text, setText] = useState<Record<string, string>>(() =>
    Object.fromEntries(fields.filter((x) => x.type === 'list').map((x) => [x.k, ((p.settings[x.k] as string[] | null) || []).join('\n')]))
  )
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  const changes = () => {
    const out: Record<string, unknown> = {}
    for (const x of fields) {
      const before = p.settings[x.k] ?? null
      if (x.type === 'list') {
        const next = entries(text[x.k] || '')
        const prev = (before as string[] | null) || []
        if (next.join('\n') !== prev.join('\n')) out[x.k] = next.length ? next : null
      } else if (vals[x.k] !== before) out[x.k] = x.type === 'num' ? Number(vals[x.k]) : x.k === 'PhishThresholdLevel' ? Number(vals[x.k]) : vals[x.k]
    }
    return out
  }

  return (
    <Modal
      wide
      title={p.name}
      sub={readOnly ? t('sePresetD') : t(`seP_${p.kind}`)}
      onClose={onClose}
      footer={
        readOnly ? (
          <button className="btn pri" onClick={onClose}>
            {t('close')}
          </button>
        ) : (
          <>
            <button className="btn" onClick={onClose}>
              {t('cancel')}
            </button>
            <button
              className="btn pri"
              disabled={busy}
              onClick={async () => {
                const c = changes()
                if (!Object.keys(c).length) return onClose()
                setBusy(true)
                setErr(null)
                try {
                  onDone(await unwrap(window.api.sec.setPolicy(tenant!.tenantId, p.kind, p.name, c)))
                } catch (e) {
                  setErr((e as Error).message)
                } finally {
                  setBusy(false)
                }
              }}
            >
              {busy && <span className="spin" />}
              {t('save')}
            </button>
          </>
        )
      }
    >
      <div className="row2">
        {fields.map((x) => (
          <div key={x.k} className="fld" style={x.type === 'list' ? { gridColumn: '1 / -1' } : undefined}>
            <label htmlFor={`pf-${x.k}`}>
              {f(x.k)} {x.mdo && <span className="lic-tag">MDO</span>}
            </label>
            {readOnly ? (
              <div className="mono">{x.type === 'list' ? ((p.settings[x.k] as string[] | null) || []).join(', ') || '—' : v(p.settings[x.k], x.k)}</div>
            ) : x.type === 'bool' ? (
              <select id={`pf-${x.k}`} value={vals[x.k] ? '1' : '0'} onChange={(e) => setVals({ ...vals, [x.k]: e.target.value === '1' })}>
                <option value="1">{t('yes')}</option>
                <option value="0">{t('no')}</option>
              </select>
            ) : x.type === 'num' ? (
              <input id={`pf-${x.k}`} type="number" min={x.min} max={x.max} value={String(vals[x.k] ?? '')} onChange={(e) => setVals({ ...vals, [x.k]: e.target.value === '' ? null : +e.target.value })} />
            ) : x.type === 'sel' ? (
              <select id={`pf-${x.k}`} value={String(vals[x.k] ?? '')} onChange={(e) => setVals({ ...vals, [x.k]: e.target.value })}>
                {!x.opts!.includes(String(vals[x.k] ?? '')) && <option value={String(vals[x.k] ?? '')}>{v(vals[x.k])}</option>}
                {x.opts!.map((o) => (
                  <option key={o} value={o}>
                    {v(o)}
                  </option>
                ))}
              </select>
            ) : (
              <textarea id={`pf-${x.k}`} className="mono" rows={3} value={text[x.k]} onChange={(e) => setText({ ...text, [x.k]: e.target.value })} />
            )}
            {x.warn && !readOnly && <small className="hint">{t('seAllowListWarn')}</small>}
          </div>
        ))}
      </div>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
