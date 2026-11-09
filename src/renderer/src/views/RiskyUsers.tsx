import { useState } from 'react'
import type { RiskyUser } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, RowMenu } from '../components/ui'
import { Tabs } from '../components/exo'
import { fmtWhen, Head, SecLoad } from '../components/security'
import { compromisedPrefill } from './Compromised'

const levelCls = (l: string) => (l === 'high' ? 'bad' : l === 'medium' ? 'warn' : l === 'low' ? 'off' : 'off')
const ACTIVE = ['atRisk', 'confirmedCompromised']

/** Entra ID Protection: risky users and risk detections (sign-ins and user risk). Needs Entra ID P2. */
export function RiskyUsersView() {
  const { t, tenant } = useApp()
  const [tab, setTab] = useState<'users' | 'detections'>('users')
  return (
    <>
      <Head title={t('riskyUsers')} />
      <div className="content">
        <Tabs<'users' | 'detections'>
          tabs={[
            ['users', t('seRuUsers')],
            ['detections', t('seRuDetections')]
          ]}
          value={tab}
          onChange={setTab}
        />
        {tab === 'users' ? <Users key={tenant!.tenantId} /> : <Detections key={tenant!.tenantId} />}
      </div>
    </>
  )
}

function Users() {
  const { t, lang, tenant, toast, go } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.sec.riskyUsers(tid), [tid])
  const [all, setAll] = useState(false)
  const [busy, setBusy] = useState(false)
  const rows = (list.data || []).filter((u) => all || ACTIVE.includes(u.riskState)).sort((a, b) => ['high', 'medium', 'low'].indexOf(a.riskLevel) - ['high', 'medium', 'low'].indexOf(b.riskLevel))

  async function act(u: RiskyUser, action: 'dismiss' | 'confirmCompromised') {
    if (!confirm(t(action === 'dismiss' ? 'seRuDismissConfirm' : 'seRuConfirmConfirm', u.displayName))) return
    setBusy(true)
    try {
      await unwrap(window.api.sec.setRisk(tid, [u.id], action))
      list.set((d) => d.map((x) => (x.id === u.id ? { ...x, riskState: action === 'dismiss' ? 'dismissed' : 'confirmedCompromised', riskLevel: action === 'dismiss' ? 'none' : 'high' } : x)))
      toast(t('seSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} consent="risk" />
      {list.data && (
        <div className="panel">
          <div className="panel-h">
            <h2>{t('seRuCount', rows.length)}</h2>
            <div className="r" style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <label className="hint" style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <input type="checkbox" className="chk" checked={all} onChange={(e) => setAll(e.target.checked)} />
                {t('seRuShowAll')}
              </label>
              <button className="btn sm" onClick={list.reload} aria-label={t('refresh')}>
                {Icon.refresh}
              </button>
            </div>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('user')}</th>
                  <th>{t('seRuLevel')}</th>
                  <th>{t('seRuState')}</th>
                  <th>{t('seRuDetail')}</th>
                  <th>{t('seRuUpdated')}</th>
                  <th style={{ width: 40 }} />
                </tr>
              </thead>
              <tbody>
                {rows.map((u) => (
                  <tr key={u.id}>
                    <td>
                      <div className="u-n">{u.displayName}</div>
                      <div className="mono">{u.upn}</div>
                    </td>
                    <td>
                      <span className={`pill ${levelCls(u.riskLevel)}`}>{t(`seLv_${u.riskLevel}` as 'seLv_high') || u.riskLevel}</span>
                    </td>
                    <td>{(t(`seRs_${u.riskState}` as 'seRs_atRisk') as string | undefined) || u.riskState}</td>
                    <td className="hint">{u.riskDetail && u.riskDetail !== 'none' ? u.riskDetail : '—'}</td>
                    <td className="mono">{fmtWhen(u.updated, lang)}</td>
                    <td>
                      <RowMenu
                        items={[
                          [t('compromised'), () => ((compromisedPrefill.upn = u.upn), go('compromised'))],
                          [t('seRuConfirm'), () => !busy && act(u, 'confirmCompromised')],
                          [t('seRuDismiss'), () => !busy && act(u, 'dismiss')]
                        ]}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!rows.length && <p className="hint" style={{ padding: 16 }}>{t('seRuNone')}</p>}
          </div>
        </div>
      )}
    </>
  )
}

function Detections() {
  const { t, lang, tenant } = useApp()
  const tid = tenant!.tenantId
  const [days, setDays] = useState(30)
  const list = useLoad(() => window.api.sec.riskDetections(tid, days), [tid, days])
  return (
    <>
      <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} consent="risk" />
      {list.data && (
        <div className="panel">
          <div className="panel-h">
            <h2>{t('seRuDetCount', list.data.length)}</h2>
            <div className="r">
              <select value={days} onChange={(e) => setDays(+e.target.value)} aria-label={t('seRuPeriod')}>
                {[7, 30, 90].map((d) => (
                  <option key={d} value={d}>
                    {t('seLastDays', d)}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('seWhen')}</th>
                  <th>{t('user')}</th>
                  <th>{t('seRuType')}</th>
                  <th>{t('seRuLevel')}</th>
                  <th>{t('seCpIp')}</th>
                  <th>{t('seLocation')}</th>
                  <th>{t('seRuState')}</th>
                </tr>
              </thead>
              <tbody>
                {list.data.map((d) => (
                  <tr key={d.id}>
                    <td className="mono">{fmtWhen(d.date, lang)}</td>
                    <td>
                      <div className="u-n">{d.user}</div>
                      <div className="mono">{d.upn}</div>
                    </td>
                    <td>
                      {(t(`seDt_${d.type}` as 'seDt_unfamiliarFeatures') as string | undefined) || d.type}
                      <div className="hint">{d.activity}</div>
                    </td>
                    <td>
                      <span className={`pill ${levelCls(d.level)}`}>{t(`seLv_${d.level}` as 'seLv_high') || d.level}</span>
                    </td>
                    <td className="mono">{d.ip}</td>
                    <td>{d.location || '—'}</td>
                    <td>{(t(`seRs_${d.state}` as 'seRs_atRisk') as string | undefined) || d.state}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!list.data.length && <p className="hint" style={{ padding: 16 }}>{t('seRuDetNone')}</p>}
          </div>
        </div>
      )}
    </>
  )
}
