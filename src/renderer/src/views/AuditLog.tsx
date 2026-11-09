import { useState } from 'react'
import type { AuditFilter, AuditRecord } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { dayEnd, dayInput, dayStart, entries, fmtWhen, Head, SecLoad } from '../components/security'
import { auditPrefill } from './Compromised'

const RECORD_TYPES = ['', 'ExchangeItem', 'ExchangeItemGroup', 'ExchangeAdmin', 'AzureActiveDirectory', 'AzureActiveDirectoryStsLogon', 'SharePointFileOperation', 'MicrosoftTeams']
/** Common questions, as the operations that answer them. */
const PRESETS: [string, string[]][] = [
  ['rules', ['New-InboxRule', 'Set-InboxRule', 'UpdateInboxRules', 'Remove-InboxRule']],
  ['forwarding', ['Set-Mailbox']],
  ['deleted', ['SoftDelete', 'HardDelete', 'MoveToDeletedItems']],
  ['access', ['MailItemsAccessed']],
  ['perms', ['Add-MailboxPermission', 'Add-RecipientPermission', 'Remove-MailboxPermission']],
  ['roles', ['Add member to role.', 'Remove member from role.']],
  ['signin', ['UserLoggedIn', 'UserLoginFailed']]
]

/** Searches the unified audit log: who did what in a mailbox or in the tenant. */
export function AuditLogView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const status = useLoad(() => window.api.sec.auditStatus(tid), [tid])
  const [draft, setDraft] = useState(() => {
    const u = auditPrefill.user
    auditPrefill.user = null
    return { from: dayInput(7), to: dayInput(0), users: u || '', operations: '', recordType: '', freeText: '' }
  })
  const [filter, setFilter] = useState<AuditFilter | null>(null)
  const list = useLoad(() => (filter ? window.api.sec.audit(tid, filter) : null), [tid, filter])
  const [open, setOpen] = useState<AuditRecord | null>(null)
  const [busy, setBusy] = useState(false)

  const apply = () =>
    setFilter({
      start: dayStart(draft.from)!,
      end: draft.to === dayInput(0) ? new Date().toISOString() : dayEnd(draft.to)!,
      users: entries(draft.users),
      operations: draft.operations
        .split(',')
        .map((x) => x.trim())
        .filter(Boolean),
      recordType: draft.recordType,
      freeText: draft.freeText.trim()
    })

  function exportCsv() {
    const cols: (keyof AuditRecord)[] = ['date', 'user', 'operation', 'recordType', 'workload', 'clientIp', 'object', 'mailbox', 'result', 'data']
    const q = (v: string) => `"${v.replace(/"/g, '""')}"`
    const csv = [cols.join(','), ...(list.data || []).map((r) => cols.map((c) => q(String(r[c] ?? ''))).join(','))].join('\r\n')
    const a = document.createElement('a')
    a.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv' }))
    a.download = `audit-${tenant!.defaultDomain}-${draft.from}-${draft.to}.csv`
    a.click()
    URL.revokeObjectURL(a.href)
  }

  return (
    <>
      <Head title={t('auditLog')} />
      <div className="content">
        <SecLoad loading={status.loading} error={status.error} onRetry={status.reload} />
        {status.data && !status.data.enabled && (
          <div className="banner bad">
            <span>{t('seAlOff')}</span>
            <button
              className="btn sm"
              disabled={busy}
              onClick={async () => {
                setBusy(true)
                try {
                  status.set(() => ({ enabled: true }))
                  await unwrap(window.api.sec.enableAudit(tid))
                  toast(t('seAlOn'))
                } catch (e) {
                  toast((e as Error).message)
                  status.reload()
                } finally {
                  setBusy(false)
                }
              }}
            >
              {t('seAlEnable')}
            </button>
          </div>
        )}
        <div className="panel" style={{ marginBottom: 16, padding: 14 }}>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 12 }}>
            {PRESETS.map(([k, ops]) => (
              <button key={k} className={`btn sm ${draft.operations === ops.join(', ') ? 'pri' : ''}`} onClick={() => setDraft({ ...draft, operations: ops.join(', ') })}>
                {t(`seAlP_${k}` as 'seAlP_rules')}
              </button>
            ))}
          </div>
          <div className="row2" style={{ gridTemplateColumns: 'repeat(3, minmax(0,1fr))' }}>
            <div className="fld">
              <label htmlFor="au">{t('seAlUsers')}</label>
              <input id="au" className="mono" value={draft.users} onChange={(e) => setDraft({ ...draft, users: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="ao">{t('seAlOps')}</label>
              <input id="ao" className="mono" value={draft.operations} onChange={(e) => setDraft({ ...draft, operations: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="ar">{t('seAlType')}</label>
              <select id="ar" value={draft.recordType} onChange={(e) => setDraft({ ...draft, recordType: e.target.value })}>
                {RECORD_TYPES.map((r) => (
                  <option key={r} value={r}>
                    {r ? (t(`seAlT_${r}` as 'seAlT_ExchangeItem') as string | undefined) || r : t('seAlAnyType')}
                  </option>
                ))}
              </select>
            </div>
            <div className="fld">
              <label htmlFor="af">{t('seFrom')}</label>
              <input id="af" type="date" min={dayInput(180)} value={draft.from} onChange={(e) => setDraft({ ...draft, from: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="at">{t('seTo')}</label>
              <input id="at" type="date" value={draft.to} onChange={(e) => setDraft({ ...draft, to: e.target.value })} />
            </div>
            <div className="fld">
              <label htmlFor="ax">{t('seAlFree')}</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input id="ax" style={{ flex: 1 }} value={draft.freeText} onChange={(e) => setDraft({ ...draft, freeText: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} />
                <button className="btn pri" onClick={apply}>
                  {Icon.search}
                  {t('seSearch')}
                </button>
              </div>
            </div>
          </div>
          <p className="hint" style={{ margin: 0 }}>
            {t('seAlHint')}
          </p>
        </div>
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} text={t('seAlSearching')} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{t('seAlCount', list.data.length)}</h2>
              <div className="r">
                <button className="btn sm" disabled={!list.data.length} onClick={exportCsv}>
                  {t('seExportCsv')}
                </button>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('seWhen')}</th>
                    <th>{t('user')}</th>
                    <th>{t('seAlOp')}</th>
                    <th>{t('seAlObject')}</th>
                    <th>{t('seCpIp')}</th>
                    <th>{t('seRpResult')}</th>
                  </tr>
                </thead>
                <tbody>
                  {list.data.map((r) => (
                    <tr key={r.id}>
                      <td className="mono">{fmtWhen(r.date, lang)}</td>
                      <td className="mono">{r.user}</td>
                      <td>
                        <button className="linkish" onClick={() => setOpen(r)}>
                          <span className="u-n">{r.operation}</span>
                        </button>
                        <div className="hint">{r.workload || r.recordType}</div>
                      </td>
                      <td className="mono" style={{ whiteSpace: 'normal', wordBreak: 'break-all', maxWidth: 340 }}>
                        {r.mailbox && r.mailbox !== r.user ? `${r.mailbox} · ` : ''}
                        {r.object}
                      </td>
                      <td className="mono">{r.clientIp}</td>
                      <td>{r.result}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!list.data.length && <p className="hint" style={{ padding: 16 }}>{t('seAlNone')}</p>}
              {list.data.length >= 5000 && <p className="hint" style={{ padding: 12 }}>{t('seMtCapped')}</p>}
            </div>
          </div>
        )}
      </div>
      {open && (
        <Modal wide title={open.operation} sub={`${open.user} · ${fmtWhen(open.date, lang)}`} onClose={() => setOpen(null)}>
          <pre className="mono" style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-all', maxHeight: 460, overflow: 'auto', fontSize: 11.5, margin: 0 }}>
            {(() => {
              try {
                return JSON.stringify(JSON.parse(open.data), null, 2)
              } catch {
                return open.data
              }
            })()}
          </pre>
        </Modal>
      )}
    </>
  )
}
