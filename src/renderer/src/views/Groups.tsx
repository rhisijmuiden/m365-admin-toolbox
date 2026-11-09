import { useState } from 'react'
import type { NewGroup } from '@shared/types'
import { unwrap, useApp, useLoad, takeSearch } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, LoadState, Modal } from '../components/ui'

export function GroupsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const groups = useLoad(() => window.api.listGroups(tid), [tid])
  const [q, setQ] = useState(takeSearch)
  const [open, setOpen] = useState(false)
  const list = (groups.data || []).filter((g) => (g.displayName + (g.mail || '')).toLowerCase().includes(q.toLowerCase()))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('m365')}
          </div>
          <h1>{t('groups')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input value={q} onChange={(e) => setQ(e.target.value)} aria-label={t('grpName')} />
          </label>
          <button className="btn" onClick={groups.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn pri" onClick={() => setOpen(true)}>
            {Icon.plus}
            {t('newGroup')}
          </button>
        </div>
      </header>
      <div className="content">
        <LoadState loading={groups.loading && !groups.data} error={groups.error} />
        {groups.data && (
          <div className="panel">
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('grpName')}</th>
                    <th>{t('grpType')}</th>
                    <th>{t('mail')}</th>
                    <th>{t('created')}</th>
                  </tr>
                </thead>
                <tbody>
                  {list.map((g) => (
                    <tr key={g.id}>
                      <td>
                        <div className="u-n">{g.displayName}</div>
                        {g.description && <div className="hint">{g.description}</div>}
                      </td>
                      <td>
                        <span className="lic-tag">{t(`kind_${g.kind}` as 'kind_m365')}</span>
                      </td>
                      <td className="mono">{g.mail || '—'}</td>
                      <td className="mono">{fmtDate(g.createdDateTime, lang, '—')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
      {open && (
        <NewGroupModal
          onClose={() => setOpen(false)}
          onCreated={(name) => {
            setOpen(false)
            toast(t('grpCreated', name))
            groups.reload()
          }}
        />
      )}
    </>
  )
}

function NewGroupModal({ onClose, onCreated }: { onClose(): void; onCreated(name: string): void }) {
  const { t, tenant } = useApp()
  const tid = tenant!.tenantId
  const users = useLoad(() => window.api.listUsers(tid), [tid])
  const [g, setG] = useState<NewGroup>({ displayName: '', description: '', mailNickname: '', kind: 'm365' })
  const [aliasTouched, setAliasTouched] = useState(false)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  const alias = (n: string) =>
    n
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '')

  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      const r = await unwrap(window.api.createGroup(tid, { ...g, ownerId: g.ownerId || undefined }))
      onCreated(r.displayName)
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('ngTitle')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !g.displayName.trim() || !g.mailNickname} onClick={submit}>
            {busy && <span className="spin" />}
            {t('createGrp')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label>{t('ngType')}</label>
        {(['m365', 'security'] as const).map((k) => (
          <label key={k} className={`opt ${g.kind === k ? 'on' : ''}`}>
            <input type="radio" name="kind" checked={g.kind === k} onChange={() => setG({ ...g, kind: k })} />
            <div>
              <b>{t(k === 'm365' ? 'ngM365' : 'ngSec')}</b>
              <small>{t(k === 'm365' ? 'ngM365D' : 'ngSecD')}</small>
            </div>
          </label>
        ))}
        <span className="hint">{t('ngDistNote')}</span>
      </div>
      <div className="fld">
        <label htmlFor="gn">{t('grpName')}</label>
        <input
          id="gn"
          autoFocus
          value={g.displayName}
          onChange={(e) => setG({ ...g, displayName: e.target.value, mailNickname: aliasTouched ? g.mailNickname : alias(e.target.value) })}
        />
      </div>
      <div className="fld">
        <label htmlFor="ga">{t('ngMail')}</label>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input
            id="ga"
            className="mono"
            value={g.mailNickname}
            onChange={(e) => (setAliasTouched(true), setG({ ...g, mailNickname: alias(e.target.value) }))}
          />
          {g.kind === 'm365' && <span className="mono">@{tenant!.defaultDomain}</span>}
        </div>
      </div>
      <div className="fld">
        <label htmlFor="gd">{t('ngDesc')}</label>
        <input id="gd" value={g.description} onChange={(e) => setG({ ...g, description: e.target.value })} />
      </div>
      <div className="fld">
        <label htmlFor="go">{t('ngOwner')}</label>
        <select id="go" value={g.ownerId || ''} onChange={(e) => setG({ ...g, ownerId: e.target.value })}>
          <option value="">{users.loading ? t('loading') : t('ngNoOwner')}</option>
          {(users.data || [])
            .filter((u) => u.accountEnabled)
            .map((u) => (
              <option key={u.id} value={u.id}>
                {u.displayName} ({u.userPrincipalName})
              </option>
            ))}
        </select>
      </div>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
