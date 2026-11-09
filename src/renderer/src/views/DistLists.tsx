import { useMemo, useState } from 'react'
import type { DistGroup, DistKind, NewDistGroup, Recipient } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, Modal } from '../components/ui'
import { AddRecipient, Chips, ExoLoad } from '../components/exo'
import { DomainPicker, useDomains } from './Mailboxes'

const KINDS: DistKind[] = ['distribution', 'security', 'dynamic']

export function DistListsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const groups = useLoad(() => window.api.exo.distGroups(tid), [tid])
  const [q, setQ] = useState('')
  const [modal, setModal] = useState<{ k: 'new' } | { k: 'detail'; g: DistGroup } | null>(null)
  const shown = (groups.data || [])
    .filter((g) => (g.displayName + g.primarySmtp).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('exchange')}
          </div>
          <h1>{t('distLists')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('searchLists')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={groups.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn pri" disabled={!groups.data} onClick={() => setModal({ k: 'new' })}>
            {Icon.plus}
            {t('dlNew')}
          </button>
        </div>
      </header>
      <div className="content">
        <ExoLoad loading={groups.loading && !groups.data} error={groups.error} onRetry={groups.reload} />
        {groups.data && (
          <div className="panel">
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('name')}</th>
                    <th>{t('typeL')}</th>
                    <th>{t('dlOwners')}</th>
                    <th>{t('dlExternal')}</th>
                    <th>{t('created')}</th>
                  </tr>
                </thead>
                <tbody>
                  {shown.map((g) => (
                    <tr key={g.id}>
                      <td>
                        <button className="linkish" onClick={() => setModal({ k: 'detail', g })}>
                          <div className="u-n">{g.displayName}</div>
                          <div className="mono">{g.primarySmtp}</div>
                        </button>
                      </td>
                      <td>
                        <span className="lic-tag">{t(`dlKind_${g.kind}`)}</span>
                      </td>
                      <td>{g.owners.length ? g.owners.join(', ') : <span className="hint">—</span>}</td>
                      <td>{g.allowExternal ? <span className="pill warn">{t('dlExtAllowed')}</span> : <span className="pill ok">{t('dlExtBlocked')}</span>}</td>
                      <td className="mono">{fmtDate(g.created, lang, '—')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!shown.length && <p className="hint" style={{ padding: 16 }}>{t('noLists')}</p>}
            </div>
          </div>
        )}
      </div>
      {modal?.k === 'new' && (
        <NewListModal
          onClose={() => setModal(null)}
          onCreated={(g) => {
            groups.set((d) => [...d, g])
            setModal(null)
            toast(t('dlCreated', g.displayName))
          }}
        />
      )}
      {modal?.k === 'detail' && (
        <ListModal
          group={modal.g}
          onClose={() => setModal(null)}
          onChanged={(g) => groups.set((d) => d.map((x) => (x.id === g.id ? g : x)))}
          onDeleted={() => {
            groups.set((d) => d.filter((x) => x.id !== modal.g.id))
            setModal(null)
            toast(t('dlDeleted', modal.g.displayName))
          }}
        />
      )}
    </>
  )
}

const aliasOf = (n: string) =>
  n
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')

const splitList = (s: string) =>
  s
    .split(',')
    .map((x) => x.trim())
    .filter(Boolean)

function NewListModal({ onClose, onCreated }: { onClose(): void; onCreated(g: DistGroup): void }) {
  const { t, tenant } = useApp()
  const domains = useDomains(useMemo(() => [...new Set([tenant!.defaultDomain.toLowerCase(), ...tenant!.domains.map((d) => d.toLowerCase())])], [tenant]))
  const [g, setG] = useState<NewDistGroup>({ kind: 'distribution', name: '', alias: '', smtp: '', members: [], owners: [], allowExternal: false, includedRecipients: 'MailboxUsers' })
  const [domain, setDomain] = useState(domains[0])
  const [touched, setTouched] = useState(false)
  const [cond, setCond] = useState({ department: '', company: '', city: '' })
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const dyn = g.kind === 'dynamic'

  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      const r = await unwrap(
        window.api.exo.newDistGroup(tenant!.tenantId, {
          ...g,
          name: g.name.trim(),
          smtp: `${g.alias}@${domain}`,
          department: splitList(cond.department),
          company: splitList(cond.company),
          city: splitList(cond.city)
        })
      )
      onCreated(r)
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('dlNew')}
      onClose={onClose}
      wide
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !g.name.trim() || !g.alias} onClick={submit}>
            {busy && <span className="spin" />}
            {t('createBtn')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label>{t('typeL')}</label>
        {KINDS.map((k) => (
          <label key={k} className={`opt ${g.kind === k ? 'on' : ''}`}>
            <input type="radio" name="dlkind" checked={g.kind === k} onChange={() => setG({ ...g, kind: k })} />
            <div>
              <b>{t(`dlKind_${k}`)}</b>
              <small>{t(`dlKindD_${k}`)}</small>
            </div>
          </label>
        ))}
      </div>
      <div className="fld">
        <label htmlFor="dl-n">{t('nmName')}</label>
        <input id="dl-n" autoFocus value={g.name} onChange={(e) => setG({ ...g, name: e.target.value, alias: touched ? g.alias : aliasOf(e.target.value) })} />
      </div>
      <div className="fld">
        <label htmlFor="dl-a">{t('nmAddress')}</label>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input id="dl-a" className="mono" value={g.alias} onChange={(e) => (setTouched(true), setG({ ...g, alias: aliasOf(e.target.value) }))} />
          <DomainPicker domains={domains} value={domain} onChange={setDomain} />
        </div>
      </div>
      {dyn ? (
        <>
          <div className="fld">
            <label>{t('dlIncluded')}</label>
            <select value={g.includedRecipients} onChange={(e) => setG({ ...g, includedRecipients: e.target.value })}>
              {(['MailboxUsers', 'AllRecipients'] as const).map((x) => (
                <option key={x} value={x}>
                  {t(`inc_${x}`)}
                </option>
              ))}
            </select>
          </div>
          <div className="row2">
            <div className="fld">
              <label>{t('dlDept')}</label>
              <input value={cond.department} onChange={(e) => setCond({ ...cond, department: e.target.value })} />
            </div>
            <div className="fld">
              <label>{t('dlCompany')}</label>
              <input value={cond.company} onChange={(e) => setCond({ ...cond, company: e.target.value })} />
            </div>
          </div>
          <div className="fld">
            <label>{t('dlCity')}</label>
            <input value={cond.city} onChange={(e) => setCond({ ...cond, city: e.target.value })} />
            <span className="hint">{t('dlCondHint')}</span>
          </div>
        </>
      ) : (
        <div className="fld">
          <label>{t('dlMembers')}</label>
          <Chips items={g.members.map((m) => ({ key: m, label: m }))} onRemove={(k) => setG({ ...g, members: g.members.filter((x) => x !== k) })} empty={t('nobody')} />
          <AddRecipient placeholder={t('addMemberPh')} onAdd={(a) => setG({ ...g, members: [...new Set([...g.members, a])] })} />
        </div>
      )}
      <div className="fld">
        <label>
          {t('dlOwners')} <span className="hint">· {t('dlOwnersHint')}</span>
        </label>
        <Chips items={g.owners.map((m) => ({ key: m, label: m }))} onRemove={(k) => setG({ ...g, owners: g.owners.filter((x) => x !== k) })} empty={t('nobody')} />
        <AddRecipient
          placeholder={t('userPh')}
          filter={(r: Recipient) => r.type === 'UserMailbox'}
          onAdd={(a) => setG({ ...g, owners: dyn ? [a] : [...new Set([...g.owners, a])] })}
        />
      </div>
      <label className="opt">
        <input type="checkbox" className="chk" checked={g.allowExternal} onChange={(e) => setG({ ...g, allowExternal: e.target.checked })} />
        <div>
          <b>{t('dlAllowExt')}</b>
        </div>
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function ListModal({ group, onClose, onChanged, onDeleted }: { group: DistGroup; onClose(): void; onChanged(g: DistGroup): void; onDeleted(): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const detail = useLoad(() => window.api.exo.distGroup(tid, group.id, group.kind), [tid, group.id])
  const [busy, setBusy] = useState<string | null>(null)
  const dyn = group.kind === 'dynamic'
  const d = detail.data

  async function step<T>(key: string, fn: () => Promise<T>, done?: (r: T) => void) {
    setBusy(key)
    try {
      const r = await fn()
      done?.(r)
      toast(t('dlSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  const ownerAddrs = (d?.owners || []).map((o) => o.primarySmtp || o.displayName)
  const setOwners = (owners: string[]) =>
    step('owners', () => unwrap(window.api.exo.setDistGroup(tid, group.id, group.kind, { owners })), (g) => {
      onChanged(g)
      detail.reload()
    })

  return (
    <Modal
      title={group.displayName}
      sub={<span className="mono">{group.primarySmtp}</span>}
      onClose={onClose}
      wide
      footer={
        <button
          className="btn"
          style={{ marginRight: 'auto', color: 'var(--bad)' }}
          disabled={!!busy}
          onClick={() => confirm(t('dlDelConfirm', group.displayName)) && step('del', () => unwrap(window.api.exo.removeDistGroup(tid, group.id, group.kind)), onDeleted)}
        >
          {t('dlDelete')}
        </button>
      }
    >
      {!d ? (
        <ExoLoad loading={detail.loading} error={detail.error} onRetry={detail.reload} />
      ) : (
        <>
          <div className="fld">
            <label>{t('typeL')}</label>
            <span>{t(`dlKind_${group.kind}`)}</span>
          </div>
          {dyn && d.group.filter && (
            <div className="fld">
              <label>{t('dlFilter')}</label>
              <code className="mono" style={{ fontSize: 12, wordBreak: 'break-word' }}>
                {d.group.filter}
              </code>
            </div>
          )}
          <div className="fld">
            <label>
              {t('dlOwners')} <span className="hint">· {t('dlOwnersHint')}</span>
            </label>
            <Chips
              items={d.owners.map((o) => ({ key: o.primarySmtp || o.displayName, label: o.primarySmtp ? `${o.displayName} · ${o.primarySmtp}` : o.displayName }))}
              onRemove={(k) => setOwners(ownerAddrs.filter((x) => x !== k))}
              empty={t('nobody')}
            />
            <AddRecipient
              placeholder={t('userPh')}
              busy={busy === 'owners'}
              filter={(r) => r.type === 'UserMailbox'}
              onAdd={(a) => setOwners(dyn ? [a] : [...new Set([...ownerAddrs, a])])}
            />
          </div>
          <label className="opt">
            <input
              type="checkbox"
              className="chk"
              checked={d.group.allowExternal}
              disabled={busy === 'ext'}
              onChange={(e) =>
                step('ext', () => unwrap(window.api.exo.setDistGroup(tid, group.id, group.kind, { allowExternal: e.target.checked })), (g) => {
                  onChanged(g)
                  detail.set((x) => ({ ...x, group: g }))
                })
              }
            />
            <div>
              <b>{t('dlAllowExt')}</b>
            </div>
          </label>
          <div className="fld">
            <label>
              {t('dlMembers')} <span className="hint">· {d.members.length}</span>
            </label>
            {dyn && <span className="hint">{t('dlDynHint')}</span>}
            <Chips
              items={d.members.map((m) => ({ key: m.primarySmtp || m.id, label: `${m.displayName}${m.primarySmtp ? ` · ${m.primarySmtp}` : ''}` }))}
              onRemove={
                dyn
                  ? undefined
                  : (k) => step('member', () => unwrap(window.api.exo.removeMember(tid, group.id, k)), (members) => detail.set((x) => ({ ...x, members })))
              }
              empty={t('nobody')}
            />
            {!dyn && (
              <AddRecipient
                placeholder={t('addMemberPh')}
                busy={busy === 'member'}
                onAdd={(a) => step('member', () => unwrap(window.api.exo.addMember(tid, group.id, a)), (members) => detail.set((x) => ({ ...x, members })))}
              />
            )}
          </div>
        </>
      )}
    </Modal>
  )
}
