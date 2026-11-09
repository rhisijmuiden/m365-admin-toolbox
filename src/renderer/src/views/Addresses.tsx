import { useMemo, useState } from 'react'
import type { Recipient } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { ExoLoad, domainOf, stripSmtp, useTypeLabel } from '../components/exo'
import { DomainPicker, useDomains } from './Mailboxes'

const isPrimary = (a: string) => a.startsWith('SMTP:')

export function AddressesView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const typeLabel = useTypeLabel()
  const list = useLoad(() => window.api.exo.recipients(tid), [tid])
  const [q, setQ] = useState('')
  const [dom, setDom] = useState('')
  const [edit, setEdit] = useState<Recipient | null>(null)

  // Addresses per domain, so you can see at a glance what still uses an old domain.
  const counts = useMemo(() => {
    const c = new Map<string, number>()
    for (const r of list.data || []) for (const a of r.addresses) c.set(domainOf(a), (c.get(domainOf(a)) || 0) + 1)
    return [...c.entries()].sort((a, b) => b[1] - a[1])
  }, [list.data])

  const shown = (list.data || [])
    .filter((r) => !dom || r.addresses.some((a) => domainOf(a) === dom))
    .filter((r) => (r.displayName + r.addresses.join(' ')).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('exchange')}
          </div>
          <h1>{t('aliases')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('searchAddr')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <ExoLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <div className="seg" style={{ flexWrap: 'wrap' }}>
                <button className={!dom ? 'on' : ''} onClick={() => setDom('')}>
                  {t('adAll')}
                </button>
                {counts.map(([d, n]) => (
                  <button key={d} className={dom === d ? 'on' : ''} onClick={() => setDom(d)}>
                    {d} <span className="hint">{n}</span>
                  </button>
                ))}
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('adRecipient')}</th>
                    <th>{t('typeL')}</th>
                    <th>{t('adAddresses')}</th>
                    <th style={{ width: 48 }} aria-label={t('actions')} />
                  </tr>
                </thead>
                <tbody>
                  {shown.map((r) => (
                    <tr key={r.id}>
                      <td>
                        <div className="u-n">{r.displayName}</div>
                      </td>
                      <td>
                        <span className="lic-tag">{typeLabel(r.type)}</span>
                      </td>
                      <td className="mono">
                        {r.addresses
                          .filter((a) => !dom || domainOf(a) === dom)
                          .map((a) => (
                            <div key={a} style={isPrimary(a) ? { fontWeight: 600 } : undefined}>
                              {stripSmtp(a)}
                            </div>
                          ))}
                      </td>
                      <td>
                        <button className="btn sm" onClick={() => setEdit(r)}>
                          {t('adEdit')}
                        </button>
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
        <EditAddresses
          r={edit}
          domains={counts.map(([d]) => d)}
          onClose={() => setEdit(null)}
          onSaved={(n) => {
            list.set((d) => d.map((x) => (x.id === n.id ? n : x)))
            setEdit(null)
            toast(t('adSaved'))
          }}
        />
      )}
    </>
  )
}

function EditAddresses({ r, domains, onClose, onSaved }: { r: Recipient; domains: string[]; onClose(): void; onSaved(r: Recipient): void }) {
  const { t, tenant } = useApp()
  const doms = useDomains(useMemo(() => [...new Set([...tenant!.domains.map((d) => d.toLowerCase()), ...domains])], [tenant, domains]))
  const current = r.addresses.map(stripSmtp)
  const [primary, setPrimary] = useState(stripSmtp(r.addresses.find(isPrimary) || r.primarySmtp))
  const [addrs, setAddrs] = useState(current)
  const [local, setLocal] = useState('')
  const [domain, setDomain] = useState(doms[0])
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const lc = (a: string) => a.toLowerCase()
  const add = addrs.filter((a) => !current.some((c) => lc(c) === lc(a)))
  const remove = current.filter((c) => !addrs.some((a) => lc(a) === lc(c)))
  const primaryChanged = lc(primary) !== lc(stripSmtp(r.addresses.find(isPrimary) || r.primarySmtp))

  async function save() {
    setBusy(true)
    setErr(null)
    try {
      onSaved(await unwrap(window.api.exo.setAddresses(tenant!.tenantId, r, { primary: primaryChanged ? primary : undefined, add, remove })))
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('adEdit')}
      sub={r.displayName}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || (!add.length && !remove.length && !primaryChanged)} onClick={save}>
            {busy && <span className="spin" />}
            {t('save')}
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: 6 }}>
        {addrs.map((a) => {
          const p = lc(a) === lc(primary)
          return (
            <div key={a} className="opt" style={{ alignItems: 'center' }}>
              <span className="mono" style={{ flex: 1, fontWeight: p ? 600 : 400 }}>
                {a}
              </span>
              {p ? (
                <span className="pill ok">{t('primary')}</span>
              ) : (
                <>
                  <button className="btn sm" onClick={() => setPrimary(a)}>
                    {t('adMakePrimary')}
                  </button>
                  <button className="btn sm" onClick={() => setAddrs(addrs.filter((x) => x !== a))}>
                    {t('remove')}
                  </button>
                </>
              )}
            </div>
          )
        })}
      </div>
      <div className="fld">
        <label>{t('adAdd')}</label>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input className="mono" value={local} onChange={(e) => setLocal(e.target.value.replace(/[\s@]/g, ''))} />
          <DomainPicker domains={doms} value={domain} onChange={setDomain} />
          <button
            className="btn sm"
            disabled={!local}
            onClick={() => {
              const a = `${local}@${domain}`
              if (!addrs.some((x) => lc(x) === lc(a))) setAddrs([...addrs, a])
              setLocal('')
            }}
          >
            {t('add')}
          </button>
        </div>
      </div>
      <span className="hint">{t('adHint')}</span>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
