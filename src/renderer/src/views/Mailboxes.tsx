import { useMemo, useState, type ReactNode } from 'react'
import type { AutoReply, CalendarSettings, Mailbox, MailboxPerms, MailboxSize, MailRight, NewMailbox, UserAccess } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate, type Key } from '../i18n'
import { Icon, Modal, RowMenu } from '../components/ui'
import { AddRecipient, Chips, ExoLoad, fmtBytes, isExternal, stripSmtp, Tabs, useTypeLabel } from '../components/exo'

export type MailboxScope = 'all' | 'shared' | 'resources'
type Filter = 'all' | 'UserMailbox' | 'SharedMailbox' | 'resources'

const RES = ['RoomMailbox', 'EquipmentMailbox']
const inScope = (m: Mailbox, f: Filter) => f === 'all' || (f === 'resources' ? RES.includes(m.type) : m.type === f)

export function MailboxesView({ scope }: { scope: MailboxScope }) {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const typeLabel = useTypeLabel()
  const list = useLoad(() => window.api.exo.mailboxes(tid), [tid])
  const [live, setLive] = useState(false)
  const [consented, setConsented] = useState(0)
  const upns = useMemo(() => (list.data?.mailboxes || []).map((m) => m.upn), [list.data])
  const sizes = useLoad(() => (list.data ? window.api.exo.sizes(tid, upns, live) : null), [tid, list.data, live, consented])
  const [q, setQ] = useState('')
  const [filter, setFilter] = useState<Filter>(scope === 'shared' ? 'SharedMailbox' : scope === 'resources' ? 'resources' : 'all')
  const [modal, setModal] = useState<{ k: 'detail'; m: Mailbox } | { k: 'convert'; m: Mailbox } | { k: 'new' } | null>(null)

  const domains = list.data?.domains || []
  const sizeOf = useMemo(() => {
    const byId = new Map<string, MailboxSize>()
    const byUpn = new Map<string, MailboxSize>()
    for (const s of sizes.data?.sizes || []) {
      if (s.id) byId.set(s.id, s)
      if (s.upn) byUpn.set(s.upn, s)
    }
    return (m: Mailbox) => byId.get(m.id) || byUpn.get(m.upn.toLowerCase())
  }, [sizes.data])

  const all = (list.data?.mailboxes || []).filter((m) => inScope(m, filter))
  const shown = all
    .filter((m) => (m.displayName + m.primarySmtp + m.alias).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
  const pct = (m: Mailbox) => {
    const s = sizeOf(m)
    const quota = m.quotaBytes ?? s?.quotaBytes ?? null
    return s?.bytes != null && quota ? (s.bytes / quota) * 100 : null
  }
  const nearFull = all.filter((m) => (pct(m) ?? 0) >= 90)
  const extFwd = all.filter((m) => isExternal(m.forwardingSmtp, domains))
  const reload = () => list.reload()

  async function toRegular(m: Mailbox) {
    if (!confirm(t('convRegConfirm', m.displayName))) return
    try {
      const n = await unwrap(window.api.exo.setMailbox(tid, m.id, { Type: 'Regular' }))
      list.set((d) => ({ ...d, mailboxes: d.mailboxes.map((x) => (x.id === n.id ? n : x)) }))
      toast(t('convRegDone', m.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  async function archive(m: Mailbox) {
    try {
      const n = await unwrap(window.api.exo.enableArchive(tid, m.id))
      list.set((d) => ({ ...d, mailboxes: d.mailboxes.map((x) => (x.id === n.id ? n : x)) }))
      toast(t('archEnabled', m.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  const title = scope === 'shared' ? t('shared') : scope === 'resources' ? t('rooms') : t('mailboxes')
  const sizeNote = sizes.loading ? (
    <span className="hint" style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
      <span className="spin" /> {t('sizesLoading')}
    </span>
  ) : sizes.data ? (
    <span className="hint">
      {sizes.data.source === 'graph' ? t('sizesGraph', fmtDate(sizes.data.reportDate || null, lang, '—')) : t('sizesExo')}
    </span>
  ) : null

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('exchange')}
          </div>
          <h1>{title}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('searchMbx')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          {scope !== 'all' && (
            <button className="btn pri" disabled={!list.data} onClick={() => setModal({ k: 'new' })}>
              {Icon.plus}
              {scope === 'shared' ? t('newShared') : t('newResource')}
            </button>
          )}
        </div>
      </header>
      <div className="content">
        <ExoLoad loading={list.loading && !list.data} error={list.error} onRetry={reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('mbxTotal')}</div>
                <div className="v">{all.length}</div>
                <div className="s">{tenant!.defaultDomain}</div>
              </div>
              <div className={`tile ${nearFull.length ? 'bad' : ''}`}>
                <div className="k">{t('mbxNearFull')}</div>
                <div className="v">{sizes.data ? nearFull.length : '…'}</div>
                <div className="s">{nearFull.map((m) => m.displayName).join(', ') || t('ok')}</div>
              </div>
              <div className={`tile ${extFwd.length ? 'warn' : ''}`}>
                <div className="k">{t('mbxExtFwd')}</div>
                <div className="v">{extFwd.length}</div>
                <div className="s">{extFwd.map((m) => m.displayName).join(', ') || t('ok')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('mbxArchives')}</div>
                <div className="v">{all.filter((m) => m.archive).length}</div>
                <div className="s">—</div>
              </div>
            </div>

            {sizes.data?.source === 'exo' && sizes.data.graphIssue === 'consent' && (
              <div className="banner">
                <span>{t('sizesConsent')}</span>
                <button
                  className="btn sm"
                  onClick={async () => {
                    try {
                      await unwrap(window.api.exo.consentReports(tid))
                      setConsented((n) => n + 1)
                    } catch (e) {
                      toast((e as Error).message)
                    }
                  }}
                >
                  {t('grantAccess')}
                </button>
              </div>
            )}
            {sizes.data?.graphIssue === 'concealed' && (
              <div className="banner">
                <span>{t('sizesConcealed')}</span>
              </div>
            )}
            {sizes.error && <ExoLoad loading={false} error={sizes.error} onRetry={sizes.reload} />}

            <div className="panel">
              <div className="panel-h">
                {scope === 'all' && (
                  <div className="seg">
                    {(['all', 'UserMailbox', 'SharedMailbox', 'resources'] as Filter[]).map((f) => (
                      <button key={f} className={filter === f ? 'on' : ''} onClick={() => setFilter(f)}>
                        {f === 'all' ? t('mbxAll') : f === 'resources' ? t('mbxResources') : typeLabel(f)}
                      </button>
                    ))}
                  </div>
                )}
                {sizeNote}
                <div className="r">
                  {sizes.data?.source === 'graph' && (
                    <button className="btn sm" onClick={() => setLive(true)}>
                      {t('liveSizes')}
                    </button>
                  )}
                </div>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('name')}</th>
                      <th>{t('typeL')}</th>
                      <th>{t('size')}</th>
                      <th>{t('archive')}</th>
                      <th>{t('forwarding')}</th>
                      <th>{t('lastAct')}</th>
                      <th style={{ width: 48 }} aria-label={t('actions')} />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((m) => {
                      const s = sizeOf(m)
                      const p = pct(m)
                      const quota = m.quotaBytes ?? s?.quotaBytes ?? null
                      const ext = isExternal(m.forwardingSmtp, domains)
                      return (
                        <tr key={m.id}>
                          <td>
                            <button className="linkish" onClick={() => setModal({ k: 'detail', m })}>
                              <div className="u-n">{m.displayName}</div>
                              <div className="mono">{m.primarySmtp}</div>
                            </button>
                          </td>
                          <td>
                            <span className="lic-tag">{typeLabel(m.type)}</span>
                          </td>
                          <td style={{ minWidth: 190 }}>
                            {s?.bytes != null ? (
                              <>
                                <div className={`bar ${p != null && p >= 90 ? 'full' : p != null && p >= 75 ? 'hi' : ''}`}>
                                  <i style={{ width: `${Math.min(100, p ?? 0)}%` }} />
                                </div>
                                <div className="hint mono" style={{ marginTop: 4 }}>
                                  {fmtBytes(s.bytes, lang)}
                                  {quota ? ` / ${fmtBytes(quota, lang)} · ${Math.round(p ?? 0)}%` : ''}
                                </div>
                              </>
                            ) : (
                              <span className="hint">{sizes.loading ? '…' : '—'}</span>
                            )}
                          </td>
                          <td>
                            {m.archive ? (
                              <span className="pill ok">{s?.archiveBytes != null ? fmtBytes(s.archiveBytes, lang) : t('yes')}</span>
                            ) : (
                              <span className="pill off">{t('no')}</span>
                            )}
                          </td>
                          <td>
                            {m.forwardingSmtp || m.forwardingAddress ? (
                              <span className={`pill ${ext ? 'bad' : 'warn'}`} title={m.forwardingSmtp || m.forwardingAddress || ''}>
                                {ext ? t('fwdExternal') : t('yes')}
                              </span>
                            ) : (
                              <span className="hint">—</span>
                            )}
                          </td>
                          <td className="mono">{fmtDate(s?.lastActivity ?? null, lang, '—')}</td>
                          <td>
                            <RowMenu
                              items={[
                                [t('details'), () => setModal({ k: 'detail', m })],
                                ...(m.type === 'UserMailbox' ? [[t('toShared'), () => setModal({ k: 'convert', m })] as [string, () => void]] : []),
                                ...(m.type === 'SharedMailbox' ? [[t('toRegular'), () => toRegular(m)] as [string, () => void]] : []),
                                ...(!m.archive && !RES.includes(m.type) ? [[t('enableArchive'), () => archive(m)] as [string, () => void]] : [])
                              ]}
                            />
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                {!shown.length && <p className="hint" style={{ padding: 16 }}>{t('noMailboxes')}</p>}
              </div>
            </div>
          </>
        )}
      </div>

      {modal?.k === 'detail' && (
        <MailboxModal
          mailbox={modal.m}
          domains={domains}
          size={sizeOf(modal.m)}
          onClose={() => setModal(null)}
          onChanged={(n) => list.set((d) => ({ ...d, mailboxes: d.mailboxes.map((x) => (x.id === n.id ? n : x)) }))}
        />
      )}
      {modal?.k === 'convert' && (
        <ConvertModal
          mailbox={modal.m}
          size={sizeOf(modal.m)}
          onClose={() => setModal(null)}
          onDone={(n) => {
            list.set((d) => ({ ...d, mailboxes: d.mailboxes.map((x) => (x.id === n.id ? n : x)) }))
            setModal(null)
          }}
        />
      )}
      {modal?.k === 'new' && (
        <NewMailboxModal
          kind={scope === 'shared' ? 'shared' : 'room'}
          domains={domains}
          onClose={() => setModal(null)}
          onCreated={(n) => {
            list.set((d) => ({ ...d, mailboxes: [...d.mailboxes, n] }))
            setModal(null)
          }}
        />
      )}
    </>
  )
}

function ConvertModal({ mailbox, size, onClose, onDone }: { mailbox: Mailbox; size?: MailboxSize; onClose(): void; onDone(m: Mailbox): void }) {
  const { t, tenant, toast } = useApp()
  const risky = mailbox.archive || mailbox.litigationHold || (size?.bytes ?? 0) > 50 * 1024 ** 3
  const [removeLic, setRemoveLic] = useState(!risky)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      const r = await unwrap(window.api.exo.convertToShared(tenant!.tenantId, mailbox.id, mailbox.userId, removeLic))
      toast(t('convDone', mailbox.displayName) + (r.licensesRemoved ? ` · ${t('convLicDone', r.licensesRemoved)}` : ''))
      onDone(r.mailbox)
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <Modal
      title={t('convTitle', mailbox.displayName)}
      sub={t('convD')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy} onClick={submit}>
            {busy && <span className="spin" />}
            {t('toShared')}
          </button>
        </>
      }
    >
      <label className="opt">
        <input type="checkbox" className="chk" checked={removeLic} onChange={(e) => setRemoveLic(e.target.checked)} />
        <div>
          <b>{t('convLic')}</b>
          <small>{t('convLicWarn')}</small>
        </div>
      </label>
      {risky && <div className="banner" style={{ marginTop: 12 }}>{t('convLicWarn')}</div>}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

const aliasOf = (n: string) =>
  n
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '.')
    .replace(/^\.|\.$/g, '')

export function DomainPicker({ domains, value, onChange }: { domains: string[]; value: string; onChange(d: string): void }) {
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)} style={{ maxWidth: 260 }}>
      {domains.map((d) => (
        <option key={d} value={d}>
          @{d}
        </option>
      ))}
    </select>
  )
}

/** Orders accepted domains with the tenant's default domain first and the onmicrosoft.com ones last. */
export function useDomains(domains: string[]) {
  const { tenant } = useApp()
  return useMemo(() => {
    const def = tenant!.defaultDomain.toLowerCase()
    return [...domains].sort((a, b) => (a === def ? -1 : b === def ? 1 : Number(a.endsWith('.onmicrosoft.com')) - Number(b.endsWith('.onmicrosoft.com')) || a.localeCompare(b)))
  }, [domains, tenant])
}

function NewMailboxModal({ kind: initial, domains, onClose, onCreated }: { kind: NewMailbox['kind']; domains: string[]; onClose(): void; onCreated(m: Mailbox): void }) {
  const { t, tenant, toast } = useApp()
  const doms = useDomains(domains)
  const [kind, setKind] = useState(initial)
  const [name, setName] = useState('')
  const [alias, setAlias] = useState('')
  const [touched, setTouched] = useState(false)
  const [domain, setDomain] = useState(doms[0] || tenant!.defaultDomain)
  const [capacity, setCapacity] = useState('')
  const [members, setMembers] = useState<string[]>([])
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      const r = await unwrap(
        window.api.exo.newMailbox(tenant!.tenantId, {
          kind,
          name: name.trim(),
          alias,
          smtp: `${alias}@${domain}`,
          capacity: capacity ? Number(capacity) : undefined,
          members
        })
      )
      toast(t('nmCreated', r.mailbox.displayName) + (r.failed.length ? ` · ${t('nmPartial', r.failed.length)}` : ''))
      onCreated(r.mailbox)
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <Modal
      title={initial === 'shared' ? t('nmTitleShared') : t('nmTitleRes')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !name.trim() || !alias} onClick={submit}>
            {busy && <span className="spin" />}
            {t('createBtn')}
          </button>
        </>
      }
    >
      {initial !== 'shared' && (
        <div className="fld">
          <label>{t('typeL')}</label>
          <div className="seg">
            {(['room', 'equipment'] as const).map((k) => (
              <button key={k} className={kind === k ? 'on' : ''} onClick={() => setKind(k)}>
                {k === 'room' ? t('nmRoom') : t('nmEquip')}
              </button>
            ))}
          </div>
        </div>
      )}
      <div className="fld">
        <label htmlFor="nm-n">{t('nmName')}</label>
        <input
          id="nm-n"
          autoFocus
          value={name}
          onChange={(e) => {
            setName(e.target.value)
            if (!touched) setAlias(aliasOf(e.target.value))
          }}
        />
      </div>
      <div className="fld">
        <label htmlFor="nm-a">{t('nmAddress')}</label>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input id="nm-a" className="mono" value={alias} onChange={(e) => (setTouched(true), setAlias(aliasOf(e.target.value)))} />
          <DomainPicker domains={doms} value={domain} onChange={setDomain} />
        </div>
      </div>
      {kind !== 'shared' && (
        <div className="fld">
          <label htmlFor="nm-c">{t('capacityL')}</label>
          <input id="nm-c" type="number" min={0} value={capacity} onChange={(e) => setCapacity(e.target.value)} style={{ maxWidth: 120 }} />
        </div>
      )}
      {kind === 'shared' && (
        <div className="fld">
          <label>{t('nmMembers')}</label>
          <Chips items={members.map((m) => ({ key: m, label: m }))} onRemove={(k) => setMembers(members.filter((x) => x !== k))} empty={t('nobody')} />
          <AddRecipient placeholder={t('userPh')} filter={(r) => r.type === 'UserMailbox'} onAdd={(a) => setMembers([...new Set([...members, a])])} />
        </div>
      )}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

type Tab = 'overview' | 'perms' | 'fwd' | 'oof' | 'archive' | 'calendar' | 'access'

function MailboxModal({ mailbox, domains, size, onClose, onChanged }: {
  mailbox: Mailbox
  domains: string[]
  size?: MailboxSize
  onClose(): void
  onChanged(m: Mailbox): void
}) {
  const { t, lang } = useApp()
  const typeLabel = useTypeLabel()
  const [m, setM] = useState(mailbox)
  const res = RES.includes(m.type)
  const [tab, setTab] = useState<Tab>('overview')
  const changed = (n: Mailbox) => (setM(n), onChanged(n))
  const tabs: [Tab, string][] = [
    ['overview', t('tabOverview')],
    ['perms', t('tabPerms')],
    ...(res ? ([['calendar', t('tabCalendar')]] as [Tab, string][]) : []),
    ['fwd', t('tabFwd')],
    ['oof', t('tabOof')],
    ...(!res ? ([['archive', t('tabArchive')]] as [Tab, string][]) : []),
    ...(m.type === 'UserMailbox' ? ([['access', t('tabAccess')]] as [Tab, string][]) : [])
  ]
  return (
    <Modal title={m.displayName} sub={<span className="mono">{m.primarySmtp}</span>} onClose={onClose} wide>
      <Tabs<Tab> tabs={tabs} value={tab} onChange={setTab} />
      {tab === 'overview' && (
        <div className="creds">
          <Row k={t('typeL')} v={typeLabel(m.type)} />
          <Row k={t('primary')} v={<span className="mono">{m.primarySmtp}</span>} />
          <Row
            k={t('aliasesL')}
            v={
              <span className="mono">
                {m.addresses
                  .filter((a) => !a.startsWith('SMTP:'))
                  .map(stripSmtp)
                  .join(', ') || '—'}
              </span>
            }
          />
          <Row k={t('size')} v={size?.bytes != null ? `${fmtBytes(size.bytes, lang)}${m.quotaBytes ? ` / ${fmtBytes(m.quotaBytes, lang)}` : ''}` : '—'} />
          <Row k={t('lastAct')} v={fmtDate(size?.lastActivity ?? null, lang, '—')} />
          {m.capacity != null && <Row k={t('capacityL')} v={String(m.capacity)} />}
          <Row k={t('hiddenGal')} v={m.hidden ? t('yes') : t('no')} />
          <Row k={t('created')} v={fmtDate(m.created, lang, '—')} />
        </div>
      )}
      {tab === 'perms' && <PermsTab m={m} />}
      {tab === 'fwd' && <ForwardTab m={m} domains={domains} onChanged={changed} />}
      {tab === 'oof' && <AutoReplyTab m={m} />}
      {tab === 'archive' && <ArchiveTab m={m} size={size} onChanged={changed} />}
      {tab === 'calendar' && <CalendarTab m={m} onChanged={changed} />}
      {tab === 'access' && <AccessTab m={m} />}
    </Modal>
  )
}

function Row({ k, v }: { k: string; v: ReactNode }) {
  return (
    <div style={{ display: 'flex', gap: 10, alignItems: 'baseline' }}>
      <span className="hint" style={{ width: 170, flex: 'none' }}>
        {k}
      </span>
      <span style={{ flex: 1, wordBreak: 'break-word' }}>{v}</span>
    </div>
  )
}

function PermsTab({ m }: { m: Mailbox }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const perms = useLoad(() => window.api.exo.permissions(tid, m.id), [tid, m.id])
  const [busy, setBusy] = useState<MailRight | null>(null)
  const [autoMap, setAutoMap] = useState(true)
  async function change(right: MailRight, user: string, add: boolean) {
    setBusy(right)
    try {
      const r = await unwrap(add ? window.api.exo.grant(tid, m.id, user, right, autoMap) : window.api.exo.revoke(tid, m.id, user, right))
      perms.set(() => r)
      toast(add ? t('permAdded') : t('permRemoved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }
  if (!perms.data) return <ExoLoad loading={perms.loading} error={perms.error} onRetry={perms.reload} />
  const rows: [MailRight, keyof MailboxPerms, Key, Key][] = [
    ['FullAccess', 'fullAccess', 'fullAccess', 'fullAccessD'],
    ['SendAs', 'sendAs', 'sendAs', 'sendAsD'],
    ['SendOnBehalf', 'sendOnBehalf', 'sendOnBehalf', 'sendOnBehalfD']
  ]
  return (
    <div style={{ display: 'grid', gap: 18 }}>
      {rows.map(([right, field, label, desc]) => (
        <div key={right} className="fld">
          <label>
            {t(label)} <span className="hint">· {t(desc)}</span>
          </label>
          <Chips items={perms.data![field].map((u) => ({ key: u, label: u }))} onRemove={(u) => change(right, u, false)} empty={t('nobody')} />
          <AddRecipient placeholder={t('userPh')} busy={busy === right} onAdd={(u) => change(right, u, true)} />
          {right === 'FullAccess' && (
            <label className="hint" style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <input type="checkbox" className="chk" checked={autoMap} onChange={(e) => setAutoMap(e.target.checked)} />
              {t('autoMap')}
            </label>
          )}
        </div>
      ))}
    </div>
  )
}

function ForwardTab({ m, domains, onChanged }: { m: Mailbox; domains: string[]; onChanged(m: Mailbox): void }) {
  const { t, tenant, toast } = useApp()
  const [to, setTo] = useState(m.forwardingSmtp || '')
  const [keep, setKeep] = useState(m.deliverAndForward || !m.forwardingSmtp)
  const [busy, setBusy] = useState(false)
  const ext = isExternal(to, domains)
  async function save(clear = false) {
    setBusy(true)
    try {
      const n = await unwrap(
        window.api.exo.setMailbox(tenant!.tenantId, m.id, clear
          ? { ForwardingSmtpAddress: null, ForwardingAddress: null, DeliverToMailboxAndForward: false }
          : { ForwardingSmtpAddress: to.trim() ? `smtp:${to.trim()}` : null, DeliverToMailboxAndForward: keep })
      )
      onChanged(n)
      if (clear) setTo('')
      toast(t('fwdSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <div>
      {m.forwardingAddress && <div className="banner" style={{ marginBottom: 12 }}>{t('fwdInternal', m.forwardingAddress)}</div>}
      <div className="fld">
        <label htmlFor="fw-to">{t('fwdTo')}</label>
        <input id="fw-to" className="mono" placeholder={t('fwdNone')} value={to} onChange={(e) => setTo(e.target.value)} />
        {ext && <div className="banner bad" style={{ marginTop: 8 }}>{t('fwdExtWarn')}</div>}
      </div>
      <label className="opt">
        <input type="checkbox" className="chk" checked={keep} onChange={(e) => setKeep(e.target.checked)} />
        <div>
          <b>{t('fwdKeep')}</b>
        </div>
      </label>
      <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
        <button className="btn pri" disabled={busy} onClick={() => save()}>
          {busy && <span className="spin" />}
          {t('save')}
        </button>
        {(m.forwardingSmtp || m.forwardingAddress) && (
          <button className="btn" disabled={busy} onClick={() => save(true)}>
            {t('fwdClear')}
          </button>
        )}
      </div>
    </div>
  )
}

const toLocal = (iso: string | null) => {
  if (!iso) return ''
  const d = new Date(iso)
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16)
}

function AutoReplyTab({ m }: { m: Mailbox }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const cfg = useLoad(() => window.api.exo.autoReply(tid, m.id), [tid, m.id])
  const [a, setA] = useState<AutoReply | null>(null)
  const [busy, setBusy] = useState(false)
  const v = a || cfg.data
  if (!v) return <ExoLoad loading={cfg.loading} error={cfg.error} onRetry={cfg.reload} />
  const set = (p: Partial<AutoReply>) => setA({ ...v, ...p })
  async function save() {
    setBusy(true)
    try {
      const r = await unwrap(window.api.exo.setAutoReply(tid, m.id, v!))
      cfg.set(() => r)
      setA(null)
      toast(t('oofSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <div>
      <div className="fld">
        <label>{t('oofState')}</label>
        <div className="seg">
          {(['Disabled', 'Enabled', 'Scheduled'] as const).map((s) => (
            <button key={s} className={v.state === s ? 'on' : ''} onClick={() => set({ state: s })}>
              {t(`oof_${s}`)}
            </button>
          ))}
        </div>
      </div>
      {v.state === 'Scheduled' && (
        <div className="row2">
          <div className="fld">
            <label>{t('oofStart')}</label>
            <input type="datetime-local" value={toLocal(v.start)} onChange={(e) => set({ start: new Date(e.target.value).toISOString() })} />
          </div>
          <div className="fld">
            <label>{t('oofEnd')}</label>
            <input type="datetime-local" value={toLocal(v.end)} onChange={(e) => set({ end: new Date(e.target.value).toISOString() })} />
          </div>
        </div>
      )}
      <div className="fld">
        <label>{t('oofInternal')}</label>
        <textarea rows={4} value={v.internalMessage} onChange={(e) => set({ internalMessage: e.target.value })} />
      </div>
      <div className="fld">
        <label>{t('oofExternal')}</label>
        <select value={v.externalAudience} onChange={(e) => set({ externalAudience: e.target.value as AutoReply['externalAudience'] })}>
          {(['None', 'Known', 'All'] as const).map((x) => (
            <option key={x} value={x}>
              {t(`oofAud_${x}`)}
            </option>
          ))}
        </select>
        {v.externalAudience !== 'None' && (
          <textarea rows={4} style={{ marginTop: 8 }} value={v.externalMessage} onChange={(e) => set({ externalMessage: e.target.value })} />
        )}
      </div>
      <button className="btn pri" disabled={busy || !a} onClick={save}>
        {busy && <span className="spin" />}
        {t('save')}
      </button>
    </div>
  )
}

function ArchiveTab({ m, size, onChanged }: { m: Mailbox; size?: MailboxSize; onChanged(m: Mailbox): void }) {
  const { t, lang, tenant, toast } = useApp()
  const [busy, setBusy] = useState(false)
  return (
    <div className="creds">
      <Row
        k={t('archive')}
        v={
          m.archive ? (
            <span className="pill ok">{t('archOn')}</span>
          ) : (
            <span style={{ display: 'inline-flex', gap: 10, alignItems: 'center' }}>
              <span className="pill off">{t('archOff')}</span>
              <button
                className="btn sm"
                disabled={busy}
                onClick={async () => {
                  setBusy(true)
                  try {
                    onChanged(await unwrap(window.api.exo.enableArchive(tenant!.tenantId, m.id)))
                    toast(t('archEnabled', m.displayName))
                  } catch (e) {
                    toast((e as Error).message)
                  } finally {
                    setBusy(false)
                  }
                }}
              >
                {busy && <span className="spin" />}
                {t('enableArchive')}
              </button>
            </span>
          )
        }
      />
      {m.archive && size?.archiveBytes != null && <Row k={t('archSize')} v={fmtBytes(size.archiveBytes, lang)} />}
      <Row k={t('retention')} v={m.retentionPolicy || '—'} />
      <Row k={t('litHold')} v={m.litigationHold ? <span className="pill warn">{t('yes')}</span> : t('no')} />
      <Row k={t('holdsL')} v={m.holds.length ? <span className="mono">{m.holds.join(', ')}</span> : '—'} />
    </div>
  )
}

function CalendarTab({ m, onChanged }: { m: Mailbox; onChanged(m: Mailbox): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const cfg = useLoad(() => window.api.exo.calendar(tid, m.id), [tid, m.id])
  const [c, setC] = useState<CalendarSettings | null>(null)
  const [cap, setCap] = useState(m.capacity != null ? String(m.capacity) : '')
  const [busy, setBusy] = useState(false)
  const v = c || cfg.data
  if (!v) return <ExoLoad loading={cfg.loading} error={cfg.error} onRetry={cfg.reload} />
  const set = (p: Partial<CalendarSettings>) => setC({ ...v, ...p })
  const capChanged = cap !== (m.capacity != null ? String(m.capacity) : '')
  async function save() {
    setBusy(true)
    try {
      if (c) {
        const r = await unwrap(window.api.exo.setCalendar(tid, m.id, c))
        cfg.set(() => r)
        setC(null)
      }
      if (capChanged) onChanged(await unwrap(window.api.exo.setMailbox(tid, m.id, { ResourceCapacity: cap ? Number(cap) : null })))
      toast(t('calSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const checks: [keyof CalendarSettings, Key][] = [
    ['allowConflicts', 'allowConf'],
    ['allowRecurringMeetings', 'allowRec'],
    ['deleteSubject', 'delSubj'],
    ['addOrganizerToSubject', 'addOrg'],
    ['processExternal', 'procExt']
  ]
  return (
    <div>
      <div className="fld">
        <label>{t('autoProc')}</label>
        <select value={v.automateProcessing} onChange={(e) => set({ automateProcessing: e.target.value as CalendarSettings['automateProcessing'] })}>
          {(['AutoAccept', 'AutoUpdate', 'None'] as const).map((x) => (
            <option key={x} value={x}>
              {t(`ap_${x}`)}
            </option>
          ))}
        </select>
      </div>
      <div className="row2">
        <div className="fld">
          <label>{t('bookWindow')}</label>
          <input type="number" min={0} value={v.bookingWindowInDays} onChange={(e) => set({ bookingWindowInDays: Number(e.target.value) })} />
        </div>
        <div className="fld">
          <label>{t('maxDur')}</label>
          <input type="number" min={0} value={v.maximumDurationInMinutes} onChange={(e) => set({ maximumDurationInMinutes: Number(e.target.value) })} />
        </div>
      </div>
      <div className="fld">
        <label>{t('capacityL')}</label>
        <input type="number" min={0} value={cap} onChange={(e) => setCap(e.target.value)} style={{ maxWidth: 120 }} />
      </div>
      <div style={{ display: 'grid', gap: 6, marginBottom: 14 }}>
        {checks.map(([k, label]) => (
          <label key={k} className="opt">
            <input type="checkbox" className="chk" checked={!!v[k]} onChange={(e) => set({ [k]: e.target.checked } as Partial<CalendarSettings>)} />
            <div>
              <b>{t(label)}</b>
            </div>
          </label>
        ))}
      </div>
      <button className="btn pri" disabled={busy || (!c && !capChanged)} onClick={save}>
        {busy && <span className="spin" />}
        {t('save')}
      </button>
    </div>
  )
}

function AccessTab({ m }: { m: Mailbox }) {
  const { t, tenant } = useApp()
  const typeLabel = useTypeLabel()
  const [go, setGo] = useState(false)
  const scan = useLoad<UserAccess[]>(() => (go ? window.api.exo.permissionsForUser(tenant!.tenantId, m.primarySmtp) : null), [go])
  if (!go)
    return (
      <button className="btn pri" onClick={() => setGo(true)}>
        {t('accScan')}
      </button>
    )
  if (!scan.data) return <ExoLoad loading={scan.loading} error={scan.error} onRetry={scan.reload} text={t('accScanning')} />
  if (!scan.data.length) return <p className="hint">{t('accNone')}</p>
  return (
    <div className="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>{t('name')}</th>
            <th>{t('typeL')}</th>
            <th>{t('tabPerms')}</th>
          </tr>
        </thead>
        <tbody>
          {scan.data.map((r) => (
            <tr key={r.id}>
              <td>
                <div className="u-n">{r.displayName}</div>
                <div className="mono">{r.primarySmtp}</div>
              </td>
              <td>{typeLabel(r.type)}</td>
              <td>
                {r.rights.map((x) => (
                  <span key={x} className="lic-tag">
                    {t(x === 'FullAccess' ? 'fullAccess' : x === 'SendAs' ? 'sendAs' : 'sendOnBehalf')}
                  </span>
                ))}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
