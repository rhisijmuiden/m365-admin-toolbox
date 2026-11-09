import { useMemo, useState } from 'react'
import type { ChannelType, NewTeam, Team, TeamActivity, TeamChannel, TeamMember, TeamTemplate, TemplateChannel } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, Modal, RowMenu } from '../components/ui'
import { Chips, Tabs } from '../components/exo'
import { AddUser, daysSince, TeamsLoad } from '../components/teams'

type Filter = 'all' | 'owners' | 'inactive' | 'archived'
export const INACTIVE_DAYS = 90
export const CHANNEL_TYPES: ChannelType[] = ['standard', 'private', 'shared']

export function TeamsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.teams.list(tid), [tid])
  const [consented, setConsented] = useState(0)
  const ids = useMemo(() => (list.data || []).map((x) => x.id), [list.data])
  const act = useLoad(() => (list.data ? window.api.teams.activity(tid, ids) : null), [tid, list.data, consented])
  const [q, setQ] = useState('')
  const [filter, setFilter] = useState<Filter>('all')
  const [modal, setModal] = useState<{ k: 'detail'; team: Team } | { k: 'new' } | null>(null)

  const activity = useMemo(() => new Map((act.data?.rows || []).map((r) => [r.teamId, r])), [act.data])
  const hasReport = !!act.data && !act.data.issue
  const teams = list.data || []
  const inactive = (x: Team) => {
    if (!hasReport) return false
    const a = activity.get(x.id)
    const d = daysSince(a?.lastActivity ?? x.created)
    return d != null && d >= INACTIVE_DAYS
  }
  const noOwner = teams.filter((x) => x.owners.length === 0)
  const oneOwner = teams.filter((x) => x.owners.length === 1)
  const archived = teams.filter((x) => x.archived)
  const stale = teams.filter((x) => !x.archived && inactive(x))
  const shown = teams
    .filter((x) =>
      filter === 'owners' ? x.owners.length < 2 : filter === 'inactive' ? !x.archived && inactive(x) : filter === 'archived' ? x.archived : true
    )
    .filter((x) => (x.displayName + (x.mail || '') + (x.description || '')).toLowerCase().includes(q.toLowerCase()))

  const replace = (n: Team) => list.set((d) => d.map((x) => (x.id === n.id ? n : x)))

  async function setArchived(x: Team, value: boolean) {
    if (value && !confirm(t('tmArchiveConfirm', x.displayName))) return
    try {
      replace(await unwrap(window.api.teams.archive(tid, x.id, value)))
      toast(value ? t('tmArchived', x.displayName) : t('tmRestored', x.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('teams')}
          </div>
          <h1>{t('teamsList')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('tmSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn pri" disabled={!list.data} onClick={() => setModal({ k: 'new' })}>
            {Icon.plus}
            {t('tmNew')}
          </button>
        </div>
      </header>
      <div className="content">
        <TeamsLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('tmTotal')}</div>
                <div className="v">{teams.length}</div>
                <div className="s">{t('tmArchivedN', archived.length)}</div>
              </div>
              <div className={`tile ${noOwner.length ? 'bad' : ''}`}>
                <div className="k">{t('tmNoOwner')}</div>
                <div className="v">{noOwner.length}</div>
                <div className="s">{noOwner.map((x) => x.displayName).join(', ') || t('ok')}</div>
              </div>
              <div className={`tile ${oneOwner.length ? 'warn' : ''}`}>
                <div className="k">{t('tmOneOwner')}</div>
                <div className="v">{oneOwner.length}</div>
                <div className="s">{oneOwner.map((x) => x.displayName).join(', ') || t('ok')}</div>
              </div>
              <div className={`tile ${stale.length ? 'warn' : ''}`}>
                <div className="k">{t('tmInactive', INACTIVE_DAYS)}</div>
                <div className="v">{act.loading ? '…' : hasReport ? stale.length : '—'}</div>
                <div className="s">{hasReport ? stale.map((x) => x.displayName).join(', ') || t('ok') : t('tmNoReport')}</div>
              </div>
            </div>

            {act.data?.issue === 'consent' && (
              <div className="banner">
                <span>{t('tmReportConsent')}</span>
                <button
                  className="btn sm"
                  onClick={async () => {
                    try {
                      await unwrap(window.api.teams.consentReports(tid))
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
            {act.data?.issue === 'concealed' && (
              <div className="banner">
                <span>{t('tmReportConcealed')}</span>
              </div>
            )}

            <div className="panel">
              <div className="panel-h">
                <div className="seg">
                  {(['all', 'owners', 'inactive', 'archived'] as Filter[]).map((f) => (
                    <button key={f} className={filter === f ? 'on' : ''} onClick={() => setFilter(f)}>
                      {t(`tmF_${f}`)}
                    </button>
                  ))}
                </div>
                {hasReport && <span className="hint">{t('tmReportDate', fmtDate(act.data!.reportDate, lang, '—'))}</span>}
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('name')}</th>
                      <th>{t('tmVisibility')}</th>
                      <th>{t('tmOwners')}</th>
                      <th>{t('tmLastActivity')}</th>
                      <th>{t('status')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((x) => (
                      <TeamRow
                        key={x.id}
                        team={x}
                        activity={activity.get(x.id)}
                        hasReport={hasReport}
                        inactive={inactive(x)}
                        onOpen={() => setModal({ k: 'detail', team: x })}
                        onArchive={(v) => setArchived(x, v)}
                      />
                    ))}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('tmNone')}
                  </p>
                )}
              </div>
            </div>
          </>
        )}
      </div>
      {modal?.k === 'new' && (
        <NewTeamModal
          onClose={() => setModal(null)}
          onCreated={(x) => {
            list.set((d) => [...d, x].sort((a, b) => a.displayName.localeCompare(b.displayName)))
            toast(t('tmCreated', x.displayName))
          }}
        />
      )}
      {modal?.k === 'detail' && <TeamModal team={modal.team} onClose={() => setModal(null)} onChanged={replace} />}
    </>
  )
}

function TeamRow({ team: x, activity, hasReport, inactive, onOpen, onArchive }: {
  team: Team
  activity?: TeamActivity
  hasReport: boolean
  inactive: boolean
  onOpen(): void
  onArchive(v: boolean): void
}) {
  const { t, lang } = useApp()
  return (
    <tr>
      <td>
        <button className="linkish" onClick={onOpen}>
          <div className="u-n">{x.displayName}</div>
          <div className="mono">{x.mail || '—'}</div>
        </button>
      </td>
      <td>
        <span className="lic-tag">{t(x.visibility === 'public' ? 'tmPublic' : 'tmPrivate')}</span>
      </td>
      <td>
        {x.owners.length === 0 ? (
          <span className="pill bad">{t('tmNoOwnerPill')}</span>
        ) : (
          <>
            {x.owners.length === 1 && <span className="pill warn" style={{ marginRight: 6 }}>{t('tmOneOwnerPill')}</span>}
            <span>{x.owners.map((o) => o.displayName).join(', ')}</span>
          </>
        )}
      </td>
      <td className="mono">
        {hasReport ? fmtDate(activity?.lastActivity ?? null, lang, t('never')) : '—'}
        {activity?.activeUsers != null && <div className="hint">{t('tmActiveUsers', activity.activeUsers)}</div>}
      </td>
      <td>
        {x.archived ? <span className="pill off">{t('tmArchivedPill')}</span> : inactive ? <span className="pill warn">{t('tmInactivePill')}</span> : <span className="pill ok">{t('active')}</span>}
      </td>
      <td style={{ width: 40 }}>
        <RowMenu
          items={[
            [t('details'), onOpen],
            x.archived ? [t('tmRestore'), () => onArchive(false)] : [t('tmArchive'), () => onArchive(true)]
          ]}
        />
      </td>
    </tr>
  )
}

// ---------- new team ----------

const blankChannel = (): TemplateChannel => ({ name: '', type: 'standard' })

export function ChannelEditor({ channels, onChange }: { channels: TemplateChannel[]; onChange(c: TemplateChannel[]): void }) {
  const { t } = useApp()
  const set = (i: number, c: Partial<TemplateChannel>) => onChange(channels.map((x, n) => (n === i ? { ...x, ...c } : x)))
  return (
    <div style={{ display: 'grid', gap: 6 }}>
      <div className="pw" style={{ alignItems: 'center', opacity: 0.7 }}>
        <input value={t('tmGeneral')} disabled />
        <select disabled value="standard">
          <option value="standard">{t('ch_standard')}</option>
        </select>
        <span style={{ width: 28 }} />
      </div>
      {channels.map((c, i) => (
        <div key={i} className="pw" style={{ alignItems: 'center' }}>
          <input placeholder={t('tmChannelName')} value={c.name} onChange={(e) => set(i, { name: e.target.value })} />
          <select value={c.type} onChange={(e) => set(i, { type: e.target.value as ChannelType })}>
            {CHANNEL_TYPES.map((k) => (
              <option key={k} value={k}>
                {t(`ch_${k}`)}
              </option>
            ))}
          </select>
          <button className="x" aria-label={t('remove')} onClick={() => onChange(channels.filter((_, n) => n !== i))}>
            ×
          </button>
        </div>
      ))}
      <button className="btn sm" style={{ justifySelf: 'start' }} onClick={() => onChange([...channels, blankChannel()])}>
        {Icon.plus}
        {t('tmAddChannel')}
      </button>
    </div>
  )
}

function NewTeamModal({ onClose, onCreated }: { onClose(): void; onCreated(t: Team): void }) {
  const { t, tenant, toast } = useApp()
  const templates = useLoad(() => window.api.teams.templates(), [])
  const [tpl, setTpl] = useState('')
  const [n, setN] = useState<NewTeam>({ displayName: '', description: '', visibility: 'private', owners: [], members: [], channels: [] })
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const [failed, setFailed] = useState<{ what: string; error: string }[] | null>(null)

  const pick = (id: string) => {
    setTpl(id)
    const x = templates.data?.find((y) => y.id === id)
    setN({ ...n, channels: x ? x.channels.map((c) => ({ ...c })) : [], visibility: x?.visibility || n.visibility, description: n.description || x?.description || '' })
  }
  const channels = n.channels.filter((c) => c.name.trim())
  const names = channels.map((c) => c.name.trim().toLowerCase())
  const dupes = names.length !== new Set(names).size

  async function submit() {
    setBusy(true)
    setErr(null)
    try {
      const r = await unwrap(window.api.teams.create(tenant!.tenantId, { ...n, displayName: n.displayName.trim(), channels }))
      onCreated(r.team)
      if (r.failed.length) setFailed(r.failed)
      else onClose()
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function saveTemplate() {
    const name = n.displayName.trim() || t('tmNewTemplate')
    const list = templates.data || []
    const item: TeamTemplate = { id: `tpl-${Date.now().toString(36)}`, name, description: n.description, visibility: n.visibility, channels }
    try {
      const saved = await unwrap(window.api.teams.saveTemplates([...list, item]))
      templates.set(() => saved)
      setTpl(item.id)
      toast(t('tplSaved', name))
    } catch (e) {
      toast((e as Error).message)
      templates.reload()
    }
  }

  if (failed)
    return (
      <Modal title={t('tmPartial')} onClose={onClose} footer={<button className="btn pri" onClick={onClose}>{t('close')}</button>}>
        <p>{t('tmPartialD')}</p>
        <ul>
          {failed.map((f, i) => (
            <li key={i}>
              <b>{f.what}</b>: <span className="hint">{f.error}</span>
            </li>
          ))}
        </ul>
      </Modal>
    )

  return (
    <Modal
      title={t('tmNew')}
      sub={t('tmNewSub')}
      onClose={onClose}
      wide
      footer={
        <>
          <button className="btn" style={{ marginRight: 'auto' }} disabled={busy || !channels.length} onClick={saveTemplate}>
            {t('tmSaveTpl')}
          </button>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || !n.displayName.trim() || !n.owners.length || dupes} onClick={submit}>
            {busy && <span className="spin" />}
            {busy ? t('tmCreating') : t('createBtn')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="tm-tpl">{t('tmTemplate')}</label>
        <select id="tm-tpl" value={tpl} onChange={(e) => pick(e.target.value)}>
          <option value="">{t('tmEmpty')}</option>
          {(templates.data || []).map((x) => (
            <option key={x.id} value={x.id}>
              {x.name}
            </option>
          ))}
        </select>
      </div>
      <div className="fld">
        <label htmlFor="tm-n">{t('name')}</label>
        <input id="tm-n" autoFocus value={n.displayName} onChange={(e) => setN({ ...n, displayName: e.target.value })} />
      </div>
      <div className="fld">
        <label htmlFor="tm-d">{t('ngDesc')}</label>
        <input id="tm-d" value={n.description} onChange={(e) => setN({ ...n, description: e.target.value })} />
      </div>
      <div className="fld">
        <label>{t('tmVisibility')}</label>
        <div className="seg">
          {(['private', 'public'] as const).map((v) => (
            <button key={v} className={n.visibility === v ? 'on' : ''} onClick={() => setN({ ...n, visibility: v })}>
              {t(v === 'public' ? 'tmPublic' : 'tmPrivate')}
            </button>
          ))}
        </div>
      </div>
      <div className="fld">
        <label>
          {t('tmOwners')} <span className="hint">· {t('tmOwnersHint')}</span>
        </label>
        <Chips items={n.owners.map((m) => ({ key: m, label: m }))} onRemove={(k) => setN({ ...n, owners: n.owners.filter((x) => x !== k) })} empty={t('nobody')} />
        <AddUser placeholder={t('userPh')} onAdd={(a) => setN({ ...n, owners: [...new Set([...n.owners, a])], members: n.members.filter((x) => x !== a) })} />
      </div>
      <div className="fld">
        <label>{t('tmMembers')}</label>
        <Chips items={n.members.map((m) => ({ key: m, label: m }))} onRemove={(k) => setN({ ...n, members: n.members.filter((x) => x !== k) })} empty={t('nobody')} />
        <AddUser placeholder={t('userPh')} onAdd={(a) => {
            if (!n.owners.includes(a)) setN({ ...n, members: [...new Set([...n.members, a])] })
          }} />
      </div>
      <div className="fld">
        <label>{t('tmChannels')}</label>
        <ChannelEditor channels={n.channels} onChange={(channels) => setN({ ...n, channels })} />
        <span className="hint">{t('tmChannelsHint')}</span>
        {dupes && <span className="err">{t('tmDupes')}</span>}
      </div>
      {busy && <p className="hint">{t('tmCreatingHint')}</p>}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

// ---------- team detail ----------

type Tab = 'members' | 'channels' | 'settings'

function TeamModal({ team, onClose, onChanged }: { team: Team; onClose(): void; onChanged(t: Team): void }) {
  const { t } = useApp()
  const [tab, setTab] = useState<Tab>('members')
  const [cur, setCur] = useState(team)
  const changed = (x: Team) => (setCur(x), onChanged(x))
  return (
    <Modal title={cur.displayName} sub={<span className="mono">{cur.mail}</span>} onClose={onClose} wide>
      {cur.archived && (
        <div className="banner">
          <span>{t('tmArchivedNote')}</span>
        </div>
      )}
      <Tabs<typeof tab>
        tabs={[
          ['members', t('tmMembers')],
          ['channels', t('tmChannels')],
          ['settings', t('tmSettings')]
        ]}
        value={tab}
        onChange={setTab}
      />
      {tab === 'members' && <MembersTab team={cur} onChanged={changed} />}
      {tab === 'channels' && <ChannelsTab team={cur} />}
      {tab === 'settings' && <SettingsTab team={cur} onChanged={changed} />}
    </Modal>
  )
}

const isOwner = (m: TeamMember) => m.roles.includes('owner')
const isGuest = (m: TeamMember) => m.roles.includes('guest')

/** Members of a team or of a private/shared channel, with role switches and removal. */
function MemberList({ load, add, setRole, remove, onOwners }: {
  load: () => ReturnType<typeof window.api.teams.members>
  add(user: string, owner: boolean): ReturnType<typeof window.api.teams.members>
  setRole(id: string, owner: boolean): ReturnType<typeof window.api.teams.members>
  remove(id: string): ReturnType<typeof window.api.teams.members>
  onOwners?(n: number): void
}) {
  const { t, toast } = useApp()
  const list = useLoad(load, [])
  const [busy, setBusy] = useState<string | null>(null)
  const [asOwner, setAsOwner] = useState(false)

  async function step(key: string, fn: () => ReturnType<typeof window.api.teams.members>) {
    setBusy(key)
    try {
      const r = await unwrap(fn())
      list.set(() => r)
      onOwners?.(r.filter(isOwner).length)
      toast(t('dlSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  if (!list.data) return <TeamsLoad loading={list.loading} error={list.error} onRetry={list.reload} />
  const owners = list.data.filter(isOwner)
  return (
    <>
      {owners.length < 2 && (
        <div className={`banner ${owners.length ? '' : 'bad'}`}>
          <span>{owners.length ? t('tmOneOwnerHint') : t('tmNoOwnerHint')}</span>
        </div>
      )}
      <div className="tbl-wrap" style={{ maxHeight: 340, overflow: 'auto', marginBottom: 12 }}>
        <table>
          <thead>
            <tr>
              <th>{t('name')}</th>
              <th>{t('tmRole')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {list.data.map((m) => (
              <tr key={m.id}>
                <td>
                  <div className="u-n">{m.displayName}</div>
                  <div className="mono">{m.email || '—'}</div>
                </td>
                <td>
                  {isGuest(m) ? (
                    <span className="pill warn">{t('tmGuest')}</span>
                  ) : (
                    <select
                      value={isOwner(m) ? 'owner' : 'member'}
                      disabled={!!busy || (isOwner(m) && owners.length === 1)}
                      onChange={(e) => step(m.id, () => setRole(m.id, e.target.value === 'owner'))}
                    >
                      <option value="owner">{t('tmOwner')}</option>
                      <option value="member">{t('tmMember')}</option>
                    </select>
                  )}
                </td>
                <td style={{ width: 90, textAlign: 'right' }}>
                  <button
                    className="btn sm"
                    disabled={!!busy || (isOwner(m) && owners.length === 1)}
                    onClick={() => confirm(t('tmRemoveConfirm', m.displayName)) && step(m.id, () => remove(m.id))}
                  >
                    {busy === m.id && <span className="spin" />}
                    {t('remove')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="fld">
        <label>{t('tmAddPerson')}</label>
        <AddUser placeholder={t('userPh')} busy={busy === 'add'} onAdd={(u) => step('add', () => add(u, asOwner))} />
        <label style={{ display: 'flex', gap: 8, alignItems: 'center', fontWeight: 400 }}>
          <input type="checkbox" className="chk" checked={asOwner} onChange={(e) => setAsOwner(e.target.checked)} />
          {t('tmAsOwner')}
        </label>
      </div>
    </>
  )
}

function MembersTab({ team, onChanged }: { team: Team; onChanged(t: Team): void }) {
  const { tenant } = useApp()
  const tid = tenant!.tenantId
  const api = window.api.teams
  // Owner count changes are reflected in the overview by re-reading the team.
  const refresh = () => unwrap(api.get(tid, team.id)).then(onChanged, () => undefined)
  return (
    <MemberList
      load={() => api.members(tid, team.id)}
      add={(u, o) => api.addMember(tid, team.id, u, o)}
      setRole={(id, o) => api.setMemberRole(tid, team.id, id, o)}
      remove={(id) => api.removeMember(tid, team.id, id)}
      onOwners={(n) => n !== team.owners.length && refresh()}
    />
  )
}

function ChannelsTab({ team }: { team: Team }) {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.teams.channels(tid, team.id), [tid, team.id])
  const [open, setOpen] = useState<TeamChannel | null>(null)
  const [nc, setNc] = useState<TemplateChannel>(blankChannel())
  const [busy, setBusy] = useState(false)

  async function add() {
    setBusy(true)
    try {
      const r = await unwrap(window.api.teams.addChannel(tid, team.id, { ...nc, name: nc.name.trim() }, team.owners.map((o) => o.upn || o.id)))
      list.set((d) => [...d, r.channel])
      setNc(blankChannel())
      toast(r.failed.length ? r.failed.map((f) => `${f.what}: ${f.error}`).join('; ') : t('tmChannelAdded', r.channel.displayName))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function remove(c: TeamChannel) {
    if (!confirm(t('tmChannelDelConfirm', c.displayName))) return
    try {
      await unwrap(window.api.teams.removeChannel(tid, team.id, c.id))
      list.set((d) => d.filter((x) => x.id !== c.id))
      if (open?.id === c.id) setOpen(null)
    } catch (e) {
      toast((e as Error).message)
    }
  }

  if (!list.data) return <TeamsLoad loading={list.loading} error={list.error} onRetry={list.reload} />
  if (open)
    return (
      <>
        <button className="btn sm" style={{ marginBottom: 12 }} onClick={() => setOpen(null)}>
          ← {t('tmChannels')}
        </button>
        <h4 style={{ margin: '0 0 10px' }}>
          {open.displayName} <span className="lic-tag">{t(`ch_${open.type}`)}</span>
        </h4>
        <MemberList
          load={() => window.api.teams.channelMembers(tid, team.id, open.id)}
          add={(u, o) => window.api.teams.addChannelMember(tid, team.id, open.id, u, o)}
          setRole={(id, o) => window.api.teams.setChannelMemberRole(tid, team.id, open.id, id, o)}
          remove={(id) => window.api.teams.removeChannelMember(tid, team.id, open.id, id)}
        />
      </>
    )
  return (
    <>
      <div className="tbl-wrap" style={{ maxHeight: 320, overflow: 'auto', marginBottom: 12 }}>
        <table>
          <thead>
            <tr>
              <th>{t('name')}</th>
              <th>{t('typeL')}</th>
              <th>{t('created')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {list.data.map((c) => (
              <tr key={c.id}>
                <td>
                  {c.type === 'standard' ? (
                    <div className="u-n">{c.displayName}</div>
                  ) : (
                    <button className="linkish" onClick={() => setOpen(c)}>
                      <div className="u-n">{c.displayName}</div>
                    </button>
                  )}
                  {c.description && <div className="hint">{c.description}</div>}
                </td>
                <td>
                  <span className={`lic-tag`}>{t(`ch_${c.type}`)}</span>
                </td>
                <td className="mono">{fmtDate(c.created, lang, '—')}</td>
                <td style={{ width: 150, textAlign: 'right' }}>
                  {c.type !== 'standard' && (
                    <button className="btn sm" onClick={() => setOpen(c)}>
                      {t('tmMembers')}
                    </button>
                  )}{' '}
                  {c.displayName !== 'General' && (
                    <button className="btn sm" onClick={() => remove(c)}>
                      {t('remove')}
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="fld">
        <label>{t('tmAddChannel')}</label>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input placeholder={t('tmChannelName')} value={nc.name} onChange={(e) => setNc({ ...nc, name: e.target.value })} />
          <select value={nc.type} onChange={(e) => setNc({ ...nc, type: e.target.value as ChannelType })}>
            {CHANNEL_TYPES.map((k) => (
              <option key={k} value={k}>
                {t(`ch_${k}`)}
              </option>
            ))}
          </select>
          <button className="btn sm" disabled={busy || !nc.name.trim()} onClick={add}>
            {busy && <span className="spin" />}
            {t('add')}
          </button>
        </div>
        {nc.type !== 'standard' && <span className="hint">{t('tmChannelOwnersHint')}</span>}
      </div>
    </>
  )
}

function SettingsTab({ team, onChanged }: { team: Team; onChanged(t: Team): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const [f, setF] = useState({ displayName: team.displayName, description: team.description || '', visibility: (team.visibility === 'public' ? 'public' : 'private') as 'private' | 'public' })
  const [ro, setRo] = useState(false)
  const [busy, setBusy] = useState<string | null>(null)

  async function step(key: string, fn: () => ReturnType<typeof window.api.teams.get>, msg: string) {
    setBusy(key)
    try {
      onChanged(await unwrap(fn()))
      toast(msg)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  const dirty = f.displayName.trim() !== team.displayName || f.description !== (team.description || '') || f.visibility !== team.visibility
  return (
    <>
      <div className="fld">
        <label htmlFor="ts-n">{t('name')}</label>
        <input id="ts-n" value={f.displayName} onChange={(e) => setF({ ...f, displayName: e.target.value })} />
      </div>
      <div className="fld">
        <label htmlFor="ts-d">{t('ngDesc')}</label>
        <input id="ts-d" value={f.description} onChange={(e) => setF({ ...f, description: e.target.value })} />
      </div>
      <div className="fld">
        <label>{t('tmVisibility')}</label>
        <div className="seg">
          {(['private', 'public'] as const).map((v) => (
            <button key={v} className={f.visibility === v ? 'on' : ''} onClick={() => setF({ ...f, visibility: v })}>
              {t(v === 'public' ? 'tmPublic' : 'tmPrivate')}
            </button>
          ))}
        </div>
      </div>
      <button
        className="btn pri"
        disabled={!!busy || !dirty || !f.displayName.trim()}
        onClick={() => step('save', () => window.api.teams.update(tid, team.id, { ...f, displayName: f.displayName.trim() }), t('dlSaved'))}
      >
        {busy === 'save' && <span className="spin" />}
        {t('save')}
      </button>
      <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '18px 0' }} />
      <div className="fld">
        <label>{t('tmArchive')}</label>
        <span className="hint">{t('tmArchiveD')}</span>
        {team.archived ? (
          <button className="btn" style={{ justifySelf: 'start' }} disabled={!!busy} onClick={() => step('arch', () => window.api.teams.archive(tid, team.id, false), t('tmRestored', team.displayName))}>
            {busy === 'arch' && <span className="spin" />}
            {t('tmRestore')}
          </button>
        ) : (
          <>
            <label style={{ display: 'flex', gap: 8, alignItems: 'center', fontWeight: 400 }}>
              <input type="checkbox" className="chk" checked={ro} onChange={(e) => setRo(e.target.checked)} />
              {t('tmSiteReadOnly')}
            </label>
            <button
              className="btn"
              style={{ justifySelf: 'start' }}
              disabled={!!busy}
              onClick={() => confirm(t('tmArchiveConfirm', team.displayName)) && step('arch', () => window.api.teams.archive(tid, team.id, true, ro), t('tmArchived', team.displayName))}
            >
              {busy === 'arch' && <span className="spin" />}
              {t('tmArchive')}
            </button>
          </>
        )}
      </div>
    </>
  )
}
