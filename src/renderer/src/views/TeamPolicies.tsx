import { useMemo, useState } from 'react'
import type { CsUser, GroupPolicyAssignment, PolicyType, TeamsPolicies } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { Tabs } from '../components/exo'
import { TeamsLoad, useGroups } from '../components/teams'

const TYPES: PolicyType[] = ['TeamsMeetingPolicy', 'TeamsMessagingPolicy', 'TeamsAppSetupPolicy']

/** Picks a policy of one type; the empty value is the Global (org-wide default) policy. */
export function PolicySelect({ catalog, type, value, onChange }: { catalog: TeamsPolicies['catalog']; type: PolicyType; value: string; onChange(v: string): void }) {
  const { t } = useApp()
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="">{t('tpGlobal')}</option>
      {(catalog[type] || [])
        .filter((p) => !p.isGlobal)
        .map((p) => (
          <option key={p.name} value={p.name} title={p.description}>
            {p.name}
          </option>
        ))}
    </select>
  )
}

export function TeamPoliciesView() {
  const { t, tenant } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.teams.policies(tid), [tid])
  const [tab, setTab] = useState<'users' | 'groups'>('users')
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('teams')}
          </div>
          <h1>{t('policies')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={data.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <TeamsLoad loading={data.loading && !data.data} error={data.error} onRetry={data.reload} />
        {data.data && (
          <>
            <Tabs<typeof tab>
              tabs={[
                ['users', t('tpPerUser')],
                ['groups', t('tpPerGroup')]
              ]}
              value={tab}
              onChange={setTab}
            />
            {tab === 'users' ? (
              <UsersTab data={data.data} onUsers={(us) => data.set((d) => ({ ...d, users: d.users.map((u) => us.find((x) => x.id === u.id) || u) }))} />
            ) : (
              <GroupsTab data={data.data} onReload={data.reload} />
            )}
          </>
        )}
      </div>
    </>
  )
}

function UsersTab({ data, onUsers }: { data: TeamsPolicies; onUsers(u: CsUser[]): void }) {
  const { t } = useApp()
  const [q, setQ] = useState('')
  const [sel, setSel] = useState<Set<string>>(new Set())
  const [modal, setModal] = useState(false)
  const shown = data.users
    .filter((u) => (u.displayName + u.upn).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
  const toggle = (id: string) => setSel((s) => (s.has(id) ? (s.delete(id), new Set(s)) : new Set(s.add(id))))
  const allOn = shown.length > 0 && shown.every((u) => sel.has(u.id))

  return (
    <div className="panel">
      <div className="panel-h">
        <label className="search">
          {Icon.search}
          <input placeholder={t('searchUsers')} value={q} onChange={(e) => setQ(e.target.value)} />
        </label>
        <span className="hint">{t('tpDirectHint')}</span>
        <div className="r">
          <button className="btn pri sm" disabled={!sel.size} onClick={() => setModal(true)}>
            {t('tpAssign', sel.size)}
          </button>
        </div>
      </div>
      <div className="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th style={{ width: 30 }}>
                <input type="checkbox" className="chk" checked={allOn} onChange={() => setSel(allOn ? new Set() : new Set(shown.map((u) => u.id)))} />
              </th>
              <th>{t('name')}</th>
              {TYPES.map((k) => (
                <th key={k}>{t(`pt_${k}`)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shown.map((u) => (
              <tr key={u.id}>
                <td>
                  <input type="checkbox" className="chk" checked={sel.has(u.id)} onChange={() => toggle(u.id)} />
                </td>
                <td>
                  <div className="u-n">{u.displayName}</div>
                  <div className="mono">{u.upn}</div>
                </td>
                {TYPES.map((k) => (
                  <td key={k}>{u.policies[k] ? <span className="lic-tag">{u.policies[k]}</span> : <span className="hint">{t('tpGlobal')}</span>}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {modal && (
        <AssignModal
          data={data}
          users={data.users.filter((u) => sel.has(u.id))}
          onClose={() => setModal(false)}
          onDone={(us) => {
            onUsers(us)
            setSel(new Set())
          }}
        />
      )}
    </div>
  )
}

function AssignModal({ data, users, onClose, onDone }: { data: TeamsPolicies; users: CsUser[]; onClose(): void; onDone(u: CsUser[]): void }) {
  const { t, tenant, toast } = useApp()
  const [type, setType] = useState<PolicyType>(TYPES[0])
  const [policy, setPolicy] = useState('')
  const [busy, setBusy] = useState(false)
  const [failed, setFailed] = useState<{ user: string; error: string }[]>([])

  async function submit() {
    setBusy(true)
    try {
      const r = await unwrap(window.api.teams.grantUsers(tenant!.tenantId, type, policy || null, users.map((u) => u.upn)))
      onDone(r.users)
      if (r.failed.length) setFailed(r.failed)
      else {
        toast(t('tpAssigned', r.users.length))
        onClose()
      }
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('tpAssignT')}
      sub={users.length === 1 ? users[0].displayName : t('tpUsersN', users.length)}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy} onClick={submit}>
            {busy && <span className="spin" />}
            {t('apply')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label>{t('tpType')}</label>
        <select value={type} onChange={(e) => (setType(e.target.value as PolicyType), setPolicy(''))}>
          {TYPES.map((k) => (
            <option key={k} value={k}>
              {t(`pt_${k}`)}
            </option>
          ))}
        </select>
      </div>
      <div className="fld">
        <label>{t('tpPolicy')}</label>
        <PolicySelect catalog={data.catalog} type={type} value={policy} onChange={setPolicy} />
        <span className="hint">{data.catalog[type]?.find((p) => p.name === (policy || 'Global'))?.description}</span>
      </div>
      {failed.length > 0 && (
        <div className="err">
          {failed.map((f) => (
            <div key={f.user}>
              {f.user}: {f.error}
            </div>
          ))}
        </div>
      )}
    </Modal>
  )
}

function GroupsTab({ data, onReload }: { data: TeamsPolicies; onReload(): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const groups = useGroups(tid)
  const name = useMemo(() => new Map(groups.map((g) => [g.id, g.displayName])), [groups])
  const [type, setType] = useState<PolicyType>(TYPES[0])
  const [add, setAdd] = useState({ group: '', policy: '', rank: 1 })
  const [busy, setBusy] = useState(false)
  const rows = data.groups.filter((g) => g.type === type).sort((a, b) => a.rank - b.rank)
  const groupId = groups.find((g) => g.displayName === add.group || g.id === add.group)?.id

  async function step(fn: () => Promise<unknown>) {
    setBusy(true)
    try {
      await fn()
      toast(t('dlSaved'))
      onReload()
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  const assignment: GroupPolicyAssignment | null = groupId && add.policy ? { groupId, type, policy: add.policy, rank: add.rank } : null
  return (
    <div className="panel">
      <div className="panel-h">
        <div className="seg">
          {TYPES.map((k) => (
            <button key={k} className={type === k ? 'on' : ''} onClick={() => setType(k)}>
              {t(`pt_${k}`)}
            </button>
          ))}
        </div>
        <span className="hint">{t('tpRankHint')}</span>
      </div>
      <div className="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>{t('tpRank')}</th>
              <th>{t('groups')}</th>
              <th>{t('tpPolicy')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {rows.map((g) => (
              <tr key={g.groupId}>
                <td className="mono">{g.rank}</td>
                <td>{name.get(g.groupId) || <span className="mono">{g.groupId}</span>}</td>
                <td>
                  <span className="lic-tag">{g.policy || t('tpGlobal')}</span>
                </td>
                <td style={{ width: 90, textAlign: 'right' }}>
                  <button
                    className="btn sm"
                    disabled={busy}
                    onClick={() => confirm(t('tpRemoveConfirm', name.get(g.groupId) || g.groupId)) && step(() => unwrap(window.api.teams.removeGroupPolicy(tid, type, g.groupId)))}
                  >
                    {t('remove')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!rows.length && (
          <p className="hint" style={{ padding: 16 }}>
            {t('tpNoGroups')}
          </p>
        )}
      </div>
      <div className="fld" style={{ padding: 14, borderTop: '1px solid var(--line)', margin: 0 }}>
        <div className="pw" style={{ alignItems: 'center' }}>
          <input list="tp-groups" placeholder={t('tpGroupPh')} value={add.group} onChange={(e) => setAdd({ ...add, group: e.target.value })} />
          <datalist id="tp-groups">
            {groups.map((g) => (
              <option key={g.id} value={g.displayName} />
            ))}
          </datalist>
          <PolicySelect catalog={data.catalog} type={type} value={add.policy} onChange={(policy) => setAdd({ ...add, policy })} />
          <input type="number" min={1} style={{ flex: '0 0 80px' }} title={t('tpRank')} value={add.rank} onChange={(e) => setAdd({ ...add, rank: Math.max(1, Number(e.target.value) || 1) })} />
          <button className="btn sm" disabled={busy || !assignment} onClick={() => assignment && step(() => unwrap(window.api.teams.grantGroup(tid, assignment)))}>
            {busy && <span className="spin" />}
            {t('add')}
          </button>
        </div>
      </div>
    </div>
  )
}
