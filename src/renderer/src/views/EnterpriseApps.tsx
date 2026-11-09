import { useEffect, useState } from 'react'
import type { AppAssignment, EnterpriseApp } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, Modal } from '../components/ui'
import { EntraLoad, OwnersPanel, pending, PrincipalInput, Tabs } from '../components/entra'

const DEFAULT_ROLE = '00000000-0000-0000-0000-000000000000'
type Filter = 'all' | 'own' | 'thirdParty' | 'microsoft' | 'disabled'

export function EnterpriseAppsView() {
  const { t, tenant } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.enterpriseApps(tid), [tid])
  const [q, setQ] = useState('')
  const [f, setF] = useState<Filter>('thirdParty')
  const [open, setOpen] = useState<EnterpriseApp | null>(null)

  useEffect(() => {
    if (pending.spId && list.data) {
      const a = list.data.find((x) => x.id === pending.spId)
      pending.spId = null
      if (a) setOpen(a)
    }
  }, [list.data])

  const rows = list.data || []
  const match = (a: EnterpriseApp) =>
    f === 'all' ? true : f === 'own' ? a.ownTenant : f === 'microsoft' ? a.microsoft : f === 'disabled' ? !a.enabled : !a.microsoft && !a.ownTenant
  const shown = rows.filter(match).filter((a) => (a.displayName + a.appId + (a.publisher || '')).toLowerCase().includes(q.toLowerCase()))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('apps')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('eaSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('apps')}</div>
                <div className="v">{rows.length}</div>
                <div className="s">{t('eaMicrosoftN', rows.filter((a) => a.microsoft).length)}</div>
              </div>
              <div className="tile">
                <div className="k">{t('eaThirdParty')}</div>
                <div className="v">{rows.filter((a) => !a.microsoft && !a.ownTenant).length}</div>
                <div className="s">{t('eaThirdPartyD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('eaAssignReq')}</div>
                <div className="v">{rows.filter((a) => a.assignmentRequired).length}</div>
                <div className="s">{t('eaAssignReqD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('eaDisabled')}</div>
                <div className="v">{rows.filter((a) => !a.enabled).length}</div>
                <div className="s">{t('eaDisabledD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <Tabs
                  tabs={(['thirdParty', 'own', 'microsoft', 'disabled', 'all'] as Filter[]).map((k) => [k, t(`eaF_${k}`)] as [Filter, string])}
                  value={f}
                  onChange={setF}
                />
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('arName')}</th>
                      <th>{t('eaPublisher')}</th>
                      <th>{t('eaSignIn')}</th>
                      <th>{t('eaAssignReq')}</th>
                      <th>{t('arAppId')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((a) => (
                      <tr key={a.id} className={a.enabled ? '' : 'muted'}>
                        <td>
                          <button className="linkish" onClick={() => setOpen(a)}>
                            <div className="u-n">{a.displayName}</div>
                          </button>
                        </td>
                        <td>{a.publisher || '—'}</td>
                        <td>{a.enabled ? <span className="pill ok">{t('eaOn')}</span> : <span className="pill bad">{t('eaOff')}</span>}</td>
                        <td>{a.assignmentRequired ? t('yes') : t('no')}</td>
                        <td className="mono">{a.appId}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('arNone')}
                  </p>
                )}
              </div>
            </div>
          </>
        )}
      </div>
      {open && <AppModal app={open} onClose={() => setOpen(null)} onChange={(a) => (setOpen(a), list.set((d) => d.map((x) => (x.id === a.id ? a : x))))} />}
    </>
  )
}

type Tab = 'props' | 'users' | 'owners' | 'perms'

function AppModal({ app, onClose, onChange }: { app: EnterpriseApp; onClose(): void; onChange(a: EnterpriseApp): void }) {
  const { t } = useApp()
  const [tab, setTab] = useState<Tab>('props')
  return (
    <Modal wide title={app.displayName} sub={<span className="mono">{app.appId}</span>} onClose={onClose}>
      <Tabs
        tabs={[
          ['props', t('eaProps')],
          ['users', t('eaUsers')],
          ['owners', t('arOwners')],
          ['perms', t('arPermissions')]
        ]}
        value={tab}
        onChange={setTab}
      />
      {tab === 'props' && <Props app={app} onChange={onChange} />}
      {tab === 'users' && <Users app={app} />}
      {tab === 'owners' && <OwnersPanel kind="servicePrincipals" id={app.id} />}
      {tab === 'perms' && <Granted app={app} />}
    </Modal>
  )
}

function Props({ app, onChange }: { app: EnterpriseApp; onChange(a: EnterpriseApp): void }) {
  const { t, tenant, toast } = useApp()
  const [busy, setBusy] = useState<string | null>(null)
  const toggle = async (k: 'enabled' | 'assignmentRequired' | 'hidden', v: boolean) => {
    if (k === 'enabled' && !v && !confirm(t('eaDisableConfirm', app.displayName))) return
    setBusy(k)
    try {
      onChange(await unwrap(window.api.entra.updateEnterpriseApp(tenant!.tenantId, app.id, { [k]: v })))
      toast(t('saved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }
  const row = (k: 'enabled' | 'assignmentRequired' | 'hidden', title: string, d: string, value: boolean) => (
    <label className={`opt ${value ? 'on' : ''}`}>
      <input type="checkbox" className="chk" checked={value} disabled={!!busy} onChange={(e) => toggle(k, e.target.checked)} />
      <span>
        {title} {busy === k && <span className="spin" />}
        <small>{d}</small>
      </span>
    </label>
  )
  return (
    <>
      {row('enabled', t('eaEnabledT'), t('eaEnabledD'), app.enabled)}
      {row('assignmentRequired', t('eaReqT'), t('eaReqD'), app.assignmentRequired)}
      {row('hidden', t('eaHiddenT'), t('eaHiddenD'), app.hidden)}
      <p className="hint">
        {t('eaPublisher')}: {app.publisher || '—'} · {t('arType')}: {app.type}
        {app.homepage && <> · {app.homepage}</>}
      </p>
    </>
  )
}

function Users({ app }: { app: EnterpriseApp }) {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.assignments(tid, app.id), [tid, app.id])
  const roles = app.appRoles.filter((r) => r.allowedMemberTypes.includes('User'))
  const [role, setRole] = useState(roles[0]?.id || DEFAULT_ROLE)
  const [busy, setBusy] = useState(false)
  const roleName = (id: string) => (id === DEFAULT_ROLE ? t('eaDefaultAccess') : app.appRoles.find((r) => r.id === id)?.displayName || id)
  const run = async (fn: () => Promise<AppAssignment[]>, msg: string) => {
    setBusy(true)
    try {
      const r = await fn()
      list.set(() => r)
      toast(msg)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <>
      <p className="hint">{app.assignmentRequired ? t('eaUsersReq') : t('eaUsersOpen')}</p>
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        {roles.length > 0 && (
          <select value={role} onChange={(e) => setRole(e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--line)', borderRadius: 7, background: 'var(--surface)' }}>
            {roles.map((r) => (
              <option key={r.id} value={r.id}>
                {r.displayName}
              </option>
            ))}
          </select>
        )}
        <div style={{ flex: 1, minWidth: 260 }}>
          <PrincipalInput groups busy={busy} label={t('eaAssign')} onPick={(p) => run(() => unwrap(window.api.entra.assign(tid, app.id, p.id, role)), t('eaAssigned', p.displayName))} />
        </div>
      </div>
      <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
      {list.data && (
        <div className="panel">
          <table>
            <thead>
              <tr>
                <th>{t('eaPrincipal')}</th>
                <th>{t('arType')}</th>
                <th>{t('eaRole')}</th>
                <th>{t('eaSince')}</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {list.data.map((a) => (
                <tr key={a.id}>
                  <td className="u-n">{a.principalName}</td>
                  <td>{a.principalType === 'Group' ? t('enGroupTag') : a.principalType === 'ServicePrincipal' ? t('eaSp') : t('enUserTag')}</td>
                  <td>{roleName(a.appRoleId)}</td>
                  <td className="mono">{fmtDate(a.created, lang, '—')}</td>
                  <td style={{ textAlign: 'right' }}>
                    <button className="btn sm danger" disabled={busy} onClick={() => confirm(t('eaUnassignConfirm', a.principalName)) && run(() => unwrap(window.api.entra.unassign(tid, app.id, a.id)), t('eaUnassigned'))}>
                      {t('remove')}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!list.data.length && (
            <p className="hint" style={{ padding: 14 }}>
              {t('eaNoUsers')}
            </p>
          )}
        </div>
      )}
    </>
  )
}

/** What the app has been granted in this tenant, read from the tenant-wide grants list. */
function Granted({ app }: { app: EnterpriseApp }) {
  const { t, tenant } = useApp()
  const list = useLoad(() => window.api.entra.grants(tenant!.tenantId), [tenant!.tenantId])
  const g = list.data?.find((x) => x.spId === app.id)
  return (
    <>
      <EntraLoad loading={list.loading} error={list.error} onRetry={list.reload} text={t('raLoading')} />
      {list.data && !g?.permissions.length && <p className="hint">{t('eaNoPerms')}</p>}
      {g && g.permissions.length > 0 && (
        <div className="panel">
          <table>
            <thead>
              <tr>
                <th>{t('arPermission')}</th>
                <th>{t('arType')}</th>
                <th>{t('raConsent')}</th>
              </tr>
            </thead>
            <tbody>
              {g.permissions.map((p) => (
                <tr key={p.grantId + p.value + p.consent}>
                  <td>
                    <span className="mono" style={{ color: 'var(--fg)' }}>{p.value}</span> {p.risky && <span className="flag external">{t('raRisky')}</span>}
                    <div className="hint">{p.resource}</div>
                  </td>
                  <td>{p.kind === 'application' ? t('arApplication') : t('arDelegated')}</td>
                  <td>{p.consent === 'admin' ? t('raAdmin') : p.consent === 'user' ? t('raUser') : t('raAppPerm')}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  )
}
