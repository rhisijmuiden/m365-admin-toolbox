import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { Settings, Tenant } from '@shared/types'
import { AppCtx, initials, unwrap, useApp, type Ctx } from './ctx'
import { makeT, type Key } from './i18n'
import { Icon } from './components/ui'
import { AddTenantModal } from './components/AddTenant'
import { InfoModal } from './components/Info'
import { UsersView } from './views/Users'
import { LicensesView } from './views/Licenses'
import { GroupsView } from './views/Groups'
import { PimView } from './views/Pim'
import { MailboxesView } from './views/Mailboxes'
import { DistListsView } from './views/DistLists'
import { AddressesView } from './views/Addresses'
import { InboxRulesView } from './views/InboxRules'
import { MailFlowView } from './views/MailFlow'
import { TeamsView } from './views/Teams'
import { TeamGuestsView } from './views/TeamGuests'
import { TeamPoliciesView } from './views/TeamPolicies'
import { TeamAccessView } from './views/TeamAccess'
import { TeamPhoneView } from './views/TeamPhone'
import { TeamTemplatesView } from './views/TeamTemplates'
import logoDark from './assets/rhis-logo.svg'
import logoLight from './assets/rhis-logo-wit.svg'

type View =
  | 'pim'
  | 'users'
  | 'licenses'
  | 'groups'
  | 'mailboxes'
  | 'shared'
  | 'rooms'
  | 'distLists'
  | 'aliases'
  | 'inboxRules'
  | 'mailflow'
  | 'teamsList'
  | 'teamGuests'
  | 'policies'
  | 'teamAccess'
  | 'teamPhone'
  | 'teamTemplates'
  | 'apps' | 'ca' | 'roles'

const NAV: { id: 'access' | 'm365' | 'exchange' | 'teams' | 'entra'; bg: string; items: [View, number][] }[] = [
  { id: 'access', bg: 'linear-gradient(135deg,#ff8a6b,#e0503c)', items: [['pim', 0]] },
  { id: 'm365', bg: 'linear-gradient(135deg,#ea5a2b,#c2381f)', items: [['users', 0], ['licenses', 0], ['groups', 0]] },
  { id: 'exchange', bg: 'linear-gradient(135deg,#2a8be0,#0a5aa6)', items: [['mailboxes', 0], ['shared', 0], ['rooms', 0], ['distLists', 0], ['aliases', 0], ['inboxRules', 0], ['mailflow', 0]] },
  { id: 'teams', bg: 'linear-gradient(135deg,#7377dd,#4b4fb3)', items: [['teamsList', 0], ['teamGuests', 0], ['policies', 0], ['teamAccess', 0], ['teamPhone', 0], ['teamTemplates', 0]] },
  { id: 'entra', bg: 'linear-gradient(135deg,#17a6a0,#0a6f75)', items: [['apps', 4], ['ca', 4], ['roles', 4]] }
]

function useDark() {
  const mq = useMemo(() => window.matchMedia('(prefers-color-scheme: dark)'), [])
  const [dark, setDark] = useState(mq.matches)
  useEffect(() => {
    const f = () => setDark(mq.matches)
    mq.addEventListener('change', f)
    return () => mq.removeEventListener('change', f)
  }, [mq])
  return dark
}

export function App() {
  const [settings, setS] = useState<Settings | null>(null)
  const [tenants, setTenants] = useState<Tenant[]>([])
  const [curId, setCurId] = useState<string | undefined>()
  const [view, setView] = useState<View>('users')
  const [modal, setModal] = useState<'add' | 'info' | null>(null)
  const [toastMsg, setToast] = useState<string | null>(null)
  const [syncedAt, setSyncedAt] = useState<Date | null>(null)
  const [, tick] = useState(0)
  const dark = useDark()

  useEffect(() => {
    Promise.all([unwrap(window.api.getSettings()), unwrap(window.api.listTenants())]).then(([s, ts]) => {
      setS(s)
      setTenants(ts)
      setCurId(ts.find((x) => x.tenantId === s.lastTenantId)?.tenantId || ts[0]?.tenantId)
    })
    const i = setInterval(() => tick((x) => x + 1), 30000)
    // A click on a "role expires soon" notification opens PIM for that tenant.
    const off = window.api.onNav((n) => {
      setCurId(n.tenantId)
      setView(n.view as View)
    })
    return () => (clearInterval(i), off())
  }, [])

  useEffect(() => {
    document.documentElement.dataset.theme = dark ? 'dark' : 'light'
  }, [dark])

  const lang = settings?.lang || 'en'
  const t = useMemo(() => makeT(lang), [lang])
  useEffect(() => {
    document.documentElement.lang = lang
  }, [lang])

  const toastTimer = useRef<number>()
  const toast = useCallback((m: string) => {
    setToast(m)
    clearTimeout(toastTimer.current)
    toastTimer.current = window.setTimeout(() => setToast(null), 3200)
  }, [])

  const setSettings = useCallback(async (patch: Partial<Settings>) => {
    setS(await unwrap(window.api.setSettings(patch)))
  }, [])

  const tenant = tenants.find((x) => x.tenantId === curId) || null
  const [epoch, setEpoch] = useState(0)

  const reauth = useCallback(async () => {
    if (!tenant) return
    try {
      const nt = await unwrap(window.api.reauthTenant(tenant.tenantId))
      setTenants((ts) => ts.map((x) => (x.tenantId === nt.tenantId ? nt : x)))
      setEpoch((e) => e + 1)
    } catch (e) {
      toast((e as Error).message)
    }
  }, [tenant, toast])

  const markSynced = useCallback(() => setSyncedAt(new Date()), [])

  const selectTenant = (id: string) => {
    setCurId(id)
    setSyncedAt(null)
    setSettings({ lastTenantId: id })
  }

  if (!settings) return null

  const goPim = () => setView('pim')
  const ctx: Ctx = { t, lang, settings, tenant, tenants, toast, setSettings, markSynced, reauth, goPim }
  const group = NAV.find((g) => g.items.some(([k]) => k === view))!
  const phase = group.items.find(([k]) => k === view)![1]
  const ago = syncedAt ? new Intl.RelativeTimeFormat(lang, { numeric: 'auto' }).format(-Math.floor((Date.now() - syncedAt.getTime()) / 60000), 'minute') : null

  return (
    <AppCtx.Provider value={ctx}>
      <div className="app">
        <aside className="rail">
          <div className="brand">
            <div className="brand-mark">M</div>
            <div>
              <div className="brand-name">{t('appName')}</div>
              <div className="brand-sub">{t('appSub')}</div>
            </div>
          </div>
          <TenantSwitcher
            tenant={tenant}
            tenants={tenants}
            onSelect={selectTenant}
            onAdd={() => setModal('add')}
            onRemove={async (x) => {
              if (!confirm(t('removeConfirm', x.name))) return
              const ts = await unwrap(window.api.removeTenant(x.tenantId))
              setTenants(ts)
              if (x.tenantId === curId) setCurId(ts[0]?.tenantId)
            }}
          />
          <nav className="nav">
            {NAV.map((g) => {
              const open = g.items.some(([k]) => k === view)
              return (
                <div key={g.id}>
                  <button className={`nav-item ${open ? 'on' : ''}`} onClick={() => setView(g.items[0][0])}>
                    <span className="ico" style={{ background: g.bg }}>
                      {Icon[g.id]}
                    </span>
                    <span>{t(g.id)}</span>
                  </button>
                  <div className="sub">
                    {g.items.map(([k, p]) => (
                      <button key={k} className={`nav-item ${k === view ? 'on' : ''}`} onClick={() => setView(k)}>
                        <span>{t(k)}</span>
                        {p > 0 && (
                          <span className="soon">
                            {t('phase')} {p}
                          </span>
                        )}
                      </button>
                    ))}
                  </div>
                </div>
              )
            })}
          </nav>
          <div className="rail-foot">
            <button onClick={() => setModal('info')} title="Info">
              {Icon.info}
              <span>Info</span>
            </button>
            <button onClick={() => setSettings({ lang: lang === 'en' ? 'nl' : 'en' })} title={t('language')}>
              {Icon.globe}
              <span>{lang.toUpperCase()}</span>
            </button>
          </div>
        </aside>

        <section className="main">
          {!tenant ? (
            <div className="content" style={{ display: 'grid', placeItems: 'center' }}>
              <div className="empty" style={{ maxWidth: 560 }}>
                <span className="ico" style={{ background: NAV[1].bg, width: 44, height: 44, margin: '0 auto', borderRadius: 11 }}>
                  {Icon.m365}
                </span>
                <h3>{t('welcomeT')}</h3>
                <p>{t('welcomeD')}</p>
                <button className="btn pri" onClick={() => setModal('add')}>
                  {Icon.plus}
                  {t('addTenant')}
                </button>
              </div>
            </div>
          ) : phase > 0 ? (
            <Soon title={t(view as Key) as string} crumb={`${tenant.name} · ${t(group.id)}`} phase={phase} bg={group.bg} icon={Icon[group.id]} />
          ) : view === 'pim' ? (
            <PimView key={tenant.tenantId + epoch} />
          ) : view === 'users' ? (
            <UsersView key={tenant.tenantId + epoch} />
          ) : view === 'licenses' ? (
            <LicensesView key={tenant.tenantId + epoch} />
          ) : view === 'mailboxes' || view === 'shared' || view === 'rooms' ? (
            <MailboxesView key={tenant.tenantId + epoch + view} scope={view === 'mailboxes' ? 'all' : view === 'shared' ? 'shared' : 'resources'} />
          ) : view === 'distLists' ? (
            <DistListsView key={tenant.tenantId + epoch} />
          ) : view === 'aliases' ? (
            <AddressesView key={tenant.tenantId + epoch} />
          ) : view === 'inboxRules' ? (
            <InboxRulesView key={tenant.tenantId + epoch} />
          ) : view === 'mailflow' ? (
            <MailFlowView key={tenant.tenantId + epoch} />
          ) : view === 'teamsList' ? (
            <TeamsView key={tenant.tenantId + epoch} />
          ) : view === 'teamGuests' ? (
            <TeamGuestsView key={tenant.tenantId + epoch} />
          ) : view === 'policies' ? (
            <TeamPoliciesView key={tenant.tenantId + epoch} />
          ) : view === 'teamAccess' ? (
            <TeamAccessView key={tenant.tenantId + epoch} />
          ) : view === 'teamPhone' ? (
            <TeamPhoneView key={tenant.tenantId + epoch} />
          ) : view === 'teamTemplates' ? (
            <TeamTemplatesView />
          ) : (
            <GroupsView key={tenant.tenantId + epoch} />
          )}
        </section>

        <footer className="status">
          <span>
            <span className={`dot ${tenant ? '' : 'off'}`} />
            {tenant ? t('connected', tenant.name) : t('notConnected')}
          </span>
          {ago && <span className="mono">{t('synced', ago)}</span>}
          <div className="powered">
            <span>{t('poweredBy')}</span>
            <img src={dark ? logoLight : logoDark} alt="RHIS" />
          </div>
        </footer>
      </div>

      {modal === 'add' && (
        <AddTenantModal
          onClose={() => setModal(null)}
          onAdded={(nt) => {
            setTenants((ts) => [...ts.filter((x) => x.tenantId !== nt.tenantId), nt].sort((a, b) => a.name.localeCompare(b.name)))
            selectTenant(nt.tenantId)
            setModal(null)
            toast(t('atDone', nt.name))
          }}
        />
      )}
      {modal === 'info' && <InfoModal onClose={() => setModal(null)} logo={dark ? logoLight : logoDark} />}
      {toastMsg && <div className="toast">{toastMsg}</div>}
    </AppCtx.Provider>
  )
}

function TenantSwitcher({ tenant, tenants, onSelect, onAdd, onRemove }: {
  tenant: Tenant | null
  tenants: Tenant[]
  onSelect(id: string): void
  onAdd(): void
  onRemove(t: Tenant): void
}) {
  const [open, setOpen] = useState(false)
  const [q, setQ] = useState('')
  const ref = useRef<HTMLDivElement>(null)
  const { t } = useApp()
  useEffect(() => {
    if (!open) return
    const f = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && setOpen(false)
    document.addEventListener('mousedown', f)
    return () => document.removeEventListener('mousedown', f)
  }, [open])
  const list = tenants.filter((x) => (x.name + x.defaultDomain).toLowerCase().includes(q.toLowerCase()))
  return (
    <div className="tenant" ref={ref}>
      <button className="tenant-btn" aria-haspopup="listbox" aria-expanded={open} onClick={() => setOpen(!open)}>
        <div className="t-av" style={{ background: tenant?.color || '#3a475a' }}>
          {tenant ? initials(tenant.name) : '+'}
        </div>
        <div className="t-meta">
          <div className="t-name">{tenant?.name || t('noTenant')}</div>
          <div className="t-dom">{tenant?.defaultDomain || t('addTenant')}</div>
        </div>
        {Icon.chev}
      </button>
      {open && (
        <div className="tenant-menu">
          {tenants.length > 5 && <input className="tm-search" autoFocus placeholder={t('searchTenants')} value={q} onChange={(e) => setQ(e.target.value)} />}
          {list.map((x) => (
            <div key={x.tenantId} className={`tm-row ${x.tenantId === tenant?.tenantId ? 'on' : ''}`}>
              <button
                className="tm-item"
                onClick={() => {
                  onSelect(x.tenantId)
                  setOpen(false)
                }}
              >
                <div className="t-av" style={{ background: x.color }}>
                  {initials(x.name)}
                </div>
                <div className="t-meta">
                  <div className="t-name">{x.name}</div>
                  <div className="t-dom">{x.defaultDomain}</div>
                </div>
                <span className={`chip ${x.mode}`}>{t(x.mode)}</span>
              </button>
              <button className="tm-x" title={t('removeTenant')} onClick={() => onRemove(x)}>
                ×
              </button>
            </div>
          ))}
          <button
            className="tm-add"
            onClick={() => {
              setOpen(false)
              onAdd()
            }}
          >
            {Icon.plus}
            {t('addTenant')}
          </button>
        </div>
      )}
    </div>
  )
}

function Soon({ title, crumb, phase, bg, icon }: { title: string; crumb: string; phase: number; bg: string; icon: JSX.Element }) {
  const { t } = useApp()
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">{crumb}</div>
          <h1>{title}</h1>
        </div>
        <div className="top-r">
          <span className="pill off">
            {t('phase')} {phase}
          </span>
        </div>
      </header>
      <div className="content">
        <div className="empty">
          <span className="ico" style={{ background: bg, width: 40, height: 40, margin: '0 auto', borderRadius: 10 }}>
            {icon}
          </span>
          <h3>{title}</h3>
          <p>{t('phaseHint')}</p>
        </div>
      </div>
    </>
  )
}
