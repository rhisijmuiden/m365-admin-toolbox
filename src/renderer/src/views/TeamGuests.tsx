import { useMemo, useState } from 'react'
import type { TeamGuest } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, RowMenu } from '../components/ui'
import { daysSince, TeamsLoad } from '../components/teams'

const STALE_DAYS = 90

/** Guests per team: who is external, since when, last sign-in, and removing them from a team or the tenant. */
export function TeamGuestsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const teams = useLoad(() => window.api.teams.list(tid), [tid])
  const list = useLoad(() => (teams.data ? window.api.teams.guests(tid, teams.data.map((x) => ({ id: x.id, displayName: x.displayName }))) : null), [tid, teams.data])
  const [q, setQ] = useState('')
  const [onlyStale, setOnlyStale] = useState(false)

  const rows = list.data || []
  const people = useMemo(() => new Set(rows.map((g) => g.userId)), [rows])
  const stale = (g: TeamGuest) => {
    const d = daysSince(g.lastSignIn)
    return g.lastSignIn ? (d ?? 0) >= STALE_DAYS : (daysSince(g.created) ?? 0) >= STALE_DAYS
  }
  const shown = rows
    .filter((g) => !onlyStale || stale(g))
    .filter((g) => (g.displayName + (g.mail || '') + g.teamName).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.teamName.localeCompare(b.teamName) || a.displayName.localeCompare(b.displayName))
  const teamsWithGuests = new Set(rows.map((g) => g.teamId)).size

  async function fromTeam(g: TeamGuest) {
    if (!confirm(t('tgRemoveConfirm', g.displayName, g.teamName))) return
    try {
      await unwrap(window.api.teams.removeGuest(tid, g.teamId, g.userId))
      list.set((d) => d.filter((x) => !(x.userId === g.userId && x.teamId === g.teamId)))
      toast(t('tgRemoved', g.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  async function account(g: TeamGuest) {
    if (!confirm(t('tgDeleteConfirm', g.displayName))) return
    try {
      await unwrap(window.api.teams.deleteGuest(tid, g.userId))
      list.set((d) => d.filter((x) => x.userId !== g.userId))
      toast(t('tgDeleted', g.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  const error = teams.error || list.error
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('teams')}
          </div>
          <h1>{t('teamGuests')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('tgSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={teams.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <TeamsLoad loading={(teams.loading || list.loading) && !list.data} error={error} onRetry={teams.reload} text={teams.data ? t('tgLoading') : undefined} />
        {list.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('tgGuests')}</div>
                <div className="v">{people.size}</div>
                <div className="s">{t('tgMemberships', rows.length)}</div>
              </div>
              <div className="tile">
                <div className="k">{t('tgTeams')}</div>
                <div className="v">{teamsWithGuests}</div>
                <div className="s">{t('tgOfTeams', teams.data?.length || 0)}</div>
              </div>
              <div className={`tile ${rows.some(stale) ? 'warn' : ''}`}>
                <div className="k">{t('tgStale', STALE_DAYS)}</div>
                <div className="v">{new Set(rows.filter(stale).map((g) => g.userId)).size}</div>
                <div className="s">{t('tgStaleD')}</div>
              </div>
              <div className="tile">
                <div className="k">{t('tgPending')}</div>
                <div className="v">{new Set(rows.filter((g) => g.state === 'PendingAcceptance').map((g) => g.userId)).size}</div>
                <div className="s">{t('tgPendingD')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="panel-h">
                <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={onlyStale} onChange={(e) => setOnlyStale(e.target.checked)} />
                  {t('tgOnlyStale')}
                </label>
              </div>
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>{t('tgGuest')}</th>
                      <th>{t('tgTeam')}</th>
                      <th>{t('tgSince')}</th>
                      <th>{t('lastSign')}</th>
                      <th>{t('status')}</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((g) => (
                      <tr key={g.teamId + g.userId}>
                        <td>
                          <div className="u-n">{g.displayName}</div>
                          <div className="mono">{g.mail || g.upn}</div>
                        </td>
                        <td>{g.teamName}</td>
                        <td className="mono">{fmtDate(g.created, lang, '—')}</td>
                        <td className="mono">{fmtDate(g.lastSignIn, lang, t('never'))}</td>
                        <td>
                          {g.state === 'PendingAcceptance' ? (
                            <span className="pill warn">{t('tgPendingPill')}</span>
                          ) : stale(g) ? (
                            <span className="pill warn">{t('tmInactivePill')}</span>
                          ) : (
                            <span className="pill ok">{t('active')}</span>
                          )}
                        </td>
                        <td style={{ width: 40 }}>
                          <RowMenu
                            items={[
                              [t('tgFromTeam'), () => fromTeam(g)],
                              [t('tgDeleteAccount'), () => account(g)]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!shown.length && (
                  <p className="hint" style={{ padding: 16 }}>
                    {t('tgNone')}
                  </p>
                )}
              </div>
            </div>
            <p className="hint">{t('tgSinceHint')}</p>
          </>
        )}
      </div>
    </>
  )
}
