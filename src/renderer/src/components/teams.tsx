import { useEffect, useState } from 'react'
import type { ExoPhase, Group, User } from '@shared/types'
import { unwrap, useApp, type ApiError } from '../ctx'
import { LoadState } from './ui'

/** Last known Teams PowerShell phase per tenant, like the Exchange one. */
const phases = new Map<string, ExoPhase>()
window.api.teams.onStatus((s) => phases.set(s.tenantId, s.phase))

function useTeamsPhase(tenantId: string) {
  const [phase, setPhase] = useState<ExoPhase | undefined>(phases.get(tenantId))
  useEffect(() => window.api.teams.onStatus((s) => s.tenantId === tenantId && setPhase(s.phase)), [tenantId])
  return phase
}

function ActionBanner({ text, button, run, onDone }: { text: string; button: string; run(): Promise<unknown>; onDone(): void }) {
  const { toast } = useApp()
  const [busy, setBusy] = useState(false)
  return (
    <div className="banner">
      <span>{text}</span>
      <button
        className="btn sm"
        disabled={busy}
        onClick={async () => {
          setBusy(true)
          try {
            await run()
            onDone()
          } catch (e) {
            toast((e as Error).message)
          } finally {
            setBusy(false)
          }
        }}
      >
        {busy && <span className="spin" />}
        {button}
      </button>
    </div>
  )
}

/**
 * LoadState for Teams screens: the one-time consent for the Teams Graph permissions, the one-time sign-in
 * for Teams PowerShell, a missing Teams role, and what the PowerShell helper is doing on first use.
 */
export function TeamsLoad({ loading, error, onRetry, text }: { loading: boolean; error: ApiError | null; onRetry(): void; text?: string }) {
  const { t, tenant, goPim } = useApp()
  const tid = tenant!.tenantId
  const phase = useTeamsPhase(tid)
  if (error?.code === 'NEEDS_CONSENT')
    return <ActionBanner text={t('tmConsentT')} button={t('grantAccess')} run={() => unwrap(window.api.teams.consent(tid))} onDone={onRetry} />
  if (error?.code === 'NEEDS_TEAMS_SIGNIN')
    return <ActionBanner text={t('tmSignInT')} button={t('tmSignInBtn')} run={() => unwrap(window.api.teams.signIn(tid))} onDone={onRetry} />
  if (error?.code === 'TEAMS_DENIED')
    return (
      <div className="banner">
        <span>
          <b>{t('noAccessT')}</b> {t('tmDeniedD')} <span className="hint">{error.message}</span>
        </span>
        <button className="btn sm" onClick={goPim}>
          {t('activateRole')}
        </button>
      </div>
    )
  if (loading)
    return (
      <p className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="spin" /> {phase === 'module' ? t('tmModule') : phase === 'connecting' ? t('tmConnecting') : text || t('loading')}
      </p>
    )
  return <LoadState loading={false} error={error} />
}

/** Graph users and groups of a tenant for pickers, loaded once per tenant per app session. */
const userCache = new Map<string, Promise<User[]>>()
const groupCache = new Map<string, Promise<Group[]>>()

function useCached<T>(cache: Map<string, Promise<T[]>>, tenantId: string, load: () => Promise<T[]>) {
  const [list, setList] = useState<T[]>([])
  useEffect(() => {
    let p = cache.get(tenantId)
    if (!p) {
      p = load()
      cache.set(tenantId, p)
      p.catch(() => cache.delete(tenantId))
    }
    let live = true
    p.then((r) => live && setList(r)).catch(() => undefined)
    return () => {
      live = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tenantId])
  return list
}

export const useUsers = (tenantId: string) => useCached(userCache, tenantId, () => unwrap(window.api.listUsers(tenantId)))
export const useGroups = (tenantId: string) => useCached(groupCache, tenantId, () => unwrap(window.api.listGroups(tenantId)))

/** UPN input with suggestions from the tenant's users (guests included), plus an Add button. */
export function AddUser({ placeholder, onAdd, busy, label }: { placeholder: string; onAdd(upn: string): Promise<void> | void; busy?: boolean; label?: string }) {
  const { t, tenant } = useApp()
  const users = useUsers(tenant!.tenantId)
  const [v, setV] = useState('')
  const [id] = useState(() => `usr-${Math.random().toString(36).slice(2, 8)}`)
  const submit = async () => {
    const a = v.trim()
    if (!a) return
    await onAdd(a)
    setV('')
  }
  return (
    <div className="pw" style={{ alignItems: 'center' }}>
      <input list={id} className="mono" placeholder={placeholder} value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && submit()} />
      <datalist id={id}>
        {users.slice(0, 1000).map((u) => (
          <option key={u.id} value={u.userPrincipalName}>
            {u.displayName}
          </option>
        ))}
      </datalist>
      <button className="btn sm" disabled={busy || !v.trim()} onClick={submit}>
        {busy && <span className="spin" />}
        {label || t('add')}
      </button>
    </div>
  )
}

export const daysSince = (iso: string | null | undefined) => (iso ? Math.floor((Date.now() - new Date(iso).getTime()) / 86400000) : null)
