import { useEffect, useState, type ReactNode } from 'react'
import type { ExoPhase, Recipient } from '@shared/types'
import { unwrap, useApp, type ApiError } from '../ctx'
import type { Key } from '../i18n'
import { LoadState } from './ui'

/** Last known Exchange connection phase per tenant, so a screen opened mid-connect shows the right text. */
const phases = new Map<string, ExoPhase>()
window.api.exo.onStatus((s) => phases.set(s.tenantId, s.phase))

export function useExoPhase(tenantId: string) {
  const [phase, setPhase] = useState<ExoPhase | undefined>(phases.get(tenantId))
  useEffect(() => window.api.exo.onStatus((s) => s.tenantId === tenantId && setPhase(s.phase)), [tenantId])
  return phase
}

/**
 * LoadState for Exchange screens: the one-time Exchange sign-in, a missing Exchange role, and what the
 * PowerShell helper is doing while the first command of a tenant waits for the module and connection.
 */
export function ExoLoad({ loading, error, onRetry, text }: { loading: boolean; error: ApiError | null; onRetry(): void; text?: string }) {
  const { t, tenant, toast, goPim } = useApp()
  const phase = useExoPhase(tenant!.tenantId)
  const [busy, setBusy] = useState(false)
  if (error?.code === 'NEEDS_EXO_SIGNIN')
    return (
      <div className="banner">
        <span>{t('exoSignInT')}</span>
        <button
          className="btn sm"
          disabled={busy}
          onClick={async () => {
            setBusy(true)
            try {
              await unwrap(window.api.exo.signIn(tenant!.tenantId))
              onRetry()
            } catch (e) {
              toast((e as Error).message)
            } finally {
              setBusy(false)
            }
          }}
        >
          {busy && <span className="spin" />}
          {t('exoSignInBtn')}
        </button>
      </div>
    )
  if (error?.code === 'EXO_DENIED')
    return (
      <div className="banner">
        <span>
          <b>{t('noAccessT')}</b> {t('exoDeniedD')} <span className="hint">{error.message}</span>
        </span>
        <button className="btn sm" onClick={goPim}>
          {t('activateRole')}
        </button>
      </div>
    )
  if (loading)
    return (
      <p className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="spin" /> {phase === 'module' ? t('exoModule') : phase === 'connecting' ? t('exoConnecting') : text || t('loading')}
      </p>
    )
  return <LoadState loading={false} error={error} />
}

export function fmtBytes(n: number | null | undefined, lang: string) {
  if (n == null) return '—'
  const u = ['B', 'KB', 'MB', 'GB', 'TB']
  let i = 0
  let v = n
  while (v >= 1024 && i < u.length - 1) (v /= 1024), i++
  return `${v.toLocaleString(lang, { maximumFractionDigits: v < 10 && i > 0 ? 1 : 0 })} ${u[i]}`
}

/** Readable label for an Exchange RecipientTypeDetails value, or the raw value when unknown. */
export function useTypeLabel() {
  const { t } = useApp()
  return (type: string) => (t(`t_${type}` as Key) as string | undefined) || type
}

export const domainOf = (a: string) => a.split('@')[1]?.toLowerCase() || ''
export const isExternal = (addr: string | null | undefined, domains: string[]) => !!addr && addr.includes('@') && !domains.includes(domainOf(addr))
export const stripSmtp = (a: string) => a.replace(/^smtp:/i, '')

export function Tabs<K extends string>({ tabs, value, onChange }: { tabs: [K, string][]; value: K; onChange(k: K): void }) {
  return (
    <div className="seg" style={{ marginBottom: 14, flexWrap: 'wrap' }}>
      {tabs.map(([k, label]) => (
        <button key={k} className={k === value ? 'on' : ''} onClick={() => onChange(k)}>
          {label}
        </button>
      ))}
    </div>
  )
}

/** Recipients of a tenant for address suggestions, loaded once per tenant per app session. */
const recipientCache = new Map<string, Promise<Recipient[]>>()
export function useRecipients(tenantId: string) {
  const [list, setList] = useState<Recipient[]>([])
  useEffect(() => {
    let p = recipientCache.get(tenantId)
    if (!p) {
      p = unwrap(window.api.exo.recipients(tenantId))
      recipientCache.set(tenantId, p)
      p.catch(() => recipientCache.delete(tenantId))
    }
    let live = true
    p.then((r) => live && setList(r)).catch(() => undefined)
    return () => {
      live = false
    }
  }, [tenantId])
  return list
}

/** Email input with suggestions from the tenant's recipients, plus an Add button. */
export function AddRecipient({ placeholder, onAdd, filter, busy }: { placeholder: string; onAdd(addr: string): Promise<void> | void; filter?(r: Recipient): boolean; busy?: boolean }) {
  const { t, tenant } = useApp()
  const all = useRecipients(tenant!.tenantId)
  const [v, setV] = useState('')
  const [id] = useState(() => `rcp-${Math.random().toString(36).slice(2, 8)}`)
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
        {all.filter(filter || (() => true)).slice(0, 500).map((r) => (
          <option key={r.id} value={r.primarySmtp}>
            {r.displayName}
          </option>
        ))}
      </datalist>
      <button className="btn sm" disabled={busy || !v.trim()} onClick={submit}>
        {busy && <span className="spin" />}
        {t('add')}
      </button>
    </div>
  )
}

/** A small removable list of addresses or names. */
export function Chips({ items, onRemove, empty }: { items: { key: string; label: ReactNode }[]; onRemove?(key: string): void; empty: string }) {
  if (!items.length) return <p className="hint">{empty}</p>
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
      {items.map((i) => (
        <span key={i.key} className="lic-tag" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
          {i.label}
          {onRemove && (
            <button className="x" style={{ fontSize: 14, lineHeight: 1, padding: 0 }} aria-label="remove" onClick={() => onRemove(i.key)}>
              ×
            </button>
          )}
        </span>
      ))}
    </div>
  )
}
