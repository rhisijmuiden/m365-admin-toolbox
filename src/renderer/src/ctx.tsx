import { createContext, useCallback, useContext, useEffect, useRef, useState } from 'react'
import type { Lang, Result, Settings, Tenant } from '@shared/types'
import type { T } from './i18n'

export interface Ctx {
  t: T
  lang: Lang
  settings: Settings
  tenant: Tenant | null
  tenants: Tenant[]
  toast(msg: string): void
  setSettings(patch: Partial<Settings>): Promise<void>
  /** Called after a successful Graph load so the status bar can show the refresh time. */
  markSynced(): void
  /** Re-runs the interactive sign-in for the current tenant. */
  reauth(): Promise<void>
}

export const AppCtx = createContext<Ctx>(null as unknown as Ctx)
export const useApp = () => useContext(AppCtx)

export class ApiError extends Error {
  constructor(
    message: string,
    public code?: string
  ) {
    super(message)
  }
}

export async function unwrap<T>(p: Promise<Result<T>>): Promise<T> {
  const r = await p
  if (!r.ok) throw new ApiError(r.error, r.code)
  return r.data
}

export interface Load<T> {
  data: T | null
  error: ApiError | null
  loading: boolean
  reload(): void
  set(fn: (d: T) => T): void
}

/** Loads data for the current tenant and reloads whenever a dependency changes. */
export function useLoad<T>(fn: () => Promise<Result<T>> | null, deps: unknown[]): Load<T> {
  const [data, setData] = useState<T | null>(null)
  const [error, setError] = useState<ApiError | null>(null)
  const [loading, setLoading] = useState(false)
  const [n, setN] = useState(0)
  const { markSynced } = useApp()
  const seq = useRef(0)
  useEffect(() => {
    const p = fn()
    if (!p) {
      setData(null)
      return
    }
    const my = ++seq.current
    setLoading(true)
    setError(null)
    unwrap(p)
      .then((d) => {
        if (my !== seq.current) return
        setData(d)
        markSynced()
      })
      .catch((e: ApiError) => my === seq.current && (setError(e), setData(null)))
      .finally(() => my === seq.current && setLoading(false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [...deps, n])
  const reload = useCallback(() => setN((x) => x + 1), [])
  const set = useCallback((f: (d: T) => T) => setData((d) => (d ? f(d) : d)), [])
  return { data, error, loading, reload, set }
}

export const initials = (n: string) =>
  n
    .replace(/^(Mr|Mw|Dhr)\.?\s+/i, '')
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((x) => x[0])
    .join('')
    .toUpperCase() || '?'

export function genPassword(len = 14) {
  const sets = ['ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%&*?']
  const all = sets.join('')
  const rnd = (n: number) => crypto.getRandomValues(new Uint32Array(1))[0] % n
  const chars = sets.map((s) => s[rnd(s.length)])
  while (chars.length < len) chars.push(all[rnd(all.length)])
  for (let i = chars.length - 1; i > 0; i--) {
    const j = rnd(i + 1)
    ;[chars[i], chars[j]] = [chars[j], chars[i]]
  }
  return chars.join('')
}
