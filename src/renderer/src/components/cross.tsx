import { useState, type ReactNode } from 'react'
import type { Tenant } from '@shared/types'
import { AppCtx, initials, unwrap, useApp, type ApiError } from '../ctx'

/** Runs a tenant-bound component for another tenant than the one selected in the switcher. */
export function TenantScope({ tenant, children }: { tenant: Tenant; children: ReactNode }) {
  const ctx = useApp()
  return <AppCtx.Provider value={{ ...ctx, tenant }}>{children}</AppCtx.Provider>
}

export function TenantTag({ tenant }: { tenant: Tenant }) {
  return (
    <span style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
      <span className="t-av" style={{ background: tenant.color, width: 22, height: 22, fontSize: 10 }}>
        {initials(tenant.name)}
      </span>
      {tenant.name}
    </span>
  )
}

export function TenantSelect({ value, onChange, id, all }: { value: string; onChange(id: string): void; id?: string; all?: string }) {
  const { tenants } = useApp()
  return (
    <select id={id} value={value} onChange={(e) => onChange(e.target.value)}>
      {all && <option value="">{all}</option>}
      {tenants.map((x) => (
        <option key={x.tenantId} value={x.tenantId}>
          {x.name} ({x.defaultDomain})
        </option>
      ))}
    </select>
  )
}

/** Runs fn for every tenant, a few at a time, so twenty tenants don't open twenty connections at once. */
export async function eachTenant<T>(list: T[], limit: number, fn: (x: T) => Promise<void>) {
  const queue = [...list]
  await Promise.all(
    Array.from({ length: Math.min(limit, queue.length) }, async () => {
      while (queue.length) await fn(queue.shift()!)
    })
  )
}

/** Fixes the errors a cross-tenant screen can fix by itself: a sign-in, a consent or the Exchange sign-in. */
export async function fixTenant(x: Tenant, e: ApiError, consent?: () => Promise<unknown>) {
  if (e.code === 'NEEDS_CONSENT' && consent) await consent()
  else if (e.code === 'NEEDS_EXO_SIGNIN') await unwrap(window.api.exo.signIn(x.tenantId))
  else await unwrap(window.api.reauthTenant(x.tenantId))
}

export const fixable = (e: ApiError | null | undefined) => !!e && ['NEEDS_CONSENT', 'NEEDS_SIGNIN', 'NEEDS_EXO_SIGNIN'].includes(e.code || '')

// CSV: Excel in Dutch saves with semicolons, English Excel with commas; both are read.

export interface Csv {
  delimiter: string
  headers: string[]
  rows: Record<string, string>[]
}

export function parseCsv(text: string): Csv {
  const src = text.replace(/^﻿/, '')
  const first = src.split(/\r?\n/, 1)[0] || ''
  const delimiter = (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : first.includes('\t') && !first.includes(',') ? '\t' : ','
  const lines: string[][] = []
  let row: string[] = []
  let cell = ''
  let quoted = false
  for (let i = 0; i < src.length; i++) {
    const c = src[i]
    if (quoted) {
      if (c === '"' && src[i + 1] === '"') (cell += '"'), i++
      else if (c === '"') quoted = false
      else cell += c
    } else if (c === '"') quoted = true
    else if (c === delimiter) row.push(cell), (cell = '')
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && src[i + 1] === '\n') i++
      row.push(cell)
      lines.push(row)
      row = []
      cell = ''
    } else cell += c
  }
  if (cell || row.length) row.push(cell), lines.push(row)
  const nonEmpty = lines.filter((l) => l.some((x) => x.trim()))
  const headers = (nonEmpty.shift() || []).map((h) => h.trim().replace(/\s+/g, '').toLowerCase())
  const rows = nonEmpty.map((l) => Object.fromEntries(headers.map((h, i) => [h, (l[i] || '').trim()])))
  return { delimiter, headers, rows }
}

/** Splits a list cell such as "E3+Teams" or "E3|Teams" (and ; or , when that isn't the column separator). */
export const listCell = (v: string | undefined, delimiter: string) =>
  (v || '')
    .split(new RegExp(`[${['|', '+', ';', ','].filter((d) => d !== delimiter).map((d) => '\\' + d).join('')}]`))
    .map((x) => x.trim())
    .filter(Boolean)

export const toCsv = (rows: string[][]) => '﻿' + rows.map((r) => r.map((v) => (/[",;\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(';')).join('\r\n')

export function download(name: string, text: string) {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }))
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 1000)
}

// Wizard steps

export type StepState = 'todo' | 'run' | 'ok' | 'fail' | 'skip'

export interface Step {
  key: string
  label: string
  /** Returns an optional note shown under the step, or throws. */
  run(): Promise<string | void>
}

export interface StepStatus {
  state: StepState
  note?: string
  error?: ApiError
}

/** Runs steps in order; a failed step doesn't stop the rest, so one missing role doesn't block the whole flow. */
export function useSteps() {
  const [status, setStatus] = useState<Record<string, StepStatus>>({})
  const [running, setRunning] = useState(false)
  async function run(steps: Step[], only?: (s: Step) => boolean) {
    setRunning(true)
    for (const s of steps) {
      if (only && !only(s)) continue
      setStatus((x) => ({ ...x, [s.key]: { state: 'run' } }))
      try {
        const note = await s.run()
        setStatus((x) => ({ ...x, [s.key]: { state: 'ok', note: note || undefined } }))
      } catch (e) {
        setStatus((x) => ({ ...x, [s.key]: { state: 'fail', error: e as ApiError } }))
      }
    }
    setRunning(false)
  }
  return { status, running, run, reset: () => setStatus({}) }
}

export function StepList({ steps, status }: { steps: Step[]; status: Record<string, StepStatus> }) {
  const { t } = useApp()
  return (
    <div style={{ display: 'grid', gap: 6 }}>
      {steps.map((s) => {
        const st = status[s.key] || { state: 'todo' }
        return (
          <div key={s.key} className={`opt ${st.state === 'fail' ? 'bad' : ''}`} style={{ cursor: 'default', alignItems: 'center' }}>
            <span style={{ width: 18, display: 'inline-flex', justifyContent: 'center' }}>
              {st.state === 'run' ? <span className="spin" /> : st.state === 'ok' ? '✓' : st.state === 'fail' ? '✕' : '•'}
            </span>
            <div style={{ minWidth: 0 }}>
              <b>{s.label}</b>
              {st.note && <small>{st.note}</small>}
              {st.error && <small style={{ color: 'var(--bad)' }}>{errText(t, st.error)}</small>}
            </div>
            <span className="avail">
              <span className={`pill ${st.state === 'ok' ? 'ok' : st.state === 'fail' ? 'bad' : st.state === 'run' ? 'warn' : 'off'}`}>{t(`crSt_${st.state}`)}</span>
            </span>
          </div>
        )
      })}
    </div>
  )
}

/** Error text with the known codes made readable. */
export function errText(t: ReturnType<typeof useApp>['t'], e: ApiError) {
  if (e.code === 'NEEDS_SIGNIN') return t('needsSignIn')
  if (e.code === 'NEEDS_CONSENT') return t('crNeedsConsent')
  if (e.code === 'NEEDS_EXO_SIGNIN') return t('exoSignInT')
  if (e.status === 403 || e.code === 'Authorization_RequestDenied') return `${t('noAccessT')} ${e.message}`
  return e.message
}
