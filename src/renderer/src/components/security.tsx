import { useEffect, useState, type ReactNode } from 'react'
import type { ExoPhase, RetractSpec } from '@shared/types'
import { unwrap, useApp, type ApiError } from '../ctx'
import { ExoLoad } from './exo'
import { EntraLoad } from './entra'

/** Set by other security screens (reported messages, message trace) so "remove from mailboxes" opens pre-filled. */
export const retractPrefill: { spec: RetractSpec | null } = { spec: null }

const sccPhases = new Map<string, ExoPhase>()
window.api.sec.onStatus((s) => sccPhases.set(s.tenantId, s.phase))

function useSccPhase(tenantId: string) {
  const [phase, setPhase] = useState<ExoPhase | undefined>(sccPhases.get(tenantId))
  useEffect(() => window.api.sec.onStatus((s) => s.tenantId === tenantId && setPhase(s.phase)), [tenantId])
  return phase
}

/**
 * LoadState for security screens. Exchange-based screens get the Exchange states; `consent` adds the
 * one-time Graph consent for that scope set; `scc` shows the Security & Compliance connection and its
 * GDAP caveat.
 */
export function SecLoad({ loading, error, onRetry, text, consent, scc }: {
  loading: boolean
  error: ApiError | null
  onRetry(): void
  text?: string
  consent?: 'risk' | 'submissions' | 'defender'
  scc?: boolean
}) {
  const { t, tenant, toast } = useApp()
  const phase = useSccPhase(tenant!.tenantId)
  const [busy, setBusy] = useState(false)
  if (consent && error?.code === 'NEEDS_CONSENT')
    return (
      <div className="banner">
        <span>{t(`seConsent_${consent}`)}</span>
        <button
          className="btn sm"
          disabled={busy}
          onClick={async () => {
            setBusy(true)
            try {
              await unwrap(window.api.sec.consent(tenant!.tenantId, consent))
              onRetry()
            } catch (e) {
              toast((e as Error).message)
            } finally {
              setBusy(false)
            }
          }}
        >
          {busy && <span className="spin" />}
          {t('grantAccess')}
        </button>
      </div>
    )
  if (consent) return <EntraLoad loading={loading} error={error} onRetry={onRetry} text={text} />
  if (error?.code === 'SCC_CONNECT')
    return (
      <div className="banner bad">
        <span>
          <b>{t('seSccFailT')}</b> {tenant!.mode === 'gdap' ? t('seSccGdap') : t('seSccFailD')} <span className="hint">{error.message}</span>
        </span>
      </div>
    )
  if (scc && loading && phase && phase !== 'ready')
    return (
      <p className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="spin" /> {phase === 'module' ? t('exoModule') : t('seSccConnecting')}
      </p>
    )
  return <ExoLoad loading={loading} error={error} onRetry={onRetry} text={text} />
}

export function Head({ title, children }: { title: string; children?: ReactNode }) {
  const { t, tenant } = useApp()
  return (
    <header className="top">
      <div>
        <div className="crumb">
          {tenant!.name} · {t('security')}
        </div>
        <h1>{title}</h1>
      </div>
      <div className="top-r">{children}</div>
    </header>
  )
}

/** yyyy-mm-dd for date inputs, n days back from today. */
export const dayInput = (daysBack = 0) => {
  const d = new Date(Date.now() - daysBack * 86400000)
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10)
}
/** Start or end of a local day as ISO. */
export const dayStart = (v: string) => (v ? new Date(`${v}T00:00:00`).toISOString() : undefined)
export const dayEnd = (v: string) => (v ? new Date(`${v}T23:59:59`).toISOString() : undefined)

export function fmtWhen(iso: string | null | undefined, lang: string) {
  if (!iso) return '—'
  return new Date(iso).toLocaleString(lang, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })
}

/** Splits a textarea of addresses, domains or URLs into entries. */
export const entries = (v: string) =>
  v
    .split(/[\s,;]+/)
    .map((x) => x.trim())
    .filter(Boolean)

/**
 * Renders a mail body without running anything in it: a sandboxed frame with no scripts and a policy that
 * only allows inline styles and embedded images, so remote images (tracking pixels) are never loaded.
 */
export function MailBody({ body, isHtml }: { body: string; isHtml: boolean }) {
  const csp = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: cid:; font-src data:">`
  const esc = (s: string) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!)
  const doc = isHtml
    ? `${csp}<base target="_blank">${body}`
    : `${csp}<pre style="white-space:pre-wrap;font:13px Segoe UI,sans-serif;margin:0">${esc(body)}</pre>`
  return <iframe title="mail" sandbox="" srcDoc={doc} style={{ width: '100%', height: 360, border: '1px solid var(--line)', borderRadius: 8, background: '#fff' }} />
}
