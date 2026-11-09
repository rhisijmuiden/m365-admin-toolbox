import { useEffect, useRef, useState, type ReactNode } from 'react'
import { useApp, type ApiError } from '../ctx'

export const Icon = {
  plus: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4">
      <path d="M12 5v14M5 12h14" />
    </svg>
  ),
  search: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="11" cy="11" r="7" />
      <path d="M20 20l-3.5-3.5" />
    </svg>
  ),
  refresh: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7" />
    </svg>
  ),
  chev: (
    <svg className="chev" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M7 10l5 5 5-5" />
    </svg>
  ),
  info: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="12" cy="12" r="9" />
      <path d="M12 11v6M12 7.5v.01" />
    </svg>
  ),
  globe: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="12" cy="12" r="9" />
      <path d="M3 12h18M12 3c3 3.5 3 14.5 0 18M12 3c-3 3.5-3 14.5 0 18" />
    </svg>
  ),
  dots: (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
      <circle cx="5" cy="12" r="1.8" />
      <circle cx="12" cy="12" r="1.8" />
      <circle cx="19" cy="12" r="1.8" />
    </svg>
  ),
  m365: (
    <svg viewBox="0 0 24 24" fill="currentColor">
      <rect x="3" y="3" width="8" height="8" rx="1.5" />
      <rect x="13" y="3" width="8" height="8" rx="1.5" opacity=".75" />
      <rect x="3" y="13" width="8" height="8" rx="1.5" opacity=".75" />
      <rect x="13" y="13" width="8" height="8" rx="1.5" opacity=".5" />
    </svg>
  ),
  exchange: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <path d="M3 7l9 6 9-6" />
    </svg>
  ),
  teams: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="9" cy="8" r="3" />
      <circle cx="17" cy="9" r="2.3" />
      <path d="M3 19c0-3.3 2.7-5 6-5s6 1.7 6 5M15 14.5c3 0 6 1.3 6 4.5" />
    </svg>
  ),
  access: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="8" cy="15" r="4" />
      <path d="M11 12l9-9M17 6l3 3M14 9l2 2" />
    </svg>
  ),
  entra: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path d="M12 3l8 3v6c0 4.5-3.4 8.2-8 9-4.6-.8-8-4.5-8-9V6z" />
      <path d="M9 12l2 2 4-4" />
    </svg>
  )
}

export function Modal({ title, sub, onClose, children, footer, wide }: {
  title: string
  sub?: ReactNode
  onClose(): void
  children: ReactNode
  footer?: ReactNode
  wide?: boolean
}) {
  const { t } = useApp()
  useEffect(() => {
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [onClose])
  return (
    <div className="scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" role="dialog" aria-modal="true" style={wide ? { width: 'min(680px,100%)' } : undefined}>
        <div className="m-h">
          <div>
            <h3>{title}</h3>
            {sub && <p>{sub}</p>}
          </div>
          <button className="x" onClick={onClose} aria-label={t('close')}>
            ×
          </button>
        </div>
        <div className="m-b">{children}</div>
        {footer && <div className="m-f">{footer}</div>}
      </div>
    </div>
  )
}

/** Shows a loading line, an expired-sign-in banner with a re-auth button, or a plain error. */
export function LoadState({ loading, error }: { loading: boolean; error: ApiError | null }) {
  const { t, reauth, goPim } = useApp()
  if (error && (error.status === 403 || error.code === 'Authorization_RequestDenied'))
    return (
      <div className="banner">
        <span>
          <b>{t('noAccessT')}</b> {t('noAccessD')}
        </span>
        <button className="btn sm" onClick={goPim}>
          {t('activateRole')}
        </button>
      </div>
    )
  if (error?.code === 'NEEDS_SIGNIN')
    return (
      <div className="banner">
        <span>{t('needsSignIn')}</span>
        <button className="btn sm" onClick={reauth}>
          {t('signInAgain')}
        </button>
      </div>
    )
  if (error)
    return (
      <div className="banner bad">
        <span>
          <b>{t('error')}:</b> {error.message}
        </span>
      </div>
    )
  if (loading)
    return (
      <p className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="spin" /> {t('loading')}
      </p>
    )
  return null
}

export function CopyButton({ text }: { text: string }) {
  const { t, toast } = useApp()
  return (
    <button type="button" className="btn sm" onClick={() => navigator.clipboard.writeText(text).then(() => toast(t('copied')))}>
      {t('copy')}
    </button>
  )
}

export function RowMenu({ items }: { items: [string, () => void][] }) {
  const { t } = useApp()
  // The menu is position:fixed so the table's overflow container doesn't clip it.
  const [pos, setPos] = useState<{ top: number; left: number } | null>(null)
  const open = !!pos
  const setOpen = (v: boolean) => !v && setPos(null)
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const f = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && setOpen(false)
    document.addEventListener('mousedown', f)
    return () => document.removeEventListener('mousedown', f)
  }, [open])
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button className="btn sm" aria-label={t('more')} title={t('more')} onClick={(e) => {
          if (open) return setPos(null)
          const r = e.currentTarget.getBoundingClientRect()
          setPos({ top: Math.min(r.bottom + 4, window.innerHeight - 190), left: r.right - 220 })
        }}>
        {Icon.dots}
      </button>
      {open && (
        <div className="tenant-menu" style={{ position: 'fixed', top: pos!.top, left: pos!.left, width: 220 }}>
          {items.map(([label, fn]) => (
            <button
              key={label}
              className="tm-item"
              style={{ width: '100%', borderRadius: 7 }}
              onClick={() => {
                setOpen(false)
                fn()
              }}
            >
              {label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
