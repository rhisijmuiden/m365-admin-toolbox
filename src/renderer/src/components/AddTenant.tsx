import { useState } from 'react'
import type { Tenant, TenantMode } from '@shared/types'
import { unwrap, useApp } from '../ctx'
import { Modal } from './ui'

export function AddTenantModal({ onClose, onAdded }: { onClose(): void; onAdded(t: Tenant): void }) {
  const { t } = useApp()
  const [mode, setMode] = useState<TenantMode>('direct')
  const [target, setTarget] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  async function go() {
    setBusy(true)
    setErr(null)
    try {
      onAdded(await unwrap(window.api.addTenant(mode, mode === 'gdap' ? target : undefined)))
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('atTitle')}
      sub={t('atSub')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={busy || (mode === 'gdap' && !target.trim())} onClick={go}>
            {busy && <span className="spin" />}
            {busy ? t('atWaiting') : t('atNext')}
          </button>
        </>
      }
    >
      <label className={`opt ${mode === 'direct' ? 'on' : ''}`}>
        <input type="radio" name="mode" checked={mode === 'direct'} onChange={() => setMode('direct')} />
        <div>
          <b>{t('atDirect')}</b>
          <small>{t('atDirectD')}</small>
        </div>
      </label>
      <label className={`opt ${mode === 'gdap' ? 'on' : ''}`}>
        <input type="radio" name="mode" checked={mode === 'gdap'} onChange={() => setMode('gdap')} />
        <div style={{ flex: 1 }}>
          <b>{t('atGdap')}</b>
          <small>{t('atGdapD')}</small>
          {mode === 'gdap' && (
            <input
              autoFocus
              className="mono"
              style={{ marginTop: 8, width: '100%', padding: '8px 10px', border: '1px solid var(--line)', borderRadius: 8, background: 'var(--surface)' }}
              placeholder="contoso.onmicrosoft.com"
              value={target}
              onChange={(e) => setTarget(e.target.value)}
              aria-label={t('atTarget')}
            />
          )}
        </div>
      </label>
      <p className="hint" style={{ margin: 0 }}>
        {t('consentHint')}
      </p>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
