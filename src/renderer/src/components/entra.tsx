import { useState } from 'react'
import type { DirectoryRef } from '@shared/types'
import { unwrap, useApp, useLoad, type ApiError } from '../ctx'
import { CopyButton, LoadState, Modal } from './ui'
import { useGroups, useUsers } from './teams'

/** Set by the expiry overview so the app registrations screen opens that app right away. */
export const pending: { appId: string | null; spId: string | null } = { appId: null, spId: null }

export const daysLeft = (iso: string | null | undefined) => (iso ? Math.floor((new Date(iso).getTime() - Date.now()) / 86400000) : null)

export function ExpiryPill({ end }: { end: string | null }) {
  const { t, lang } = useApp()
  const d = daysLeft(end)
  if (d === null) return <span className="pill off">{t('enNoEnd')}</span>
  const date = new Date(end!).toLocaleDateString(lang, { day: 'numeric', month: 'short', year: 'numeric' })
  if (d < 0) return <span className="pill bad">{t('enExpired', date)}</span>
  const cls = d <= 30 ? 'bad' : d <= 90 ? 'warn' : 'ok'
  return <span className={`pill ${cls}`} title={date}>{t('enDaysLeft', d)}</span>
}

/**
 * LoadState for Entra screens: the one-time consent for the Entra Graph permissions, the partner sign-in
 * for GDAP data, a missing role, or a missing Entra ID licence.
 */
export function EntraLoad({ loading, error, onRetry, text }: { loading: boolean; error: ApiError | null; onRetry(): void; text?: string }) {
  const { t, tenant, toast } = useApp()
  const [busy, setBusy] = useState(false)
  const tid = tenant!.tenantId
  const act = (fn: () => Promise<unknown>) => async () => {
    setBusy(true)
    try {
      await fn()
      onRetry()
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  if (error?.code === 'NEEDS_CONSENT')
    return (
      <div className="banner">
        <span>{t('enConsentT')}</span>
        <button className="btn sm" disabled={busy} onClick={act(() => unwrap(window.api.entra.consent(tid)))}>
          {busy && <span className="spin" />}
          {t('grantAccess')}
        </button>
      </div>
    )
  if (error?.code === 'NEEDS_HOME_SIGNIN')
    return (
      <div className="banner">
        <span>{t('homeSignIn')}</span>
        <button className="btn sm" disabled={busy} onClick={act(() => unwrap(window.api.reauthHome(tid)))}>
          {t('homeSignInBtn')}
        </button>
      </div>
    )
  if (error && /license|licence|premium|P1|P2/i.test(error.message) && error.status !== 401)
    return (
      <div className="banner">
        <span>
          <b>{t('enLicenseT')}</b> <span className="hint">{error.message}</span>
        </span>
      </div>
    )
  if (loading)
    return (
      <p className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="spin" /> {text || t('loading')}
      </p>
    )
  return <LoadState loading={false} error={error} />
}

/** Input with suggestions for users (and optionally groups) of the tenant; resolves to a directory object. */
export function PrincipalInput({ onPick, groups, label, busy, placeholder }: {
  onPick(p: DirectoryRef): Promise<void> | void
  groups?: boolean
  label?: string
  busy?: boolean
  placeholder?: string
}) {
  const { t, tenant, toast } = useApp()
  const users = useUsers(tenant!.tenantId)
  const gs = useGroups(tenant!.tenantId)
  const [v, setV] = useState('')
  const [id] = useState(() => `pr-${Math.random().toString(36).slice(2, 8)}`)
  const submit = async () => {
    const q = v.trim().toLowerCase()
    if (!q) return
    const u = users.find((x) => x.userPrincipalName.toLowerCase() === q || x.displayName.toLowerCase() === q)
    const g = groups ? gs.find((x) => (x.mail || '').toLowerCase() === q || x.displayName.toLowerCase() === q) : undefined
    const p: DirectoryRef | null = u
      ? { id: u.id, displayName: u.displayName, detail: u.userPrincipalName, type: 'user' }
      : g
        ? { id: g.id, displayName: g.displayName, detail: g.mail, type: 'group' }
        : null
    if (!p) return toast(t('enNotFound', v.trim()))
    await onPick(p)
    setV('')
  }
  return (
    <div className="pw" style={{ alignItems: 'center' }}>
      <input list={id} placeholder={placeholder || (groups ? t('enUserOrGroup') : t('enUserPh'))} value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && submit()} />
      <datalist id={id}>
        {users.slice(0, 1500).map((u) => (
          <option key={u.id} value={u.userPrincipalName}>
            {u.displayName}
          </option>
        ))}
        {groups &&
          gs.slice(0, 1500).map((g) => (
            <option key={g.id} value={g.displayName}>
              {t('enGroupTag')}
            </option>
          ))}
      </datalist>
      <button type="button" className="btn sm" disabled={busy || !v.trim()} onClick={submit}>
        {busy && <span className="spin" />}
        {label || t('add')}
      </button>
    </div>
  )
}

/** A list of chosen users and groups with remove buttons, plus the picker. */
export function PrincipalList({ items, onChange, groups }: { items: DirectoryRef[]; onChange(v: DirectoryRef[]): void; groups?: boolean }) {
  const { t } = useApp()
  return (
    <div style={{ display: 'grid', gap: 6 }}>
      {items.map((p) => (
        <div key={p.id} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <span className="lic-tag">{p.type === 'group' ? t('enGroupTag') : t('enUserTag')}</span>
          <span style={{ flex: 1 }}>
            {p.displayName} {p.detail && <span className="mono">{p.detail}</span>}
          </span>
          <button type="button" className="btn sm danger" onClick={() => onChange(items.filter((x) => x.id !== p.id))}>
            {t('remove')}
          </button>
        </div>
      ))}
      <PrincipalInput groups={groups} onPick={(p) => onChange(items.some((x) => x.id === p.id) ? items : [...items, p])} />
    </div>
  )
}

/** Owners of an app registration or enterprise app. */
export function OwnersPanel({ kind, id }: { kind: 'applications' | 'servicePrincipals'; id: string }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.owners(tid, kind, id), [tid, id])
  const [busy, setBusy] = useState(false)
  const run = async (fn: () => Promise<DirectoryRef[]>) => {
    setBusy(true)
    try {
      const r = await fn()
      list.set(() => r)
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <div style={{ display: 'grid', gap: 10 }}>
      <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
      {list.data && !list.data.length && <p className="hint">{t('enNoOwners')}</p>}
      {list.data?.map((o) => (
        <div key={o.id} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <span style={{ flex: 1 }}>
            <span className="u-n">{o.displayName}</span> <span className="mono">{o.detail}</span>
          </span>
          <button className="btn sm danger" disabled={busy} onClick={() => confirm(t('enOwnerRemoveConfirm', o.displayName)) && run(() => unwrap(window.api.entra.removeOwner(tid, kind, id, o.id)))}>
            {t('remove')}
          </button>
        </div>
      ))}
      <PrincipalInput busy={busy} label={t('enAddOwner')} onPick={(p) => run(() => unwrap(window.api.entra.addOwner(tid, kind, id, p.id)))} />
    </div>
  )
}

export const SECRET_MONTHS = [3, 6, 12, 18, 24]

/** New client secret: description and lifetime, then the value shown once with a copy button. */
export function SecretModal({ kind, objectId, appName, onClose, onDone }: { kind: 'app' | 'sp'; objectId: string; appName: string; onClose(): void; onDone(): void }) {
  const { t, tenant } = useApp()
  const [name, setName] = useState('')
  const [months, setMonths] = useState(12)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const [res, setRes] = useState<{ secret: string; keyId: string; end: string } | null>(null)
  if (res)
    return (
      <Modal title={t('enSecretMade')} sub={t('enSecretOnce')} onClose={onClose} footer={<button className="btn pri" onClick={onClose}>{t('close')}</button>}>
        <div className="creds">
          {[[t('enSecretValue'), res.secret], [t('enSecretId'), res.keyId]].map(([k, v]) => (
            <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <span className="hint" style={{ width: 110 }}>{k}</span>
              <b style={{ flex: 1, wordBreak: 'break-all' }}>{v}</b>
              <CopyButton text={v} />
            </div>
          ))}
        </div>
        <p className="hint">{t('enSecretValid', new Date(res.end).toLocaleDateString())}</p>
      </Modal>
    )
  return (
    <Modal
      title={t('enNewSecret')}
      sub={appName}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                setRes(await unwrap(window.api.entra.addSecret(tenant!.tenantId, kind, objectId, name.trim() || `Toolbox ${new Date().toISOString().slice(0, 10)}`, months)))
                onDone()
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('enCreateSecret')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="sn">{t('enDescription')}</label>
        <input id="sn" value={name} onChange={(e) => setName(e.target.value)} placeholder={t('enSecretNamePh')} />
      </div>
      <div className="fld">
        <label htmlFor="sm">{t('enExpires')}</label>
        <select id="sm" value={months} onChange={(e) => setMonths(+e.target.value)}>
          {SECRET_MONTHS.map((m) => (
            <option key={m} value={m}>
              {t('enMonths', m)}
            </option>
          ))}
        </select>
      </div>
      <p className="hint">{t('enSecretHint')}</p>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

/** Local date-time input value (yyyy-mm-ddThh:mm) to ISO, and back. */
export const toIso = (v: string) => (v ? new Date(v).toISOString() : undefined)
export const localInput = (d: Date) => new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16)

export function Tabs<K extends string>({ tabs, value, onChange }: { tabs: [NoInfer<K>, string][]; value: K; onChange(k: NoInfer<K>): void }) {
  return (
    <div className="seg" role="tablist">
      {tabs.map(([k, label]) => (
        <button key={k} role="tab" aria-selected={k === value} className={k === value ? 'on' : ''} onClick={() => onChange(k)}>
          {label}
        </button>
      ))}
    </div>
  )
}
