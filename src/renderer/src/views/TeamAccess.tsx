import { useEffect, useState } from 'react'
import type { TeamsFederation, TeamsGuestSettings } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import type { Key } from '../i18n'
import { Icon } from '../components/ui'
import { Chips } from '../components/exo'
import { TeamsLoad } from '../components/teams'

const DOMAIN = /^(\*\.)?[a-z0-9-]+(\.[a-z0-9-]+)+$/i

function Toggle({ label, hint, checked, onChange, disabled }: { label: string; hint?: string; checked: boolean; onChange(v: boolean): void; disabled?: boolean }) {
  return (
    <label className={`opt ${checked ? 'on' : ''}`}>
      <input type="checkbox" className="chk" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <div>
        <b>{label}</b>
        {hint && <small>{hint}</small>}
      </div>
    </label>
  )
}

function DomainList({ items, onChange, placeholder, empty }: { items: string[]; onChange(v: string[]): void; placeholder: string; empty: string }) {
  const { t } = useApp()
  const [v, setV] = useState('')
  const ok = DOMAIN.test(v.trim())
  const add = () => ok && (onChange([...new Set([...items, v.trim().toLowerCase()])]), setV(''))
  return (
    <>
      <Chips items={items.map((d) => ({ key: d, label: d }))} onRemove={(k) => onChange(items.filter((x) => x !== k))} empty={empty} />
      <div className="pw" style={{ alignItems: 'center' }}>
        <input className="mono" placeholder={placeholder} value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && add()} />
        <button className="btn sm" disabled={!ok} onClick={add}>
          {t('add')}
        </button>
      </div>
    </>
  )
}

/** Tenant-wide external access (federation, Teams personal accounts) and guest access settings. */
export function TeamAccessView() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const data = useLoad(() => window.api.teams.access(tid), [tid])
  const [fed, setFed] = useState<TeamsFederation | null>(null)
  const [guest, setGuest] = useState<TeamsGuestSettings | null>(null)
  const [busy, setBusy] = useState<string | null>(null)

  useEffect(() => {
    if (data.data) {
      setFed(data.data.federation)
      setGuest(data.data.guest)
    }
  }, [data.data])

  async function save(key: 'fed' | 'guest') {
    setBusy(key)
    try {
      // Guest settings: only what changed, so an unknown current value is never written back.
      const changed = Object.fromEntries(Object.entries(guest!).filter(([k, v]) => data.data!.guest[k as keyof TeamsGuestSettings] !== v))
      const r = await unwrap(key === 'fed' ? window.api.teams.setFederation(tid, fed!) : window.api.teams.setGuest(tid, changed))
      data.set(() => r)
      toast(t('dlSaved'))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  const fedDirty = JSON.stringify(fed) !== JSON.stringify(data.data?.federation)
  const guestDirty = JSON.stringify(guest) !== JSON.stringify(data.data?.guest)
  const fedInvalid = !!fed && fed.allowFederated && !fed.allowAll && !fed.allowed.length

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('teams')}
          </div>
          <h1>{t('teamAccess')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={data.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
        </div>
      </header>
      <div className="content">
        <TeamsLoad loading={data.loading && !data.data} error={data.error} onRetry={data.reload} />
        {fed && guest && (
          <div style={{ display: 'grid', gap: 18, gridTemplateColumns: 'repeat(auto-fit, minmax(380px, 1fr))', alignItems: 'start' }}>
            <div className="panel" style={{ padding: 16, display: 'grid', gap: 12, marginTop: 0 }}>
              <h3 style={{ margin: 0 }}>{t('taExternal')}</h3>
              <span className="hint">{t('taExternalD')}</span>
              <Toggle label={t('taFederated')} hint={t('taFederatedD')} checked={fed.allowFederated} onChange={(v) => setFed({ ...fed, allowFederated: v })} />
              {fed.allowFederated && (
                <>
                  <div className="seg">
                    <button className={fed.allowAll ? 'on' : ''} onClick={() => setFed({ ...fed, allowAll: true })}>
                      {t('taAllowAll')}
                    </button>
                    <button className={!fed.allowAll ? 'on' : ''} onClick={() => setFed({ ...fed, allowAll: false })}>
                      {t('taAllowList')}
                    </button>
                  </div>
                  {fed.allowAll ? (
                    <div className="fld">
                      <label>{t('taBlocked')}</label>
                      <DomainList items={fed.blocked} onChange={(blocked) => setFed({ ...fed, blocked })} placeholder="contoso.com" empty={t('none')} />
                    </div>
                  ) : (
                    <div className="fld">
                      <label>{t('taAllowed')}</label>
                      <DomainList items={fed.allowed} onChange={(allowed) => setFed({ ...fed, allowed })} placeholder="contoso.com" empty={t('none')} />
                      {fedInvalid && <span className="err">{t('taNeedDomain')}</span>}
                    </div>
                  )}
                </>
              )}
              <Toggle label={t('taConsumer')} hint={t('taConsumerD')} checked={fed.consumer} onChange={(v) => setFed({ ...fed, consumer: v })} />
              {fed.consumer && <Toggle label={t('taConsumerIn')} checked={fed.consumerInbound} onChange={(v) => setFed({ ...fed, consumerInbound: v })} />}
              <div style={{ display: 'flex', gap: 8 }}>
                <button className="btn pri" disabled={!fedDirty || fedInvalid || !!busy} onClick={() => save('fed')}>
                  {busy === 'fed' && <span className="spin" />}
                  {t('save')}
                </button>
                {fedDirty && (
                  <button className="btn" onClick={() => setFed(data.data!.federation)}>
                    {t('cancel')}
                  </button>
                )}
              </div>
            </div>

            <div className="panel" style={{ padding: 16, display: 'grid', gap: 12, marginTop: 0 }}>
              <h3 style={{ margin: 0 }}>{t('taGuest')}</h3>
              <span className="hint">{t('taGuestD')}</span>
              <Toggle label={t('taAllowGuest')} checked={guest.allowGuest} onChange={(v) => setGuest({ ...guest, allowGuest: v })} />
              {(
                [
                  ['privateCalling', 'taCalling'],
                  ['ipVideo', 'taVideo'],
                  ['meetNow', 'taMeetNow'],
                  ['chat', 'taChat'],
                  ['editMessages', 'taEdit'],
                  ['deleteMessages', 'taDelete'],
                  ['giphy', 'taGiphy']
                ] as [keyof TeamsGuestSettings, Key][]
              ).map(([k, label]) => (
                <Toggle key={k} label={t(label) as string} disabled={!guest.allowGuest} checked={!!guest[k]} onChange={(v) => setGuest({ ...guest, [k]: v })} />
              ))}
              <div className="fld">
                <label>{t('taScreen')}</label>
                <select value={guest.screenSharing} disabled={!guest.allowGuest} onChange={(e) => setGuest({ ...guest, screenSharing: e.target.value })}>
                  {['EntireScreen', 'SingleApplication', 'Disabled'].map((m) => (
                    <option key={m} value={m}>
                      {t(`ss_${m}` as Key)}
                    </option>
                  ))}
                </select>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className="btn pri" disabled={!guestDirty || !!busy} onClick={() => save('guest')}>
                  {busy === 'guest' && <span className="spin" />}
                  {t('save')}
                </button>
                {guestDirty && (
                  <button className="btn" onClick={() => setGuest(data.data!.guest)}>
                    {t('cancel')}
                  </button>
                )}
              </div>
              <span className="hint">{t('taGuestNote')}</span>
            </div>
          </div>
        )}
      </div>
    </>
  )
}
