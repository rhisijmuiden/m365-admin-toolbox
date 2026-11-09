import { useState } from 'react'
import type { NamedLocation } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { Icon, Modal, RowMenu } from '../components/ui'
import { EntraLoad } from '../components/entra'

type Draft = Omit<NamedLocation, 'modified' | 'id'> & { id?: string }
const empty: Draft = { displayName: '', kind: 'ip', trusted: true, ranges: [], countries: [], includeUnknown: false }

const CIDR4 = /^(\d{1,3}\.){3}\d{1,3}\/\d{1,2}$/
const CIDR6 = /^[0-9a-f:]+\/\d{1,3}$/i

/** Named locations for Conditional Access: trusted office IP ranges and country lists. */
export function NamedLocationsView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.entra.locations(tid), [tid])
  const [edit, setEdit] = useState<Draft | null>(null)
  const rows = list.data || []

  async function remove(l: NamedLocation) {
    if (!confirm(t('nlDeleteConfirm', l.displayName))) return
    try {
      const r = await unwrap(window.api.entra.deleteLocation(tid, l.id))
      list.set(() => r)
      toast(t('nlDeleted'))
    } catch (e) {
      toast((e as Error).message)
      list.reload()
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('entra')}
          </div>
          <h1>{t('locations')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn" onClick={() => setEdit({ ...empty, kind: 'country', trusted: false, displayName: '' })}>
            {t('nlNewCountry')}
          </button>
          <button className="btn pri" onClick={() => setEdit({ ...empty })}>
            {Icon.plus}
            {t('nlNewIp')}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ margin: '0 0 14px', maxWidth: '80ch' }}>
          {t('nlIntro')}
        </p>
        <EntraLoad loading={list.loading && !list.data} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{t('arName')}</th>
                    <th>{t('arType')}</th>
                    <th>{t('nlTrusted')}</th>
                    <th>{t('nlRanges')}</th>
                    <th>{t('caModified')}</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {rows.map((l) => (
                    <tr key={l.id}>
                      <td className="u-n">{l.displayName}</td>
                      <td>{l.kind === 'ip' ? t('nlIp') : t('nlCountry')}</td>
                      <td>{l.kind === 'ip' ? (l.trusted ? <span className="pill ok">{t('yes')}</span> : <span className="pill off">{t('no')}</span>) : '—'}</td>
                      <td className="mono" style={{ whiteSpace: 'normal', maxWidth: 420 }}>
                        {(l.kind === 'ip' ? l.ranges : l.countries).join(', ')}
                      </td>
                      <td className="mono">{fmtDate(l.modified, lang, '—')}</td>
                      <td style={{ width: 40 }}>
                        <RowMenu
                          items={[
                            [t('edit'), () => setEdit({ ...l })],
                            [t('remove'), () => remove(l)]
                          ]}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!rows.length && (
                <p className="hint" style={{ padding: 16 }}>
                  {t('nlNone')}
                </p>
              )}
            </div>
          </div>
        )}
      </div>
      {edit && (
        <LocationModal
          initial={edit}
          onClose={() => setEdit(null)}
          onSave={async (d) => {
            const r = await unwrap(window.api.entra.saveLocation(tid, d))
            list.set(() => r)
            setEdit(null)
            toast(t('saved'))
          }}
        />
      )}
    </>
  )
}

function LocationModal({ initial, onClose, onSave }: { initial: Draft; onClose(): void; onSave(d: Draft): Promise<void> }) {
  const { t } = useApp()
  const [d, setD] = useState(initial)
  const [text, setText] = useState((initial.kind === 'ip' ? initial.ranges : initial.countries).join('\n'))
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const items = text.split(/[\s,;]+/).map((x) => x.trim()).filter(Boolean)
  const bad = d.kind === 'ip' ? items.filter((x) => !CIDR4.test(x) && !CIDR6.test(x)) : items.filter((x) => !/^[A-Za-z]{2}$/.test(x))
  return (
    <Modal
      title={initial.id ? t('nlEdit') : d.kind === 'ip' ? t('nlNewIp') : t('nlNewCountry')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !d.displayName.trim() || !items.length || bad.length > 0}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                await onSave(d.kind === 'ip' ? { ...d, displayName: d.displayName.trim(), ranges: items } : { ...d, displayName: d.displayName.trim(), countries: items.map((x) => x.toUpperCase()) })
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('save')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="ln">{t('arName')}</label>
        <input id="ln" value={d.displayName} onChange={(e) => setD({ ...d, displayName: e.target.value })} placeholder={d.kind === 'ip' ? t('nlNamePh') : ''} />
      </div>
      <div className="fld">
        <label htmlFor="lr">{d.kind === 'ip' ? t('nlRangesLabel') : t('nlCountriesLabel')}</label>
        <textarea id="lr" className="mono" value={text} onChange={(e) => setText(e.target.value)} placeholder={d.kind === 'ip' ? '203.0.113.10/32' : 'NL\nBE'} />
        {bad.length > 0 && <div className="err">{t('nlInvalid', bad.join(', '))}</div>}
      </div>
      {d.kind === 'ip' ? (
        <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input type="checkbox" className="chk" checked={d.trusted} onChange={(e) => setD({ ...d, trusted: e.target.checked })} />
          {t('nlMarkTrusted')}
        </label>
      ) : (
        <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input type="checkbox" className="chk" checked={d.includeUnknown} onChange={(e) => setD({ ...d, includeUnknown: e.target.checked })} />
          {t('nlUnknown')}
        </label>
      )}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
