import { useMemo, useRef, useState } from 'react'
import type { Recipient, Sku, Tenant } from '@shared/types'
import { genPassword, unwrap, useApp, type ApiError } from '../ctx'
import { CopyButton } from '../components/ui'
import { download, errText, listCell, parseCsv, TenantSelect, TenantTag, toCsv } from '../components/cross'
import { slug } from './Users'

type Action = 'users' | 'licenses' | 'aliases'
const ACTIONS: Action[] = ['users', 'licenses', 'aliases']

const COLUMNS: Record<Action, { required: string[]; optional: string[]; example: string[][] }> = {
  users: {
    required: ['displayname', 'userprincipalname'],
    optional: ['givenname', 'surname', 'department', 'usagelocation', 'password', 'licenses', 'tenant'],
    example: [
      ['displayName', 'givenName', 'surname', 'userPrincipalName', 'department', 'usageLocation', 'licenses'],
      ['Jan Jansen', 'Jan', 'Jansen', 'jan.jansen@contoso.nl', 'Sales', 'NL', 'O365_BUSINESS_PREMIUM'],
      ['Eva de Vries', 'Eva', 'de Vries', 'eva.devries@contoso.nl', 'Finance', 'NL', 'O365_BUSINESS_PREMIUM+POWER_BI_PRO']
    ]
  },
  licenses: {
    required: ['userprincipalname'],
    optional: ['add', 'remove', 'tenant'],
    example: [
      ['userPrincipalName', 'add', 'remove'],
      ['jan.jansen@contoso.nl', 'SPE_E3', 'O365_BUSINESS_PREMIUM'],
      ['eva.devries@contoso.nl', 'POWER_BI_PRO', '']
    ]
  },
  aliases: {
    required: ['identity', 'alias'],
    optional: ['primary', 'tenant'],
    example: [
      ['identity', 'alias', 'primary'],
      ['jan.jansen@contoso.nl', 'j.jansen@contoso.nl', 'no'],
      ['info@contoso.nl', 'contact@contoso.nl', 'no']
    ]
  }
}

type RowState = { state: 'todo' | 'run' | 'ok' | 'fail'; note?: string }
type Planned = { n: number; data: Record<string, string>; tenant: Tenant | null; problem: string | null }

const yes = (v: string | undefined) => /^(1|y|yes|j|ja|true|waar)$/i.test((v || '').trim())

/** CSV bulk actions: create users, swap licenses, add aliases; one file may cover several tenants. */
export function BulkView() {
  const { t, tenant, tenants } = useApp()
  const [action, setAction] = useState<Action>('users')
  const [defTid, setDefTid] = useState(tenant!.tenantId)
  const [text, setText] = useState('')
  const [results, setResults] = useState<Record<number, RowState>>({})
  const [running, setRunning] = useState(false)
  const [passwords, setPasswords] = useState<Record<number, string>>({})
  const file = useRef<HTMLInputElement>(null)
  const cols = COLUMNS[action]

  const csv = useMemo(() => (text.trim() ? parseCsv(text) : null), [text])
  const missing = csv ? cols.required.filter((c) => !csv.headers.includes(c)) : []

  /** The tenant of a row: the tenant column, else the domain of the address, else the chosen default. */
  function tenantFor(r: Record<string, string>): Tenant | null {
    const tv = (r.tenant || '').toLowerCase()
    if (tv) return tenants.find((x) => [x.tenantId, x.name, x.defaultDomain, ...x.domains].some((v) => v.toLowerCase() === tv)) || null
    const addr = r.userprincipalname || r.identity || ''
    const dom = addr.split('@')[1]?.toLowerCase()
    return (dom && tenants.find((x) => x.domains.some((d) => d.toLowerCase() === dom))) || tenants.find((x) => x.tenantId === defTid) || null
  }

  const plan: Planned[] = (csv?.rows || []).map((data, i) => {
    const x = tenantFor(data)
    let problem: string | null = null
    if (!x) problem = t('buNoTenant', data.tenant || '')
    else if (cols.required.some((c) => !data[c])) problem = t('buMissing', cols.required.filter((c) => !data[c]).join(', '))
    else if (action === 'users' && !/^[^@\s]+@[^@\s]+$/.test(data.userprincipalname)) problem = t('buBadUpn')
    else if (action === 'licenses' && !data.add && !data.remove) problem = t('buNoLic')
    return { n: i + 1, data, tenant: x, problem }
  })
  const valid = plan.filter((p) => !p.problem)

  async function run() {
    setRunning(true)
    setResults({})
    setPasswords({})
    const skus = new Map<string, Promise<Sku[]>>()
    const recips = new Map<string, Promise<Recipient[]>>()
    const skusOf = (id: string) => (skus.has(id) ? skus : skus.set(id, unwrap(window.api.listSkus(id)))).get(id)!
    const recipsOf = (id: string) => (recips.has(id) ? recips : recips.set(id, unwrap(window.api.exo.recipients(id)))).get(id)!
    const resolve = async (id: string, names: string[]) => {
      const list = await skusOf(id)
      return names.map((nm) => {
        const s = list.find((x) => [x.skuPartNumber, x.name, x.skuId].some((v) => v.toLowerCase() === nm.toLowerCase()))
        if (!s) throw new Error(t('buUnknownSku', nm))
        return s
      })
    }
    const setRow = (n: number, s: RowState) => setResults((r) => ({ ...r, [n]: s }))

    for (const p of valid) {
      const id = p.tenant!.tenantId
      const d = p.data
      setRow(p.n, { state: 'run' })
      try {
        if (action === 'users') {
          const lic = await resolve(id, listCell(d.licenses, csv!.delimiter))
          const pw = d.password || genPassword()
          const upn = d.userprincipalname
          await unwrap(
            window.api.createUser(id, {
              givenName: d.givenname || '',
              surname: d.surname || '',
              displayName: d.displayname,
              mailNickname: slug(upn.split('@')[0]).replace(/\./g, '') || slug(d.displayname),
              userPrincipalName: upn,
              password: pw,
              forceChangePasswordNextSignIn: true,
              usageLocation: (d.usagelocation || 'NL').toUpperCase(),
              department: d.department || undefined,
              skuIds: lic.map((s) => s.skuId)
            })
          )
          setPasswords((x) => ({ ...x, [p.n]: pw }))
          setRow(p.n, { state: 'ok', note: lic.length ? t('obLicN', lic.length) : undefined })
        } else if (action === 'licenses') {
          const add = await resolve(id, listCell(d.add, csv!.delimiter))
          const remove = await resolve(id, listCell(d.remove, csv!.delimiter))
          const r = await unwrap(
            window.api.changeLicenses(id, [{ userId: d.userprincipalname, add: add.map((s) => s.skuId), remove: remove.map((s) => s.skuId) }])
          )
          if (r.failed.length) throw new Error(r.failed[0].error)
          setRow(p.n, { state: 'ok', note: [...add.map((s) => `+${s.name}`), ...remove.map((s) => `−${s.name}`)].join(', ') })
        } else {
          const who = d.identity.toLowerCase()
          const r = (await recipsOf(id)).find(
            (x) => x.primarySmtp.toLowerCase() === who || x.id.toLowerCase() === who || x.addresses.some((a) => a.replace(/^smtp:/i, '').toLowerCase() === who)
          )
          if (!r) throw new Error(t('buNoRecipient', d.identity))
          const primary = yes(d.primary)
          await unwrap(window.api.exo.setAddresses(id, r, { add: [d.alias], primary: primary ? d.alias : undefined }))
          setRow(p.n, { state: 'ok', note: primary ? t('buPrimary') : undefined })
        }
      } catch (e) {
        setRow(p.n, { state: 'fail', note: errText(t, e as ApiError) })
      }
    }
    setRunning(false)
  }

  function loadFile(f: File | undefined) {
    if (!f) return
    const rd = new FileReader()
    rd.onload = () => (setText(String(rd.result || '')), setResults({}))
    rd.readAsText(f)
  }

  const okN = Object.values(results).filter((r) => r.state === 'ok').length
  const failN = Object.values(results).filter((r) => r.state === 'fail').length
  const shownCols = [...cols.required, ...cols.optional.filter((c) => c !== 'tenant' && c !== 'password' && csv?.headers.includes(c))]

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('bulk')}</h1>
        </div>
      </header>
      <div className="content stack">
        <div className="panel" style={{ padding: 16, display: 'grid', gap: 12 }}>
          <div className="seg" style={{ justifySelf: 'start' }}>
            {ACTIONS.map((a) => (
              <button key={a} className={action === a ? 'on' : ''} disabled={running} onClick={() => (setAction(a), setResults({}))}>
                {t(`buA_${a}`)}
              </button>
            ))}
          </div>
          <p className="hint" style={{ margin: 0 }}>
            {t(`buD_${action}`)}
          </p>
          <div className="hint">
            {t('buCols')} <span className="mono">{cols.required.join(', ')}</span> · {t('buOptional')} <span className="mono">{cols.optional.join(', ')}</span>
          </div>
          <div className="row2">
            <div className="fld">
              <label htmlFor="buT">{t('buDefault')}</label>
              <TenantSelect id="buT" value={defTid} onChange={setDefTid} />
            </div>
            <div className="fld">
              <label>{t('buFile')}</label>
              <div className="pw">
                <button className="btn sm" disabled={running} onClick={() => file.current?.click()}>
                  {t('buOpen')}
                </button>
                <button className="btn sm" onClick={() => download(`toolbox-${action}-example.csv`, toCsv(cols.example))}>
                  {t('buExample')}
                </button>
                <input ref={file} type="file" accept=".csv,.txt" hidden onChange={(e) => (loadFile(e.target.files?.[0]), (e.target.value = ''))} />
              </div>
            </div>
          </div>
          <textarea
            rows={6}
            className="mono"
            placeholder={t('buPaste')}
            value={text}
            disabled={running}
            onChange={(e) => (setText(e.target.value), setResults({}))}
            style={{ width: '100%', fontSize: 12 }}
          />
          {missing.length > 0 && <div className="err">{t('buMissingCols', missing.join(', '))}</div>}
        </div>

        {csv && !missing.length && (
          <div className="panel">
            <div className="panel-h">
              <b>{t('buPreview', plan.length, valid.length)}</b>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                {(okN > 0 || failN > 0) && <span className="hint">{t('buResult', okN, failN)}</span>}
                {action === 'users' && Object.keys(passwords).length > 0 && (
                  <CopyButton
                    text={plan
                      .filter((p) => passwords[p.n])
                      .map((p) => `${p.data.userprincipalname};${passwords[p.n]}`)
                      .join('\n')}
                  />
                )}
                <button className="btn pri" disabled={running || !valid.length} onClick={() => confirm(t('buConfirm', valid.length)) && run()}>
                  {running && <span className="spin" />}
                  {t('buRun', valid.length)}
                </button>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th>{t('status')}</th>
                    <th>{t('scTenant')}</th>
                    {shownCols.map((c) => (
                      <th key={c}>{c}</th>
                    ))}
                    {action === 'users' && Object.keys(passwords).length > 0 && <th>{t('pw')}</th>}
                  </tr>
                </thead>
                <tbody>
                  {plan.map((p) => {
                    const r = results[p.n]
                    return (
                      <tr key={p.n}>
                        <td className="hint">{p.n}</td>
                        <td style={{ maxWidth: 320, whiteSpace: 'normal' }}>
                          {p.problem ? (
                            <span className="pill bad">{p.problem}</span>
                          ) : r?.state === 'run' ? (
                            <span className="spin" />
                          ) : r ? (
                            <>
                              <span className={`pill ${r.state === 'ok' ? 'ok' : 'bad'}`}>{t(`crSt_${r.state}`)}</span>
                              {r.note && <div className="hint">{r.note}</div>}
                            </>
                          ) : (
                            <span className="pill off">{t('buReady')}</span>
                          )}
                        </td>
                        <td>{p.tenant ? <TenantTag tenant={p.tenant} /> : '—'}</td>
                        {shownCols.map((c) => (
                          <td key={c} className={c.includes('@') || /principal|identity|alias/.test(c) ? 'mono' : ''}>
                            {p.data[c]}
                          </td>
                        ))}
                        {action === 'users' && Object.keys(passwords).length > 0 && <td className="mono">{passwords[p.n] || ''}</td>}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
        <p className="hint">{t('buHint')}</p>
      </div>
    </>
  )
}
