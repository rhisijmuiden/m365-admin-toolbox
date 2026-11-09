import { useState } from 'react'
import type { Sku, User } from '@shared/types'
import { genPassword, initials, unwrap, useApp, useLoad } from '../ctx'
import { fmtDate } from '../i18n'
import { CopyButton, Icon, LoadState, Modal, RowMenu } from '../components/ui'

const LOCATIONS = ['NL', 'BE', 'DE', 'LU', 'FR', 'GB', 'US']

export function UsersView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const users = useLoad(() => window.api.listUsers(tid), [tid])
  const skus = useLoad(() => window.api.listSkus(tid), [tid])
  const [q, setQ] = useState('')
  const [sel, setSel] = useState<Set<string>>(new Set())
  const [modal, setModal] = useState<
    | { k: 'create' }
    | { k: 'assign' }
    | { k: 'creds'; title: string; lines: [string, string][]; note: string }
    | { k: 'reset'; user: User }
    | { k: 'tap'; user: User }
    | null
  >(null)

  const list = users.data || []
  const skuList = skus.data || []
  const skuById = new Map(skuList.map((s) => [s.skuId, s]))
  const shown = list.filter((u) => (u.displayName + u.userPrincipalName + (u.department || '')).toLowerCase().includes(q.toLowerCase()))
  const unl = list.filter((u) => !u.skuIds.length).length
  const shortSkus = skuList.filter((s) => s.shortage > 0)
  const shortage = shortSkus.reduce((a, s) => a + s.shortage, 0)
  const reload = () => {
    users.reload()
    skus.reload()
  }

  async function act(u: User, kind: 'block' | 'unblock' | 'revoke') {
    if (kind === 'block' && !confirm(t('blockConfirm', u.displayName))) return
    if (kind === 'revoke' && !confirm(t('revokeConfirm', u.displayName))) return
    try {
      if (kind === 'revoke') {
        await unwrap(window.api.revokeSessions(tid, u.id))
        toast(t('revokedT', u.displayName))
        return
      }
      const enabled = kind === 'unblock'
      await unwrap(window.api.setSignIn(tid, u.id, enabled))
      // Blocking alone leaves existing tokens valid for up to an hour, so also revoke sessions.
      if (!enabled) await window.api.revokeSessions(tid, u.id)
      users.set((d) => d.map((x) => (x.id === u.id ? { ...x, accountEnabled: enabled } : x)))
      toast(t(enabled ? 'unblockedT' : 'blockedT', u.displayName))
    } catch (e) {
      toast((e as Error).message)
    }
  }

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant!.name} · {t('m365')}
          </div>
          <h1>{t('users')}</h1>
        </div>
        <div className="top-r">
          <label className="search">
            {Icon.search}
            <input placeholder={t('searchUsers')} value={q} onChange={(e) => setQ(e.target.value)} />
          </label>
          <button className="btn" onClick={reload} title={t('refresh')} aria-label={t('refresh')}>
            {Icon.refresh}
          </button>
          <button className="btn" disabled={!sel.size || !skus.data} onClick={() => setModal({ k: 'assign' })}>
            {t('assign')}
            {sel.size ? ` (${sel.size})` : ''}
          </button>
          <button className="btn pri" disabled={!skus.data} onClick={() => setModal({ k: 'create' })}>
            {Icon.plus}
            {t('newUser')}
          </button>
        </div>
      </header>
      <div className="content">
        <LoadState loading={users.loading && !users.data} error={users.error || skus.error} />
        {users.data && (
          <>
            <div className="tiles">
              <div className="tile">
                <div className="k">{t('totUsers')}</div>
                <div className="v">{list.length}</div>
                <div className="s">{tenant!.defaultDomain}</div>
              </div>
              <div className="tile">
                <div className="k">{t('licensed')}</div>
                <div className="v">{list.length - unl}</div>
                <div className="s">{list.length ? Math.round(((list.length - unl) / list.length) * 100) : 0}%</div>
              </div>
              <div className={`tile ${unl ? 'warn' : ''}`}>
                <div className="k">{t('unlicensed')}</div>
                <div className="v">{unl}</div>
                <div className="s">—</div>
              </div>
              <div className={`tile ${shortage ? 'bad' : ''}`}>
                <div className="k">{t('short')}</div>
                <div className="v">{shortage}</div>
                <div className="s">{shortage ? shortSkus.map((s) => s.name.replace('Microsoft 365 ', '')).join(', ') : t('ok')}</div>
              </div>
            </div>
            <div className="panel">
              <div className="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style={{ width: 36 }}>
                        <input
                          type="checkbox"
                          className="chk"
                          aria-label="Select all"
                          checked={shown.length > 0 && shown.every((u) => sel.has(u.id))}
                          onChange={(e) => setSel(e.target.checked ? new Set(shown.map((u) => u.id)) : new Set())}
                        />
                      </th>
                      <th>{t('name')}</th>
                      <th>{t('dept')}</th>
                      <th>{t('lics')}</th>
                      <th>{t('status')}</th>
                      <th>{t('lastSign')}</th>
                      <th style={{ width: 48 }} aria-label={t('actions')} />
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((u) => (
                      <tr key={u.id}>
                        <td>
                          <input
                            type="checkbox"
                            className="chk"
                            checked={sel.has(u.id)}
                            aria-label={u.displayName}
                            onChange={(e) => {
                              const n = new Set(sel)
                              e.target.checked ? n.add(u.id) : n.delete(u.id)
                              setSel(n)
                            }}
                          />
                        </td>
                        <td>
                          <div className="u">
                            <div className="u-av">{initials(u.displayName)}</div>
                            <div>
                              <div className="u-n">{u.displayName}</div>
                              <div className="mono">{u.userPrincipalName}</div>
                            </div>
                          </div>
                        </td>
                        <td>{u.department || '—'}</td>
                        <td>
                          {u.skuIds.length ? (
                            u.skuIds.map((id) => (
                              <span key={id} className="lic-tag" title={skuById.get(id)?.name}>
                                {skuById.get(id)?.skuPartNumber || id.slice(0, 8)}
                              </span>
                            ))
                          ) : (
                            <span className="pill warn">{t('none')}</span>
                          )}
                        </td>
                        <td>{u.accountEnabled ? <span className="pill ok">{t('active')}</span> : <span className="pill off">{t('blocked')}</span>}</td>
                        <td className="mono">{fmtDate(u.lastSignIn, lang, '—')}</td>
                        <td>
                          <RowMenu
                            items={[
                              u.accountEnabled ? [t('block'), () => act(u, 'block')] : [t('unblock'), () => act(u, 'unblock')],
                              [t('resetPw'), () => setModal({ k: 'reset', user: u })],
                              [t('tap'), () => setModal({ k: 'tap', user: u })],
                              [t('revoke'), () => act(u, 'revoke')]
                            ]}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>

      {modal?.k === 'create' && (
        <CreateUserModal
          skus={skuList}
          onClose={() => setModal(null)}
          onCreated={(u, pw) => {
            users.set((d) => [u, ...d])
            skus.reload()
            setModal({ k: 'creds', title: t('createdT'), lines: [[t('uname'), u.userPrincipalName], [t('pw'), pw]], note: t('createdD') })
          }}
        />
      )}
      {modal?.k === 'assign' && (
        <AssignModal
          users={list.filter((u) => sel.has(u.id))}
          skus={skuList}
          onClose={() => setModal(null)}
          onDone={() => {
            setModal(null)
            setSel(new Set())
            reload()
          }}
        />
      )}
      {modal?.k === 'reset' && (
        <ResetModal
          user={modal.user}
          onClose={() => setModal(null)}
          onDone={(pw) =>
            setModal({ k: 'creds', title: t('resetDone'), lines: [[t('uname'), modal.user.userPrincipalName], [t('pw'), pw]], note: t('createdD') })
          }
        />
      )}
      {modal?.k === 'tap' && (
        <TapModal
          user={modal.user}
          onClose={() => setModal(null)}
          onDone={(pass, minutes) =>
            setModal({ k: 'creds', title: t('tapDone'), lines: [[t('uname'), modal.user.userPrincipalName], ['TAP', pass]], note: t('tapValid', minutes) })
          }
        />
      )}
      {modal?.k === 'creds' && (
        <Modal title={modal.title} sub={modal.note} onClose={() => setModal(null)} footer={<button className="btn pri" onClick={() => setModal(null)}>{t('close')}</button>}>
          <div className="creds">
            {modal.lines.map(([k, v]) => (
              <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span className="hint" style={{ width: 110 }}>{k}</span>
                <b style={{ flex: 1, wordBreak: 'break-all' }}>{v}</b>
                <CopyButton text={v} />
              </div>
            ))}
          </div>
        </Modal>
      )}
    </>
  )
}

function LicenseOptions({ skus, checked, mixed, onToggle }: {
  skus: Sku[]
  checked: Set<string>
  mixed?: Set<string>
  onToggle(id: string): void
}) {
  const { t } = useApp()
  if (!skus.length) return <p className="hint">{t('noSkus')}</p>
  return (
    <div style={{ display: 'grid', gap: 6 }}>
      {skus.map((s) => (
        <label key={s.skuId} className="opt">
          <input
            type="checkbox"
            className="chk"
            checked={checked.has(s.skuId)}
            ref={(el) => {
              if (el) el.indeterminate = !!mixed?.has(s.skuId) && !checked.has(s.skuId)
            }}
            onChange={() => onToggle(s.skuId)}
          />
          <div>
            <b>{s.name}</b>
            <small className="mono">{s.skuPartNumber}</small>
            {mixed?.has(s.skuId) && !checked.has(s.skuId) && <small>{t('mixed')}</small>}
          </div>
          <span className="avail">
            <span className={`pill ${s.available > 0 ? 'ok' : 'bad'}`}>
              {s.available} {t('avail').toLowerCase()}
            </span>
          </span>
        </label>
      ))}
    </div>
  )
}

const slug = (s: string) =>
  s
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9.]+/g, '')

function CreateUserModal({ skus, onClose, onCreated }: { skus: Sku[]; onClose(): void; onCreated(u: User, pw: string): void }) {
  const { t, tenant } = useApp()
  const [first, setFirst] = useState('')
  const [last, setLast] = useState('')
  const [display, setDisplay] = useState('')
  const [uname, setUname] = useState('')
  const [touched, setTouched] = useState({ display: false, uname: false })
  const [domain, setDomain] = useState(tenant!.defaultDomain)
  const [pw, setPw] = useState(genPassword())
  const [force, setForce] = useState(true)
  const [loc, setLoc] = useState('NL')
  const [dept, setDept] = useState('')
  const [lics, setLics] = useState<Set<string>>(new Set())
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  const upd = (f: string, l: string) => {
    if (!touched.display) setDisplay(`${f} ${l}`.trim())
    if (!touched.uname) setUname([slug(f), slug(l)].filter(Boolean).join('.'))
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    setBusy(true)
    setErr(null)
    try {
      const u = await unwrap(
        window.api.createUser(tenant!.tenantId, {
          givenName: first.trim(),
          surname: last.trim(),
          displayName: display.trim(),
          mailNickname: uname.replace(/\./g, '') || slug(display),
          userPrincipalName: `${uname}@${domain}`,
          password: pw,
          forceChangePasswordNextSignIn: force,
          usageLocation: loc,
          department: dept.trim() || undefined,
          skuIds: [...lics]
        })
      )
      onCreated(u, pw)
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal title={t('cuTitle')} sub={t('cuSub', tenant!.name)} onClose={onClose}>
      <form id="cu" onSubmit={submit} style={{ display: 'grid', gap: 12 }}>
        <div className="row2">
          <div className="fld">
            <label htmlFor="cuF">{t('first')}</label>
            <input id="cuF" autoFocus value={first} onChange={(e) => (setFirst(e.target.value), upd(e.target.value, last))} />
          </div>
          <div className="fld">
            <label htmlFor="cuL">{t('last')}</label>
            <input id="cuL" value={last} onChange={(e) => (setLast(e.target.value), upd(first, e.target.value))} />
          </div>
        </div>
        <div className="row2">
          <div className="fld">
            <label htmlFor="cuD">{t('display')}</label>
            <input id="cuD" required value={display} onChange={(e) => (setDisplay(e.target.value), setTouched({ ...touched, display: true }))} />
          </div>
          <div className="fld">
            <label htmlFor="cuDept">{t('dept')}</label>
            <input id="cuDept" value={dept} onChange={(e) => setDept(e.target.value)} />
          </div>
        </div>
        <div className="row2">
          <div className="fld">
            <label htmlFor="cuU">{t('uname')}</label>
            <input
              id="cuU"
              required
              className="mono"
              pattern="[A-Za-z0-9._\-]+"
              value={uname}
              onChange={(e) => (setUname(e.target.value), setTouched({ ...touched, uname: true }))}
            />
          </div>
          <div className="fld">
            <label htmlFor="cuDom">{t('domain')}</label>
            <select id="cuDom" value={domain} onChange={(e) => setDomain(e.target.value)}>
              {tenant!.domains.map((d) => (
                <option key={d} value={d}>
                  @{d}
                </option>
              ))}
            </select>
          </div>
        </div>
        <div className="fld">
          <label htmlFor="cuPw">{t('pw')}</label>
          <div className="pw">
            <input id="cuPw" required minLength={8} value={pw} onChange={(e) => setPw(e.target.value)} />
            <button type="button" className="btn sm" onClick={() => setPw(genPassword())}>
              {t('gen')}
            </button>
            <CopyButton text={pw} />
          </div>
        </div>
        <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input type="checkbox" className="chk" checked={force} onChange={(e) => setForce(e.target.checked)} />
          {t('forceChange')}
        </label>
        <div className="fld">
          <label htmlFor="cuLoc">{t('location')}</label>
          <select id="cuLoc" value={loc} onChange={(e) => setLoc(e.target.value)}>
            {LOCATIONS.map((l) => (
              <option key={l} value={l}>
                {new Intl.DisplayNames([document.documentElement.lang], { type: 'region' }).of(l)} ({l})
              </option>
            ))}
          </select>
        </div>
        <div className="fld">
          <label>{t('licPick')}</label>
          <LicenseOptions
            skus={skus}
            checked={lics}
            onToggle={(id) => {
              const n = new Set(lics)
              n.has(id) ? n.delete(id) : n.add(id)
              setLics(n)
            }}
          />
        </div>
        {err && <div className="err">{err}</div>}
      </form>
      <div className="m-f" style={{ margin: '6px -20px -12px' }}>
        <button type="button" className="btn" onClick={onClose}>
          {t('cancel')}
        </button>
        <button className="btn pri" type="submit" form="cu" disabled={busy}>
          {busy && <span className="spin" />}
          {busy ? t('creating') : t('create')}
        </button>
      </div>
    </Modal>
  )
}

function AssignModal({ users, skus, onClose, onDone }: { users: User[]; skus: Sku[]; onClose(): void; onDone(): void }) {
  const { t, tenant, toast } = useApp()
  const all = new Set(skus.filter((s) => users.every((u) => u.skuIds.includes(s.skuId))).map((s) => s.skuId))
  const some = new Set(skus.filter((s) => users.some((u) => u.skuIds.includes(s.skuId))).map((s) => s.skuId))
  const [checked, setChecked] = useState<Set<string>>(all)
  // A mixed SKU the admin never touched stays as it is per user.
  const [touched, setTouched] = useState<Set<string>>(new Set())
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  async function apply() {
    setBusy(true)
    setErr(null)
    const changes = users.map((u) => ({
      userId: u.id,
      add: [...checked].filter((id) => !u.skuIds.includes(id)),
      remove: skus
        .map((s) => s.skuId)
        .filter((id) => u.skuIds.includes(id) && !checked.has(id) && (all.has(id) || touched.has(id)))
    }))
    try {
      const r = await unwrap(window.api.changeLicenses(tenant!.tenantId, changes))
      if (r.failed.length) {
        const names = new Map(users.map((u) => [u.id, u.displayName]))
        setErr(`${t('partial', r.failed.length)}:\n` + r.failed.map((f) => `${names.get(f.userId)}: ${f.error}`).join('\n'))
      } else {
        toast(t('assignedT'))
        onDone()
      }
    } catch (e) {
      setErr((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('alTitle')}
      sub={t('alSub', users.length)}
      onClose={err ? onDone : onClose}
      footer={
        <>
          <button className="btn" onClick={err ? onDone : onClose}>
            {err ? t('close') : t('cancel')}
          </button>
          <button className="btn pri" disabled={busy} onClick={apply}>
            {busy && <span className="spin" />}
            {busy ? t('applying') : t('apply')}
          </button>
        </>
      }
    >
      <LicenseOptions
        skus={skus}
        checked={checked}
        mixed={new Set([...some].filter((id) => !all.has(id) && !touched.has(id)))}
        onToggle={(id) => {
          const n = new Set(checked)
          n.has(id) ? n.delete(id) : n.add(id)
          setChecked(n)
          setTouched(new Set(touched).add(id))
        }}
      />
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function ResetModal({ user, onClose, onDone }: { user: User; onClose(): void; onDone(pw: string): void }) {
  const { t, tenant } = useApp()
  const [pw, setPw] = useState(genPassword())
  const [force, setForce] = useState(true)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('resetPw')}
      sub={t('resetPwSub', user.displayName)}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || pw.length < 8}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                await unwrap(window.api.resetPassword(tenant!.tenantId, user.id, pw, force))
                onDone(pw)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('resetPw')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="rpw">{t('pw')}</label>
        <div className="pw">
          <input id="rpw" value={pw} onChange={(e) => setPw(e.target.value)} />
          <button type="button" className="btn sm" onClick={() => setPw(genPassword())}>
            {t('gen')}
          </button>
        </div>
      </div>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={force} onChange={(e) => setForce(e.target.checked)} />
        {t('forceChange')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}

function TapModal({ user, onClose, onDone }: { user: User; onClose(): void; onDone(pass: string, minutes: number): void }) {
  const { t, tenant } = useApp()
  const [minutes, setMinutes] = useState(60)
  const [once, setOnce] = useState(true)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  return (
    <Modal
      title={t('tap')}
      sub={t('tapSub', user.displayName)}
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
                const r = await unwrap(window.api.createTap(tenant!.tenantId, user.id, minutes, once))
                onDone(r.pass, r.lifetimeInMinutes)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {t('tapCreate')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="tapm">{t('tapLifetime')}</label>
        <select id="tapm" value={minutes} onChange={(e) => setMinutes(+e.target.value)}>
          {[60, 120, 480, 1440].map((m) => (
            <option key={m} value={m}>
              {t('minutes', m)}
            </option>
          ))}
        </select>
      </div>
      <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input type="checkbox" className="chk" checked={once} onChange={(e) => setOnce(e.target.checked)} />
        {t('tapOnce')}
      </label>
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
