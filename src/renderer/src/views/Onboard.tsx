import { useEffect, useState } from 'react'
import type { Group, Mailbox, Sku, Team } from '@shared/types'
import { genPassword, unwrap, useApp, type ApiError } from '../ctx'
import { CopyButton } from '../components/ui'
import { useGroups, useUsers } from '../components/teams'
import { errText, fixable, fixTenant, StepList, TenantScope, TenantSelect, useSteps, type Step } from '../components/cross'
import { LicenseOptions, LOCATIONS, slug } from './Users'

/** New employee: account, licenses, groups, Teams and mailbox access in one go, in any tenant. */
export function OnboardView() {
  const { t, tenant, tenants } = useApp()
  const [tid, setTid] = useState(tenant!.tenantId)
  const x = tenants.find((y) => y.tenantId === tid) || tenant!
  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {t('enAllTenants', tenants.length)}
          </div>
          <h1>{t('onboard')}</h1>
        </div>
        <div className="top-r">
          <label className="hint" htmlFor="obT">
            {t('scTenant')}
          </label>
          <TenantSelect id="obT" value={tid} onChange={setTid} />
        </div>
      </header>
      <TenantScope tenant={x}>
        <OnboardForm key={tid} />
      </TenantScope>
    </>
  )
}

type Picked<T> = Map<string, T>
const toggle = <T,>(m: Picked<T>, id: string, v: T) => {
  const n = new Map(m)
  n.has(id) ? n.delete(id) : n.set(id, v)
  return n
}

function OnboardForm() {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const users = useUsers(tid)
  const groups = useGroups(tid)
  const [skus, setSkus] = useState<Sku[] | null>(null)
  const [skuErr, setSkuErr] = useState<ApiError | null>(null)

  const [first, setFirst] = useState('')
  const [last, setLast] = useState('')
  const [display, setDisplay] = useState('')
  const [uname, setUname] = useState('')
  const [touched, setTouched] = useState({ display: false, uname: false })
  const [domain, setDomain] = useState(tenant!.defaultDomain)
  const [dept, setDept] = useState('')
  const [loc, setLoc] = useState('NL')
  const [pw, setPw] = useState(genPassword())
  const [force, setForce] = useState(true)
  const [lics, setLics] = useState<Set<string>>(new Set())
  const [grp, setGrp] = useState<Picked<Group>>(new Map())
  const [gq, setGq] = useState('')
  const [teams, setTeams] = useState<Team[] | null>(null)
  const [teamErr, setTeamErr] = useState<ApiError | null>(null)
  const [tms, setTms] = useState<Picked<Team>>(new Map())
  const [mbx, setMbx] = useState<Mailbox[] | null>(null)
  const [mbxErr, setMbxErr] = useState<ApiError | null>(null)
  const [shared, setShared] = useState<Picked<Mailbox>>(new Map())
  const [sendAs, setSendAs] = useState(false)
  const [tap, setTap] = useState(false)
  const [copyFrom, setCopyFrom] = useState('')
  const [copying, setCopying] = useState(false)
  const [loadingExtra, setLoadingExtra] = useState<'teams' | 'mbx' | null>(null)

  const [created, setCreated] = useState<{ id: string; upn: string } | null>(null)
  const [tapPass, setTapPass] = useState<string | null>(null)
  const steps = useSteps()

  useEffect(() => {
    unwrap(window.api.listSkus(tid)).then(setSkus, setSkuErr)
  }, [tid])

  const upn = `${uname}@${domain}`
  const upd = (f: string, l: string) => {
    if (!touched.display) setDisplay(`${f} ${l}`.trim())
    if (!touched.uname) setUname([slug(f), slug(l)].filter(Boolean).join('.'))
  }

  async function loadTeams() {
    setLoadingExtra('teams')
    setTeamErr(null)
    try {
      setTeams(await unwrap(window.api.teams.list(tid)))
    } catch (e) {
      setTeamErr(e as ApiError)
    } finally {
      setLoadingExtra(null)
    }
  }
  async function loadMbx() {
    setLoadingExtra('mbx')
    setMbxErr(null)
    try {
      setMbx((await unwrap(window.api.exo.mailboxes(tid))).mailboxes.filter((m) => m.type === 'SharedMailbox'))
    } catch (e) {
      setMbxErr(e as ApiError)
    } finally {
      setLoadingExtra(null)
    }
  }
  async function fix(e: ApiError, again: () => void, consent?: () => Promise<unknown>) {
    try {
      await fixTenant(tenant!, e, consent)
      again()
    } catch (err) {
      toast((err as Error).message)
    }
  }

  /** Takes over the licenses and groups of an existing colleague. */
  async function copy() {
    setCopying(true)
    try {
      const u = await unwrap(window.api.cross.findUser(tid, copyFrom.trim()))
      const g = await unwrap(window.api.cross.userGroups(tid, u.id))
      setLics(new Set(u.skuIds.filter((id) => skus?.some((s) => s.skuId === id))))
      setGrp(new Map(g.filter((x) => !x.readOnly).map((x) => [x.id, x])))
      toast(t('obCopied', u.displayName))
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setCopying(false)
    }
  }

  // The user id is known only after the first step, so later steps read it from this ref-like holder.
  const ctx = { id: created?.id || '', upn: created?.upn || upn }
  const plan: Step[] = [
    {
      key: 'user',
      label: t('obStUser', upn),
      run: async () => {
        if (ctx.id) return t('obExists')
        const u = await unwrap(
          window.api.createUser(tid, {
            givenName: first.trim(),
            surname: last.trim(),
            displayName: display.trim(),
            mailNickname: uname.replace(/\./g, '') || slug(display),
            userPrincipalName: upn,
            password: pw,
            forceChangePasswordNextSignIn: force,
            usageLocation: loc,
            department: dept.trim() || undefined,
            skuIds: [...lics]
          })
        )
        ctx.id = u.id
        ctx.upn = u.userPrincipalName
        setCreated({ id: u.id, upn: u.userPrincipalName })
        return lics.size ? t('obLicN', lics.size) : undefined
      }
    }
  ]
  if (grp.size)
    plan.push({
      key: 'groups',
      label: t('obStGroups', grp.size),
      run: async () => {
        const r = await unwrap(window.api.cross.addToGroups(tid, ctx.id, ctx.upn, [...grp.values()]))
        if (r.failed.length) throw new Error(r.failed.map((f) => `${f.id}: ${f.error}`).join(' · '))
      }
    })
  for (const tm of tms.values())
    plan.push({
      key: `team-${tm.id}`,
      label: t('obStTeam', tm.displayName),
      run: async () => void (await unwrap(window.api.teams.addMember(tid, tm.id, ctx.id, false)))
    })
  for (const m of shared.values())
    plan.push({
      key: `mbx-${m.id}`,
      label: t('obStMbx', m.displayName),
      run: async () => {
        await unwrap(window.api.exo.grant(tid, m.id, ctx.upn, 'FullAccess', true))
        if (sendAs) await unwrap(window.api.exo.grant(tid, m.id, ctx.upn, 'SendAs'))
      }
    })
  if (tap)
    plan.push({
      key: 'tap',
      label: t('obStTap'),
      run: async () => {
        const r = await unwrap(window.api.createTap(tid, ctx.id, 480, false))
        setTapPass(r.pass)
      }
    })

  const started = Object.keys(steps.status).length > 0
  const failed = plan.filter((s) => steps.status[s.key]?.state === 'fail')
  const valid = display.trim() && /^[A-Za-z0-9._-]+$/.test(uname) && pw.length >= 8

  const gShown = groups.filter((g) => g.displayName.toLowerCase().includes(gq.toLowerCase())).slice(0, 60)

  if (started)
    return (
      <div className="content stack">
        <div className="panel" style={{ padding: 16, display: 'grid', gap: 12 }}>
          <b>{t('obRunT', display, tenant!.name)}</b>
          <StepList steps={plan} status={steps.status} />
          {!steps.running && failed.length > 0 && (
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <button className="btn" onClick={() => steps.run(plan, (s) => steps.status[s.key]?.state === 'fail')}>
                {t('obRetry', failed.length)}
              </button>
              <span className="hint">{t('obRetryHint')}</span>
            </div>
          )}
          {created && (
            <div className="creds">
              {([[t('uname'), created.upn], [t('pw'), pw], ...(tapPass ? [[t('obTap'), tapPass]] : [])] as [string, string][]).map(([k, v]) => (
                <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span className="hint" style={{ width: 170 }}>{k}</span>
                  <b style={{ flex: 1, wordBreak: 'break-all' }}>{v}</b>
                  <CopyButton text={v} />
                </div>
              ))}
            </div>
          )}
          {!steps.running && (
            <div>
              <button
                className="btn pri"
                onClick={() => {
                  steps.reset()
                  setCreated(null)
                  setTapPass(null)
                  setFirst('')
                  setLast('')
                  setDisplay('')
                  setUname('')
                  setTouched({ display: false, uname: false })
                  setPw(genPassword())
                }}
              >
                {t('obNext')}
              </button>
            </div>
          )}
        </div>
      </div>
    )

  return (
    <div className="content stack">
      <div className="row2" style={{ alignItems: 'start' }}>
        <div className="panel" style={{ padding: 16, display: 'grid', gap: 12 }}>
          <b>{t('obAccount')}</b>
          <div className="row2">
            <div className="fld">
              <label htmlFor="obF">{t('first')}</label>
              <input id="obF" autoFocus value={first} onChange={(e) => (setFirst(e.target.value), upd(e.target.value, last))} />
            </div>
            <div className="fld">
              <label htmlFor="obL">{t('last')}</label>
              <input id="obL" value={last} onChange={(e) => (setLast(e.target.value), upd(first, e.target.value))} />
            </div>
          </div>
          <div className="row2">
            <div className="fld">
              <label htmlFor="obD">{t('display')}</label>
              <input id="obD" value={display} onChange={(e) => (setDisplay(e.target.value), setTouched({ ...touched, display: true }))} />
            </div>
            <div className="fld">
              <label htmlFor="obDept">{t('dept')}</label>
              <input id="obDept" value={dept} onChange={(e) => setDept(e.target.value)} />
            </div>
          </div>
          <div className="row2">
            <div className="fld">
              <label htmlFor="obU">{t('uname')}</label>
              <input id="obU" className="mono" value={uname} onChange={(e) => (setUname(e.target.value), setTouched({ ...touched, uname: true }))} />
            </div>
            <div className="fld">
              <label htmlFor="obDom">{t('domain')}</label>
              <select id="obDom" value={domain} onChange={(e) => setDomain(e.target.value)}>
                {tenant!.domains.map((d) => (
                  <option key={d} value={d}>
                    @{d}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div className="fld">
            <label htmlFor="obPw">{t('pw')}</label>
            <div className="pw">
              <input id="obPw" value={pw} onChange={(e) => setPw(e.target.value)} />
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
            <label htmlFor="obLoc">{t('location')}</label>
            <select id="obLoc" value={loc} onChange={(e) => setLoc(e.target.value)}>
              {LOCATIONS.map((l) => (
                <option key={l} value={l}>
                  {new Intl.DisplayNames([document.documentElement.lang], { type: 'region' }).of(l)} ({l})
                </option>
              ))}
            </select>
          </div>
          <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <input type="checkbox" className="chk" checked={tap} onChange={(e) => setTap(e.target.checked)} />
            {t('obTapOpt')}
          </label>
        </div>

        <div style={{ display: 'grid', gap: 12 }}>
          <div className="panel" style={{ padding: 16, display: 'grid', gap: 10 }}>
            <b>{t('obCopyT')}</b>
            <span className="hint">{t('obCopyD')}</span>
            <div className="pw">
              <input list="obUsers" className="mono" placeholder={t('obCopyPh')} value={copyFrom} onChange={(e) => setCopyFrom(e.target.value)} />
              <datalist id="obUsers">
                {users.slice(0, 1000).map((u) => (
                  <option key={u.id} value={u.userPrincipalName}>
                    {u.displayName}
                  </option>
                ))}
              </datalist>
              <button className="btn sm" disabled={!copyFrom.trim() || copying || !skus} onClick={copy}>
                {copying && <span className="spin" />}
                {t('obCopy')}
              </button>
            </div>
          </div>

          <div className="panel" style={{ padding: 16, display: 'grid', gap: 10 }}>
            <b>{t('lics')}</b>
            {skuErr ? <span className="err">{errText(t, skuErr)}</span> : skus ? (
              <LicenseOptions skus={skus} checked={lics} onToggle={(id) => setLics((s) => (s.has(id) ? new Set([...s].filter((y) => y !== id)) : new Set([...s, id])))} />
            ) : (
              <span className="spin" />
            )}
          </div>

          <div className="panel" style={{ padding: 16, display: 'grid', gap: 10 }}>
            <b>
              {t('groups')} {grp.size > 0 && <span className="hint">({grp.size})</span>}
            </b>
            <input placeholder={t('obGroupSearch')} value={gq} onChange={(e) => setGq(e.target.value)} />
            <div style={{ display: 'grid', gap: 4, maxHeight: 220, overflow: 'auto' }}>
              {[...grp.values(), ...gShown.filter((g) => !grp.has(g.id))].map((g) => (
                <label key={g.id} className={`opt ${grp.has(g.id) ? 'on' : ''}`} style={{ padding: '6px 10px' }}>
                  <input type="checkbox" className="chk" checked={grp.has(g.id)} onChange={() => setGrp((m) => toggle(m, g.id, g))} />
                  <div>
                    {g.displayName}
                    <small>{t(`kind_${g.kind}` as 'kind_m365')}</small>
                  </div>
                </label>
              ))}
            </div>
          </div>

          <div className="panel" style={{ padding: 16, display: 'grid', gap: 10 }}>
            <b>Teams</b>
            {teams ? (
              <div style={{ display: 'grid', gap: 4, maxHeight: 200, overflow: 'auto' }}>
                {teams
                  .filter((tm) => tm.archived !== true)
                  .map((tm) => (
                    <label key={tm.id} className={`opt ${tms.has(tm.id) ? 'on' : ''}`} style={{ padding: '6px 10px' }}>
                      <input type="checkbox" className="chk" checked={tms.has(tm.id)} onChange={() => setTms((m) => toggle(m, tm.id, tm))} />
                      <div>{tm.displayName}</div>
                    </label>
                  ))}
              </div>
            ) : teamErr ? (
              <div className="banner">
                <span>{errText(t, teamErr)}</span>
                {fixable(teamErr) && (
                  <button className="btn sm" onClick={() => fix(teamErr, loadTeams, () => unwrap(window.api.teams.consent(tid)))}>
                    {teamErr.code === 'NEEDS_CONSENT' ? t('grantAccess') : t('signInAgain')}
                  </button>
                )}
              </div>
            ) : (
              <button className="btn sm" style={{ justifySelf: 'start' }} disabled={loadingExtra === 'teams'} onClick={loadTeams}>
                {loadingExtra === 'teams' && <span className="spin" />}
                {t('obLoadTeams')}
              </button>
            )}
          </div>

          <div className="panel" style={{ padding: 16, display: 'grid', gap: 10 }}>
            <b>{t('shared')}</b>
            {mbx ? (
              <>
                <div style={{ display: 'grid', gap: 4, maxHeight: 200, overflow: 'auto' }}>
                  {mbx.map((m) => (
                    <label key={m.id} className={`opt ${shared.has(m.id) ? 'on' : ''}`} style={{ padding: '6px 10px' }}>
                      <input type="checkbox" className="chk" checked={shared.has(m.id)} onChange={() => setShared((s) => toggle(s, m.id, m))} />
                      <div>
                        {m.displayName}
                        <small className="mono">{m.primarySmtp}</small>
                      </div>
                    </label>
                  ))}
                  {!mbx.length && <span className="hint">{t('obNoShared')}</span>}
                </div>
                <label className="hint" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" className="chk" checked={sendAs} onChange={(e) => setSendAs(e.target.checked)} />
                  {t('obSendAs')}
                </label>
              </>
            ) : mbxErr ? (
              <div className="banner">
                <span>{errText(t, mbxErr)}</span>
                {fixable(mbxErr) && (
                  <button className="btn sm" onClick={() => fix(mbxErr, loadMbx)}>
                    {mbxErr.code === 'NEEDS_EXO_SIGNIN' ? t('exoSignInBtn') : t('signInAgain')}
                  </button>
                )}
              </div>
            ) : (
              <button className="btn sm" style={{ justifySelf: 'start' }} disabled={loadingExtra === 'mbx'} onClick={loadMbx}>
                {loadingExtra === 'mbx' && <span className="spin" />}
                {loadingExtra === 'mbx' ? t('exoConnecting') : t('obLoadShared')}
              </button>
            )}
          </div>
        </div>
      </div>
      <div className="panel" style={{ padding: 14, display: 'flex', gap: 12, alignItems: 'center', justifyContent: 'space-between' }}>
        <span className="hint">{t('obSummary', plan.length)}</span>
        <button className="btn pri" disabled={!valid} onClick={() => steps.run(plan)}>
          {t('obRun')}
        </button>
      </div>
    </div>
  )
}
