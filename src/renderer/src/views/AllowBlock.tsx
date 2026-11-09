import { useState } from 'react'
import type { AllowBlockItem, AllowBlockList, NewAllowBlock } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal } from '../components/ui'
import { Tabs } from '../components/exo'
import { entries, Head, SecLoad } from '../components/security'

const LISTS: AllowBlockList[] = ['Sender', 'Url', 'FileHash', 'Spoof']

/** The Tenant Allow/Block List: senders and domains, URLs, file hashes and spoofed senders. */
export function AllowBlockView() {
  const { t, lang, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const list = useLoad(() => window.api.sec.allowBlock(tid), [tid])
  const [tab, setTab] = useState<AllowBlockList>('Sender')
  const [q, setQ] = useState('')
  const [sel, setSel] = useState<Set<string>>(new Set())
  const [adding, setAdding] = useState(false)
  const [busy, setBusy] = useState(false)
  const rows = (list.data || []).filter((i) => i.listType === tab && (i.value + i.notes).toLowerCase().includes(q.toLowerCase()))
  const chosen = rows.filter((i) => sel.has(i.id))

  async function remove() {
    if (!confirm(t('seAbRemoveConfirm', chosen.length))) return
    setBusy(true)
    try {
      const r = await unwrap(window.api.sec.removeAllowBlock(tid, tab, chosen.map((i) => i.id)))
      list.set(() => r)
      setSel(new Set())
      toast(t('seAbRemoved', chosen.length))
    } catch (e) {
      toast((e as Error).message)
      list.reload()
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <Head title={t('allowBlock')}>
        <label className="search">
          {Icon.search}
          <input placeholder={t('seAbSearch')} value={q} onChange={(e) => setQ(e.target.value)} />
        </label>
        <button className="btn" onClick={list.reload} title={t('refresh')} aria-label={t('refresh')}>
          {Icon.refresh}
        </button>
        <button className="btn pri" onClick={() => setAdding(true)}>
          {Icon.plus}
          {t('seAbAdd')}
        </button>
      </Head>
      <div className="content">
        <p className="hint" style={{ maxWidth: 760, marginTop: 0 }}>
          {t('seAbIntro')}
        </p>
        <Tabs
          tabs={LISTS.map((l) => [l, `${t(`seAb_${l}`)} (${(list.data || []).filter((i) => i.listType === l).length})`] as [AllowBlockList, string])}
          value={tab}
          onChange={(k) => (setTab(k), setSel(new Set()))}
        />
        <SecLoad loading={list.loading} error={list.error} onRetry={list.reload} />
        {list.data && (
          <div className="panel">
            <div className="panel-h">
              <h2>{t(`seAb_${tab}`)}</h2>
              <div className="r">
                <button className="btn sm danger" disabled={!chosen.length || busy} onClick={remove}>
                  {busy && <span className="spin" />}
                  {t('remove')} {chosen.length > 0 && `(${chosen.length})`}
                </button>
              </div>
            </div>
            <div className="tbl-wrap">
              <table>
                <thead>
                  <tr>
                    <th style={{ width: 30 }} />
                    <th>{t('seAbEntry')}</th>
                    <th>{t('seAbAction')}</th>
                    <th>{t('seExpires')}</th>
                    <th>{t('seAbNotes')}</th>
                    <th>{t('seAbChanged')}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((i: AllowBlockItem) => (
                    <tr key={i.id}>
                      <td>
                        <input
                          type="checkbox"
                          className="chk"
                          aria-label={i.value}
                          checked={sel.has(i.id)}
                          onChange={(e) => {
                            const s = new Set(sel)
                            e.target.checked ? s.add(i.id) : s.delete(i.id)
                            setSel(s)
                          }}
                        />
                      </td>
                      <td className="mono" style={{ wordBreak: 'break-all', whiteSpace: 'normal' }}>
                        {i.value}
                      </td>
                      <td>{i.action === 'Block' ? <span className="pill bad">{t('seBlock')}</span> : <span className="pill ok">{t('seAllow')}</span>}</td>
                      <td className="mono">{i.expires ? new Date(i.expires).toLocaleDateString(lang) : t('seNever')}</td>
                      <td className="hint" style={{ whiteSpace: 'normal' }}>
                        {i.notes}
                      </td>
                      <td className="mono">{i.modified ? new Date(i.modified).toLocaleDateString(lang) : '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!rows.length && <p className="hint" style={{ padding: 16 }}>{t('seAbEmpty')}</p>}
            </div>
          </div>
        )}
      </div>
      {adding && (
        <AddModal
          initial={tab === 'Spoof' ? 'Sender' : tab}
          onClose={() => setAdding(false)}
          onDone={(r, lt) => {
            list.set(() => r)
            setTab(lt)
            setAdding(false)
          }}
        />
      )}
    </>
  )
}

function AddModal({ initial, onClose, onDone }: { initial: NewAllowBlock['listType']; onClose(): void; onDone(r: AllowBlockItem[], lt: AllowBlockList): void }) {
  const { t, tenant, toast } = useApp()
  const [n, setN] = useState<NewAllowBlock>({ listType: initial, entries: [], action: 'Block', days: null, notes: '' })
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const list = entries(text)
  const days = n.action === 'Allow' ? [null, 7, 30] : [null, 7, 30, 90]
  return (
    <Modal
      title={t('seAbAdd')}
      sub={t('seAbAddSub')}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button
            className="btn pri"
            disabled={busy || !list.length}
            onClick={async () => {
              setBusy(true)
              setErr(null)
              try {
                const r = await unwrap(window.api.sec.newAllowBlock(tenant!.tenantId, { ...n, entries: list }))
                toast(t('seAbAdded', list.length))
                onDone(r, n.listType)
              } catch (e) {
                setErr((e as Error).message)
              } finally {
                setBusy(false)
              }
            }}
          >
            {busy && <span className="spin" />}
            {n.action === 'Block' ? t('seBlock') : t('seAllow')} {list.length > 0 && `(${list.length})`}
          </button>
        </>
      }
    >
      <div className="fld">
        <label>{t('seAbType')}</label>
        <div className="seg">
          {(['Sender', 'Url', 'FileHash'] as const).map((l) => (
            <button key={l} className={n.listType === l ? 'on' : ''} onClick={() => setN({ ...n, listType: l })}>
              {t(`seAb_${l}`)}
            </button>
          ))}
        </div>
      </div>
      <div className="fld">
        <label>{t('seAbAction')}</label>
        <div className="seg">
          {(['Block', 'Allow'] as const).map((a) => (
            <button key={a} className={n.action === a ? 'on' : ''} onClick={() => setN({ ...n, action: a, days: null })}>
              {a === 'Block' ? t('seBlock') : t('seAllow')}
            </button>
          ))}
        </div>
      </div>
      <div className="fld">
        <label htmlFor="abe">{t(`seAbEntries_${n.listType}`)}</label>
        <textarea id="abe" className="mono" rows={5} value={text} onChange={(e) => setText(e.target.value)} placeholder={t(`seAbPh_${n.listType}`)} />
      </div>
      <div className="row2">
        <div className="fld">
          <label htmlFor="abd">{t('seExpires')}</label>
          <select id="abd" value={n.days ?? ''} onChange={(e) => setN({ ...n, days: e.target.value ? +e.target.value : null })}>
            {days.map((d) => (
              <option key={d ?? 'x'} value={d ?? ''}>
                {d ? t('seAfterDays', d) : n.action === 'Allow' ? t('seAllowDefault') : t('seNever')}
              </option>
            ))}
          </select>
        </div>
        <div className="fld">
          <label htmlFor="abn">{t('seAbNotes')}</label>
          <input id="abn" value={n.notes} onChange={(e) => setN({ ...n, notes: e.target.value })} />
        </div>
      </div>
      {n.action === 'Allow' && <p className="hint">{t('seAbAllowHint')}</p>}
      {err && <div className="err">{err}</div>}
    </Modal>
  )
}
