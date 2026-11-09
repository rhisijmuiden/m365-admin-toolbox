import { useState } from 'react'
import type { TeamTemplate } from '@shared/types'
import { unwrap, useApp, useLoad } from '../ctx'
import { Icon, Modal, RowMenu } from '../components/ui'
import { ChannelEditor } from './Teams'

const newId = () => `tpl-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`

/** The template library: stored on this PC and offered in "New team" for every tenant. */
export function TeamTemplatesView() {
  const { t, tenant, toast } = useApp()
  const list = useLoad(() => window.api.teams.templates(), [])
  const [edit, setEdit] = useState<TeamTemplate | null>(null)
  const [fromTeam, setFromTeam] = useState(false)

  async function save(items: TeamTemplate[], msg: string) {
    try {
      list.set(() => items)
      const r = await unwrap(window.api.teams.saveTemplates(items))
      list.set(() => r)
      toast(msg)
    } catch (e) {
      toast((e as Error).message)
      list.reload()
    }
  }

  const items = list.data || []
  const upsert = (x: TeamTemplate) => save(items.some((y) => y.id === x.id) ? items.map((y) => (y.id === x.id ? x : y)) : [...items, x], t('tplSaved', x.name))

  return (
    <>
      <header className="top">
        <div>
          <div className="crumb">
            {tenant ? `${tenant.name} · ` : ''}
            {t('teams')}
          </div>
          <h1>{t('teamTemplates')}</h1>
        </div>
        <div className="top-r">
          <button className="btn" disabled={!tenant} onClick={() => setFromTeam(true)}>
            {t('tplFromTeam')}
          </button>
          <button className="btn pri" onClick={() => setEdit({ id: newId(), name: '', description: '', visibility: 'private', channels: [] })}>
            {Icon.plus}
            {t('tplNew')}
          </button>
        </div>
      </header>
      <div className="content">
        <p className="hint" style={{ marginTop: 0 }}>
          {t('tplIntro')}
        </p>
        <div className="panel">
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>{t('name')}</th>
                  <th>{t('tmVisibility')}</th>
                  <th>{t('tmChannels')}</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {items.map((x) => (
                  <tr key={x.id}>
                    <td>
                      <button className="linkish" onClick={() => setEdit(x)}>
                        <div className="u-n">{x.name}</div>
                        <div className="hint">{x.description}</div>
                      </button>
                    </td>
                    <td>
                      <span className="lic-tag">{t(x.visibility === 'public' ? 'tmPublic' : 'tmPrivate')}</span>
                    </td>
                    <td>
                      {x.channels.map((c) => (
                        <span key={c.name} className="lic-tag">
                          {c.name}
                          {c.type !== 'standard' && ` · ${t(`ch_${c.type}`)}`}
                        </span>
                      ))}
                    </td>
                    <td style={{ width: 40 }}>
                      <RowMenu
                        items={[
                          [t('details'), () => setEdit(x)],
                          [t('tplDuplicate'), () => setEdit({ ...x, id: newId(), name: `${x.name} (2)`, channels: x.channels.map((c) => ({ ...c })) })],
                          [t('remove'), () => confirm(t('tplDelConfirm', x.name)) && save(items.filter((y) => y.id !== x.id), t('tplDeleted', x.name))]
                        ]}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!items.length && list.data && (
              <p className="hint" style={{ padding: 16 }}>
                {t('tplNone')}
              </p>
            )}
          </div>
        </div>
      </div>
      {edit && (
        <TemplateModal
          tpl={edit}
          onClose={() => setEdit(null)}
          onSave={(x) => {
            upsert(x)
            setEdit(null)
          }}
        />
      )}
      {fromTeam && (
        <FromTeamModal
          onClose={() => setFromTeam(false)}
          onPick={(x) => {
            setFromTeam(false)
            setEdit(x)
          }}
        />
      )}
    </>
  )
}

function TemplateModal({ tpl, onClose, onSave }: { tpl: TeamTemplate; onClose(): void; onSave(t: TeamTemplate): void }) {
  const { t } = useApp()
  const [x, setX] = useState(tpl)
  const channels = x.channels.filter((c) => c.name.trim())
  const names = channels.map((c) => c.name.trim().toLowerCase())
  const dupes = names.length !== new Set(names).size
  return (
    <Modal
      title={tpl.name || t('tplNew')}
      onClose={onClose}
      wide
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={!x.name.trim() || dupes} onClick={() => onSave({ ...x, name: x.name.trim(), channels: channels.map((c) => ({ ...c, name: c.name.trim() })) })}>
            {t('save')}
          </button>
        </>
      }
    >
      <div className="fld">
        <label htmlFor="tp-n">{t('name')}</label>
        <input id="tp-n" autoFocus value={x.name} onChange={(e) => setX({ ...x, name: e.target.value })} />
      </div>
      <div className="fld">
        <label htmlFor="tp-d">{t('ngDesc')}</label>
        <input id="tp-d" value={x.description} onChange={(e) => setX({ ...x, description: e.target.value })} />
      </div>
      <div className="fld">
        <label>{t('tmVisibility')}</label>
        <div className="seg">
          {(['private', 'public'] as const).map((v) => (
            <button key={v} className={x.visibility === v ? 'on' : ''} onClick={() => setX({ ...x, visibility: v })}>
              {t(v === 'public' ? 'tmPublic' : 'tmPrivate')}
            </button>
          ))}
        </div>
      </div>
      <div className="fld">
        <label>{t('tmChannels')}</label>
        <ChannelEditor channels={x.channels} onChange={(c) => setX({ ...x, channels: c })} />
        {dupes && <span className="err">{t('tmDupes')}</span>}
      </div>
    </Modal>
  )
}

/** Starts a template from the channels of an existing team in the current tenant. */
function FromTeamModal({ onClose, onPick }: { onClose(): void; onPick(t: TeamTemplate): void }) {
  const { t, tenant, toast } = useApp()
  const tid = tenant!.tenantId
  const teams = useLoad(() => window.api.teams.list(tid), [tid])
  const [id, setId] = useState('')
  const [busy, setBusy] = useState(false)

  async function pick() {
    const team = teams.data!.find((x) => x.id === id)!
    setBusy(true)
    try {
      const ch = await unwrap(window.api.teams.channels(tid, id))
      onPick({
        id: newId(),
        name: team.displayName,
        description: team.description || '',
        visibility: team.visibility === 'public' ? 'public' : 'private',
        channels: ch.filter((c) => c.displayName !== 'General').map((c) => ({ name: c.displayName, description: c.description || undefined, type: c.type }))
      })
    } catch (e) {
      toast((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Modal
      title={t('tplFromTeam')}
      sub={tenant!.name}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="btn pri" disabled={!id || busy} onClick={pick}>
            {busy && <span className="spin" />}
            {t('tplUse')}
          </button>
        </>
      }
    >
      {!teams.data ? (
        <p className="hint">{teams.error ? teams.error.message : t('loading')}</p>
      ) : (
        <div className="fld">
          <label htmlFor="ft-t">{t('tgTeam')}</label>
          <select id="ft-t" value={id} onChange={(e) => setId(e.target.value)}>
            <option value="">—</option>
            {teams.data.map((x) => (
              <option key={x.id} value={x.id}>
                {x.displayName}
              </option>
            ))}
          </select>
        </div>
      )}
    </Modal>
  )
}
