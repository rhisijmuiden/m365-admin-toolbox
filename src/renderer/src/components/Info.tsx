import { useEffect, useState } from 'react'
import type { AppInfo, Channel, Theme, UpdateState } from '@shared/types'
import { unwrap, useApp } from '../ctx'
import { Modal } from './ui'

export function InfoModal({ onClose, logo }: { onClose(): void; logo: string }) {
  const { t, settings, setSettings, toast, lang } = useApp()
  const [info, setInfo] = useState<AppInfo | null>(null)
  const [upd, setUpd] = useState<UpdateState>({ status: 'idle' })
  const [clientId, setClientId] = useState(settings.clientId)

  useEffect(() => {
    unwrap(window.api.appInfo()).then(setInfo)
    unwrap(window.api.updateState()).then(setUpd)
    return window.api.onUpdate(setUpd)
  }, [])

  const check = async () => {
    setUpd({ status: 'checking' })
    const r = await window.api.checkUpdate()
    if (!r.ok) setUpd({ status: 'error', error: r.error })
  }

  const switchChannel = async (c: Channel) => {
    if (c === settings.channel) return
    await setSettings({ channel: c })
    toast(t('chSwitched', c === 'prod' ? t('prod') : t('test')))
    check()
  }

  const status = (() => {
    switch (upd.status) {
      case 'checking':
        return <span className="hint">{t('checking')}</span>
      case 'none':
        return <span className="pill ok">{t('upToDate')}</span>
      case 'available':
        return (
          <>
            <span className="pill warn">{t('updAvail', upd.version || '')}</span>
            <button className="btn pri sm" onClick={() => window.api.downloadUpdate()}>
              {t('download')}
            </button>
          </>
        )
      case 'downloading':
        return <span className="hint">{t('downloading', upd.percent || 0)}</span>
      case 'ready':
        return (
          <button className="btn pri sm" onClick={() => window.api.installUpdate()}>
            {t('install')}
          </button>
        )
      case 'error':
        return (
          <span className="pill bad" title={upd.error}>
            {t('updError')}
          </span>
        )
      case 'dev':
        return <span className="hint">{t('devMode')}</span>
      default:
        return null
    }
  })()

  return (
    <Modal title={t('infoTitle')} onClose={onClose} footer={<button className="btn" onClick={onClose}>{t('close')}</button>}>
      <div className="ver">
        <div>
          <div className="hint">{t('version')}</div>
          <div className="big">{info?.version || '…'}</div>
        </div>
        <div className="r">
          {status}
          {['idle', 'none', 'error', 'dev'].includes(upd.status) && (
            <button className="btn sm" onClick={check}>
              {t('checkUpd')}
            </button>
          )}
        </div>
      </div>

      <div className="fld">
        <label>{t('channel')}</label>
        <div className="channel">
          {(['prod', 'test'] as Channel[]).map((c) => (
            <button key={c} className={`ch ${settings.channel === c ? 'on' : ''}`} onClick={() => switchChannel(c)}>
              <b>{t(c)}</b>
              <small>{t(c === 'prod' ? 'prodD' : 'testD')}</small>
            </button>
          ))}
        </div>
      </div>

      <div className="row2">
        <div className="fld">
          <label>{t('language')}</label>
          <div className="seg">
            {(['en', 'nl'] as const).map((l) => (
              <button key={l} className={lang === l ? 'on' : ''} onClick={() => setSettings({ lang: l })}>
                {l === 'en' ? 'English' : 'Nederlands'}
              </button>
            ))}
          </div>
        </div>
        <div className="fld">
          <label>{t('theme')}</label>
          <div className="seg">
            {(['system', 'light', 'dark'] as Theme[]).map((th) => (
              <button key={th} className={settings.theme === th ? 'on' : ''} onClick={() => setSettings({ theme: th })}>
                {t(`th_${th}` as 'th_system')}
              </button>
            ))}
          </div>
        </div>
      </div>

      <details>
        <summary>{t('advanced')}</summary>
        <div className="fld" style={{ marginTop: 10 }}>
          <label htmlFor="cid">{t('clientId')}</label>
          <div className="pw">
            <input id="cid" value={clientId} onChange={(e) => setClientId(e.target.value)} />
            <button
              className="btn sm"
              onClick={async () => {
                await setSettings({ clientId: clientId.trim() })
                toast(t('saved'))
              }}
            >
              {t('save')}
            </button>
          </div>
          <span className="hint">{t('clientIdHint')}</span>
        </div>
      </details>

      <div className="made">
        <img src={logo} alt="RHIS" />
        <span className="hint">
          {t('madeBy')} ·{' '}
          <a href={`https://github.com/${info?.repo}`} target="_blank" rel="noreferrer" style={{ color: 'var(--accent)' }}>
            GitHub
          </a>{' '}
          · Electron {info?.electron}
        </span>
      </div>
    </Modal>
  )
}
