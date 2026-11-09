import { app, type BrowserWindow } from 'electron'
import { autoUpdater } from 'electron-updater'
import type { Channel, UpdateState } from '@shared/types'

let win: BrowserWindow | null = null
let state: UpdateState = { status: app.isPackaged ? 'idle' : 'dev' }

function emit(next: UpdateState) {
  state = next
  win?.webContents.send('update:state', state)
}

/**
 * Production follows full GitHub releases built from main; test follows pre-releases built from the
 * test branch (versions like 0.2.0-beta.7). Downgrades are allowed so switching back to production
 * installs the latest stable build even when it is older than the test build.
 */
export function applyChannel(channel: Channel) {
  autoUpdater.allowPrerelease = channel === 'test'
  autoUpdater.channel = channel === 'test' ? 'beta' : 'latest'
  autoUpdater.allowDowngrade = true
}

export function initUpdater(w: BrowserWindow, channel: Channel) {
  win = w
  autoUpdater.autoDownload = false
  autoUpdater.autoInstallOnAppQuit = true
  applyChannel(channel)
  autoUpdater.on('checking-for-update', () => emit({ status: 'checking' }))
  autoUpdater.on('update-available', (i) => emit({ status: 'available', version: i.version }))
  autoUpdater.on('update-not-available', () => emit({ status: 'none' }))
  autoUpdater.on('download-progress', (p) => emit({ ...state, status: 'downloading', percent: Math.round(p.percent) }))
  autoUpdater.on('update-downloaded', (i) => emit({ status: 'ready', version: i.version }))
  autoUpdater.on('error', (e) => emit({ status: 'error', error: e.message }))
  if (app.isPackaged) setTimeout(() => checkForUpdates().catch(() => {}), 5000)
}

export const updateState = () => state

export async function checkForUpdates() {
  if (!app.isPackaged) {
    emit({ status: 'dev' })
    return state
  }
  await autoUpdater.checkForUpdates()
  return state
}

export async function downloadUpdate() {
  emit({ ...state, status: 'downloading', percent: 0 })
  await autoUpdater.downloadUpdate()
}

export function installUpdate() {
  autoUpdater.quitAndInstall(false, true)
}
