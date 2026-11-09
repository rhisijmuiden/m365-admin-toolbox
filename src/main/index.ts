import { app, BrowserWindow, ipcMain, nativeTheme, shell } from 'electron'
import { join } from 'path'
import type { AppInfo, LicenseChange, NewGroup, NewUser, PimKind, PimRequest, Result, Settings, Tenant, TenantMode } from '@shared/types'
import { forget, homeTenantOf, signInInteractive } from './auth'
import * as graph from './graph'
import * as pim from './pim'
import { initPimWatch } from './pimWatch'
import { settings, tenants } from './store'
import { applyChannel, checkForUpdates, downloadUpdate, initUpdater, installUpdate, updateState } from './updater'

const REPO = 'rhisijmuiden/m365-admin-toolbox'
const COLORS = ['#1f6feb', '#c2381f', '#0a6f75', '#5b5fc7', '#a15c00', '#1a7f4b', '#b4237a', '#3a475a']

let win: BrowserWindow | null = null

function createWindow() {
  const s = settings.get()
  nativeTheme.themeSource = s.theme
  win = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 980,
    minHeight: 640,
    show: false,
    title: 'M365 Admin Toolbox',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#0e141c' : '#f3f5f8',
    autoHideMenuBar: true,
    // Packaged builds take the icon from the exe; in dev point at the source PNG.
    icon: app.isPackaged ? undefined : join(__dirname, '../../build/icon.png'),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  })
  win.once('ready-to-show', () => win?.show())
  // Links from the UI (GitHub, docs) open in the default browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https:\/\//.test(url)) shell.openExternal(url)
    return { action: 'deny' }
  })
  if (!app.isPackaged && process.env.ELECTRON_RENDERER_URL) win.loadURL(process.env.ELECTRON_RENDERER_URL)
  else win.loadFile(join(__dirname, '../renderer/index.html'))
  initUpdater(win, s.channel)
  initPimWatch(win)
}

/** Wraps a handler so errors reach the renderer as a Result instead of a rejected invoke. */
function handle<A extends unknown[], T>(name: string, fn: (...args: A) => Promise<T> | T) {
  ipcMain.handle(name, async (_e, ...args: unknown[]): Promise<Result<T>> => {
    try {
      return { ok: true, data: await fn(...(args as A)) }
    } catch (e) {
      const err = e as Error & { code?: string; status?: number }
      return { ok: false, error: err.message || String(e), code: err.code, status: err.status }
    }
  })
}

async function addTenant(mode: TenantMode, target?: string): Promise<Tenant> {
  const authority = mode === 'gdap' ? (target || '').trim() : 'organizations'
  if (!authority) throw new Error('Tenant domain or ID is required')
  const r = await signInInteractive(authority)
  if (!r.account) throw new Error('Sign-in returned no account')
  const org = await graph.organization(r.accessToken)
  const existing = tenants.get(r.tenantId)
  const t: Tenant = {
    tenantId: r.tenantId,
    name: org.name,
    defaultDomain: org.defaultDomain,
    domains: org.domains,
    mode,
    homeAccountId: r.account.homeAccountId,
    username: r.account.username,
    color: existing?.color || COLORS[tenants.list().length % COLORS.length],
    addedAt: existing?.addedAt || new Date().toISOString()
  }
  tenants.upsert(t)
  settings.set({ lastTenantId: t.tenantId })
  return t
}

/** Signs in again for a tenant whose refresh token expired or was revoked. */
async function reauth(tenantId: string): Promise<Tenant> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error('Unknown tenant')
  const r = await signInInteractive(tenantId, t.username)
  if (!r.account) throw new Error('Sign-in returned no account')
  if (r.tenantId !== tenantId) throw new Error('Signed in to a different tenant')
  const next = { ...t, homeAccountId: r.account.homeAccountId, username: r.account.username }
  tenants.upsert(next)
  return next
}

/** For GDAP tenants: sign in to the partner tenant, which PIM for groups needs (consent there too). */
async function reauthHome(tenantId: string) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error('Unknown tenant')
  await signInInteractive(homeTenantOf(tenantId), t.username)
}

function registerIpc() {
  handle('app:info', (): AppInfo => ({
    version: app.getVersion(),
    channel: settings.get().channel,
    isPackaged: app.isPackaged,
    electron: process.versions.electron,
    repo: REPO
  }))
  handle('settings:get', () => settings.get())
  handle('settings:set', (patch: Partial<Settings>) => {
    const next = settings.set(patch)
    if (patch.theme) nativeTheme.themeSource = next.theme
    if (patch.channel) applyChannel(next.channel)
    return next
  })
  handle('tenants:list', () => tenants.list())
  handle('tenants:add', (mode: TenantMode, target?: string) => addTenant(mode, target))
  handle('tenants:reauth', (id: string) => reauth(id))
  handle('tenants:remove', async (id: string) => {
    const t = tenants.get(id)
    tenants.remove(id)
    if (t) await forget(t.homeAccountId)
    return tenants.list()
  })
  handle('users:list', (id: string) => graph.listUsers(id))
  handle('users:create', (id: string, u: NewUser) => graph.createUser(id, u))
  handle('users:setSignIn', (id: string, userId: string, enabled: boolean) => graph.setSignIn(id, userId, enabled))
  handle('users:resetPassword', (id: string, userId: string, pw: string, force: boolean) =>
    graph.resetPassword(id, userId, pw, force)
  )
  handle('users:createTap', (id: string, userId: string, minutes: number, once: boolean) =>
    graph.createTap(id, userId, minutes, once)
  )
  handle('users:revokeSessions', (id: string, userId: string) => graph.revokeSessions(id, userId))
  handle('licenses:list', (id: string) => graph.listSkus(id))
  handle('licenses:change', (id: string, changes: LicenseChange[]) => graph.changeLicenses(id, changes))
  handle('groups:list', (id: string) => graph.listGroups(id))
  handle('groups:create', (id: string, g: NewGroup) => graph.createGroup(id, g))
  handle('tenants:reauthHome', (id: string) => reauthHome(id))
  handle('pim:load', (id: string) => pim.load(id))
  handle('pim:activate', (id: string, r: PimRequest) => pim.activate(id, r))
  handle('pim:extend', (id: string, r: PimRequest) => pim.extend(id, r))
  handle('pim:deactivate', (id: string, kind: PimKind, targetId: string) => pim.deactivate(id, kind, targetId))
  handle('pim:cancel', (id: string, kind: PimKind, requestId: string) => pim.cancel(id, kind, requestId))
  handle('update:state', () => updateState())
  handle('update:check', () => checkForUpdates())
  handle('update:download', () => downloadUpdate())
  handle('update:install', () => installUpdate())
}

if (!app.requestSingleInstanceLock()) app.quit()
else {
  app.on('second-instance', () => {
    if (win) {
      if (win.isMinimized()) win.restore()
      win.focus()
    }
  })
  app.setAppUserModelId('nl.rhis.m365toolbox')
  app.whenReady().then(() => {
    registerIpc()
    createWindow()
  })
  app.on('window-all-closed', () => app.quit())
}
