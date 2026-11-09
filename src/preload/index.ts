import { contextBridge, ipcRenderer } from 'electron'
import type { UpdateState } from '@shared/types'

const inv = (ch: string) => (...args: unknown[]) => ipcRenderer.invoke(ch, ...args)

const api = {
  appInfo: inv('app:info'),
  getSettings: inv('settings:get'),
  setSettings: inv('settings:set'),
  listTenants: inv('tenants:list'),
  addTenant: inv('tenants:add'),
  reauthTenant: inv('tenants:reauth'),
  removeTenant: inv('tenants:remove'),
  listUsers: inv('users:list'),
  createUser: inv('users:create'),
  setSignIn: inv('users:setSignIn'),
  resetPassword: inv('users:resetPassword'),
  createTap: inv('users:createTap'),
  revokeSessions: inv('users:revokeSessions'),
  listSkus: inv('licenses:list'),
  changeLicenses: inv('licenses:change'),
  listGroups: inv('groups:list'),
  createGroup: inv('groups:create'),
  updateState: inv('update:state'),
  checkUpdate: inv('update:check'),
  downloadUpdate: inv('update:download'),
  installUpdate: inv('update:install'),
  onUpdate(cb: (s: UpdateState) => void) {
    const fn = (_e: unknown, s: UpdateState) => cb(s)
    ipcRenderer.on('update:state', fn)
    return () => ipcRenderer.removeListener('update:state', fn)
  }
}

contextBridge.exposeInMainWorld('api', api)
