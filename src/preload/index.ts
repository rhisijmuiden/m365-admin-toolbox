import { contextBridge, ipcRenderer } from 'electron'
import type { ExoPhase, UpdateState } from '@shared/types'

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
  reauthHome: inv('tenants:reauthHome'),
  pimLoad: inv('pim:load'),
  pimActivate: inv('pim:activate'),
  pimExtend: inv('pim:extend'),
  pimDeactivate: inv('pim:deactivate'),
  pimCancel: inv('pim:cancel'),
  onNav(cb: (n: { tenantId: string; view: string }) => void) {
    const fn = (_e: unknown, n: { tenantId: string; view: string }) => cb(n)
    ipcRenderer.on('nav', fn)
    return () => ipcRenderer.removeListener('nav', fn)
  },
  exo: {
    signIn: inv('exo:signIn'),
    consentReports: inv('exo:consentReports'),
    mailboxes: inv('exo:mailboxes'),
    mailbox: inv('exo:mailbox'),
    sizes: inv('exo:sizes'),
    setMailbox: inv('exo:setMailbox'),
    enableArchive: inv('exo:enableArchive'),
    convertToShared: inv('exo:convertToShared'),
    newMailbox: inv('exo:newMailbox'),
    permissions: inv('exo:permissions'),
    permissionsForUser: inv('exo:permissionsForUser'),
    grant: inv('exo:grant'),
    revoke: inv('exo:revoke'),
    autoReply: inv('exo:autoReply'),
    setAutoReply: inv('exo:setAutoReply'),
    calendar: inv('exo:calendar'),
    setCalendar: inv('exo:setCalendar'),
    distGroups: inv('exo:distGroups'),
    distGroup: inv('exo:distGroup'),
    newDistGroup: inv('exo:newDistGroup'),
    setDistGroup: inv('exo:setDistGroup'),
    addMember: inv('exo:addMember'),
    removeMember: inv('exo:removeMember'),
    removeDistGroup: inv('exo:removeDistGroup'),
    recipients: inv('exo:recipients'),
    setAddresses: inv('exo:setAddresses'),
    inboxRules: inv('exo:inboxRules'),
    setInboxRule: inv('exo:setInboxRule'),
    mailFlow: inv('exo:mailFlow'),
    onStatus(cb: (s: { tenantId: string; phase: ExoPhase }) => void) {
      const fn = (_e: unknown, s: { tenantId: string; phase: ExoPhase }) => cb(s)
      ipcRenderer.on('exo:status', fn)
      return () => ipcRenderer.removeListener('exo:status', fn)
    }
  },
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
