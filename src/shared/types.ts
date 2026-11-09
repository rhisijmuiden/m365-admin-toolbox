export type TenantMode = 'gdap' | 'direct'
export type Lang = 'en' | 'nl'
export type Theme = 'system' | 'light' | 'dark'
export type Channel = 'prod' | 'test'

export interface Tenant {
  tenantId: string
  name: string
  defaultDomain: string
  domains: string[]
  mode: TenantMode
  homeAccountId: string
  username: string
  color: string
  addedAt: string
}

export interface Settings {
  lang: Lang
  theme: Theme
  channel: Channel
  clientId: string
  lastTenantId?: string
}

export interface User {
  id: string
  displayName: string
  userPrincipalName: string
  department: string | null
  accountEnabled: boolean
  skuIds: string[]
  lastSignIn: string | null
  createdDateTime: string | null
}

export interface Sku {
  skuId: string
  skuPartNumber: string
  name: string
  enabled: number
  suspended: number
  warning: number
  consumed: number
  available: number
  shortage: number
  capabilityStatus: string
}

export interface Group {
  id: string
  displayName: string
  description: string | null
  kind: 'm365' | 'security' | 'distribution' | 'mailSecurity'
  mail: string | null
  createdDateTime: string | null
}

export interface NewUser {
  givenName: string
  surname: string
  displayName: string
  mailNickname: string
  userPrincipalName: string
  password: string
  forceChangePasswordNextSignIn: boolean
  usageLocation: string
  department?: string
  skuIds: string[]
}

export interface NewGroup {
  displayName: string
  description?: string
  mailNickname: string
  kind: 'm365' | 'security'
  ownerId?: string
}

export interface LicenseChange {
  userId: string
  add: string[]
  remove: string[]
}

export interface UpdateState {
  status: 'idle' | 'checking' | 'none' | 'available' | 'downloading' | 'ready' | 'error' | 'dev'
  version?: string
  percent?: number
  error?: string
}

export interface AppInfo {
  version: string
  channel: Channel
  isPackaged: boolean
  electron: string
  repo: string
}

/** Result envelope so errors cross IPC with a usable code and message. */
export type Result<T> = { ok: true; data: T } | { ok: false; error: string; code?: string }
