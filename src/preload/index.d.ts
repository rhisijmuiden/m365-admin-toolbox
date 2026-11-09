import type {
  AppInfo,
  Group,
  LicenseChange,
  NewGroup,
  NewUser,
  Result,
  Settings,
  Sku,
  Tenant,
  TenantMode,
  UpdateState,
  User
} from '../shared/types'

type R<T> = Promise<Result<T>>

export interface Api {
  appInfo(): R<AppInfo>
  getSettings(): R<Settings>
  setSettings(patch: Partial<Settings>): R<Settings>
  listTenants(): R<Tenant[]>
  addTenant(mode: TenantMode, target?: string): R<Tenant>
  reauthTenant(tenantId: string): R<Tenant>
  removeTenant(tenantId: string): R<Tenant[]>
  listUsers(tenantId: string): R<User[]>
  createUser(tenantId: string, u: NewUser): R<User>
  setSignIn(tenantId: string, userId: string, enabled: boolean): R<void>
  resetPassword(tenantId: string, userId: string, password: string, forceChange: boolean): R<void>
  createTap(tenantId: string, userId: string, lifetimeInMinutes: number, isUsableOnce: boolean): R<{ pass: string; lifetimeInMinutes: number; start: string }>
  revokeSessions(tenantId: string, userId: string): R<void>
  listSkus(tenantId: string): R<Sku[]>
  changeLicenses(tenantId: string, changes: LicenseChange[]): R<{ failed: { userId: string; error: string }[] }>
  listGroups(tenantId: string): R<Group[]>
  createGroup(tenantId: string, g: NewGroup): R<Group>
  updateState(): R<UpdateState>
  checkUpdate(): R<UpdateState>
  downloadUpdate(): R<void>
  installUpdate(): R<void>
  onUpdate(cb: (s: UpdateState) => void): () => void
}

declare global {
  interface Window {
    api: Api
  }
}
