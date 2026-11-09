import { app, safeStorage } from 'electron'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs'
import { join } from 'path'
import type { Settings, Tenant } from '@shared/types'

/** Microsoft Graph Command Line Tools, a first-party public client that supports loopback redirects. */
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

const dir = () => {
  const d = app.getPath('userData')
  if (!existsSync(d)) mkdirSync(d, { recursive: true })
  return d
}

function readJson<T>(file: string, fallback: T): T {
  try {
    return { ...fallback, ...JSON.parse(readFileSync(join(dir(), file), 'utf8')) }
  } catch {
    return fallback
  }
}

function writeJson(file: string, data: unknown) {
  writeFileSync(join(dir(), file), JSON.stringify(data, null, 2))
}

const defaults: Settings = {
  lang: app.getLocale().toLowerCase().startsWith('nl') ? 'nl' : 'en',
  theme: 'system',
  channel: app.getVersion().includes('-') ? 'test' : 'prod',
  clientId: DEFAULT_CLIENT_ID
}

export const settings = {
  get: (): Settings => readJson('settings.json', defaults),
  set(patch: Partial<Settings>): Settings {
    const next = { ...settings.get(), ...patch }
    if (!next.clientId?.trim()) next.clientId = DEFAULT_CLIENT_ID
    writeJson('settings.json', next)
    return next
  }
}

export const tenants = {
  list: (): Tenant[] => readJson<{ items: Tenant[] }>('tenants.json', { items: [] }).items,
  save(items: Tenant[]) {
    writeJson('tenants.json', { items })
  },
  upsert(t: Tenant) {
    const items = tenants.list().filter((x) => x.tenantId !== t.tenantId)
    items.push(t)
    items.sort((a, b) => a.name.localeCompare(b.name))
    tenants.save(items)
  },
  remove(id: string) {
    tenants.save(tenants.list().filter((x) => x.tenantId !== id))
  },
  get: (id: string) => tenants.list().find((x) => x.tenantId === id)
}

/** MSAL token cache, encrypted with the Windows user's DPAPI key via safeStorage. */
export const tokenCache = {
  path: () => join(dir(), 'msal-cache.bin'),
  read(): string {
    const p = tokenCache.path()
    if (!existsSync(p)) return ''
    const buf = readFileSync(p)
    if (safeStorage.isEncryptionAvailable()) {
      try {
        return safeStorage.decryptString(buf)
      } catch {
        return ''
      }
    }
    return buf.toString('utf8')
  },
  write(data: string) {
    const out = safeStorage.isEncryptionAvailable() ? safeStorage.encryptString(data) : Buffer.from(data, 'utf8')
    writeFileSync(tokenCache.path(), out)
  }
}
