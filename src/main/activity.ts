import { BrowserWindow, dialog } from 'electron'
import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'fs'
import { join } from 'path'
import type { ActivityEntry } from '@shared/types'
import { tenants, userDataDir } from './store'

/**
 * Local log of every change made through the app. It is written for every IPC call that changes
 * something in a tenant, so new screens are covered without extra code.
 */
const file = () => join(userDataDir(), 'activity.jsonl')
const MAX_BYTES = 5 * 1024 * 1024
const KEEP_LINES = 20000

const VERBS =
  /^(create|set|add|remove|delete|reset|revoke|grant|assign|unassign|change|convert|new|enable|update|archive|release|purge|start|discard|remediate|invite|review|activate|extend|deactivate|cancel|secure|save|quarantine(Release|Delete)|enterpriseVoice|pimGroup(Assign|Remove))([A-Z]|$)/
const SKIP = /^(settings|update|app|activity):/
const SECRET = /pass|secret|password|token|cert|content|body|message/i

export const shouldLog = (channel: string) => !SKIP.test(channel) && VERBS.test(channel.split(':')[1] || '')

function redact(v: unknown, depth = 0): unknown {
  if (Array.isArray(v)) return v.length > 20 ? [...v.slice(0, 20).map((x) => redact(x, depth + 1)), `+${v.length - 20}`] : v.map((x) => redact(x, depth + 1))
  if (v && typeof v === 'object' && depth < 4)
    return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, SECRET.test(k) ? '***' : redact(x, depth + 1)]))
  return v
}

/** The tenant a call is about: its first argument when that is a known tenant id. */
export function tenantOf(args: unknown[]) {
  const first = args[0]
  const t = typeof first === 'string' ? tenants.get(first) : undefined
  return t ? { tenantId: t.tenantId, tenantName: t.name, rest: args.slice(1) } : { tenantId: '', tenantName: '', rest: args }
}

export function record(channel: string, pre: ReturnType<typeof tenantOf>, ok: boolean, error?: string, result?: unknown) {
  let { tenantId, tenantName } = pre
  const { rest } = pre
  // Adding a tenant has no tenant id yet; take it from the result.
  const r = result as { tenantId?: string; name?: string } | undefined
  if (!tenantId && r?.tenantId) (tenantId = r.tenantId), (tenantName = r.name || '')
  // Bulk calls report per-item failures in a successful result; the log shows them as failed.
  const failed = (result as { failed?: { error?: string }[] } | undefined)?.failed
  if (ok && Array.isArray(failed) && failed.length) (ok = false), (error = `${failed.length} failed: ${failed.map((f) => f.error).join('; ')}`)
  const details = rest
    .map((a) => (typeof a === 'string' ? a : JSON.stringify(redact(a))))
    .join(' · ')
    .slice(0, 600)
  const e: ActivityEntry = { at: new Date().toISOString(), tenantId, tenantName, action: channel, details, ok, error: error?.slice(0, 400) }
  try {
    appendFileSync(file(), JSON.stringify(e) + '\n')
    if (statSync(file()).size > MAX_BYTES) writeFileSync(file(), readFileSync(file(), 'utf8').split('\n').filter(Boolean).slice(-KEEP_LINES).join('\n') + '\n')
  } catch {
    // Logging must never break the action itself.
  }
}

export function list(tenantId?: string): ActivityEntry[] {
  if (!existsSync(file())) return []
  const out: ActivityEntry[] = []
  for (const line of readFileSync(file(), 'utf8').split('\n')) {
    if (!line) continue
    try {
      const e = JSON.parse(line) as ActivityEntry
      if (!tenantId || e.tenantId === tenantId) out.push(e)
    } catch {
      // A line cut off by a crash is skipped.
    }
  }
  return out.reverse()
}

const csvCell = (v: string) => (/[";\n,]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)

/** Saves the log (one tenant or all) as CSV; returns the path, or null when the dialog was cancelled. */
export async function exportCsv(win: BrowserWindow | null, tenantId?: string) {
  const items = list(tenantId)
  const name = tenantId ? tenants.get(tenantId)?.defaultDomain || tenantId : 'all-tenants'
  const opts = { defaultPath: `toolbox-activity-${name}-${new Date().toISOString().slice(0, 10)}.csv`, filters: [{ name: 'CSV', extensions: ['csv'] }] }
  const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts)
  if (r.canceled || !r.filePath) return null
  const rows = [['time', 'tenant', 'tenantId', 'action', 'details', 'result', 'error']].concat(
    items.map((e) => [e.at, e.tenantName, e.tenantId, e.action, e.details, e.ok ? 'ok' : 'failed', e.error || ''])
  )
  // BOM so Excel opens the file as UTF-8.
  writeFileSync(r.filePath, '﻿' + rows.map((x) => x.map(csvCell).join(',')).join('\r\n'))
  return r.filePath
}
