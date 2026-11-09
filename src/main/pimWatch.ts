import { Notification, type BrowserWindow } from 'electron'
import type { PimItem } from '@shared/types'
import { settings, tenants } from './store'

/** Warns this many minutes before a PIM activation ends. */
const WARN_MIN = 10

type Watched = { tenantId: string; item: PimItem; warned: boolean }
const watched = new Map<string, Watched>()
let win: BrowserWindow | null = null

const key = (tenantId: string, i: PimItem) => `${tenantId}:${i.kind}:${i.targetId}`

/** Records the timed activations of a tenant. With `merge`, adds to what is known instead of replacing it. */
export function track(tenantId: string, items: PimItem[], merge = false) {
  if (!merge) for (const k of [...watched.keys()]) if (k.startsWith(`${tenantId}:`)) watched.delete(k)
  for (const i of items) {
    if (i.state !== 'active' || !i.end) continue
    const k = key(tenantId, i)
    const prev = watched.get(k)
    watched.set(k, { tenantId, item: i, warned: prev?.item.end === i.end ? prev.warned : false })
  }
}

function tick() {
  const now = Date.now()
  const nl = settings.get().lang === 'nl'
  for (const [k, w] of watched) {
    const end = new Date(w.item.end!).getTime()
    if (end < now) {
      watched.delete(k)
      continue
    }
    const mins = Math.round((end - now) / 60000)
    if (w.warned || mins > WARN_MIN) continue
    w.warned = true
    const tenant = tenants.get(w.tenantId)?.name || ''
    const n = new Notification({
      title: nl ? `${w.item.name} verloopt over ${mins} min` : `${w.item.name} expires in ${mins} min`,
      body: nl ? `${tenant}: verleng of activeer opnieuw in M365 Admin Toolbox.` : `${tenant}: extend or activate again in M365 Admin Toolbox.`
    })
    n.on('click', () => {
      if (!win) return
      if (win.isMinimized()) win.restore()
      win.focus()
      win.webContents.send('nav', { tenantId: w.tenantId, view: 'pim' })
    })
    n.show()
  }
}

export function initPimWatch(w: BrowserWindow) {
  win = w
  setInterval(tick, 30_000)
}
