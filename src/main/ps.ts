import { spawn, type ChildProcessWithoutNullStreams } from 'child_process'
import { existsSync, readFileSync, writeFileSync } from 'fs'
import { join } from 'path'
import type { BrowserWindow } from 'electron'
import type { ExoPhase } from '@shared/types'
import { userDataDir } from './store'

/*
 * A long-lived Windows PowerShell helper per tenant (Exchange Online, Microsoft Teams). Each helper script
 * reads one JSON request per line and answers with one marked JSON line, so only the first command per
 * tenant pays for loading the module and connecting. Requests are queued per tenant because a PowerShell
 * runspace runs one command at a time.
 */

const MARK = '@@TB@@'
const IDLE_MS = 20 * 60000

let win: BrowserWindow | null = null
export const setPsWindow = (w: BrowserWindow) => (win = w)

export class PsError extends Error {
  constructor(
    message: string,
    public code?: string,
    public status?: number
  ) {
    super(message)
  }
}

type Reply = { id: number; ok: boolean; data?: unknown; error?: string; code?: string | null }

class Proc {
  private proc: ChildProcessWithoutNullStreams
  private buf = ''
  private seq = 0
  private pending = new Map<number, { resolve(v: unknown): void; reject(e: Error): void }>()
  private stderr = ''
  dead = false
  onExit?: () => void

  constructor(script: string) {
    this.proc = spawn('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script], {
      windowsHide: true
    })
    this.proc.stdout.setEncoding('utf8')
    this.proc.stderr.setEncoding('utf8')
    this.proc.stdout.on('data', (d: string) => this.read(d))
    this.proc.stderr.on('data', (d: string) => (this.stderr = (this.stderr + d).slice(-4000)))
    this.proc.on('error', (e) => this.fail(e.message))
    this.proc.on('exit', (code) => this.fail(`PowerShell exited (${code})${this.stderr ? ': ' + this.stderr.trim().slice(-600) : ''}`))
  }

  private read(d: string) {
    this.buf += d
    let i: number
    while ((i = this.buf.indexOf('\n')) >= 0) {
      const line = this.buf.slice(0, i).replace(/\r$/, '')
      this.buf = this.buf.slice(i + 1)
      const at = line.indexOf(MARK)
      if (at < 0) continue
      let r: Reply
      try {
        r = JSON.parse(line.slice(at + MARK.length))
      } catch {
        continue
      }
      const p = this.pending.get(r.id)
      if (!p) continue
      this.pending.delete(r.id)
      if (r.ok) p.resolve(r.data)
      else p.reject(new PsError(r.error || 'PowerShell error', r.code || undefined, r.code?.endsWith('_DENIED') ? 403 : undefined))
    }
  }

  private fail(msg: string) {
    if (this.dead) return
    this.dead = true
    for (const p of this.pending.values()) p.reject(new PsError(msg))
    this.pending.clear()
    this.onExit?.()
  }

  call<T>(op: string, args: Record<string, unknown> = {}): Promise<T> {
    if (this.dead) return Promise.reject(new PsError('PowerShell is not running'))
    const id = ++this.seq
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject })
      this.proc.stdin.write(JSON.stringify({ id, op, args }) + '\n')
    })
  }

  close() {
    if (this.dead) return
    this.proc.stdin.end()
    setTimeout(() => !this.dead && this.proc.kill(), 3000)
  }
}

export interface HelperDef {
  /** Short name, used for the script file and the `<name>:status` event to the renderer. */
  name: string
  source: string
  /**
   * Fresh connect arguments for a tenant. `key` changes whenever a token was renewed, which makes the
   * session reconnect before the next command.
   */
  connect(tenantId: string): Promise<{ key: string; args: Record<string, unknown> }>
}

export function psHelper(def: HelperDef) {
  const status = (tenantId: string, phase: ExoPhase) => win?.webContents.send(`${def.name}:status`, { tenantId, phase })

  function scriptPath() {
    const p = join(userDataDir(), `${def.name}-helper.ps1`)
    // Rewritten when the app ships a new version of the script. The BOM makes Windows PowerShell read it as UTF-8.
    const want = '﻿' + def.source
    if (!existsSync(p) || readFileSync(p, 'utf8') !== want) writeFileSync(p, want, 'utf8')
    return p
  }

  /** Checks for the PowerShell module and installs it for the current user when missing. Runs once per app start. */
  let moduleReady: Promise<{ version: string; installed: boolean }> | null = null
  function ensureModule() {
    if (!moduleReady) {
      const p = new Proc(scriptPath())
      moduleReady = p.call<{ version: string; installed: boolean }>('Tb-Init').finally(() => p.close())
      // A failed check is retried on the next use instead of being cached.
      moduleReady.catch(() => (moduleReady = null))
    }
    return moduleReady
  }

  class Session {
    private proc: Proc | null = null
    private key = ''
    private chain: Promise<unknown> = Promise.resolve()
    private idle?: NodeJS.Timeout

    constructor(private tenantId: string) {}

    private async exec<T>(op: string, args: Record<string, unknown>): Promise<T> {
      clearTimeout(this.idle)
      try {
        const c = await def.connect(this.tenantId)
        if (!this.proc || this.proc.dead) {
          status(this.tenantId, 'module')
          await ensureModule()
          this.proc = new Proc(scriptPath())
          this.proc.onExit = () => (this.key = '')
          await this.proc.call('Tb-Init')
          this.key = ''
        }
        // MSAL hands out a new token shortly before the old one expires; reconnect with it.
        if (c.key !== this.key) {
          status(this.tenantId, 'connecting')
          await this.proc.call('Tb-Connect', c.args)
          this.key = c.key
        }
        status(this.tenantId, 'ready')
        return await this.proc.call<T>(op, args)
      } finally {
        this.idle = setTimeout(() => this.close(), IDLE_MS)
      }
    }

    run<T>(op: string, args: Record<string, unknown>): Promise<T> {
      const r = this.chain.then(() => this.exec<T>(op, args))
      this.chain = r.catch(() => undefined)
      return r
    }

    close() {
      clearTimeout(this.idle)
      this.proc?.close()
      this.proc = null
      this.key = ''
    }
  }

  const sessions = new Map<string, Session>()

  return {
    ensureModule,
    run<T>(tenantId: string, op: string, args: Record<string, unknown> = {}): Promise<T> {
      let s = sessions.get(tenantId)
      if (!s) sessions.set(tenantId, (s = new Session(tenantId)))
      return s.run<T>(op, args)
    },
    closeSession(tenantId: string) {
      sessions.get(tenantId)?.close()
      sessions.delete(tenantId)
    },
    closeAll() {
      for (const s of sessions.values()) s.close()
      sessions.clear()
    }
  }
}
