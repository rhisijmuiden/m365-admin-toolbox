import { shell } from 'electron'
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type ICachePlugin
} from '@azure/msal-node'
import { settings, tenants, tokenCache } from './store'

export const SCOPES = [
  'User.ReadWrite.All',
  'Group.ReadWrite.All',
  'Directory.ReadWrite.All',
  'Organization.Read.All',
  'AuditLog.Read.All',
  'User-PasswordProfile.ReadWrite.All',
  'UserAuthenticationMethod.ReadWrite.All',
  'User.RevokeSessions.All'
]

const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(ctx) {
    ctx.tokenCache.deserialize(tokenCache.read())
  },
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) tokenCache.write(ctx.tokenCache.serialize())
  }
}

let pca: PublicClientApplication | null = null
let pcaClientId = ''

function client(): PublicClientApplication {
  const { clientId } = settings.get()
  if (!pca || pcaClientId !== clientId) {
    pca = new PublicClientApplication({
      auth: { clientId, authority: 'https://login.microsoftonline.com/organizations' },
      cache: { cachePlugin }
    })
    pcaClientId = clientId
  }
  return pca
}

const page = (title: string, body: string) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:Segoe UI,sans-serif;background:#f3f5f8;color:#18202b;display:grid;place-items:center;height:100vh;margin:0"><div style="text-align:center"><h2 style="margin:0 0 8px">${title}</h2><p style="color:#5c6878">${body}</p></div></body>`

/**
 * Opens the system browser for a Microsoft sign-in. `authority` is "organizations" for a tenant admin
 * signing in with their own account, or the customer tenant (domain or id) for GDAP partner access.
 */
export async function signInInteractive(authority: string, loginHint?: string): Promise<AuthenticationResult> {
  return client().acquireTokenInteractive({
    scopes: SCOPES,
    authority: `https://login.microsoftonline.com/${authority}`,
    prompt: 'select_account',
    loginHint,
    openBrowser: async (url) => {
      await shell.openExternal(url)
    },
    successTemplate: page('Signed in / Aangemeld', 'You can close this tab and return to M365 Admin Toolbox.'),
    errorTemplate: page('Sign-in failed / Aanmelden mislukt', 'Return to M365 Admin Toolbox for details.')
  })
}

async function accountFor(homeAccountId: string): Promise<AccountInfo | null> {
  return client().getTokenCache().getAccountByHomeId(homeAccountId)
}

export class NeedsSignIn extends Error {
  code = 'NEEDS_SIGNIN'
}

/** Graph access token for a registered tenant, silently from cache or refresh token. */
export async function tokenFor(tenantId: string): Promise<string> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const account = await accountFor(t.homeAccountId)
  if (!account) throw new NeedsSignIn('No cached account for this tenant')
  try {
    const r = await client().acquireTokenSilent({
      account,
      scopes: SCOPES,
      authority: `https://login.microsoftonline.com/${tenantId}`
    })
    return r.accessToken
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) throw new NeedsSignIn(e.message)
    throw e
  }
}

/** Removes the cached account unless another registered tenant still uses it (common with GDAP). */
export async function forget(homeAccountId: string) {
  if (tenants.list().some((t) => t.homeAccountId === homeAccountId)) return
  const account = await accountFor(homeAccountId)
  if (account) await client().getTokenCache().removeAccount(account)
}
