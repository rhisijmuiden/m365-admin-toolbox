import { shell } from 'electron'
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type ICachePlugin
} from '@azure/msal-node'
import { exoTokenCache, settings, tenants, tokenCache } from './store'

export const SCOPES = [
  'User.ReadWrite.All',
  'Group.ReadWrite.All',
  'Directory.ReadWrite.All',
  'Organization.Read.All',
  'AuditLog.Read.All',
  'User-PasswordProfile.ReadWrite.All',
  'UserAuthenticationMethod.ReadWrite.All',
  'User.RevokeSessions.All',
  // PIM: directory roles in the customer tenant, PIM for groups and GDAP relationships in the partner tenant.
  'RoleManagement.Read.Directory',
  'RoleEligibilitySchedule.Read.Directory',
  'RoleAssignmentSchedule.ReadWrite.Directory',
  'PrivilegedEligibilitySchedule.Read.AzureADGroup',
  'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup',
  'DelegatedAdminRelationship.Read.All'
]

const pluginFor = (file: typeof tokenCache): ICachePlugin => ({
  async beforeCacheAccess(ctx) {
    ctx.tokenCache.deserialize(file.read())
  },
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) file.write(ctx.tokenCache.serialize())
  }
})

const cachePlugin = pluginFor(tokenCache)

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

export class NeedsHomeSignIn extends Error {
  code = 'NEEDS_HOME_SIGNIN'
}

async function silent(homeAccountId: string, authorityTenant: string, Err: typeof NeedsSignIn | typeof NeedsHomeSignIn) {
  const account = await accountFor(homeAccountId)
  if (!account) throw new Err('No cached account for this tenant')
  try {
    const r = await client().acquireTokenSilent({
      account,
      scopes: SCOPES,
      authority: `https://login.microsoftonline.com/${authorityTenant}`
    })
    return r.accessToken
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) throw new Err(e.message)
    throw e
  }
}

/** Graph access token for a registered tenant, silently from cache or refresh token. */
export async function tokenFor(tenantId: string): Promise<string> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  return silent(t.homeAccountId, tenantId, NeedsSignIn)
}

/** The tenant the signed-in account belongs to; for GDAP that is the partner (RHIS) tenant. */
export function homeTenantOf(tenantId: string): string {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  return t.homeAccountId.split('.')[1] || tenantId
}

/** Token for the partner tenant of a GDAP tenant, used for PIM for groups and GDAP relationships. */
export async function tokenForHome(tenantId: string): Promise<string> {
  const t = tenants.get(tenantId)!
  return silent(t.homeAccountId, homeTenantOf(tenantId), NeedsHomeSignIn)
}

/**
 * Interactive step-up for PIM: a claims challenge (Conditional Access authentication context) or a
 * fresh MFA sign-in. Returns the new token directly, because a silent call could still serve the old one.
 */
export async function stepUp(authorityTenant: string, loginHint: string, claims?: string): Promise<string> {
  const r = await client().acquireTokenInteractive({
    scopes: SCOPES,
    authority: `https://login.microsoftonline.com/${authorityTenant}`,
    loginHint,
    claims,
    prompt: claims ? undefined : 'login',
    extraQueryParameters: claims ? undefined : { amr_values: 'mfa' },
    openBrowser: async (url) => {
      await shell.openExternal(url)
    },
    successTemplate: page('Verified / Geverifieerd', 'You can close this tab and return to M365 Admin Toolbox.'),
    errorTemplate: page('Verification failed / Verificatie mislukt', 'Return to M365 Admin Toolbox for details.')
  })
  return r.accessToken
}

/** Removes the cached account unless another registered tenant still uses it (common with GDAP). */
export async function forget(homeAccountId: string) {
  if (tenants.list().some((t) => t.homeAccountId === homeAccountId)) return
  const account = await accountFor(homeAccountId)
  if (account) await client().getTokenCache().removeAccount(account)
  const exoAccount = await exoClient().getTokenCache().getAccountByHomeId(homeAccountId)
  if (exoAccount) await exoClient().getTokenCache().removeAccount(exoAccount)
}

export class NeedsConsent extends Error {
  code = 'NEEDS_CONSENT'
}

/** Graph token with scopes outside the base set (e.g. Reports.Read.All), which not every tenant has consented to. */
export async function tokenForScopes(tenantId: string, scopes: string[]): Promise<string> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const account = await accountFor(t.homeAccountId)
  if (!account) throw new NeedsSignIn('No cached account for this tenant')
  try {
    const r = await client().acquireTokenSilent({ account, scopes, authority: `https://login.microsoftonline.com/${tenantId}` })
    return r.accessToken
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) throw new NeedsConsent(e.message)
    throw e
  }
}

/** Interactive sign-in that asks consent for extra Graph scopes. */
export async function consentScopes(tenantId: string, scopes: string[]) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  await client().acquireTokenInteractive({
    scopes,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    loginHint: t.username,
    openBrowser: async (url) => {
      await shell.openExternal(url)
    },
    successTemplate: page('Signed in / Aangemeld', 'You can close this tab and return to M365 Admin Toolbox.'),
    errorTemplate: page('Sign-in failed / Aanmelden mislukt', 'Return to M365 Admin Toolbox for details.')
  })
}

/*
 * Exchange Online: Connect-ExchangeOnline -AccessToken needs a token for outlook.office365.com issued to the
 * Exchange Online PowerShell app itself. It is a separate public client with its own consent and cache;
 * the user signs in once per tenant (usually SSO in the browser) and refresh tokens do the rest.
 */
export const EXO_CLIENT_ID = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'
const EXO_SCOPES = ['https://outlook.office365.com/.default']

let exoPca: PublicClientApplication | null = null
function exoClient(): PublicClientApplication {
  if (!exoPca)
    exoPca = new PublicClientApplication({
      auth: { clientId: EXO_CLIENT_ID, authority: 'https://login.microsoftonline.com/organizations' },
      cache: { cachePlugin: pluginFor(exoTokenCache) }
    })
  return exoPca
}

export class NeedsExoSignIn extends Error {
  code = 'NEEDS_EXO_SIGNIN'
}

export async function exoToken(tenantId: string): Promise<{ token: string; expiresOn: number }> {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const account = await exoClient().getTokenCache().getAccountByHomeId(t.homeAccountId)
  if (!account) throw new NeedsExoSignIn('Sign in to Exchange Online first')
  try {
    const r = await exoClient().acquireTokenSilent({ account, scopes: EXO_SCOPES, authority: `https://login.microsoftonline.com/${tenantId}` })
    return { token: r.accessToken, expiresOn: r.expiresOn?.getTime() || Date.now() + 3000000 }
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) throw new NeedsExoSignIn(e.message)
    throw e
  }
}

export async function exoSignIn(tenantId: string) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const r = await exoClient().acquireTokenInteractive({
    scopes: EXO_SCOPES,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    loginHint: t.username,
    openBrowser: async (url) => {
      await shell.openExternal(url)
    },
    successTemplate: page('Exchange Online: signed in / aangemeld', 'You can close this tab and return to M365 Admin Toolbox.'),
    errorTemplate: page('Sign-in failed / Aanmelden mislukt', 'Return to M365 Admin Toolbox for details.')
  })
  if (r.account && r.account.homeAccountId !== t.homeAccountId) {
    // Another account was picked in the browser; keep it out of this tenant's cache entry.
    await exoClient().getTokenCache().removeAccount(r.account)
    throw new Error(`Sign in as ${t.username}`)
  }
}
