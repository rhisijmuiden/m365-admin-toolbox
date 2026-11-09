import { shell } from 'electron'
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type ICachePlugin
} from '@azure/msal-node'
import { exoTokenCache, settings, teamsTokenCache, tenants, tokenCache } from './store'

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
  for (const c of [exoClient(), teamsClient()]) {
    const a = await c.getTokenCache().getAccountByHomeId(homeAccountId)
    if (a) await c.getTokenCache().removeAccount(a)
  }
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
 * PowerShell modules: Connect-ExchangeOnline and Connect-MicrosoftTeams take access tokens issued to the
 * module's own first-party app, because a Microsoft app only gets tokens for the Microsoft APIs it is
 * pre-authorized for. Each is a separate public client with its own consent and cache; the user signs in
 * once per tenant (usually SSO in the browser) and refresh tokens do the rest.
 */
function moduleClient(clientId: string, cache: typeof tokenCache) {
  let pca: PublicClientApplication | null = null
  return () => {
    if (!pca)
      pca = new PublicClientApplication({
        auth: { clientId, authority: 'https://login.microsoftonline.com/organizations' },
        cache: { cachePlugin: pluginFor(cache) }
      })
    return pca
  }
}

async function moduleToken(c: PublicClientApplication, tenantId: string, scopes: string[], Err: new (m: string) => Error) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const account = await c.getTokenCache().getAccountByHomeId(t.homeAccountId)
  if (!account) throw new Err('Sign in first')
  try {
    const r = await c.acquireTokenSilent({ account, scopes, authority: `https://login.microsoftonline.com/${tenantId}` })
    return { token: r.accessToken, expiresOn: r.expiresOn?.getTime() || Date.now() + 3000000 }
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) throw new Err(e.message)
    throw e
  }
}

async function moduleSignIn(c: PublicClientApplication, tenantId: string, scopes: string[], extra: string[], title: string) {
  const t = tenants.get(tenantId)
  if (!t) throw new Error(`Unknown tenant ${tenantId}`)
  const r = await c.acquireTokenInteractive({
    scopes,
    extraScopesToConsent: extra.length ? extra : undefined,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    loginHint: t.username,
    openBrowser: async (url) => {
      await shell.openExternal(url)
    },
    successTemplate: page(`${title}: signed in / aangemeld`, 'You can close this tab and return to M365 Admin Toolbox.'),
    errorTemplate: page('Sign-in failed / Aanmelden mislukt', 'Return to M365 Admin Toolbox for details.')
  })
  if (r.account && r.account.homeAccountId !== t.homeAccountId) {
    // Another account was picked in the browser; keep it out of this tenant's cache entry.
    await c.getTokenCache().removeAccount(r.account)
    throw new Error(`Sign in as ${t.username}`)
  }
}

export const EXO_CLIENT_ID = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'
const EXO_SCOPES = ['https://outlook.office365.com/.default']
const exoClient = moduleClient(EXO_CLIENT_ID, exoTokenCache)

export class NeedsExoSignIn extends Error {
  code = 'NEEDS_EXO_SIGNIN'
}

export const exoToken = (tenantId: string) => moduleToken(exoClient(), tenantId, EXO_SCOPES, NeedsExoSignIn)
export const exoSignIn = (tenantId: string) => moduleSignIn(exoClient(), tenantId, EXO_SCOPES, [], 'Exchange Online')

/** Microsoft Teams PowerShell connects with a Graph token and one for the Skype and Teams Tenant Admin API. */
export const TEAMS_CLIENT_ID = '12128f48-ec9e-42f0-b203-ea49fb6af367'
const TEAMS_ADMIN_SCOPES = ['48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default']
const TEAMS_GRAPH_SCOPES = ['https://graph.microsoft.com/.default']
const teamsClient = moduleClient(TEAMS_CLIENT_ID, teamsTokenCache)

export class NeedsTeamsSignIn extends Error {
  code = 'NEEDS_TEAMS_SIGNIN'
}

export async function teamsTokens(tenantId: string) {
  const admin = await moduleToken(teamsClient(), tenantId, TEAMS_ADMIN_SCOPES, NeedsTeamsSignIn)
  const graph = await moduleToken(teamsClient(), tenantId, TEAMS_GRAPH_SCOPES, NeedsTeamsSignIn)
  return { admin: admin.token, graph: graph.token }
}

export const teamsSignIn = (tenantId: string) => moduleSignIn(teamsClient(), tenantId, TEAMS_ADMIN_SCOPES, TEAMS_GRAPH_SCOPES, 'Microsoft Teams')
