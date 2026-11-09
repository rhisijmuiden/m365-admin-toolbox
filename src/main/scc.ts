import helperSource from './scc/helper.ps1?raw'
import { sccTokens } from './auth'
import { psHelper } from './ps'
import { tenants } from './store'

/** Security & Compliance PowerShell (compliance search and purge): one process per tenant, like Exchange. */

function organization(tenantId: string) {
  const t = tenants.get(tenantId)!
  return t.domains.find((d) => /^[^.]+\.onmicrosoft\.com$/i.test(d)) || t.defaultDomain
}

const helper = psHelper({
  name: 'scc',
  source: helperSource,
  async connect(tenantId) {
    const { tokens, key } = await sccTokens(tenantId)
    return { key, args: { Tokens: tokens, Org: organization(tenantId) } }
  }
})

export const { run, closeSession, closeAll } = helper
