import helperSource from './exo/helper.ps1?raw'
import { exoToken } from './auth'
import { psHelper } from './ps'
import { tenants } from './store'

/** Exchange Online: one PowerShell process per tenant, connected with Connect-ExchangeOnline -AccessToken. */

function organization(tenantId: string) {
  const t = tenants.get(tenantId)!
  return t.domains.find((d) => /^[^.]+\.onmicrosoft\.com$/i.test(d)) || t.defaultDomain
}

const helper = psHelper({
  name: 'exo',
  source: helperSource,
  async connect(tenantId) {
    const { token } = await exoToken(tenantId)
    return { key: token, args: { Token: token, Org: organization(tenantId) } }
  }
})

export const { ensureModule, run, closeSession, closeAll } = helper
