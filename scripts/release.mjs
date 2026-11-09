// Builds the Windows installer and publishes it as a GitHub Release.
//   test branch -> pre-release  v<version>-beta.<commit count>  (test update channel)
//   main branch -> release      v<version>                      (production update channel)
// Needs GH_TOKEN, or a logged-in GitHub CLI (`gh auth token`).
import { execSync } from 'child_process'
import { readFileSync } from 'fs'

const sh = (c) => execSync(c, { encoding: 'utf8' }).trim()
const run = (c, env = {}) => execSync(c, { stdio: 'inherit', env: { ...process.env, ...env } })

const branch = process.env.RELEASE_BRANCH || sh('git rev-parse --abbrev-ref HEAD')
if (!['main', 'test'].includes(branch)) {
  console.error(`Releases are only built from main or test (current: ${branch})`)
  process.exit(1)
}
if (sh('git status --porcelain')) {
  console.error('Working tree is not clean, commit first')
  process.exit(1)
}

const base = JSON.parse(readFileSync('package.json', 'utf8')).version
const version = branch === 'test' ? `${base}-beta.${sh('git rev-list --count HEAD')}` : base
const token = process.env.GH_TOKEN || sh('gh auth token')

console.log(`Releasing ${version} from ${branch}`)
run('npm run build')
run(
  `npx electron-builder --win nsis --publish always -c.extraMetadata.version=${version} -c.publish.releaseType=${branch === 'test' ? 'prerelease' : 'release'}`,
  { GH_TOKEN: token }
)
