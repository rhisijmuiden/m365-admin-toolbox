# M365 Admin Toolbox

Multi-tenant Microsoft 365 admin desktop app for Windows, by RHIS.

## What it does (phase 1)

- Add tenants with a Microsoft sign-in, either with your own partner account through **GDAP** or with the tenant's own admin account. Switch between tenants from the dropdown top-left.
- **Role activation (PIM)**: see eligible, active and permanent admin roles, activate with a reason, duration and ticket number (with MFA or authentication-context step-up), extend, deactivate, cancel pending requests, and get a notification 10 minutes before a role expires. For GDAP tenants this is PIM for groups in the RHIS partner tenant, narrowed to the groups of that customer's GDAP relationship. A "no permission" error offers role activation directly.
- **Users**: list, create with a generated password, block or allow sign-in, reset password, create a Temporary Access Pass, revoke sessions.
- **Licenses**: assign and remove licenses for one or more users, see purchased versus assigned per subscription, and spot shortages in one tenant or across all tenants.
- **Groups**: list, create Microsoft 365 and security groups.
- **Info**: version, update check and install, production or test update channel, language (English/Nederlands), theme.

Exchange, Teams and Entra modules are shown in the menu with the phase they are planned for.

## How it works

- Electron + React + TypeScript (electron-vite).
- Sign-in with MSAL (`@azure/msal-node`) in the system browser. The token cache is encrypted with Windows DPAPI (Electron `safeStorage`) under `%APPDATA%\M365 Admin Toolbox`.
- Data through Microsoft Graph v1.0.
- Default app (client) ID is *Microsoft Graph Command Line Tools* (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). It can be changed under Info → Advanced. The first sign-in per tenant asks an admin to consent to the requested Graph permissions.
- Temporary Access Pass only works when that method is enabled in the tenant's authentication methods policy.

## Branches, channels and releases

| Branch | GitHub Release | Version | Update channel in the app |
| --- | --- | --- | --- |
| `test` | pre-release | `0.1.0-beta.<commits>` | Test |
| `main` | release | `0.1.0` | Production |

The app checks for updates through `electron-updater` with the GitHub provider. Switching back from Test to Production installs the newest stable release, even if it is older.

## Development

```sh
npm install
npm run dev        # run the app with hot reload
npm run typecheck
npm run dist       # build the Windows installer into dist/ without publishing
npm run release    # build and publish a GitHub Release for the current branch (main or test)
npm run icon       # regenerate build/icon.png from the RHIS mark
```

`npm run release` uses `GH_TOKEN`, or the token of the logged-in GitHub CLI.

---

Powered by RHIS.
