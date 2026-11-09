# M365 Admin Toolbox

Multi-tenant Microsoft 365 admin desktop app for Windows, by RHIS.

## What it does

- Add tenants with a Microsoft sign-in, either with your own partner account through **GDAP** or with the tenant's own admin account. Switch between tenants from the dropdown top-left.
- **Role activation (PIM)**: see eligible, active and permanent admin roles, activate with a reason, duration and ticket number (with MFA or authentication-context step-up), extend, deactivate, cancel pending requests, and get a notification 10 minutes before a role expires. For GDAP tenants this is PIM for groups in the RHIS partner tenant, narrowed to the groups of that customer's GDAP relationship. A "no permission" error offers role activation directly.
- **Users**: list, create with a generated password, block or allow sign-in, reset password, create a Temporary Access Pass, revoke sessions.
- **Licenses**: assign and remove licenses for one or more users, see purchased versus assigned per subscription, and spot shortages in one tenant or across all tenants.
- **Groups**: list, create Microsoft 365 and security groups.
- **Info**: version, update check and install, production or test update channel, language (English/Nederlands), theme.

### Exchange (phase 2)

- **Mailboxes**: size, quota, percentage full (warning above 90%), archive, forwarding (external forwarding in red) and last activity. Per mailbox: Full Access, Send As and Send on behalf, forwarding, automatic replies, archive, retention policy and holds, and for a user the other mailboxes they can open.
- **Shared mailboxes**: create with members (Full Access and Send As), convert a user mailbox to shared and remove its licenses, or back.
- **Rooms & equipment**: create, capacity, and booking settings (auto accept, conflicts, booking window, maximum duration).
- **Distribution lists**: distribution lists, mail-enabled security groups and dynamic lists; members, owners and external senders.
- **Addresses & aliases**: all addresses per domain, change the primary address, add and remove aliases.
- **Inbox rule check**: scans every mailbox for rules that forward, redirect, delete or hide mail, plus mailbox forwarding to external addresses; disable or remove rules.
- **Mail flow**: read-only list of mail flow rules and accepted domains.

Teams and Entra modules are shown in the menu with the phase they are planned for.

## How it works

- Electron + React + TypeScript (electron-vite).
- Sign-in with MSAL (`@azure/msal-node`) in the system browser. The token cache is encrypted with Windows DPAPI (Electron `safeStorage`) under `%APPDATA%\M365 Admin Toolbox`.
- Data through Microsoft Graph v1.0.
- Default app (client) ID is *Microsoft Graph Command Line Tools* (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). It can be changed under Info → Advanced. The first sign-in per tenant asks an admin to consent to the requested Graph permissions.
- Exchange runs through the ExchangeOnlineManagement PowerShell module (v3). The app checks for it at start and installs it for the current Windows user when missing (no admin rights needed). One hidden Windows PowerShell process per tenant stays connected for 20 minutes after the last command. It connects with `Connect-ExchangeOnline -AccessToken`, using a token for the Exchange Online PowerShell app (`fb78d390-0c51-40cd-8e17-fdbfab77341b`); the first Exchange screen per tenant asks for that sign-in once. The role needed is Exchange Administrator.
- Mailbox sizes come from the Microsoft 365 usage report (Graph, `Reports.Read.All`, one or two days behind) when the tenant allows it, and otherwise live from Exchange.
- Temporary Access Pass only works when that method is enabled in the tenant's authentication methods policy.

## Branches, channels and releases

| Branch | GitHub Release | Version | Update channel in the app |
| --- | --- | --- | --- |
| `test` | pre-release | `<version>-beta.<commits>` | Test |
| `main` | release | `<version>` | Production |

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

Set `TOOLBOX_USER_DATA` to a folder to run `npm run dev` with a separate profile next to an installed copy.

`npm run release` uses `GH_TOKEN`, or the token of the logged-in GitHub CLI.

---

Powered by RHIS.
