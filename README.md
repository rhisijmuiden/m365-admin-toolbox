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

### Teams (phase 3)

- **Teams & channels**: all teams with owners, privacy, last activity (Teams usage report) and archive state; filters for teams without an owner or with only one, inactive teams (90+ days) and archived teams. Create a team, empty or from a template, with owners, members and standard, private and shared channels. Per team: members and roles, channels and their members, name and privacy, archive and restore.
- **Guests**: every guest in every team, with when the guest account was created, last sign-in and invitation state; remove a guest from a team or delete the guest account.
- **Policies**: meeting, messaging and app setup policies per user (also several users at once) and per group with rank.
- **External & guest access**: federation with other organizations (all except blocked, or only allowed domains), Teams personal accounts, and what guests can do.
- **Phone**: phone numbers and who has them, assign and unassign numbers, Enterprise Voice, calling and voice routing policy (for tenants with Teams Phone).
- **Templates**: a template library (privacy and channels) stored on this PC and used for every tenant; start one from an existing team.

### Entra ID (phase 4)

- **App registrations**: register an app (supported accounts, redirect URIs, API permissions from Microsoft Graph, Exchange Online or SharePoint) and grant admin consent right away. Per app: IDs, redirect URIs per platform, API permissions with their consent state, client secrets (value shown once) and certificate upload (.cer/.crt/.pem), app roles and owners.
- **Expiring secrets**: secrets and certificates of app registrations and enterprise apps (SAML signing certificates) across all tenants, colored for expired, 30, 60 and 90 days; create a new secret straight from the list.
- **Enterprise apps**: enable or disable sign-in, assignment required, hide from My Apps; assign users and groups to app roles; owners; granted permissions.
- **Risky permissions**: apps with delegated or application permissions on mail, files or the whole directory, user consents; revoke a permission or block the app.
- **Conditional Access**: policies with state and grant, details with names instead of ids, turn on, report-only or off, and new policies from a form or from Microsoft's baseline templates (MFA for everyone, block legacy authentication, MFA for admins, Azure management, guests, outside trusted locations), always report-only by default.
- **Named locations**: IP ranges (trusted or not) and country lists.
- **Admin roles**: who has which role (permanent, time-bound or eligible), assign and remove with start and end dates, and for GDAP customers the relationships with their end date, roles and security groups, checked against the roles the toolbox needs.
- **PIM management**: approve or deny activation requests, role settings (maximum duration, MFA, reason, ticket, approval and approvers), PIM for groups (eligible members and owners) and the PIM audit log.
- **MFA methods**: registered methods per user (registration report, or per user without Entra ID P1), users without MFA or with phone only; remove methods, reset MFA and create a Temporary Access Pass.
- **Guest users**: invite guests, resend invitations, block sign-in, and clean up guests who never accepted (30 days) or did not sign in for 90 days.

## How it works

- Electron + React + TypeScript (electron-vite).
- Sign-in with MSAL (`@azure/msal-node`) in the system browser. The token cache is encrypted with Windows DPAPI (Electron `safeStorage`) under `%APPDATA%\M365 Admin Toolbox`.
- Data through Microsoft Graph v1.0.
- Default app (client) ID is *Microsoft Graph Command Line Tools* (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). It can be changed under Info → Advanced. The first sign-in per tenant asks an admin to consent to the requested Graph permissions.
- Exchange runs through the ExchangeOnlineManagement PowerShell module (v3). The app checks for it at start and installs it for the current Windows user when missing (no admin rights needed). One hidden Windows PowerShell process per tenant stays connected for 20 minutes after the last command. It connects with `Connect-ExchangeOnline -AccessToken`, using a token for the Exchange Online PowerShell app (`fb78d390-0c51-40cd-8e17-fdbfab77341b`); the first Exchange screen per tenant asks for that sign-in once. The role needed is Exchange Administrator.
- Teams, channels, members and guests go through Graph. The Teams endpoints need a few extra delegated scopes (`Team.Create`, `TeamMember.ReadWrite.All`, `TeamSettings.ReadWrite.All`, `Channel.Create`, `Channel.Delete.All`, `ChannelMember.ReadWrite.All`), which the Teams screen asks consent for once per tenant.
- Teams policies, external and guest access and phone numbers run through the MicrosoftTeams PowerShell module (v5 or newer), installed for the current Windows user on first use, in its own helper process per tenant like Exchange. It connects with `Connect-MicrosoftTeams -AccessTokens` using tokens for the Microsoft Teams PowerShell app (`12128f48-ec9e-42f0-b203-ea49fb6af367`); the first of these screens per tenant asks for that sign-in once. The role needed is Teams Administrator.
- Entra ID runs entirely through Graph. Its write scopes (`Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`, `DelegatedPermissionGrant.ReadWrite.All`, `Policy.Read.All`, `Policy.ReadWrite.ConditionalAccess`, `RoleManagement.ReadWrite.Directory`, the PIM schedule and policy scopes, `User.Invite.All`) are asked for once per tenant from the Entra screens. Roles needed: Application Administrator, Conditional Access Administrator, Privileged Role Administrator and Authentication Administrator. Conditional Access needs Entra ID P1, PIM needs Entra ID P2. Approving directory role requests uses the Graph beta endpoint, because v1.0 has no role approval API yet.
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
