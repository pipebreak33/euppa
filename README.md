# Euppa

**English** · [Svenska](README.sv.md)

<p align="center">
  <img src="assets/euppa.svg" alt="Euppa: Hi! I'm Euppa." width="720">
</p>

Everything you need to do in **Entra ID**, in a few keystrokes instead of hunting in the portal or writing Graph calls.
Works the same on **Windows, macOS and Linux**. One of the tools in the troupe with [Uppa](https://github.com/pipebreak33/uppa),
[Puppa](https://github.com/pipebreak33/puppa) and [Muppa](https://github.com/pipebreak33/muppa), and it has its own
page in [Ruppa](https://github.com/pipebreak33/ruppa). The mascot is an E with eyes: it suggests the next step, blinks and
looks around while something runs, is happy when it went well and worried before something changes.

## Install (or update: run the same line again)

**Windows** (PowerShell):

```powershell
irm https://raw.githubusercontent.com/pipebreak33/euppa/main/install.ps1 | iex
```

**macOS / Linux:**

```bash
curl -fsSL https://raw.githubusercontent.com/pipebreak33/euppa/main/install.sh | sh
```

The installer finds (or on Windows installs) Python 3.8+, puts Euppa in `~/.euppa/app` and creates the command
`euppa` (`~/.local/bin/euppa`, or `%USERPROFILE%\.euppa\bin\euppa.cmd` + PATH on Windows). Your tenants and
settings in `~/.euppa` are never touched by an update. It warns if PowerShell 7 or the module is missing.
With [Wuppa](https://github.com/pipebreak33/wuppa) or [Ruppa](https://github.com/pipebreak33/ruppa), Euppa comes
along automatically.

Choose the source with `EUPPA_REPO=owner/repo` and `EUPPA_BRANCH=branch`. From a local copy: `./install.sh` / `./install.ps1`.

Needs PowerShell 7 (`pwsh`) and **only one module**: `Microsoft.Graph.Authentication`. Everything goes through
`Invoke-MgGraphRequest` to Microsoft Graph, so none of the heavy Graph submodules are needed.
`ImportExcel` is only needed for the Excel export (otherwise you get CSV).

## Uninstall

**Windows** (PowerShell):

```powershell
irm https://raw.githubusercontent.com/pipebreak33/euppa/main/uninstall.ps1 | iex
```

**macOS / Linux:**

```bash
curl -fsSL https://raw.githubusercontent.com/pipebreak33/euppa/main/uninstall.sh | sh
```

Removes the program and the `euppa` command (and the PATH entry on Windows). Your tenants and settings in `~/.euppa`, and the reports in `~/Documents/Euppa`, are kept.
`EUPPA_PURGE=1` removes your data too, and `EUPPA_YES=1` skips the question.

## How it works

- **Users first.** *Users…* lists everyone in the tenant. Type to fuzzy-search, tick one or more with
  space (`ctrl+a` ticks everyone shown) and choose what should happen: add to a group, remove from a group,
  assign a license, block, reset the password or MFA, give a role … Everything runs at once and is confirmed once.
- **Areas.** 83 actions in 11 areas: users, guests, groups, devices, apps, roles and
  PIM, Conditional Access, logs, security and the tenant. Start typing in any menu to search everything.
- **Search the tenant.** When an action asks for a user, group, device, app, role, policy or license, part
  of the name is enough. One match is picked directly, several give a list (checkboxes when several can be chosen).
- **The right scopes per task.** Euppa signs in with `Connect-MgGraph` and exactly the scopes the action needs, and
  adds more when needed without losing the old ones. *Sign in…* gives read access to almost everything at once.
- **Tenants as in Puppa.** Signing in saves the tenant (domain, id, account) in *Tenants…*, so you can switch between
  customers. The first time, Euppa borrows Puppa's list. Choose **browser** or **device code** (good over SSH) per tenant.
- **Reading first, changes with a yes.** Read actions run directly. Actions marked **changes** show the whole command and
  which tenant it applies to, and need a **y** (or `--yes` on the command line). `-n` only shows what would run.
- **Excel or CSV.** After a read action: `x` for Excel, `c` for CSV, in `~/Documents/Euppa`.
- **No keychain prompt.** On macOS the sign-in is kept only during the run by default (`-ContextScope
  Process`). Turn on *Remember sign-ins* (or `euppa remember on`) to avoid signing in every time.

## Language

Euppa speaks **English** by default and **Swedish** on request: menus, actions, results and error messages. The choice is
shared by the whole troupe: `euppa lang sv` or `euppa lang en`, or the **Language** row in the menu. `TRUPPA_LANG=sv` sets it
for a single run. Swedish commands, action ids and field names (`euppa kör anvandare-sparra Anvandare=…`) still work.

## Commands

| | |
|---|---|
| `euppa` | the menu |
| `euppa users [UPN …]` | pick users (multi-select) and do something with them |
| `euppa actions [--json]` | every action per area (`--json`: areas, actions and fields for Ruppa) |
| `euppa run ACTION [Name=value …]` | run an action; asks for anything missing |
| `  --yes` | confirm actions that change something without asking (required without a terminal) |
| `  --excel` / `--csv` / `--json` | save the result in `~/Documents/Euppa`, or print it as JSON |
| `euppa lookup TYPE [search] [--all] [--json]` | search the tenant: `user`, `group`, `device`, `app`, `sp`, `role`, `ca`, `sku`, `deleted` |
| `euppa login [--scopes A,B] [--device-code] [--tenant X]` | sign in to Microsoft Graph |
| `euppa tenants` / `euppa tenant NAME` / `euppa new` | saved tenants, switch, add |
| `euppa remember on\|off` | keep sign-ins between runs |
| `euppa logout` | `Disconnect-MgGraph` |
| `euppa lang [en\|sv]` | language for the whole troupe |
| `-n`, `--dry-run` | show the commands without running anything |
| `--no-anim` | no animations (also off with `NO_COLOR` and when the output is piped) |

Several values are written comma-separated. Examples:

```bash
euppa run groups-add Users=anna@contoso.com,bo@contoso.com Group=Sales --yes
euppa run apps-expiring-secrets Days=60 --excel
euppa run ca-state Policy="Require MFA for admins" State=Report-only --yes
euppa lookup user --all --json
```

## What Euppa can do

### Users

| Action | | What it does |
|---|---|---|
| `users-list` | read | **List users.** Every user with status, type, title, department and number of licenses. |
| `users-info` | read | **User info.** Everything about one or more users: contact, manager, status, sync, licenses, groups and last sign-in. |
| `users-inactive` | read | **Inactive users.** Accounts that haven't signed in for X days (or ever). Needs Entra ID P1 for sign-in dates. |
| `users-create` | changes | **Create user.** New cloud user. Leave the password empty and Euppa makes up a strong one and shows it once. |
| `users-edit` | changes | **Edit user field.** Set job title, department, office, mobile, company etc. on one or more users. An empty value clears the field. |
| `users-block` | changes | **Block and sign out.** Blocks sign-in and revokes every session (token) right away. Your own account is skipped. |
| `users-enable` | changes | **Enable account.** Turns sign-in back on for blocked accounts. |
| `users-reset-password` | changes | **Reset password.** New password for one or more users. Empty = a unique, strong password per person, shown once. |
| `users-sign-out` | changes | **Sign out everywhere.** Revokes every session and refresh token. The user has to sign in again on every device. |
| `users-mfa` | read | **MFA methods.** Which sign-in methods the users have registered: Authenticator, phone, FIDO2 key, Windows Hello … |
| `users-mfa-reset` | changes | **Reset MFA.** Removes every registered MFA method (the password stays) and signs out. The user registers again at next sign-in. |
| `users-tap` | changes | **Temporary Access Pass.** Creates a Temporary Access Pass, e.g. to register MFA or a new PC. Needs TAP to be enabled in the tenant. |
| `users-licenses` | read | **User's licenses.** Licenses per user, and whether they are assigned directly or through a group. |
| `users-license-assign` | changes | **Assign license.** Gives the users a license. Anyone without a usage location gets one first (required for licenses). |
| `users-license-remove` | changes | **Remove license.** Removes a directly assigned license. Group-based licenses go away by removing the user from the group. |
| `users-manager` | changes | **Set manager.** Sets the manager for one or more users. |
| `users-groups` | read | **User's groups and roles.** Groups, admin roles and administrative units the users are direct members of. |
| `users-devices` | read | **User's devices.** Devices the users own or have registered. |
| `users-delete` | changes | **Delete user.** Deletes the users. They stay under Deleted users for 30 days and can be restored. Your own account is skipped. |
| `users-deleted` | read | **Deleted users.** Users in the recycle bin and how many days are left before they're gone for good. |
| `users-restore` | changes | **Restore deleted user.** Brings deleted users back with licenses and group memberships. |

### Guests

| Action | | What it does |
|---|---|---|
| `guests-list` | read | **List guests.** Every guest account: invitation accepted or not, when they were invited and last sign-in. |
| `guests-pending` | read | **Guests who haven't accepted.** Invitations left unaccepted for more than X days – good candidates to clean up. |
| `guests-invite` | changes | **Invite guest.** Invites one or more external users (comma-separated email addresses) and shows the redemption link. |

### Groups

| Action | | What it does |
|---|---|---|
| `groups-list` | read | **List groups.** Every group with type (Microsoft 365, security, distribution), dynamic or assigned membership and email. |
| `groups-members` | read | **Group members.** Who is in the group – directly, or also through nested groups. |
| `groups-owners` | read | **Group owners.** The owners of one or more groups. |
| `groups-without-owners` | read | **Groups without owners.** Groups nobody owns – a common clean-up item. |
| `groups-dynamic` | read | **Dynamic groups.** Groups with dynamic membership, their rule and whether the rule is on or paused. |
| `groups-create` | changes | **Create group.** New security group or Microsoft 365 group, with owners and members right away. |
| `groups-add` | changes | **Add to group.** Adds the users as members of one or more groups. Already a member counts as done. |
| `groups-remove` | changes | **Remove from group.** Removes the users from one or more groups. |
| `groups-owners-add` | changes | **Add owner.** Makes the users owners of one or more groups. |
| `groups-owners-remove` | changes | **Remove owner.** Removes the users as owners. A group with owners must keep at least one. |
| `groups-dynamic-rule` | changes | **Set dynamic rule.** Makes the group dynamic with a rule, e.g. (user.department -eq "Sales"). The rule replaces assigned members. |
| `groups-delete` | changes | **Delete group.** Deletes the groups. Microsoft 365 groups can be restored for 30 days, security groups can't. |

### Devices

| Action | | What it does |
|---|---|---|
| `devices-list` | read | **List devices.** Every device in Entra ID: OS, join type (Entra joined, hybrid, registered), status and last sign-in. |
| `devices-inactive` | read | **Inactive devices.** Devices that haven't signed in for X days – candidates to disable and then delete. |
| `devices-disable` | changes | **Disable device.** Disables the devices in Entra ID (they can't sign in) without deleting them. |
| `devices-enable` | changes | **Enable device.** Turns disabled devices back on. |
| `devices-delete` | changes | **Delete device.** Deletes the devices from Entra ID. BitLocker keys and LAPS passwords go with them – read them out first. |
| `devices-bitlocker` | read | **BitLocker keys.** Recovery keys for the devices. Every lookup is recorded in the audit log. |
| `devices-laps` | read | **LAPS password.** The local admin password (Windows LAPS in Entra ID) for the devices. Every lookup is logged. |

### Apps

| Action | | What it does |
|---|---|---|
| `apps-list` | read | **App registrations.** Every app registration with the number of secrets and certificates and when the first one expires. |
| `apps-expiring-secrets` | read | **Expiring secrets and certificates.** Client secrets and certificates (SAML signing on enterprise apps too) that expire within X days or already have. |
| `apps-permissions` | read | **App permissions.** API permissions the app requests, delegated and application, and whether they have consent. |
| `apps-secret-new` | changes | **New client secret.** Creates a new secret on the app. The value is shown only this once – copy it right away. |
| `apps-secrets-cleanup` | changes | **Remove expired secrets.** Removes client secrets that have already expired from the apps. |
| `enterprise-apps-list` | read | **Enterprise apps.** Enterprise applications: on/off, whether assignment is required, and whether the app is your own or external. |
| `enterprise-apps-consents` | read | **Consents (delegated).** Which apps have delegated permissions, granted by an admin for everyone or by individual users. |
| `enterprise-apps-access` | read | **Who has access to the app.** Users, groups and apps assigned to the enterprise app, with role. |
| `enterprise-apps-grant-consent` | changes | **Grant admin consent.** Grants consent for the whole organisation to every permission the app requests (like the button in the portal). |
| `enterprise-apps-revoke-consent` | changes | **Revoke consent.** Removes every consent (delegated and/or application permissions) the enterprise app has received. |
| `enterprise-apps-toggle` | changes | **Turn an enterprise app off or on.** Off = nobody can sign in to the app (good for a suspicious app). On = back again. |

### Roles and PIM

| Action | | What it does |
|---|---|---|
| `roles-list` | read | **Who has admin roles.** Every active role assignment: role, who and scope. |
| `roles-user` | read | **User's roles.** Active and eligible (PIM) roles for the users. |
| `pim-eligible` | read | **Eligible roles (PIM).** Everyone who can activate a role in PIM, and for how long. Needs Entra ID P2. |
| `pim-active` | read | **Active PIM assignments.** Roles that are active right now – permanent and activated through PIM – with end time. Needs Entra ID P2. |
| `roles-assign` | changes | **Assign role.** Gives the users an admin role – preferably as eligible in PIM (activated when needed) instead of permanent. |
| `roles-remove` | changes | **Remove role.** Removes the role from the users – both active and eligible (PIM) assignments. |
| `pim-activate` | changes | **Activate my role (PIM).** Activates a role you're eligible for, for a few hours. Often needs MFA and a justification. |

### Conditional Access

| Action | | What it does |
|---|---|---|
| `ca-list` | read | **Conditional Access policies.** Every policy: on, off or report-only, who and which apps it applies to and what it requires. |
| `ca-detail` | read | **Policy in plain words.** A policy broken down into who, apps, conditions and controls – with names instead of ids. |
| `ca-state` | changes | **Change policy state.** Put policies in report-only (logs what would have happened), on or off. |
| `ca-exclude` | changes | **Exclude users from policy.** Adds users as exclusions in policies, e.g. break-glass accounts. |
| `ca-locations` | read | **Named locations.** IP ranges and countries used in the policies, and which ones are trusted. |
| `ca-location-new` | changes | **New named location.** Creates a location from IP ranges (CIDR, e.g. 203.0.113.0/24) or countries (SE, NO …). |

### Logs

| Action | | What it does |
|---|---|---|
| `logs-signins` | read | **Sign-ins.** The sign-in log: who, app, IP, location, result and Conditional Access. Needs Entra ID P1 (max 999 rows). |
| `logs-failed` | read | **Failed sign-ins per user.** Summary: who failed, how many times, with which error and from how many IPs. |
| `logs-audit` | read | **Audit log.** Who changed what in the directory: users, groups, roles, apps, policies (max 999 rows). |

### Security

| Action | | What it does |
|---|---|---|
| `risk-users` | read | **Risky users.** Users Entra ID Protection has flagged as risky. Needs Entra ID P2. |
| `risk-detections` | read | **Risk detections.** Detected risks: unfamiliar location, leaked credentials, impossible travel etc. Needs Entra ID P2. |
| `risk-dismiss` | changes | **Dismiss user risk.** Marks the risk as handled (e.g. after a password change). |
| `risk-compromised` | changes | **Confirm compromised.** Sets the user risk to high (confirmed compromised) so risk policies kick in. Block the account too. |
| `mfa-status` | read | **MFA status in the tenant.** Who has MFA registered, with which methods, and whether they are admins. Needs Entra ID P1. |

### Tenant

| Action | | What it does |
|---|---|---|
| `tenant-licenses` | read | **License overview.** Every license: purchased, used, available and status. |
| `tenant-license-users` | read | **Who has the license.** Every user with a given license, and whether it's direct or through a group. |
| `tenant-domains` | read | **Domains.** The domains in the tenant: verified, default, federated and what they're used for. |
| `tenant-info` | read | **Organisation info.** Name, tenant id, country, created, directory sync and whether security defaults are on. |
| `tenant-settings` | read | **Organisation settings.** What users are allowed to do: register apps, create groups and tenants, invite guests, consent to apps etc. |
| `tenant-setting-change` | changes | **Change organisation setting.** Turn on or off: app registration, creating groups and tenants, guest invitations, user consent, security defaults. |

### Local

| Action | | What it does |
|---|---|---|
| `local-session` | local | **My sign-in.** Who you're signed in as, in which tenant and with which scopes. |
| `local-modules` | local | **Modules.** PowerShell and the modules Euppa uses, with version. |

## Permissions and licenses

Euppa uses the app *Microsoft Graph Command Line Tools* with delegated permissions: you can never do more than your
account is allowed to. The first time a new scope is needed, Microsoft asks for consent (a Global Admin can approve it for
the whole organisation). Sign-in and audit logs, MFA status and Conditional Access need **Entra ID P1**;
risky users, risk detections and PIM need **Entra ID P2**. Euppa tells you in plain words when something is missing.

## In Ruppa

[Ruppa](https://github.com/pipebreak33/ruppa) shows Euppa as its own page (⌘5) with areas, cards and forms.
Ruppa is only a shell: it reads areas, actions and fields from the installed `euppa.py` and runs
`euppa run … --json`, `euppa lookup … --json` and `euppa actions --json`. New actions in Euppa show up there directly.

## Technology

One Python file (standard library only, 3.8+) that keeps a PowerShell 7 process open, so the Graph sign-in
lives between actions. Every action is a small PowerShell script in `ACTIONS` in `euppa.py`, which talks to Graph
through a thin helper (`Graph 'users?$select=…' -All`) with paging, plain-language error messages and lookups
from names to ids. Data in `~/.euppa` (`kunder.json`, `config.json`): no secrets, no tokens.

## License

MIT
