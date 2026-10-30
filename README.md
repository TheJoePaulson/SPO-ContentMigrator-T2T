# Content Migrator T2T - Community Edition

A Windows desktop tool for SharePoint Online list and library migrations, within one tenant or
from one Microsoft 365 tenant to another (**T2T**, tenant-to-tenant): inventory, copy and move with folder structure and
metadata, CSV-driven runs, list item copy with attachments, retry from log, and
"create from existing" / "modernize" for lists.

![Content Migrator T2T - Main App Window](https://i.imgur.com/cra3sEB.png "Content Migrator T2T - Main App Window")

Built on .NET 8 (WPF) with Microsoft's own libraries only:

| Layer | Library |
|---|---|
| UI | WPF, dark theme (`Theme/Dark.xaml`) |
| SharePoint | `Microsoft.SharePointOnline.CSOM` |
| Sign-in | `Microsoft.Identity.Client` (MSAL), interactive, system browser |

No PnP, no PowerShell, no third-party packages.

## 1. Register an app in Entra ID (required, once per tenant)

The Community Edition ships **without an App ID**. Each tenant uses its own app registration:

1. Entra admin center → **App registrations → New registration**. Any name (e.g. "Content Migrator T2T"),
   *Accounts in this organizational directory only*.
2. **Authentication → Add a platform → Mobile and desktop applications** → add `http://localhost`.
   Set **Allow public client flows** to **Yes**.
3. **API permissions → Add → SharePoint → Delegated** → `AllSites.FullControl`, then
   **Grant admin consent**.
4. Copy the **Application (client) ID** into the app: **Settings → Source tenant → App (client) ID**.

For cross-tenant runs, do the same in the destination tenant and put that ID in **Settings → Destination
tenant**. Alternatively, register one *multi-tenant* app, consent it in both tenants, and leave the
destination App ID blank so the source one is reused.

No client secret or certificate is needed. The operator signs in interactively.

## 2. Build

Requirements: Windows, the .NET 8 SDK, and NuGet access.

```powershell
cd ContentMigrator.CommunityEdition
dotnet publish -c Release -r win-x64 --self-contained false -o .\publish
```

Run `publish\ContentMigrator-T2T-CE.exe`. Machines that run it need the **.NET 8 Desktop Runtime (x64)**.
Package versions float (`16.1.*`, `4.*`); pin them in the `.csproj` after your first restore if you
want repeatable builds.

## 3. First run

`appsettings.json` is created **next to the .exe** and Settings opens. Enter:

- **Source tenant → Tenant name**: `contoso`, `contoso.onmicrosoft.com` or the tenant GUID.
- **App (client) ID**: from step 1.
- **Sign-in account** (optional): pre-fills the sign-in page; the log warns if a run uses another account.
- **Folders**: Output root (default `C:\ContentMigrator`), plus optional separate log and report folders.

Settings can be edited in the app or by hand; use **Reload from file** after a hand edit.
`appsettings.sample.json` lists every key with its default. The folder holding the .exe must be
writable for Save to work. No secrets are stored.

## 4. Permissions: be a Site Collection Admin

The signed-in account must be a **Site Collection Admin on every source and destination site**. The
SharePoint Administrator role alone is not enough, and lists with unique permissions are invisible
to non-admins. The app checks this on every site it connects to and stops with a clear message
before touching data (Settings → Advanced → *Require site collection admin*). In a cross-tenant run
this applies to the source account on the source site **and** the destination account on the
destination site.

## 5. Cross-tenant migration (tenant A → tenant B)

1. **Settings → Destination tenant**: tenant name (required), App ID and redirect URI (blank = same
   as the source), and optionally the destination sign-in account.
2. **Settings → Cross-tenant**: how people are matched (see below), the upload chunk size, and
   whether to keep the staged files.
3. On the main window, enter the destination site's **full URL** (or `/sites/Name`, resolved against
   the destination tenant) and tick **Destination is in another tenant**. The footer shows both tenants.
4. Run. You sign in twice, once per tenant; each sign-in is remembered until the app closes.

How it differs from a same-tenant run:

| | Same tenant | Cross-tenant |
|---|---|---|
| Files | SharePoint copy/move jobs, version history kept | Downloaded from the source and uploaded to the destination through this PC. **Latest version only.** Size verified after upload. |
| Move (2, 4, 7) | Available | **Blocked.** Nothing is ever deleted from the source tenant; use Copy. |
| People columns, Created By, Modified By | Same accounts | **Mapped**: user mapping CSV → domain mapping → fallback account |
| Lists (6, 9, 10) | | Same as same-tenant, with people mapped |

**People mapping** (Settings → Cross-tenant), in this order:
1. **User mapping CSV**: columns `SourceUser,DestinationUser`. *Generate CSV Template* writes
   `UserMapping_Template.csv` as a starting point.
2. **Domain mappings**: one per line, e.g. `contoso.com=fabrikam.com`, so `jane@contoso.com` becomes
   `jane@fabrikam.com`.
3. **Fallback account**: used when the mapped account doesn't exist in the destination tenant. Blank =
   leave the person empty.

Every distinct person is written once to `UserMappingLog.csv` (method and result), and anyone
unmatched is also in `UnresolvedUsers.csv`.

Throughput depends on this PC's connection, because every byte passes through it. Run large
migrations from a machine with a fast, stable link.

## Operations

| # | Operation | Notes |
|---|---|---|
| 1 | Inventory a library | CSV of every file and folder |
| 2 / 3 | Move / copy a full library or folder | Same site: `MoveCopyUtil`. Cross site: SharePoint copy jobs. Cross-tenant: download/upload (copy only). Metadata replayed. |
| 4 / 5 | Move / copy from CSV | Use *Generate CSV Template* for the format |
| 6 | Copy list items | Column values, Created/Modified/Author/Editor, attachments |
| 7 | Retry failed file operations | Reads `OperationLog.csv`; Copy vs Move taken from the log |
| 8 | Generate CSV template | |
| 9 | Create destination from existing | Same base template, columns, content types, views. Review-only saves the schema XML only. |
| 10 | Modernize a list | New modern list, supported columns and views, then the items |

Safety built in: same-source/same-destination guard, site admin check, duplicate check, review-only
provisioning, throttling retries (HTTP 429/503 with Retry-After), and cancel at any time with logs saved.

## Output

Each run writes to a `<yyyyMMdd_HHmmss>_<Operation>` folder:

| Folder | Files |
|---|---|
| Output root | AttachmentStaging, TransferStaging (cross-tenant), ProvisioningSchema_*.xml |
| Log location | OperationLog, ErrorLog, MetadataLog, FolderCreationLog, UnresolvedUsers, UserMappingLog, ProvisioningLog, AttachmentLog, RunMetadata (.csv/.json), RunSummary.txt, RunLog.txt |
| Report location | Inventory_*.csv, MetadataCompatibility_*.csv, CSV templates (incl. UserMapping_Template.csv), RetryInput_*.csv |

Leave Log and Report location blank to keep everything in one run folder.

## Customising the look

All colours are the brushes at the top of `Theme/Dark.xaml`; change them there and rebuild.
The icon is `Resources/AppIcon.ico` (exe, title bar, taskbar) and `Resources/AppIcon.png` (banner).
