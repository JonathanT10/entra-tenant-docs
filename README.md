# Entra Tenant Docs

Living documentation for your Entra ID tenant. One read-only PowerShell script, four outputs from the same timestamped snapshot:

| Output | Who it's for |
|---|---|
| `docs/` — 9 numbered Markdown files | Admins, auditors, change records. Deterministic output with no timestamps in the section files, so **committing the folder to git turns every re-run into a config-drift diff.** Includes a computed change log. |
| `tenant.json` | Anything you build on top — the complete structured snapshot. |
| `report.html` | Everyone else. A self-contained, timestamped report — KPI tiles, trend sparklines, a "what changed" feed, license meters, credential-expiry status, Conditional Access at a glance. No server, no dependencies; open it in a browser, drop it on an intranet share. |
| `history/` | One JSON snapshot archived per run — the raw material behind trends and the change log. |

![Report screenshot](screenshot.png)

<sub>Dark mode: [screenshot-dark.png](screenshot-dark.png)</sub>

Built by an IT manager who got tired of tenant knowledge living in twelve portal blades and zero documents.

## Try it in 30 seconds (no tenant needed)

```powershell
.\Export-EntraTenantDocs.ps1 -SampleData -OutputPath .\demo
```

Renders the bundled sample tenant — open `demo\report.html` and browse `demo\docs\`.

## Document a real tenant

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser   # the only module it needs
.\Export-EntraTenantDocs.ps1                                       # connects, collects, writes .\tenant-docs\
```

Required scopes (all read-only): `Directory.Read.All`, `Policy.Read.All`, `RoleManagement.Read.Directory`, `Application.Read.All`. Everything is raw `Invoke-MgGraphRequest` calls — no dependency on the twenty-odd Graph submodules.

Re-render offline from a previous snapshot (no connection, byte-identical output):

```powershell
.\Export-EntraTenantDocs.ps1 -FromJson .\tenant-docs\tenant.json -OutputPath .\rerender
```

### History, trends, and the change log

Every run archives its snapshot to `history/` (default: under the output folder; `-HistoryPath` moves it, `-NoHistory` skips archiving). From two snapshots onward:

- **Trends** appear in the report — sparklines for members, guests, enforced CA policies, role assignments, app registrations, and credentials in the renewal window, each with its change since the previous snapshot.
- **What changed** appears in the report and in `docs/08-changelog.md` — the computed diff between snapshots, in English: who got a role, which CA policy flipped state, which dynamic rule was edited, what got bought.

Schedule it weekly and the change log writes itself:

```powershell
cd tenant-repo
Export-EntraTenantDocs.ps1 -OutputPath .
git add docs tenant.json history; git commit -m "Tenant snapshot $(Get-Date -Format yyyy-MM-dd)"
git diff HEAD~1 -- docs      # the same drift, as a raw diff
```

## What gets documented (the identity plane)

1. **Tenant** — org info, verified domains, license SKUs (purchased/assigned/available)
2. **Conditional Access** — every policy rendered readable (users, groups, roles, apps, platforms, locations, client apps, risk, grant and session controls — GUIDs resolved to names), plus named locations
3. **Directory roles** — permanent assignments per role
4. **Groups** — counts by type, role-assignable groups, every dynamic membership rule verbatim
5. **Authentication methods policy** — which methods are enabled, for whom
6. **User & guest settings** — the authorization policy in plain English (who can invite guests, who can register apps, guest access level)
7. **App registrations** — sign-in audience and credential expiry, soonest first
8. **Change log** — computed by diffing consecutive snapshots: CA policies added/removed/state-changed, role assignments granted/revoked, purchased-license changes, dynamic-rule edits, auth-method toggles, setting flips, app registrations added/removed

## How this differs from the existing tools

- [**EntraExporter**](https://github.com/microsoft/EntraExporter) (Microsoft) exports the tenant as raw JSON — excellent for backup and version tracking, but nobody can hand `conditionalAccessPolicies.json` to a manager. This tool's output is documentation first.
- [**Microsoft365DSC**](https://microsoft365dsc.com/) manages configuration as code — far more powerful, and a correspondingly bigger commitment. This tool is one script you can run in five minutes with read-only scopes.

If you need backup/restore or config enforcement, use those. If you need *current, readable docs and a shareable report*, this is the gap this fills.

## Honesty notes

- **Read-only.** Every call is a GET (plus one `getByIds` POST that only resolves object IDs to display names). The script changes nothing.
- Tested for syntax and structure against mocked Graph data; **not yet run against a production tenant** — dev tenant first, as with anything that touches Graph.
- **PIM eligible assignments are not included** — the roles section reads permanent assignments only. Group-based role assignments are listed as the group, not expanded to members.
- Secret *values* never appear anywhere — Graph doesn't return them; only credential names and expiry dates are documented.
- The identity plane only, for now. Exchange, Intune, SharePoint, and Teams settings live behind different APIs and are on the roadmap, not in the script.

## Roadmap

- Intune: compliance policies and configuration profiles
- Conditional Access gap analysis (see also the planned CA analyzer)
- `-Anonymize` switch for sharing output with consultants
- Scheduled-run wrapper with change alerts (email / Teams webhook)

## Related tools

Same author, same philosophy (small, honest, read-only by default): [entra-lifecycle-toolkit](https://github.com/JonathanT10/entra-lifecycle-toolkit) · [m365-license-waste-report](https://github.com/JonathanT10/m365-license-waste-report) · [entra-security-snapshot](https://github.com/JonathanT10/entra-security-snapshot) · [print-fleet-dashboard](https://github.com/JonathanT10/print-fleet-dashboard)

## License

MIT — see [LICENSE](LICENSE).
