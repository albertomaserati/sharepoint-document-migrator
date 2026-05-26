# SharePoint Document Migrator

A .NET 10 console application that migrates documents between SharePoint Online document libraries **within the same Microsoft 365 tenant** via the Microsoft Graph API, preserving the full folder structure and — optionally — column metadata.

## What it does

The tool performs a one-way migration from a **source** SharePoint library to a **target** library:

1. **Resolves** source and target libraries by site path and library name.
2. **Enumerates** the full folder tree and file list of the source, starting from an optional root subfolder.
3. **Recreates** the folder structure on the target.
4. **Downloads** each source file to a local temporary file, then **uploads** it to the target using Graph's chunked upload session (handles large files automatically).
5. **Migrates metadata** (column values) from the source list item to the target list item after each upload, with optional field name remapping between the two libraries.

Files that already exist on the target can be skipped or overwritten depending on configuration.  
A **dry-run mode** lets you preview the full migration plan without making any changes.

Transient Graph API errors (429, 502, 503, 504) are retried automatically with exponential backoff, honouring the `Retry-After` response header when present.

---

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- An **Azure AD application registration** with the following Microsoft Graph **application** permissions granted and admin-consented:
  - `Sites.ReadWrite.All`
  - `Files.ReadWrite.All`

---

## Configuration

All settings live in `appsettings.json` inside the project folder. Sensitive values (tenant ID, client secret, etc.) should be placed in an environment-specific override file — e.g. `appsettings.Development.json` — which is merged on top of the base file at startup and should not be committed to source control.

The active environment is selected via the `DOTNET_ENVIRONMENT` environment variable (defaults to `Production`). Launch profiles are defined in `Properties/launchSettings.json`.

### Authentication section

```json
"Authentication": {
  "TenantId": "<azure-ad-tenant-guid>",
  "ClientId": "<app-registration-client-id>",
  "UseDeviceCode": false,
  "ClientSecret": "<client-secret>"
}
```

| Property | Description |
|---|---|
| `TenantId` | Azure AD tenant GUID. |
| `ClientId` | Client ID of the app registration. |
| `UseDeviceCode` | `true` — interactive device-code flow (opens a browser prompt, useful for testing). `false` — client credentials flow (app-only, for unattended runs). |
| `ClientSecret` | Required when `UseDeviceCode` is `false`. |

### Migration section

```json
"Migration": {
  "Source": {
    "Hostname": "contoso.sharepoint.com",
    "SitePath": "/sites/source-site",
    "LibraryName": "Documents",
    "RootFolder": "Archive/2023"
  },
  "Target": {
    "Hostname": "contoso.sharepoint.com",
    "SitePath": "/sites/target-site",
    "LibraryName": "Documents",
    "RootFolder": ""
  },
  "DryRun": true,
  "OverwriteExistingFiles": false,
  "EnableMetadataMigration": false,
  "MetadataFieldMappings": {},
  "MaxRetryCount": 5,
  "UploadChunkSizeKiB": 3200
}
```

#### Source / Target location

Both `Source` and `Target` share the same four properties:

| Property | Description |
|---|---|
| `Hostname` | SharePoint tenant hostname, e.g. `contoso.sharepoint.com`. |
| `SitePath` | Server-relative path of the site, e.g. `/sites/my-site` or `/teams/department/it`. |
| `LibraryName` | Display name of the document library. |
| `RootFolder` | Optional subfolder within the library used as the migration root. Leave empty to migrate the entire library. |

#### Migration behaviour

| Property | Type | Default | Description |
|---|---|---|---|
| `DryRun` | bool | `true` | When `true`, enumerates source items and logs what would happen without touching the target. **Always start with a dry run.** |
| `OverwriteExistingFiles` | bool | `false` | When `false`, files already present on the target are skipped. When `true`, they are overwritten. |
| `EnableMetadataMigration` | bool | `false` | When `true`, column values from each source list item are written to the corresponding target list item after upload. Source and target schemas are validated for compatibility before the migration starts and the tool aborts if they are incompatible. |
| `MetadataFieldMappings` | object | `{}` | Maps source internal field names to different target internal field names. Fields not listed here are assumed to share the same internal name in both libraries. Example: `{ "CategoriaProgetto": "Category" }`. |
| `MaxRetryCount` | int | `5` | Maximum number of retry attempts for transient Graph API errors. |
| `UploadChunkSizeKiB` | int | `3200` | Upload chunk size in KiB. Must be a positive multiple of `320` (Graph API requirement). |

---

## Running the application

```bash
# Dry run — safe, no changes written to the target
dotnet run --project SharePointDocumentMigrator

# Specify an environment (loads appsettings.Production.json overrides)
DOTNET_ENVIRONMENT=Production dotnet run --project SharePointDocumentMigrator

# Publish and run the binary directly
dotnet publish -c Release -o ./publish
./publish/SharePointDocumentMigrator
```

Press **Ctrl+C** at any time to cancel. The in-progress upload is abandoned and the process exits cleanly.

---

## Output

The application writes structured log messages to the console throughout execution.

### During migration

| Event | Level | Sample message |
|---|---|---|
| Dry-run file found | `Information` | `Source file name: Reports/Q1.xlsx. Url: https://...` |
| Dry-run folder found | `Information` | `[DRY RUN] Folder found — source: Reports | target: Reports` |
| Transient error, retrying | `Warning` | `Graph request failed with 429. Retrying (attempt 1/5) in 00:00:02.` |
| File failed permanently | `Error` | `Failed to migrate 'Reports/Q1.xlsx'.` + exception detail |

Failed files are logged individually and **do not stop the migration** — the tool continues with the remaining files.

### Final summary

At the end of every run, a single summary line is logged at `Information` level.

**Dry run:**
```
Dry run completed. Planned files: 142, planned folders: 18.
```

**Live run:**
```
Migration completed. Migrated files: 138, skipped: 3, failures: 1, updated metadata: 138, folders created: 18.
```

### Exit code

| Code | Meaning |
|---|---|
| `0` | Run completed with no failures (includes successful dry runs). |
| `1` | One or more files failed to migrate, or a startup/configuration error occurred. |
