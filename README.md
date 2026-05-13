# SharePoint Document Migrator

Console application `.NET 10` in C# per migrare documenti da una document library SharePoint a un'altra, nello stesso tenant Microsoft 365, preservando il contenuto dei file e i metadati scrivibili compatibili.

## Cosa fa

- accede a SharePoint tramite **login interattivo** (`Device Code` di default, oppure browser interattivo)
- legge una singola document library sorgente per esecuzione
- ricrea la struttura cartelle nel target
- copia i file nel target
- valida che source e target abbiano **schema colonne compatibile**
- riscrive i metadati SharePoint aggiornabili tramite Microsoft Graph
- supporta `dry-run`, retry su errori Graph transienti e upload chunked per file oltre 4 MB

## Limiti noti

- alcuni metadati di sistema SharePoint/Microsoft 365 restano di sola lettura e non possono essere ripristinati via API
- la prima versione assume **stesso schema colonne** tra libreria sorgente e libreria destinazione
- il tool lavora su **una sola document library per job**

## Prerequisiti

1. SDK `.NET 10`
2. App registration Microsoft Entra ID configurata come **public client**
3. Permessi delegati Microsoft Graph almeno:
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
   - `User.Read`
4. Consenso amministrativo, se richiesto dalla tenant policy

## Configurazione

Modifica `src\SharePointDocumentMigrator\appsettings.json` oppure passa i valori via command line / environment variables.

Esempio:

```json
{
  "Authentication": {
    "TenantId": "YOUR-TENANT-ID",
    "ClientId": "YOUR-APP-ID",
    "UseDeviceCode": true,
    "Scopes": [
      "https://graph.microsoft.com/Sites.ReadWrite.All",
      "https://graph.microsoft.com/Files.ReadWrite.All",
      "https://graph.microsoft.com/User.Read"
    ]
  },
  "Migration": {
    "Source": {
      "Hostname": "contoso.sharepoint.com",
      "SitePath": "/sites/SourceSite",
      "LibraryName": "Documents",
      "RootFolder": ""
    },
    "Target": {
      "Hostname": "contoso.sharepoint.com",
      "SitePath": "/sites/TargetSite",
      "LibraryName": "Documents",
      "RootFolder": ""
    },
    "DryRun": true,
    "OverwriteExistingFiles": false,
    "MaxRetryCount": 5,
    "UploadChunkSizeKiB": 3200
  }
}
```

### Parametri principali

| Chiave | Significato |
| --- | --- |
| `Source.Hostname` / `Target.Hostname` | Host SharePoint, ad esempio `contoso.sharepoint.com` |
| `SitePath` | Percorso relativo del sito, ad esempio `/sites/Finance` |
| `LibraryName` | Nome visualizzato della document library |
| `RootFolder` | Sottocartella opzionale da usare come radice della migrazione |
| `DryRun` | Se `true`, non scrive nulla sul target |
| `OverwriteExistingFiles` | Se `true`, sostituisce i file gia' presenti |
| `UploadChunkSizeKiB` | Chunk size per upload session, multiplo positivo di `320` |

## Esecuzione

```powershell
dotnet run --project .\src\SharePointDocumentMigrator\SharePointDocumentMigrator.csproj
```

Per un override rapido:

```powershell
dotnet run --project .\src\SharePointDocumentMigrator\SharePointDocumentMigrator.csproj -- `
  --Migration:DryRun=false `
  --Migration:OverwriteExistingFiles=true
```

## Test

```powershell
dotnet test .\SharePointDocumentMigrator.slnx
```
