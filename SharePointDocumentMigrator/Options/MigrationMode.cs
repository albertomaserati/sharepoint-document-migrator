namespace SharePointDocumentMigrator.Options;

/// <summary>
/// Determines which HTTP client strategy the migrator uses to access the source library.
/// </summary>
public enum MigrationMode
{
    /// <summary>
    /// Uses the Microsoft Graph API for both source enumeration and target upload.
    /// Works for standard SharePoint libraries visible to Graph's /lists and /drives endpoints.
    /// </summary>
    Graph,

    /// <summary>
    /// Uses the SharePoint REST API (/_api/web/) for source resolution, enumeration and download,
    /// and the Microsoft Graph API for target folder creation and file upload.
    /// Required when the source library is hidden from Graph's collection endpoints
    /// (e.g. SharePoint Publishing libraries, Italian-localized sites).
    /// </summary>
    Http
}
