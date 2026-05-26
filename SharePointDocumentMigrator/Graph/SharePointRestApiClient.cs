using Azure.Core;
using Microsoft.Extensions.Logging;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Graph;

/// <summary>
/// HTTP client for the SharePoint REST API (<c>/_api/web/</c>).
/// Used as a fallback when the Microsoft Graph API cannot resolve hidden document libraries
/// (e.g. SharePoint Publishing libraries excluded from Graph's /lists and /drives collections).
/// </summary>
/// <remarks>
/// Authentication uses a SharePoint-scoped token (<c>https://{hostname}/.default</c>),
/// which is separate from the Microsoft Graph token used by <see cref="GraphApiClient"/>.
/// </remarks>
public sealed class SharePointRestApiClient(
    HttpClient httpClient,
    TokenCredential credential,
    ILogger<SharePointRestApiClient> logger)
{
    private readonly HttpClient _httpClient = httpClient;
    private readonly TokenCredential _credential = credential;
    private readonly ILogger<SharePointRestApiClient> _logger = logger;

    /// <summary>
    /// Attempts to resolve a document library by its display title using
    /// <c>/_api/web/lists/getbytitle('{title}')</c>.
    /// Unlike Graph's <c>/lists</c> collection, this endpoint exposes hidden libraries.
    /// </summary>
    /// <returns>
    /// A <see cref="ListInfo"/> whose <c>Id</c> is the list GUID (without curly braces)
    /// and whose <c>WebUrl</c> is the server-relative URL of the library root folder.
    /// Returns <c>null</c> if the library is not found.
    /// </returns>
    public async Task<ListInfo?> TryResolveListAsync(
        string hostname,
        string siteAbsoluteUrl,
        string libraryTitle,
        CancellationToken cancellationToken)
    {
        var escapedTitle = EscapeSingleQuotes(libraryTitle);
        var requestUrl = $"{siteAbsoluteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{escapedTitle}')" +
                         "?$expand=RootFolder&$select=Id,Title,RootFolder/ServerRelativeUrl";

        var document = await GetJsonAsync(hostname, requestUrl, cancellationToken, allowNotFound: true);
        if (document is null)
        {
            _logger.LogInformation(
                "Library '{LibraryTitle}' not found via SharePoint REST API at '{SiteUrl}'.",
                libraryTitle, siteAbsoluteUrl);
            return null;
        }

        // REST API returns GUIDs with surrounding braces: {xxxxxxxx-...} — strip them for Graph compatibility.
        var rawId = document["Id"]?.GetValue<string>()
            ?? throw new InvalidOperationException("SharePoint REST API response missing list Id.");
        var listId = rawId.Trim('{', '}');

        var displayName = document["Title"]?.GetValue<string>() ?? libraryTitle;
        var serverRelativeUrl = document["RootFolder"]?["ServerRelativeUrl"]?.GetValue<string>() ?? string.Empty;

        _logger.LogInformation(
            "Library '{LibraryTitle}' resolved via SharePoint REST API (listId: {ListId}, root: '{Root}').",
            libraryTitle, listId, serverRelativeUrl);

        return new ListInfo(listId, displayName, serverRelativeUrl);
    }

    /// <summary>
    /// Enumerates all files and subfolders within a library, starting from an optional root subfolder.
    /// Traverses the folder tree recursively. The <c>Forms</c> system subfolder is automatically excluded.
    /// </summary>
    /// <param name="hostname">SharePoint tenant hostname, e.g. <c>contoso.sharepoint.com</c>.</param>
    /// <param name="siteAbsoluteUrl">Absolute URL of the SharePoint site, e.g. <c>https://contoso.sharepoint.com/teams/intranet/it</c>.</param>
    /// <param name="libraryServerRelativeUrl">Server-relative URL of the library root, e.g. <c>/teams/intranet/it/documenti</c>.</param>
    /// <param name="rootFolderRelativePath">Optional subfolder path within the library to use as enumeration root.</param>
    public async Task<(IReadOnlyList<DriveItemInfo> Folders, IReadOnlyList<DriveItemInfo> Files)> EnumerateLibraryAsync(
        string hostname,
        string siteAbsoluteUrl,
        string libraryServerRelativeUrl,
        string? rootFolderRelativePath,
        CancellationToken cancellationToken)
    {
        var normalizedLibraryUrl = libraryServerRelativeUrl.TrimEnd('/');
        var startServerRelativePath = string.IsNullOrEmpty(rootFolderRelativePath)
            ? normalizedLibraryUrl
            : $"{normalizedLibraryUrl}/{rootFolderRelativePath.Trim('/')}";

        var folders = new List<DriveItemInfo>();
        var files = new List<DriveItemInfo>();

        var pending = new Queue<(string ServerPath, string RelPath)>();
        pending.Enqueue((startServerRelativePath, string.Empty));

        while (pending.Count > 0)
        {
            var (serverPath, relPath) = pending.Dequeue();

            var fileNodes = await GetFolderChildrenAsync(hostname, siteAbsoluteUrl, serverPath, "Files", cancellationToken);

            foreach (var file in fileNodes)
            {
                var name = file["Name"]?.GetValue<string>() ?? string.Empty;
                var serverRelUrl = file["ServerRelativeUrl"]?.GetValue<string>() ?? string.Empty;

                long.TryParse(file["Length"]?.GetValue<string>(), out long size);

                files.Add(new DriveItemInfo(
                    Id: file["UniqueId"]?.GetValue<string>() ?? string.Empty,
                    Name: name,
                    RelativePath: CombinePath(relPath, name),
                    IsFolder: false,
                    WebUrl: $"https://{hostname}{serverRelUrl}",
                    Size: size > 0 ? size : null));
            }

            var folderNodes = await GetFolderChildrenAsync(hostname, siteAbsoluteUrl, serverPath, "Folders", cancellationToken);

            foreach (var subfolder in folderNodes)
            {
                var name = subfolder["Name"]?.GetValue<string>() ?? string.Empty;

                if (string.Equals(name, "Forms", StringComparison.OrdinalIgnoreCase)) continue;

                var serverRelUrl = subfolder["ServerRelativeUrl"]?.GetValue<string>() ?? string.Empty;
                var folderRelPath = CombinePath(relPath, name);

                folders.Add(new DriveItemInfo(
                    Id: subfolder["UniqueId"]?.GetValue<string>() ?? string.Empty,
                    Name: name,
                    RelativePath: folderRelPath,
                    IsFolder: true,
                    WebUrl: $"https://{hostname}{serverRelUrl}",
                    Size: null));

                pending.Enqueue((serverRelUrl, folderRelPath));
            }
        }

        return (folders, files);
    }

    /// <summary>
    /// Downloads a file using its absolute SharePoint URL, authenticated via a SharePoint-scoped token.
    /// </summary>
    public async Task DownloadFileAsync(
        string hostname,
        string fileAbsoluteUrl,
        string destinationFilePath,
        CancellationToken cancellationToken)
    {
        var token = await GetTokenAsync(hostname, cancellationToken);

        using var request = new HttpRequestMessage(HttpMethod.Get, fileAbsoluteUrl);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using var response = await _httpClient.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new InvalidOperationException(
                $"SharePoint REST API download failed for '{fileAbsoluteUrl}': {(int)response.StatusCode} — {body}");
        }

        await using var sourceStream = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var destinationStream = File.Create(destinationFilePath);
        await sourceStream.CopyToAsync(destinationStream, cancellationToken);
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private async Task<IReadOnlyList<JsonObject>> GetFolderChildrenAsync(
        string hostname,
        string siteAbsoluteUrl,
        string serverRelativeFolderPath,
        string childType, // "Files" or "Folders"
        CancellationToken cancellationToken)
    {
        var escapedPath = EscapeSingleQuotes(serverRelativeFolderPath);
        var select = childType == "Files"
            ? "$select=Name,UniqueId,ServerRelativeUrl,Length"
            : "$select=Name,UniqueId,ServerRelativeUrl";

        var requestUrl = $"{siteAbsoluteUrl.TrimEnd('/')}/_api/web" +
                         $"/GetFolderByServerRelativeUrl('{escapedPath}')/{childType}" +
                         $"?{select}&$top=5000";

        var items = new List<JsonObject>();
        string? nextUrl = requestUrl;

        while (!string.IsNullOrEmpty(nextUrl))
        {
            var document = await GetJsonAsync(hostname, nextUrl, cancellationToken);
            if (document is null) break;

            var values = document["value"]?.AsArray() ?? [];
            foreach (var node in values)
            {
                if (node is JsonObject obj)
                    items.Add(obj);
            }

            // Handle pagination via OData next-link (present when result set exceeds $top).
            nextUrl = document["odata.nextLink"]?.GetValue<string>()
                      ?? document["@odata.nextLink"]?.GetValue<string>();
        }

        return items;
    }

    private async Task<JsonObject?> GetJsonAsync(
        string hostname,
        string requestUrl,
        CancellationToken cancellationToken,
        bool allowNotFound = false)
    {
        var token = await GetTokenAsync(hostname, cancellationToken);

        using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        // Request clean JSON without OData metadata wrappers.
        request.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

        using var response = await _httpClient.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);

        if (allowNotFound && response.StatusCode == HttpStatusCode.NotFound)
            return null;

        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new InvalidOperationException(
                $"SharePoint REST API request to '{requestUrl}' failed with {(int)response.StatusCode}: {body}");
        }

        var content = await response.Content.ReadAsStringAsync(cancellationToken);
        return JsonNode.Parse(content) as JsonObject;
    }

    private async Task<string> GetTokenAsync(string hostname, CancellationToken cancellationToken)
    {
        // SharePoint REST API requires a token for the SharePoint resource, not Microsoft Graph.
        var scope = $"https://{hostname}/.default";
        var context = new TokenRequestContext([scope]);
        var token = await _credential.GetTokenAsync(context, cancellationToken);
        return token.Token;
    }

    /// <summary>Escapes single quotes for OData string literals (SharePoint REST convention).</summary>
    private static string EscapeSingleQuotes(string value) => value.Replace("'", "''");

    private static string CombinePath(string parent, string child) =>
        string.IsNullOrEmpty(parent) ? child : $"{parent}/{child}";
}
