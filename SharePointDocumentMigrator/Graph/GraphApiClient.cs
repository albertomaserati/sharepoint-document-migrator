using Azure.Core;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharePointDocumentMigrator.Migration;
using SharePointDocumentMigrator.Options;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Graph;

public sealed class GraphApiClient(
    HttpClient httpClient,
    TokenCredential credential,
    IOptions<AuthenticationOptions> authenticationOptions,
    IOptions<MigrationOptions> migrationOptions,
    ILogger<GraphApiClient> logger)
{
    private const string DISPLAY_NAME = "displayName";
    private const string FOLDER = "folder";
    private const string ID = "id";
    private const string MICROSOFT_GRAPH_CONFLICT_BEHAVIOR = "@microsoft.graph.conflictBehavior";
    private const string NAME = "name";
    private const string SIZE = "size";
    private const string WEB_URL = "webUrl";

    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = false
    };

    private static readonly HashSet<HttpStatusCode> RetriableStatusCodes =
    [
        HttpStatusCode.TooManyRequests,
        HttpStatusCode.BadGateway,
        HttpStatusCode.ServiceUnavailable,
        HttpStatusCode.GatewayTimeout
    ];

    private readonly HttpClient _httpClient = httpClient;
    private readonly TokenCredential _credential = credential;
    private readonly AuthenticationOptions _authenticationOptions = authenticationOptions.Value;
    private readonly MigrationOptions _migrationOptions = migrationOptions.Value;
    private readonly ILogger<GraphApiClient> _logger = logger;
    private readonly ConcurrentDictionary<string, DriveItemInfo> _folderCache = new(StringComparer.OrdinalIgnoreCase);

    public async Task<ResolvedLibraryContext> ResolveLibraryContextAsync(
        SharePointLocationOptions options,
        CancellationToken cancellationToken)
    {
        var site = await ResolveSiteAsync(options, cancellationToken);

        ListInfo list;
        DriveInfo drive;

        var listFromEndpoint = await TryResolveListFromListsEndpointAsync(
            site.Id,
            options.LibraryName,
            cancellationToken);

        if (listFromEndpoint is not null)
        {
            list = listFromEndpoint;
            drive = await ResolveDriveAsync(site.Id, list.Id, cancellationToken);
        }
        else
        {
            // Hidden lists are excluded from collection queries but are often still reachable
            // via GET /sites/{id}/lists/{url-name} when the caller has sufficient permissions.

            _logger.LogInformation(
                "Library '{LibraryName}' not found via /lists — trying direct URL-name access.",
                options.LibraryName);

            var listByName = await TryResolveListByUrlNameAsync(site.Id, options.LibraryName, cancellationToken);

            if (listByName is not null)
            {
                list = listByName;
                drive = await ResolveDriveAsync(site.Id, list.Id, cancellationToken);
            }
            else
            {
                // search /drives collection (catches libraries exposed as drives
                // but hidden from /lists, common in SharePoint Publishing sites).

                _logger.LogInformation(
                    "Library '{LibraryName}' not found via direct URL-name access — falling back to /drives search.",
                    options.LibraryName);

                (list, drive) = await ResolveListAndDriveFromDrivesAsync(
                    site.Id, options.LibraryName, cancellationToken);
            }
        }

        var columns = await GetColumnsAsync(site.Id, list.Id, cancellationToken);

        return new ResolvedLibraryContext(site, list, drive, columns);
    }

    public async Task<SiteInfo> ResolveSiteAsync(
        SharePointLocationOptions options,
        CancellationToken cancellationToken)
    {
        var sitePath = PathUtilities.NormalizeSitePath(options.SitePath);

        var requestUri = $"sites/{options.Hostname}:/{PathUtilities.EncodeGraphPath(sitePath)}?$select=id,displayName,webUrl";

        var document = await GetJsonAsync(requestUri, cancellationToken)
            ?? throw new InvalidOperationException("Graph response missing site payload.");

        return new SiteInfo(
            document[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing site id."),
            document[DISPLAY_NAME]?.GetValue<string>() ?? options.SitePath,
            document[WEB_URL]?.GetValue<string>() ?? string.Empty);
    }

    private async Task<ListInfo?> TryResolveListByUrlNameAsync(
        string siteId,
        string libraryName,
        CancellationToken cancellationToken)
    {
        // Graph API allows addressing a list by its URL-segment name, not just its GUID.
        // This succeeds for hidden lists that are excluded from the /lists collection query.
        var document = await GetJsonAsync(
            $"sites/{siteId}/lists/{libraryName}?$select=id,displayName,webUrl",
            cancellationToken,
            allowNotFound: true);

        if (document is null)
        {
            return null;
        }

        _logger.LogInformation("Library '{LibraryName}' resolved via direct URL-name access.", libraryName);

        return new ListInfo(
            document[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing list id in direct URL-name lookup."),
            document[DISPLAY_NAME]?.GetValue<string>() ?? libraryName,
            document[WEB_URL]?.GetValue<string>() ?? string.Empty);
    }

    private async Task<ListInfo?> TryResolveListFromListsEndpointAsync(
        string siteId,
        string libraryName,
        CancellationToken cancellationToken)
    {
        var requestUri = $"sites/{siteId}/lists?$select=id,displayName,webUrl";
        var collection = await GetCollectionAsync(requestUri, cancellationToken);

        _logger.LogInformation(
            "Available lists on site '{SiteId}':{NewLine}{Lists}",
            siteId,
            Environment.NewLine,
            string.Join(Environment.NewLine, collection.Select(n =>
                $"  displayName='{n[DISPLAY_NAME]?.GetValue<string>()}' | webUrl='{n[WEB_URL]?.GetValue<string>()}'")));

        // Match by the last segment of webUrl only (e.g. "Documents", "documenti") avoiding match by display names
        // which can collide with unrelated libraries 
        var listNode = collection.FirstOrDefault(node =>
            MatchesUrlSegment(node?[WEB_URL]?.GetValue<string>(), libraryName));

        if (listNode is null)
        {
            return null;
        }

        return new ListInfo(
            listNode[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing list id."),
            listNode[DISPLAY_NAME]?.GetValue<string>() ?? libraryName,
            listNode[WEB_URL]?.GetValue<string>() ?? string.Empty);
    }

    private async Task<(ListInfo List, DriveInfo Drive)> ResolveListAndDriveFromDrivesAsync(
        string siteId,
        string libraryName,
        CancellationToken cancellationToken)
    {
        var drivesUri = $"sites/{siteId}/drives?$select=id,name,webUrl";
        var drives = await GetCollectionAsync(drivesUri, cancellationToken);

        _logger.LogInformation(
            "Available drives on site '{SiteId}':{NewLine}{Drives}",
            siteId,
            Environment.NewLine,
            string.Join(Environment.NewLine, drives.Select(d =>
                $"  name='{d[NAME]?.GetValue<string>()}' | webUrl='{d[WEB_URL]?.GetValue<string>()}'")));

        var driveNode = drives.FirstOrDefault(d =>
            MatchesUrlSegment(d[WEB_URL]?.GetValue<string>(), libraryName)) ?? throw new InvalidOperationException(
                $"Document library '{libraryName}' was not found in site '{siteId}' via Graph API");

        var driveId = driveNode[ID]?.GetValue<string>()
            ?? throw new InvalidOperationException("Graph response missing drive id in /drives fallback.");

        _logger.LogInformation(
            "Found library '{LibraryName}' via /drives fallback (driveId: {DriveId}).",
            libraryName, driveId);

        var listDocument = await GetJsonAsync($"drives/{driveId}/list?$select=id,displayName,webUrl", cancellationToken)
            ?? throw new InvalidOperationException("Graph response missing list payload in /drives fallback.");

        var list = new ListInfo(
            listDocument[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing list id in /drives fallback."),
            listDocument[DISPLAY_NAME]?.GetValue<string>() ?? libraryName,
            listDocument[WEB_URL]?.GetValue<string>() ?? string.Empty);

        // DriveInfo is built from the node already in memory — no additional API call needed.
        var drive = new DriveInfo(
            driveId,
            driveNode[NAME]?.GetValue<string>() ?? string.Empty,
            driveNode[WEB_URL]?.GetValue<string>() ?? string.Empty);

        return (list, drive);
    }

    public async Task<DriveInfo> ResolveDriveAsync(
        string siteId,
        string listId,
        CancellationToken cancellationToken)
    {
        var document = await GetJsonAsync($"sites/{siteId}/lists/{listId}/drive?$select=id,name,webUrl", cancellationToken)
            ?? throw new InvalidOperationException("Graph response missing drive payload.");

        return new DriveInfo(
            document[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing drive id."),
            document["name"]?.GetValue<string>() ?? string.Empty,
            document[WEB_URL]?.GetValue<string>() ?? string.Empty);
    }

    public async Task<IReadOnlyList<SharePointColumnDefinition>> GetColumnsAsync(
        string siteId,
        string listId,
        CancellationToken cancellationToken)
    {
        var requestUri = $"sites/{siteId}/lists/{listId}/columns?$select=name,displayName,hidden,readOnly,required,text,choice,number,currency,dateTime,lookup,personOrGroup,boolean,hyperlinkOrPicture,term,calculated";
        var collection = await GetCollectionAsync(requestUri, cancellationToken);

        return [.. collection
            .Select(column => new SharePointColumnDefinition(
                column["name"]?.GetValue<string>() ?? string.Empty,
                column[DISPLAY_NAME]?.GetValue<string>() ?? string.Empty,
                DetermineColumnType(column),
                column["hidden"]?.GetValue<bool>() ?? false,
                column["readOnly"]?.GetValue<bool>() ?? false,
                column["required"]?.GetValue<bool>() ?? false))
            .Where(column => !string.IsNullOrWhiteSpace(column.InternalName))];
    }

    public async Task<DriveItemInfo> GetDriveItemByPathAsync(
        string driveId,
        string? relativePath,
        CancellationToken cancellationToken)
    {
        var normalizedPath = PathUtilities.NormalizeRelativePath(relativePath);
        var requestUri = string.IsNullOrEmpty(normalizedPath)
            ? $"drives/{driveId}/root?$select=id,name,webUrl,size,folder"
            : $"drives/{driveId}/root:/{PathUtilities.EncodeGraphPath(normalizedPath)}?$select=id,name,webUrl,size,folder";

        var document = await GetJsonAsync(requestUri, cancellationToken)
            ?? throw new InvalidOperationException("Graph response missing drive item payload.");
        return ToDriveItem(document, normalizedPath);
    }

    public async Task<DriveItemInfo?> TryGetDriveItemByPathAsync(
        string driveId,
        string? relativePath,
        CancellationToken cancellationToken)
    {
        var normalizedPath = PathUtilities.NormalizeRelativePath(relativePath);
        var requestUri = string.IsNullOrEmpty(normalizedPath)
            ? $"drives/{driveId}/root?$select=id,name,webUrl,size,folder"
            : $"drives/{driveId}/root:/{PathUtilities.EncodeGraphPath(normalizedPath)}?$select=id,name,webUrl,size,folder";

        var document = await GetJsonAsync(requestUri, cancellationToken, allowNotFound: true);
        return document is null ? null : ToDriveItem(document, normalizedPath);
    }

    public async Task<IReadOnlyList<DriveItemInfo>> ListChildrenAsync(
        string driveId,
        string itemId,
        string parentRelativePath,
        CancellationToken cancellationToken)
    {
        var requestUri = $"drives/{driveId}/items/{itemId}/children?$select=id,name,webUrl,size,folder";
        var collection = await GetCollectionAsync(requestUri, cancellationToken);

        return [.. collection
            .Select(node =>
            {
                var name = node["name"]?.GetValue<string>() ?? string.Empty;
                var relativePath = PathUtilities.CombineRelativePath(parentRelativePath, name);
                return ToDriveItem(node, relativePath);
            })];
    }

    public async Task DownloadFileAsync(
        string driveId,
        string itemId,
        string destinationFilePath,
        CancellationToken cancellationToken)
    {
        using var response = await SendGraphAsync(
            HttpMethod.Get,
            $"drives/{driveId}/items/{itemId}/content",
            contentFactory: null,
            cancellationToken);

        await using var sourceStream = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var destinationStream = File.Create(destinationFilePath);
        await sourceStream.CopyToAsync(destinationStream, cancellationToken);
    }

    public async Task<JsonObject> GetListItemFieldsAsync(
        string driveId,
        string itemId,
        CancellationToken cancellationToken)
    {
        return await GetJsonAsync($"drives/{driveId}/items/{itemId}/listItem/fields", cancellationToken)
            ?? [];
    }

    public async Task<string> GetListItemIdForDriveItemAsync(
        string driveId,
        string itemId,
        CancellationToken cancellationToken)
    {
        var document = await GetJsonAsync($"drives/{driveId}/items/{itemId}/listItem?$select=id", cancellationToken)
            ?? throw new InvalidOperationException("Graph response missing list item payload.");
        return document[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing list item id.");
    }

    public async Task UpdateListItemFieldsAsync(
        string siteId,
        string listId,
        string listItemId,
        JsonObject fields,
        CancellationToken cancellationToken)
    {
        var payload = fields.ToJsonString(SerializerOptions);
        await SendGraphAsync(
            HttpMethod.Patch,
            $"sites/{siteId}/lists/{listId}/items/{listItemId}/fields",
            () => new StringContent(payload, Encoding.UTF8, "application/json"),
            cancellationToken);
    }

    public async Task EnsureFolderExistsAsync(
        string driveId,
        string? relativePath,
        CancellationToken cancellationToken)
    {
        var normalizedPath = PathUtilities.NormalizeRelativePath(relativePath);
        if (string.IsNullOrEmpty(normalizedPath))
        {
            return;
        }

        if (_folderCache.ContainsKey($"{driveId}:{normalizedPath}"))
        {
            return;
        }

        var currentPath = string.Empty;
        DriveItemInfo currentFolder = await GetDriveItemByPathAsync(driveId, currentPath, cancellationToken);

        foreach (var segment in normalizedPath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            currentPath = PathUtilities.CombineRelativePath(currentPath, segment);
            var cacheKey = $"{driveId}:{currentPath}";
            if (_folderCache.TryGetValue(cacheKey, out var cachedFolder))
            {
                currentFolder = cachedFolder;
                continue;
            }

            var existingFolder = await TryGetDriveItemByPathAsync(driveId, currentPath, cancellationToken);
            if (existingFolder is not null)
            {
                currentFolder = existingFolder;
                _folderCache.TryAdd(cacheKey, currentFolder);
                continue;
            }

            var payload = new JsonObject
            {
                [NAME] = segment,
                [FOLDER] = new JsonObject(),
                [MICROSOFT_GRAPH_CONFLICT_BEHAVIOR] = "fail"
            };

            var createdFolder = await PostJsonAsync(
                $"drives/{driveId}/items/{currentFolder.Id}/children",
                payload,
                cancellationToken);

            currentFolder = ToDriveItem(createdFolder, currentPath);
            _folderCache.TryAdd(cacheKey, currentFolder);
        }
    }

    public async Task<UploadResult> UploadFileAsync(
        string driveId,
        string destinationRelativePath,
        string sourceFilePath,
        bool overwrite,
        CancellationToken cancellationToken)
    {
        var normalizedPath = PathUtilities.NormalizeRelativePath(destinationRelativePath);
        if (string.IsNullOrEmpty(normalizedPath))
        {
            throw new InvalidOperationException("Destination path for file upload cannot be empty.");
        }

        var existingItem = await TryGetDriveItemByPathAsync(driveId, normalizedPath, cancellationToken);
        if (existingItem is not null && !overwrite)
        {
            return new UploadResult(existingItem, WasSkipped: true);
        }

        var fileInfo = new FileInfo(sourceFilePath);
        if (!fileInfo.Exists)
        {
            throw new FileNotFoundException("Source file not found for upload.", sourceFilePath);
        }

        if (fileInfo.Length <= 4 * 1024 * 1024)
        {
            return await UploadSmallFileAsync(driveId, normalizedPath, sourceFilePath, cancellationToken);
        }

        return await UploadWithSessionAsync(driveId, normalizedPath, sourceFilePath, overwrite, cancellationToken);
    }

    private async Task<UploadResult> UploadSmallFileAsync(
        string driveId,
        string destinationRelativePath,
        string sourceFilePath,
        CancellationToken cancellationToken)
    {
        var requestUri = $"drives/{driveId}/root:/{PathUtilities.EncodeGraphPath(destinationRelativePath)}:/content";
        var bytes = await File.ReadAllBytesAsync(sourceFilePath, cancellationToken);

        using var response = await SendGraphAsync(
            HttpMethod.Put,
            requestUri,
            () =>
            {
                var content = new ByteArrayContent(bytes);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                return content;
            },
            cancellationToken);

        var document = await DeserializeObjectAsync(response, cancellationToken);
        return new UploadResult(ToDriveItem(document, destinationRelativePath), WasSkipped: false);
    }

    private async Task<UploadResult> UploadWithSessionAsync(
        string driveId,
        string destinationRelativePath,
        string sourceFilePath,
        bool overwrite,
        CancellationToken cancellationToken)
    {
        var payload = new JsonObject
        {
            ["item"] = new JsonObject
            {
                [MICROSOFT_GRAPH_CONFLICT_BEHAVIOR] = overwrite ? "replace" : "fail",
                [NAME] = Path.GetFileName(destinationRelativePath)
            }
        };

        var session = await PostJsonAsync(
            $"drives/{driveId}/root:/{PathUtilities.EncodeGraphPath(destinationRelativePath)}:/createUploadSession",
            payload,
            cancellationToken);

        var uploadUrl = session["uploadUrl"]?.GetValue<string>()
            ?? throw new InvalidOperationException("Graph upload session did not return an uploadUrl.");

        var fileInfo = new FileInfo(sourceFilePath);
        await using var fileStream = File.OpenRead(sourceFilePath);

        var chunkSize = _migrationOptions.UploadChunkSizeKiB * 1024;
        var buffer = new byte[chunkSize];
        long totalLength = fileInfo.Length;
        long offset = 0;

        while (offset < totalLength)
        {
            var bytesRead = await fileStream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken);
            if (bytesRead == 0)
            {
                break;
            }

            var chunk = new byte[bytesRead];
            Buffer.BlockCopy(buffer, 0, chunk, 0, bytesRead);

            using var response = await SendAbsoluteAsync(
                HttpMethod.Put,
                uploadUrl,
                () =>
                {
                    var content = new ByteArrayContent(chunk);
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    content.Headers.ContentRange = new ContentRangeHeaderValue(offset, offset + bytesRead - 1, totalLength);
                    return content;
                },
                includeAuthorization: false,
                cancellationToken);
            if (response.StatusCode == HttpStatusCode.Created || response.StatusCode == HttpStatusCode.OK)
            {
                var document = await DeserializeObjectAsync(response, cancellationToken);
                return new UploadResult(ToDriveItem(document, destinationRelativePath), WasSkipped: false);
            }

            if (response.StatusCode != HttpStatusCode.Accepted)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                throw new GraphApiException(
                    $"Chunk upload failed for '{destinationRelativePath}'.",
                    (int)response.StatusCode,
                    uploadUrl,
                    body);
            }

            offset += bytesRead;
        }

        throw new InvalidOperationException($"Upload session for '{destinationRelativePath}' ended without a final drive item.");
    }

    private async Task<JsonObject?> GetJsonAsync(
        string requestUri,
        CancellationToken cancellationToken,
        bool allowNotFound = false)
    {
        using var response = await SendGraphAsync(
            HttpMethod.Get,
            requestUri,
            contentFactory: null,
            cancellationToken,
            allowNotFound);

        return allowNotFound && response.StatusCode == HttpStatusCode.NotFound
            ? null
            : await DeserializeObjectAsync(response, cancellationToken);
    }

    private async Task<JsonObject> PostJsonAsync(
        string requestUri,
        JsonObject payload,
        CancellationToken cancellationToken)
    {
        using var response = await SendGraphAsync(
            HttpMethod.Post,
            requestUri,
            () => new StringContent(payload.ToJsonString(SerializerOptions), Encoding.UTF8, "application/json"),
            cancellationToken);

        return await DeserializeObjectAsync(response, cancellationToken);
    }

    private async Task<IReadOnlyList<JsonObject>> GetCollectionAsync(
        string requestUri,
        CancellationToken cancellationToken)
    {
        var items = new List<JsonObject>();
        var nextUri = requestUri;

        while (!string.IsNullOrEmpty(nextUri))
        {
            var document = await GetJsonAsync(nextUri, cancellationToken);
            var values = document?["value"]?.AsArray() ?? [];
            foreach (var node in values)
            {
                if (node is JsonObject obj)
                {
                    items.Add(obj);
                }
            }

            nextUri = document?["@odata.nextLink"]?.GetValue<string>();
        }

        return items;
    }

    private async Task<HttpResponseMessage> SendGraphAsync(
        HttpMethod method,
        string requestUri,
        Func<HttpContent?>? contentFactory,
        CancellationToken cancellationToken,
        bool allowNotFound = false)
    {
        return await SendAbsoluteAsync(
            method,
            requestUri,
            contentFactory,
            includeAuthorization: true,
            cancellationToken,
            allowNotFound: allowNotFound);
    }

    private async Task<HttpResponseMessage> SendAbsoluteAsync(
        HttpMethod method,
        string requestUri,
        Func<HttpContent?>? contentFactory,
        bool includeAuthorization,
        CancellationToken cancellationToken,
        bool allowNotFound = false)
    {
        var attempt = 0;
        while (true)
        {
            using var request = new HttpRequestMessage(method, requestUri)
            {
                Content = contentFactory?.Invoke()
            };

            if (includeAuthorization)
            {
                var accessToken = await _credential.GetTokenAsync(
                    new TokenRequestContext(_authenticationOptions.Scopes),
                    cancellationToken);

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);
            }

            var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

            if (allowNotFound && response.StatusCode == HttpStatusCode.NotFound)
            {
                return response;
            }

            if (response.IsSuccessStatusCode)
            {
                return response;
            }

            if (attempt < _migrationOptions.MaxRetryCount && RetriableStatusCodes.Contains(response.StatusCode))
            {
                var delay = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(Math.Pow(2, attempt + 1));
                _logger.LogWarning(
                    "Graph request {Method} {RequestUri} failed with {StatusCode}. Retrying in {Delay}.",
                    method,
                    requestUri,
                    (int)response.StatusCode,
                    delay);

                response.Dispose();
                await Task.Delay(delay, cancellationToken);

                attempt++;

                continue;
            }

            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            var message = $"Graph request {method} {requestUri} failed with status {(int)response.StatusCode}.";
            response.Dispose();

            throw new GraphApiException(message, (int)response.StatusCode, requestUri, body);
        }
    }

    private static async Task<JsonObject> DeserializeObjectAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        var content = await response.Content.ReadAsStringAsync(cancellationToken);
        var node = JsonNode.Parse(content) as JsonObject;
        return node ?? [];
    }

    private static string DetermineColumnType(JsonObject column)
    {
        string[] supportedTypeProperties =
        [
            "text",
            "choice",
            "number",
            "currency",
            "dateTime",
            "lookup",
            "personOrGroup",
            "boolean",
            "hyperlinkOrPicture",
            "term",
            "calculated"
        ];

        return supportedTypeProperties.FirstOrDefault(column.ContainsKey) ?? "unknown";
    }

    private static bool MatchesUrlSegment(string? webUrl, string libraryName)
    {
        if (string.IsNullOrEmpty(webUrl))
        {
            return false;
        }

        // Checks whether the last segment of the webUrl matches the library name (e.g. ".../documenti" → "documenti")
        var lastSegment = webUrl.TrimEnd('/').Split('/').LastOrDefault() ?? string.Empty;
        return string.Equals(lastSegment, libraryName, StringComparison.OrdinalIgnoreCase);
    }

    private static DriveItemInfo ToDriveItem(JsonObject node, string relativePath)
    {
        return new DriveItemInfo(
            node[ID]?.GetValue<string>() ?? throw new InvalidOperationException("Graph response missing drive item id."),
            node[NAME]?.GetValue<string>() ?? string.Empty,
            PathUtilities.NormalizeRelativePath(relativePath),
            node[FOLDER] is not null,
            node[WEB_URL]?.GetValue<string>() ?? string.Empty,
            node[SIZE]?.GetValue<long>());
    }
}
