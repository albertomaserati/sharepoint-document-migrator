using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Graph;

public sealed record SiteInfo(string Id, string DisplayName, string WebUrl);

public sealed record ListInfo(string Id, string DisplayName, string WebUrl);

public sealed record DriveInfo(string Id, string Name, string WebUrl);

public sealed record SharePointColumnDefinition(
    string InternalName,
    string DisplayName,
    string Type,
    bool Hidden,
    bool ReadOnly,
    bool Required);

public sealed record DriveItemInfo(
    string Id,
    string Name,
    string RelativePath,
    bool IsFolder,
    string WebUrl,
    long? Size);

public sealed record UploadResult(DriveItemInfo Item, bool WasSkipped);

public sealed record SchemaValidationResult(
    IReadOnlyDictionary<string, SharePointColumnDefinition> WritableColumns,
    IReadOnlyList<string> Errors)
{
    public bool IsValid => Errors.Count == 0;
}

public sealed record ResolvedLibraryContext(
    SiteInfo Site,
    ListInfo List,
    DriveInfo Drive,
    IReadOnlyList<SharePointColumnDefinition> Columns);

public sealed record FileFailure(string RelativePath, string ErrorMessage);

public sealed record MetadataSelectionResult(JsonObject Fields, IReadOnlyList<string> SkippedFields);
