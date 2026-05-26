namespace SharePointDocumentMigrator.Options;

public sealed class MigrationOptions
{
    public const string SectionName = "Migration";

    public SharePointLocationOptions Source { get; init; } = new();

    public SharePointLocationOptions Target { get; init; } = new();

    public bool DryRun { get; init; } = true;

    public bool OverwriteExistingFiles { get; init; }

    public bool EnableMetadataMigration { get; init; } = true;

    /// <summary>
    /// Optional mapping between source and target internal field names.
    /// Fields not listed here are assumed to share the same internal name in both libraries.
    /// Example: { "CategoriaProgetto": "Category" }
    /// </summary>
    public Dictionary<string, string> MetadataFieldMappings { get; init; } = [];

    public int MaxRetryCount { get; init; } = 5;

    public int UploadChunkSizeKiB { get; init; } = 3200;
}
