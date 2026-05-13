namespace SharePointDocumentMigrator.Options;

public sealed class SharePointLocationOptions
{
    public string Hostname { get; init; } = string.Empty;

    public string SitePath { get; init; } = string.Empty;

    public string LibraryName { get; init; } = string.Empty;

    public string RootFolder { get; init; } = string.Empty;
}
