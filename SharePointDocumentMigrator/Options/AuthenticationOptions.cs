namespace SharePointDocumentMigrator.Options;

public sealed class AuthenticationOptions
{
    public const string SectionName = "Authentication";

    public string TenantId { get; init; } = string.Empty;

    public string ClientId { get; init; } = string.Empty;

    public string ClientSecret { get; init; } = string.Empty;

    public bool UseDeviceCode { get; init; } = true;

    public string[] Scopes { get; set; } = [];

    public void ApplyCustomScopes(params string[] customScopes)
    {
        if (customScopes == null || customScopes.Length == 0)
            return;
        Scopes = customScopes;
    }
}
