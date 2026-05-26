namespace SharePointDocumentMigrator.Migration;

public static class PathUtilities
{
    public static string NormalizeSitePath(string sitePath)
    {
        if (string.IsNullOrWhiteSpace(sitePath))
            throw new InvalidOperationException("SharePoint site path cannot be empty.");

        var normalized = sitePath.Replace('\\', '/').Trim();
        normalized = normalized.Trim('/');
        return normalized;
    }

    public static string NormalizeRelativePath(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
            return string.Empty;

        var normalized = relativePath.Replace('\\', '/').Trim();
        normalized = normalized.Trim('/');
        return normalized;
    }

    public static string CombineRelativePath(params string?[] pathSegments)
    {
        return string.Join(
            '/',
            pathSegments
                .Select(NormalizeRelativePath)
                .Where(segment => !string.IsNullOrEmpty(segment)));
    }

    public static string GetParentPath(string relativePath)
    {
        var normalized = NormalizeRelativePath(relativePath);

        if (string.IsNullOrEmpty(normalized))
            return string.Empty;

        var lastSlashIndex = normalized.LastIndexOf('/');
        return lastSlashIndex < 0 ? string.Empty : normalized[..lastSlashIndex];
    }

    public static string EncodeGraphPath(string relativePath)
    {
        var normalized = NormalizeRelativePath(relativePath);
        return string.Join(
            '/',
            normalized
                .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(Uri.EscapeDataString));
    }
}
