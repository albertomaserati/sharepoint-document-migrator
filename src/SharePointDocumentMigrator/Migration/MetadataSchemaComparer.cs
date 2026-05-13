using SharePointDocumentMigrator.Graph;

namespace SharePointDocumentMigrator.Migration;

public sealed class MetadataSchemaComparer
{
    private static readonly HashSet<string> IgnoredColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "Attachments",
        "Author",
        "Created",
        "Created_x0020_Date",
        "Editor",
        "FileLeafRef",
        "FileRef",
        "GUID",
        "ID",
        "Modified",
        "ContentType",
        "ContentTypeId"
    };

    public SchemaValidationResult Compare(
        IReadOnlyList<SharePointColumnDefinition> sourceColumns,
        IReadOnlyList<SharePointColumnDefinition> targetColumns)
    {
        var targetLookup = targetColumns.ToDictionary(column => column.InternalName, StringComparer.OrdinalIgnoreCase);
        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase);
        var errors = new List<string>();

        foreach (var sourceColumn in sourceColumns.Where(IsCandidate))
        {
            if (!targetLookup.TryGetValue(sourceColumn.InternalName, out var targetColumn))
            {
                errors.Add($"Target library is missing writable column '{sourceColumn.InternalName}'.");
                continue;
            }

            if (!string.Equals(sourceColumn.Type, targetColumn.Type, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"Column '{sourceColumn.InternalName}' has incompatible types: source '{sourceColumn.Type}', target '{targetColumn.Type}'.");
                continue;
            }

            if (targetColumn.Hidden || targetColumn.ReadOnly)
            {
                errors.Add($"Target column '{sourceColumn.InternalName}' is not writable.");
                continue;
            }

            writableColumns[sourceColumn.InternalName] = sourceColumn;
        }

        return new SchemaValidationResult(writableColumns, errors);
    }

    private static bool IsCandidate(SharePointColumnDefinition column)
    {
        return !column.Hidden
            && !column.ReadOnly
            && !IgnoredColumns.Contains(column.InternalName);
    }
}
