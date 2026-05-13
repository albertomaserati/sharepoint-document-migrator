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
        IReadOnlyList<SharePointColumnDefinition> targetColumns,
        IReadOnlyDictionary<string, string>? fieldMappings = null)
    {
        var targetLookup = targetColumns.ToDictionary(column => column.InternalName, StringComparer.OrdinalIgnoreCase);
        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase);
        var errors = new List<string>();

        foreach (var sourceColumn in sourceColumns.Where(IsCandidate))
        {
            // Resolve the target column name: use the mapping if present, otherwise fall back to the source name
            var targetName = fieldMappings is not null && fieldMappings.TryGetValue(sourceColumn.InternalName, out var mapped)
                ? mapped
                : sourceColumn.InternalName;

            if (!targetLookup.TryGetValue(targetName, out var targetColumn))
            {
                var hint = targetName != sourceColumn.InternalName ? $" (mappato da '{sourceColumn.InternalName}')" : string.Empty;
                errors.Add($"Target library is missing writable column '{targetName}'{hint}.");
                continue;
            }

            if (!string.Equals(sourceColumn.Type, targetColumn.Type, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"Column '{sourceColumn.InternalName}' → '{targetName}' has incompatible types: source '{sourceColumn.Type}', target '{targetColumn.Type}'.");
                continue;
            }

            if (targetColumn.Hidden || targetColumn.ReadOnly)
            {
                errors.Add($"Target column '{targetName}' is not writable.");
                continue;
            }

            // WritableColumns is keyed by SOURCE name: the selector uses it to determine which
            // source fields are allowed; renaming to the target key happens in MetadataFieldSelector.
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
