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

    public async Task<SchemaValidationResult> Compare(
        IReadOnlyList<SharePointColumnDefinition> sourceColumns,
        IReadOnlyList<SharePointColumnDefinition> targetColumns,
        IReadOnlyDictionary<string, string>? fieldMappings = null)
    {
        Dictionary<string, SharePointColumnDefinition> targetLookup = targetColumns.ToDictionary(column => column.InternalName, StringComparer.OrdinalIgnoreCase);
        Dictionary<string, SharePointColumnDefinition> writableColumns = new(StringComparer.OrdinalIgnoreCase);
        var errors = new List<string>();

        await ExportColumnsOnTsvAsync(sourceColumns, targetColumns, Path.Combine(".", "metadata-columns"));

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

    /// <summary>
    /// Exports three TSV files describing the columns of the source and target libraries:
    /// <list type="bullet">
    ///   <item><description><c>source_columns.tsv</c> — all source columns (InternalName, DisplayName)</description></item>
    ///   <item><description><c>target_columns.tsv</c> — all target columns (InternalName, DisplayName)</description></item>
    ///   <item><description><c>columns_diff.tsv</c> — columns not shared between the two libraries, annotated with their origin (InternalName, DisplayName, Library)</description></item>
    /// </list>
    /// </summary>
    private static async Task ExportColumnsOnTsvAsync(
        IReadOnlyList<SharePointColumnDefinition> sourceColumns,
        IReadOnlyList<SharePointColumnDefinition> targetColumns,
        string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);

        await WriteColumnsTsv(
            Path.Combine(outputDirectory, "source_columns.tsv"),
            sourceColumns);

        await WriteColumnsTsv(
            Path.Combine(outputDirectory, "target_columns.tsv"),
            targetColumns);

        var sourceNames = sourceColumns
            .Select(c => c.InternalName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var targetNames = targetColumns
            .Select(c => c.InternalName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var onlyInSource = sourceColumns.Where(c => !targetNames.Contains(c.InternalName));
        var onlyInTarget = targetColumns.Where(c => !sourceNames.Contains(c.InternalName));

        using var diffWriter = new StreamWriter(Path.Combine(outputDirectory, "columns_diff.tsv"));
        await diffWriter.WriteLineAsync("InternalName\tDisplayName\tLibrary");

        foreach (var column in onlyInSource)
            await diffWriter.WriteLineAsync($"{column.InternalName}\t{column.DisplayName}\tSource");

        foreach (var column in onlyInTarget)
            await diffWriter.WriteLineAsync($"{column.InternalName}\t{column.DisplayName}\tTarget");
    }

    private static async Task WriteColumnsTsv(string filePath, IReadOnlyList<SharePointColumnDefinition> columns)
    {
        using var writer = new StreamWriter(filePath);
        await writer.WriteLineAsync("InternalName\tDisplayName");

        foreach (var column in columns)
            await writer.WriteLineAsync($"{column.InternalName}\t{column.DisplayName}");
    }
}
