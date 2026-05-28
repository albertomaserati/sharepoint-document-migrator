using SharePointDocumentMigrator.Graph;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Migration;

public static class MetadataFieldSelector
{
    private static readonly HashSet<string> ExplicitlyIgnoredFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "Attachments",
        "Author",
        "Created",
        "Created_x0020_Date",
        "DocIcon",
        "Editor",
        "FileLeafRef",
        "FileRef",
        "File_x0020_Type",
        "FolderChildCount",
        "GUID",
        "ID",
        "InstanceID",
        "ItemChildCount",
        "Modified",
        "Order",
        "Path",
        "SortBehavior",
        "_CommentCount",
        "_LikeCount",
        "_ModerationComments",
        "_ModerationStatus",
        "_UIVersionString"
    };

    public static MetadataSelectionResult SelectUpdatableFields(
        JsonObject fields,
        IReadOnlyDictionary<string, SharePointColumnDefinition> writableColumns,
        IReadOnlyDictionary<string, string>? fieldMappings = null,
        Func<string, JsonNode?, JsonNode?>? fieldValueTransformer = null)
    {
        var selectedFields = new JsonObject();
        var skippedFields = new List<string>();

        foreach (KeyValuePair<string, JsonNode?> property in fields)
        {
            if (property.Value is null || property.Key.StartsWith('@'))
            {
                continue;
            }

            if (ExplicitlyIgnoredFields.Contains(property.Key) || !writableColumns.ContainsKey(property.Key))
            {
                skippedFields.Add(property.Key);
                continue;
            }

            var targetKey = fieldMappings is not null && fieldMappings.TryGetValue(property.Key, out var mapped)
                ? mapped
                : property.Key;

            var value = fieldValueTransformer?.Invoke(targetKey, property.Value) ?? property.Value?.DeepClone();
            selectedFields[targetKey] = value;
        }

        return new MetadataSelectionResult(selectedFields, skippedFields);
    }

}
