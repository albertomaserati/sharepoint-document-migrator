using System.Text.Json.Nodes;
using SharePointDocumentMigrator.Graph;

namespace SharePointDocumentMigrator.Migration;

public sealed class MetadataFieldSelector
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

    public MetadataSelectionResult SelectUpdatableFields(
        JsonObject fields,
        IReadOnlyDictionary<string, SharePointColumnDefinition> writableColumns,
        IReadOnlyDictionary<string, string>? fieldMappings = null)
    {
        var selectedFields = new JsonObject();
        var skippedFields = new List<string>();

        foreach (var property in fields)
        {
            if (property.Value is null)
            {
                continue;
            }

            if (property.Key.StartsWith('@'))
            {
                continue;
            }

            if (ExplicitlyIgnoredFields.Contains(property.Key))
            {
                skippedFields.Add(property.Key);
                continue;
            }

            if (!writableColumns.ContainsKey(property.Key))
            {
                skippedFields.Add(property.Key);
                continue;
            }

            // If a mapping exists for this field, write it to the target using the mapped name;
            // otherwise keep the original source name.
            var targetKey = fieldMappings is not null && fieldMappings.TryGetValue(property.Key, out var mapped)
                ? mapped
                : property.Key;

            selectedFields[targetKey] = property.Value.DeepClone();
        }

        return new MetadataSelectionResult(selectedFields, skippedFields);
    }
}
