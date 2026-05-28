using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Migration;

public interface IMigrationCustomizations
{
    /// <summary>
    /// Transforms the source file relative path into the target file relative path.
    /// Used to rename or clean file names before upload (e.g. strip version suffixes).
    /// Return the input unchanged to perform no transformation.
    /// </summary>
    string TransformTargetPath(string sourceRelativePath);

    /// <summary>
    /// Transforms a metadata field value before it is written to the target list item.
    /// Return the input value unchanged to perform no transformation.
    /// </summary>
    JsonNode? TransformFieldValue(string fieldName, JsonNode? value);
}
