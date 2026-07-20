using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Tests;

public sealed class MetadataFieldSelectorTests
{
    [Fact]
    public void SelectUpdatableFields_FiltersOutSystemAndUnknownFields()
    {
        var fields = new JsonObject
        {
            ["Title"] = "Document A",
            ["Category"] = "Contracts",
            ["Modified"] = "2026-01-01T00:00:00Z",
            ["UnknownColumn"] = "ignored"
        };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = new("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false),
            ["Category"] = new("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns);

        Assert.Equal(2, result.Fields.Count);
        Assert.Equal("Document A", result.Fields["Title"]?.GetValue<string>());
        Assert.Equal("Contracts", result.Fields["Category"]?.GetValue<string>());
        Assert.Contains("Modified", result.SkippedFields);
        Assert.Contains("UnknownColumn", result.SkippedFields);
    }

    [Fact]
    public void SelectUpdatableFields_RenamesFieldKey_WhenMappingIsProvided()
    {
        var fields = new JsonObject
        {
            ["CategoriaProgetto"] = "Infrastruttura",
            ["Title"] = "Doc"
        };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["CategoriaProgetto"] = new("CategoriaProgetto", "Categoria", "choice", Hidden: false, ReadOnly: false, Required: false),
            ["Title"] = new("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["CategoriaProgetto"] = "Category" };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns, mappings);

        // The field must appear under the target name "Category", not under the source name
        Assert.True(result.Fields.ContainsKey("Category"));
        Assert.False(result.Fields.ContainsKey("CategoriaProgetto"));
        Assert.Equal("Infrastruttura", result.Fields["Category"]?.GetValue<string>());

        // Fields without a mapping keep their original name
        Assert.True(result.Fields.ContainsKey("Title"));
    }

    [Fact]
    public void SelectUpdatableFields_KeepsOriginalKey_WhenNoMappingForField()
    {
        var fields = new JsonObject { ["Title"] = "Doc" };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = new("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        // Mapping is present but does not cover "Title"
        var mappings = new Dictionary<string, string> { ["OtherField"] = "OtherTarget" };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns, mappings);

        Assert.True(result.Fields.ContainsKey("Title"));
        Assert.Equal("Doc", result.Fields["Title"]?.GetValue<string>());
    }

    [Fact]
    public void SelectUpdatableFields_SkipsFieldsWithNullValue()
    {
        var fields = new JsonObject
        {
            ["Title"]    = "Doc",
            ["NullField"] = (JsonNode?)null
        };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"]     = new("Title",     "Title",      "text", Hidden: false, ReadOnly: false, Required: false),
            ["NullField"] = new("NullField", "Null Field", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns);

        // null-valued field is silently skipped (not in Fields, not in SkippedFields)
        Assert.Single(result.Fields);
        Assert.True(result.Fields.ContainsKey("Title"));
        Assert.False(result.Fields.ContainsKey("NullField"));
    }

    [Fact]
    public void SelectUpdatableFields_SkipsAtPrefixedKeys()
    {
        var fields = new JsonObject
        {
            ["Title"]           = "Doc",
            ["@odata.type"]     = "SP.Data.Item",
            ["@odata.editLink"] = "https://example.com"
        };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = new("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns);

        Assert.Single(result.Fields);
        Assert.True(result.Fields.ContainsKey("Title"));
    }

    [Fact]
    public void SelectUpdatableFields_AppliesValueTransformer_AndUsesReturnValue()
    {
        var fields = new JsonObject { ["Category"] = "original" };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Category"] = new("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = MetadataFieldSelector.SelectUpdatableFields(
            fields,
            writableColumns,
            fieldValueTransformer: (_, _) => JsonValue.Create("transformed"));

        Assert.Equal("transformed", result.Fields["Category"]?.GetValue<string>());
    }

    [Fact]
    public void SelectUpdatableFields_ValueTransformer_ReceivesTargetKeyAfterMapping()
    {
        var fields = new JsonObject { ["SourceField"] = "value" };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["SourceField"] = new("SourceField", "Source", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["SourceField"] = "TargetField" };

        string? receivedKey = null;
        MetadataFieldSelector.SelectUpdatableFields(
            fields,
            writableColumns,
            mappings,
            fieldValueTransformer: (key, val) => { receivedKey = key; return val?.DeepClone(); });

        Assert.Equal("TargetField", receivedKey);
    }

    [Fact]
    public void SelectUpdatableFields_WithNoMappingsArgument_UsesOriginalKeys()
    {
        var fields = new JsonObject { ["Title"] = "Doc", ["Score"] = "42" };

        var writableColumns = new Dictionary<string, SharePointColumnDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = new("Title", "Title", "text",   Hidden: false, ReadOnly: false, Required: false),
            ["Score"] = new("Score", "Score", "number", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = MetadataFieldSelector.SelectUpdatableFields(fields, writableColumns);

        Assert.True(result.Fields.ContainsKey("Title"));
        Assert.True(result.Fields.ContainsKey("Score"));
        Assert.Empty(result.SkippedFields);
    }
}
