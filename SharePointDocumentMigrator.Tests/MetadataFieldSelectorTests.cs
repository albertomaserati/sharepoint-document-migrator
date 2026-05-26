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
}
