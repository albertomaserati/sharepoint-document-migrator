using System.Text.Json.Nodes;
using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class MetadataFieldSelectorTests
{
    [Fact]
    public void SelectUpdatableFields_FiltersOutSystemAndUnknownFields()
    {
        var selector = new MetadataFieldSelector();
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

        var result = selector.SelectUpdatableFields(fields, writableColumns);

        Assert.Equal(2, result.Fields.Count);
        Assert.Equal("Document A", result.Fields["Title"]?.GetValue<string>());
        Assert.Equal("Contracts", result.Fields["Category"]?.GetValue<string>());
        Assert.Contains("Modified", result.SkippedFields);
        Assert.Contains("UnknownColumn", result.SkippedFields);
    }
}
