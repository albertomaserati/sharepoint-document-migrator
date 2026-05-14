using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class MetadataSchemaComparerTests
{
    [Fact]
    public async Task Compare_ReturnsWritableColumns_WhenSchemasMatch()
    {
        var comparer = new MetadataSchemaComparer();
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.Compare(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.Equal(2, result.WritableColumns.Count);
    }

    [Fact]
    public async Task Compare_ReturnsErrors_WhenTargetColumnIsMissing()
    {
        var comparer = new MetadataSchemaComparer();
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.Compare(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains("Category", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Compare_ResolvesTargetColumnViaMapping_WhenNamesAreDifferent()
    {
        var comparer = new MetadataSchemaComparer();
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("CategoriaProgetto", "Categoria Progetto", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["CategoriaProgetto"] = "Category" };

        var result = await MetadataSchemaComparer.Compare(sourceColumns, targetColumns, mappings);

        Assert.True(result.IsValid);
        // WritableColumns is keyed by source name
        Assert.True(result.WritableColumns.ContainsKey("CategoriaProgetto"));
    }

    [Fact]
    public async Task Compare_ReturnsError_WhenMappedTargetColumnIsMissing()
    {
        var comparer = new MetadataSchemaComparer();
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("CategoriaProgetto", "Categoria Progetto", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["CategoriaProgetto"] = "NonExistentColumn" };

        var result = await MetadataSchemaComparer.Compare(sourceColumns, targetColumns, mappings);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains("NonExistentColumn", StringComparison.Ordinal));
    }
}
