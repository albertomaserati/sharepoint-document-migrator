using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class MetadataSchemaComparerTests
{
    [Fact]
    public async Task Compare_ReturnsWritableColumns_WhenSchemasMatch()
    {
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

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.Equal(2, result.WritableColumns.Count);
    }

    [Fact]
    public async Task Compare_ReturnsErrors_WhenTargetColumnIsMissing()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains("Category", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Compare_ResolvesTargetColumnViaMapping_WhenNamesAreDifferent()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("CategoriaProgetto", "Categoria Progetto", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Category", "Category", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["CategoriaProgetto"] = "Category" };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns, mappings);

        Assert.True(result.IsValid);
        // WritableColumns is keyed by source name
        Assert.True(result.WritableColumns.ContainsKey("CategoriaProgetto"));
    }

    [Fact]
    public async Task Compare_ReturnsError_WhenMappedTargetColumnIsMissing()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("CategoriaProgetto", "Categoria Progetto", "choice", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var mappings = new Dictionary<string, string> { ["CategoriaProgetto"] = "NonExistentColumn" };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns, mappings);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains("NonExistentColumn", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Compare_ExcludesIgnoredColumns_FromWritableColumns()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Author",   "Author",   "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Modified", "Modified", "text", Hidden: false, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Title",    "Title",    "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.False(result.WritableColumns.ContainsKey("Author"));
        Assert.False(result.WritableColumns.ContainsKey("Modified"));
        Assert.True(result.WritableColumns.ContainsKey("Title"));
    }

    [Fact]
    public async Task Compare_ExcludesHiddenSourceColumns()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("HiddenField", "Hidden", "text", Hidden: true, ReadOnly: false, Required: false),
            new SharePointColumnDefinition("Title",       "Title",  "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.False(result.WritableColumns.ContainsKey("HiddenField"));
    }

    [Fact]
    public async Task Compare_ExcludesReadOnlySourceColumns()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("ReadOnlyField", "RO", "text", Hidden: false, ReadOnly: true, Required: false),
            new SharePointColumnDefinition("Title",         "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.False(result.WritableColumns.ContainsKey("ReadOnlyField"));
    }

    [Fact]
    public async Task Compare_ReturnsError_WhenTypesAreDifferent()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Score", "Score", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Score", "Score", "number", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("Score", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("Teoresi_x0020_Legal_x0020_Entity")]
    [InlineData("Language")]
    [InlineData("Document_x0020_type")]
    [InlineData("Reference_x0020_process")]
    public async Task Compare_AllowsTypeMismatch_ForExplicitlyAllowedLookupFields(string targetFieldName)
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition(targetFieldName, targetFieldName, "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition(targetFieldName, targetFieldName, "lookup", Hidden: false, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.True(result.WritableColumns.ContainsKey(targetFieldName));
    }

    [Fact]
    public async Task Compare_ReturnsError_WhenTargetColumnIsHidden()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: true, ReadOnly: false, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("Title", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Compare_ReturnsError_WhenTargetColumnIsReadOnly()
    {
        var sourceColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: false, Required: false)
        };

        var targetColumns = new[]
        {
            new SharePointColumnDefinition("Title", "Title", "text", Hidden: false, ReadOnly: true, Required: false)
        };

        var result = await MetadataSchemaComparer.CompareAsync(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("Title", StringComparison.Ordinal));
    }
}
