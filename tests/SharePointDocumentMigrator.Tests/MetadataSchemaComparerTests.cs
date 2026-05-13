using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class MetadataSchemaComparerTests
{
    [Fact]
    public void Compare_ReturnsWritableColumns_WhenSchemasMatch()
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

        var result = comparer.Compare(sourceColumns, targetColumns);

        Assert.True(result.IsValid);
        Assert.Equal(2, result.WritableColumns.Count);
    }

    [Fact]
    public void Compare_ReturnsErrors_WhenTargetColumnIsMissing()
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

        var result = comparer.Compare(sourceColumns, targetColumns);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains("Category", StringComparison.Ordinal));
    }
}
