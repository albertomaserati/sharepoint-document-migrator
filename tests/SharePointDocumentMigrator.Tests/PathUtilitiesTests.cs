using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class PathUtilitiesTests
{
    [Theory]
    [InlineData("folder/sub", "folder/sub")]
    [InlineData("/folder/sub/", "folder/sub")]
    [InlineData("folder\\sub", "folder/sub")]
    public void NormalizeRelativePath_NormalizesSeparators(string input, string expected)
    {
        Assert.Equal(expected, PathUtilities.NormalizeRelativePath(input));
    }

    [Fact]
    public void CombineRelativePath_JoinsNonEmptySegments()
    {
        var combined = PathUtilities.CombineRelativePath("/target", "", "folder\\child", "file.txt");

        Assert.Equal("target/folder/child/file.txt", combined);
    }

    [Fact]
    public void GetParentPath_ReturnsFolderPortion()
    {
        Assert.Equal("target/folder", PathUtilities.GetParentPath("target/folder/file.txt"));
    }
}
