using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class PathUtilitiesTests
{
    // ──────────────────────────── NormalizeRelativePath ──────────────────────────

    [Theory]
    [InlineData("folder/sub",   "folder/sub")]
    [InlineData("/folder/sub/", "folder/sub")]
    [InlineData("folder\\sub",  "folder/sub")]
    public void NormalizeRelativePath_NormalizesSeparators(string input, string expected)
    {
        Assert.Equal(expected, PathUtilities.NormalizeRelativePath(input));
    }

    [Fact]
    public void NormalizeRelativePath_WithNull_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, PathUtilities.NormalizeRelativePath(null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void NormalizeRelativePath_WithEmptyOrWhitespace_ReturnsEmpty(string input)
    {
        Assert.Equal(string.Empty, PathUtilities.NormalizeRelativePath(input));
    }

    // ──────────────────────────── NormalizeSitePath ──────────────────────────────

    [Theory]
    [InlineData("/teams/intranet/it/", "teams/intranet/it")]
    [InlineData("teams\\intranet\\it", "teams/intranet/it")]
    [InlineData("sites/root",          "sites/root")]
    public void NormalizeSitePath_NormalizesInput(string input, string expected)
    {
        Assert.Equal(expected, PathUtilities.NormalizeSitePath(input));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NormalizeSitePath_WithEmptyOrWhitespace_Throws(string? input)
    {
        Assert.Throws<InvalidOperationException>(() => PathUtilities.NormalizeSitePath(input!));
    }

    // ──────────────────────────── CombineRelativePath ────────────────────────────

    [Fact]
    public void CombineRelativePath_JoinsNonEmptySegments()
    {
        var combined = PathUtilities.CombineRelativePath("/target", "", "folder\\child", "file.txt");

        Assert.Equal("target/folder/child/file.txt", combined);
    }

    [Fact]
    public void CombineRelativePath_AllEmptySegments_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, PathUtilities.CombineRelativePath("", null, "  "));
    }

    [Fact]
    public void CombineRelativePath_SingleSegment_ReturnsNormalizedSegment()
    {
        Assert.Equal("folder", PathUtilities.CombineRelativePath("/folder/"));
    }

    // ──────────────────────────── GetParentPath ──────────────────────────────────

    [Fact]
    public void GetParentPath_ReturnsFolderPortion()
    {
        Assert.Equal("target/folder", PathUtilities.GetParentPath("target/folder/file.txt"));
    }

    [Fact]
    public void GetParentPath_RootLevelFile_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, PathUtilities.GetParentPath("file.txt"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void GetParentPath_WithEmptyOrNull_ReturnsEmpty(string? input)
    {
        Assert.Equal(string.Empty, PathUtilities.GetParentPath(input!));
    }

    // ──────────────────────────── EncodeGraphPath ────────────────────────────────

    [Fact]
    public void EncodeGraphPath_EncodesSpacesInSegments()
    {
        var encoded = PathUtilities.EncodeGraphPath("My Documents/file name.pdf");

        Assert.Equal("My%20Documents/file%20name.pdf", encoded);
    }

    [Fact]
    public void EncodeGraphPath_EncodesSpecialCharacters()
    {
        var encoded = PathUtilities.EncodeGraphPath("folder/file (1).pdf");

        Assert.Contains("%28", encoded);   // (
        Assert.Contains("%29", encoded);   // )
    }

    [Fact]
    public void EncodeGraphPath_DoesNotDoubleEncodeSlashes()
    {
        var encoded = PathUtilities.EncodeGraphPath("a/b/c");

        Assert.Equal("a/b/c", encoded);
    }
}
