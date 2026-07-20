using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class TemporaryFileTests
{
    [Fact]
    public void Create_WithNoArgument_CreatesFileWithTmpExtension()
    {
        var path = TemporaryFile.Create();
        try
        {
            Assert.True(File.Exists(path));
            Assert.Equal(".tmp", Path.GetExtension(path));
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_WithNullExtension_CreatesFileWithTmpExtension()
    {
        var path = TemporaryFile.Create(null);
        try
        {
            Assert.Equal(".tmp", Path.GetExtension(path));
            Assert.True(File.Exists(path));
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_WithWhitespaceExtension_CreatesFileWithTmpExtension()
    {
        var path = TemporaryFile.Create("   ");
        try
        {
            Assert.Equal(".tmp", Path.GetExtension(path));
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_WithExtensionWithoutDot_PrependsDot()
    {
        var path = TemporaryFile.Create("pdf");
        try
        {
            Assert.Equal(".pdf", Path.GetExtension(path));
            Assert.True(File.Exists(path));
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_WithExtensionWithDot_UsesExtensionAsIs()
    {
        var path = TemporaryFile.Create(".docx");
        try
        {
            Assert.Equal(".docx", Path.GetExtension(path));
            Assert.True(File.Exists(path));
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_CreatesEmptyFile()
    {
        var path = TemporaryFile.Create();
        try
        {
            Assert.Equal(0, new FileInfo(path).Length);
        }
        finally
        {
            TemporaryFile.Delete(path);
        }
    }

    [Fact]
    public void Create_EachCall_ReturnsDistinctPath()
    {
        var path1 = TemporaryFile.Create();
        var path2 = TemporaryFile.Create();
        try
        {
            Assert.NotEqual(path1, path2);
        }
        finally
        {
            TemporaryFile.Delete(path1);
            TemporaryFile.Delete(path2);
        }
    }

    [Fact]
    public void Delete_RemovesExistingFile()
    {
        var path = TemporaryFile.Create();
        Assert.True(File.Exists(path));

        TemporaryFile.Delete(path);

        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Delete_OnNonExistentPath_DoesNotThrow()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid():N}.tmp");

        var ex = Record.Exception(() => TemporaryFile.Delete(path));

        Assert.Null(ex);
    }
}
