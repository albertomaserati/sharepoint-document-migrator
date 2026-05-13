namespace SharePointDocumentMigrator.Migration;

public sealed class TemporaryFile : IDisposable
{
    private TemporaryFile(string path)
    {
        Path = path;
    }

    public string Path { get; }

    public static TemporaryFile Create(string? extension = null)
    {
        extension = string.IsNullOrWhiteSpace(extension)
            ? ".tmp"
            : extension.StartsWith('.') ? extension : $".{extension}";

        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");
        using (File.Create(path))
        {
        }

        return new TemporaryFile(path);
    }

    public void Dispose()
    {
        if (File.Exists(Path))
        {
            File.Delete(Path);
        }
    }
}
