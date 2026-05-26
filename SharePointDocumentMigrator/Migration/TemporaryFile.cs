namespace SharePointDocumentMigrator.Migration;

public static class TemporaryFile
{
    public static string Create(string? extension = null)
    {
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".tmp";
        }
        else if (!extension.StartsWith('.'))
        {
            extension = $".{extension}";
        }

        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");

        File.WriteAllBytes(path, []);

        return path;
    }

    public static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
