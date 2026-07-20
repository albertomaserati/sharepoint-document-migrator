using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;

namespace SharePointDocumentMigrator.Logging;

[ProviderAlias("File")]
public sealed class FileLoggerProvider : ILoggerProvider
{
    private readonly string _filePath;
    private readonly object _writeLock = new();
    private readonly ConcurrentDictionary<string, FileLogger> _loggers = new();

    public FileLoggerProvider(string filePath)
    {
        _filePath = filePath;

        var directory = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);
    }

    public ILogger CreateLogger(string categoryName) =>
        _loggers.GetOrAdd(categoryName, name => new FileLogger(name, _filePath, _writeLock));

    public void Dispose() => _loggers.Clear();
}
