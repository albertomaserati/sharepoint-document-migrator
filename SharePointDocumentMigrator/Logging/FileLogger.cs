using Microsoft.Extensions.Logging;

namespace SharePointDocumentMigrator.Logging;

internal sealed class FileLogger(string categoryName, string filePath, object writeLock) : ILogger
{
    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

    public bool IsEnabled(LogLevel logLevel) => logLevel != LogLevel.None;

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        if (!IsEnabled(logLevel))
            return;

        var message = formatter(state, exception);
        if (string.IsNullOrEmpty(message) && exception is null)
            return;

        var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{logLevel}] {categoryName}: {message}";
        if (exception is not null)
            line += Environment.NewLine + exception;

        lock (writeLock)
        {
            File.AppendAllText(filePath, line + Environment.NewLine);
        }
    }
}
