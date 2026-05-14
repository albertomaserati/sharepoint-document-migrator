namespace SharePointDocumentMigrator.Graph;

public sealed class GraphApiException(
    string message,
    int statusCode,
    string requestUri,
    string? responseBody = null) : Exception(message)
{
    public int StatusCode { get; } = statusCode;

    public string RequestUri { get; } = requestUri;

    public string? ResponseBody { get; } = responseBody;
}
