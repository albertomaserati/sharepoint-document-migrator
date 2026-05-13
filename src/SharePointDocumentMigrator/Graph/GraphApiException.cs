namespace SharePointDocumentMigrator.Graph;

public sealed class GraphApiException : Exception
{
    public GraphApiException(string message, int statusCode, string requestUri, string? responseBody = null)
        : base(message)
    {
        StatusCode = statusCode;
        RequestUri = requestUri;
        ResponseBody = responseBody;
    }

    public int StatusCode { get; }

    public string RequestUri { get; }

    public string? ResponseBody { get; }
}
