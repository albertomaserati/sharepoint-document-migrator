using Microsoft.Extensions.Options;

namespace SharePointDocumentMigrator.Options;

public sealed class AuthenticationOptionsValidator : IValidateOptions<AuthenticationOptions>
{
    public ValidateOptionsResult Validate(string? name, AuthenticationOptions options)
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(options.TenantId))
        {
            errors.Add("Authentication:TenantId is required.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            errors.Add("Authentication:ClientId is required.");
        }

        if (options.Scopes.Length == 0)
        {
            errors.Add("Authentication:Scopes must contain at least one Microsoft Graph delegated scope.");
        }

        return errors.Count > 0 ? ValidateOptionsResult.Fail(errors) : ValidateOptionsResult.Success;
    }
}
