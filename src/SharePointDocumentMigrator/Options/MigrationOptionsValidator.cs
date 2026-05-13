using Microsoft.Extensions.Options;

namespace SharePointDocumentMigrator.Options;

public sealed class MigrationOptionsValidator : IValidateOptions<MigrationOptions>
{
    public ValidateOptionsResult Validate(string? name, MigrationOptions options)
    {
        var errors = new List<string>();

        ValidateLocation("Migration:Source", options.Source, errors);
        ValidateLocation("Migration:Target", options.Target, errors);

        if (options.MaxRetryCount < 0)
        {
            errors.Add("Migration:MaxRetryCount must be zero or greater.");
        }

        if (options.UploadChunkSizeKiB <= 0 || options.UploadChunkSizeKiB % 320 != 0)
        {
            errors.Add("Migration:UploadChunkSizeKiB must be a positive multiple of 320.");
        }

        return errors.Count > 0 ? ValidateOptionsResult.Fail(errors) : ValidateOptionsResult.Success;
    }

    private static void ValidateLocation(string prefix, SharePointLocationOptions location, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(location.Hostname))
        {
            errors.Add($"{prefix}:Hostname is required.");
        }

        if (string.IsNullOrWhiteSpace(location.SitePath))
        {
            errors.Add($"{prefix}:SitePath is required.");
        }

        if (string.IsNullOrWhiteSpace(location.LibraryName))
        {
            errors.Add($"{prefix}:LibraryName is required.");
        }
    }
}
