using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Options;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Migration;

public sealed class MigrationOrchestrator(
    GraphApiClient graphApiClient,
    IMigrationCustomizations customizations,
    IOptions<MigrationOptions> migrationOptions,
    ILogger<MigrationOrchestrator> logger)
{
    private readonly GraphApiClient _graphApiClient = graphApiClient;
    private readonly IMigrationCustomizations _customizations = customizations;
    private readonly MigrationOptions _migrationOptions = migrationOptions.Value;
    private readonly ILogger<MigrationOrchestrator> _logger = logger;

    private MigrationReport _report = null!;

    public async Task<MigrationReport> ExecuteAsync(CancellationToken cancellationToken)
    {
        _report = new MigrationReport(_migrationOptions.DryRun);

        _logger.LogInformation("Resolving source and target SharePoint locations.");
        ResolvedLibraryContext sourceContext = await _graphApiClient.ResolveLibraryContextAsync(_migrationOptions.Source, cancellationToken);
        ResolvedLibraryContext targetContext = await _graphApiClient.ResolveLibraryContextAsync(_migrationOptions.Target, cancellationToken);
        SchemaValidationResult? schemaValidation = await ValidateSchemaAsync(sourceContext, targetContext);

        string sourceRoot = PathUtilities.NormalizeRelativePath(_migrationOptions.Source.RootFolder);
        string targetRoot = PathUtilities.NormalizeRelativePath(_migrationOptions.Target.RootFolder);
        LibraryInventory inventory = await EnumerateGraphAsync(sourceContext.Drive.Id, sourceRoot, cancellationToken);

        foreach (string? folderRelativePath in inventory.Folders.Select(f => f.RelativePath))
        {
            string targetFolderPath = PathUtilities.CombineRelativePath(targetRoot, folderRelativePath);

            if (_migrationOptions.DryRun)
            {
                _logger.LogInformation("[DRY RUN] Folder found — source: {SourcePath} | target: {TargetPath}", folderRelativePath, targetFolderPath);
                _report.RecordDryRunFolder(targetFolderPath);
                continue;
            }

            await _graphApiClient.EnsureFolderExistsAsync(targetContext.Drive.Id, targetFolderPath, cancellationToken);

            _report.RecordFolderEnsured();
        }

        foreach (DriveItemInfo file in inventory.Files)
        {
            string targetFilePath = PathUtilities.CombineRelativePath(targetRoot, _customizations.TransformTargetPath(file.RelativePath));

            try
            {
                if (_migrationOptions.DryRun)
                {
                    _logger.LogInformation("Source file name: {SourceFile}. Url: {WebUrl}", file.RelativePath, file.WebUrl);
                    _report.RecordDryRunFile(targetFilePath);

                    continue;
                }

                await _graphApiClient.EnsureFolderExistsAsync(
                    targetContext.Drive.Id,
                    PathUtilities.GetParentPath(targetFilePath),
                    cancellationToken);

                var tempPath = TemporaryFile.Create(Path.GetExtension(file.Name));

                await _graphApiClient.DownloadFileAsync(
                    sourceContext.Drive.Id,
                    file.Id,
                    tempPath,
                    cancellationToken);

                UploadResult uploadResult = await _graphApiClient.UploadFileAsync(
                    targetContext.Drive.Id,
                    targetFilePath,
                    tempPath,
                    _migrationOptions.OverwriteExistingFiles,
                    cancellationToken);

                TemporaryFile.Delete(tempPath);

                if (uploadResult.WasSkipped)
                {
                    _report.RecordSkippedFile(targetFilePath);
                    continue;
                }

                if (_migrationOptions.EnableMetadataMigration)
                {
                    await ProcessMetadataAsync(
                        sourceContext,
                        targetContext,
                        file,
                        schemaValidation!,
                        uploadResult,
                        cancellationToken);
                }

                _report.RecordMigratedFile(targetFilePath);
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Failed to migrate '{RelativePath}'.", file.RelativePath);
                _report.RecordFailure(file.RelativePath, exception.Message);
            }
        }

        _logger.LogInformation(_report.PrintReport());
        return _report;
    }

    private async Task<SchemaValidationResult?> ValidateSchemaAsync(ResolvedLibraryContext sourceContext, ResolvedLibraryContext targetContext)
    {
        SchemaValidationResult? res = null;

        if (_migrationOptions.EnableMetadataMigration)
        {
            res = await MetadataSchemaComparer.CompareAsync(
                sourceContext.Columns,
                targetContext.Columns,
                _migrationOptions.MetadataFieldMappings);

            if (!_migrationOptions.DryRun && !res.IsValid)
            {
                throw new InvalidOperationException(
                    "Source and target libraries are not schema-compatible:" + Environment.NewLine
                    + string.Join(Environment.NewLine, res.Errors.Select(e => $" - {e}")));
            }
        }

        return res;
    }

    private async Task<LibraryInventory> EnumerateGraphAsync(string driveId, string sourceRoot, CancellationToken cancellationToken)
    {
        var root = await _graphApiClient.GetDriveItemByPathAsync(driveId, sourceRoot, cancellationToken);
        if (!root.IsFolder)
        {
            throw new InvalidOperationException("The configured SharePoint source root must be a folder.");
        }

        List<DriveItemInfo> folders = [];
        List<DriveItemInfo> files = [];
        Queue<DriveItemInfo> pending = new();

        pending.Enqueue(new DriveItemInfo(
            root.Id,
            root.Name,
            string.Empty,
            true,
            root.WebUrl,
            root.Size));

        while (pending.Count > 0)
        {
            var current = pending.Dequeue();
            var children = await _graphApiClient.ListChildrenAsync(driveId, current.Id, current.RelativePath, cancellationToken);

            foreach (var child in children)
            {
                if (child.IsFolder)
                {
                    folders.Add(child);
                    pending.Enqueue(child);
                }
                else
                {
                    files.Add(child);
                }
            }
        }

        return new LibraryInventory(folders, files);
    }

    private async Task ProcessMetadataAsync(
        ResolvedLibraryContext sourceContext,
        ResolvedLibraryContext targetContext,
        DriveItemInfo file,
        SchemaValidationResult schemaValidation,
        UploadResult uploadResult,
        CancellationToken cancellationToken)
    {
        var sourceFields = await _graphApiClient.GetListItemFieldsAsync(sourceContext.Drive.Id, file.Id, cancellationToken);

        var selectedMetadata = MetadataFieldSelector.SelectUpdatableFields(
            sourceFields,
            schemaValidation.WritableColumns,
            _migrationOptions.MetadataFieldMappings,
            _customizations.TransformFieldValue);

        if (selectedMetadata.Fields.Count > 0)
        {
            var targetListItemId = await _graphApiClient.GetListItemIdForDriveItemAsync(
                targetContext.Drive.Id,
                uploadResult.Item.Id,
                cancellationToken);

            var resolvedFields = await ResolveLookupFieldsAsync(
                selectedMetadata.Fields,
                targetContext.Site.Id,
                targetContext.Columns,
                cancellationToken);

            if (resolvedFields.Count > 0)
            {
                await _graphApiClient.UpdateListItemFieldsAsync(
                    targetContext.Site.Id,
                    targetContext.List.Id,
                    targetListItemId,
                    resolvedFields,
                    cancellationToken);

                _report.RecordMetadataUpdated();
            }
        }
    }

    /// <summary>
    /// Replaces lookup-type fields in <paramref name="fields"/> with their integer
    /// <c>{FieldName}LookupId</c> equivalents required by the Graph PATCH endpoint.
    /// Non-lookup fields are passed through unchanged.
    /// </summary>
    private async Task<JsonObject> ResolveLookupFieldsAsync(
        JsonObject fields,
        string siteId,
        IReadOnlyList<SharePointColumnDefinition> targetColumns,
        CancellationToken cancellationToken)
    {
        Dictionary<string, SharePointColumnDefinition> columnsByName =
            targetColumns.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);

        JsonObject result = [];

        foreach (var property in fields)
        {
            if (!columnsByName.TryGetValue(property.Key, out SharePointColumnDefinition? column)
                || column.Lookup is null)
            {
                result[property.Key] = property.Value?.DeepClone();
                continue;
            }

            string? displayValue = ExtractDisplayValue(property.Value);

            if (string.IsNullOrEmpty(displayValue))
            {
                _logger.LogDebug("Skipping lookup field '{Field}': value is empty.", property.Key);
                continue;
            }

            int? lookupId = await _graphApiClient.ResolveLookupIdAsync(
                siteId,
                column.Lookup,
                displayValue,
                cancellationToken);

            if (lookupId is null)
            {
                _logger.LogWarning(
                    "Could not resolve lookup value '{Value}' for field '{Field}' — field will be skipped.",
                    displayValue,
                    property.Key);

                continue;
            }

            result[$"{property.Key}LookupId"] = lookupId.Value;
        }

        return result;
    }

    private static string? ExtractDisplayValue(JsonNode? value) =>
        value switch
        {
            JsonValue v when v.TryGetValue<string>(out var s) => s,
            JsonArray { Count: > 0 } arr => arr[0]?.GetValue<string>(),
            _ => null
        };

    private sealed record LibraryInventory(IReadOnlyList<DriveItemInfo> Folders, IReadOnlyList<DriveItemInfo> Files);
}