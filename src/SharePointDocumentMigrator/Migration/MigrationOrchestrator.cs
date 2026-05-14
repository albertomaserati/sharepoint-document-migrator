using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Options;

namespace SharePointDocumentMigrator.Migration;

public sealed class MigrationOrchestrator(
    GraphApiClient graphApiClient,
    MetadataFieldSelector fieldSelector,
    IOptions<MigrationOptions> migrationOptions,
    ILogger<MigrationOrchestrator> logger)
{
    private readonly GraphApiClient _graphApiClient = graphApiClient;
    private readonly MetadataFieldSelector _fieldSelector = fieldSelector;
    private readonly MigrationOptions _migrationOptions = migrationOptions.Value;
    private readonly ILogger<MigrationOrchestrator> _logger = logger;

    private MigrationReport _report;

    public async Task<MigrationReport> ExecuteAsync(CancellationToken cancellationToken)
    {
        _report = new MigrationReport(_migrationOptions.DryRun);

        _logger.LogInformation("Resolving source and target SharePoint locations.");
        ResolvedLibraryContext sourceContext = await _graphApiClient.ResolveLibraryContextAsync(_migrationOptions.Source, cancellationToken);
        ResolvedLibraryContext targetContext = await _graphApiClient.ResolveLibraryContextAsync(_migrationOptions.Target, cancellationToken);

        SchemaValidationResult? schemaValidation = await MetadataSchemaComparer.Compare(
            sourceContext.Columns,
            targetContext.Columns,
            _migrationOptions.MetadataFieldMappings);

        if (!_migrationOptions.DryRun && _migrationOptions.EnableMetadataMigration && !schemaValidation.IsValid)
        {
            throw new InvalidOperationException(
                "Source and target libraries are not schema-compatible:" + Environment.NewLine
                + string.Join(Environment.NewLine, schemaValidation.Errors.Select(error => $" - {error}")));
        }

        var sourceRoot = PathUtilities.NormalizeRelativePath(_migrationOptions.Source.RootFolder);
        var targetRoot = PathUtilities.NormalizeRelativePath(_migrationOptions.Target.RootFolder);
        var inventory = await EnumerateAsync(sourceContext.Drive.Id, sourceRoot, cancellationToken);

        foreach (var folder in inventory.Folders)
        {
            var targetFolderPath = PathUtilities.CombineRelativePath(targetRoot, folder.RelativePath);
            if (_migrationOptions.DryRun)
            {
                _report.RecordDryRunFolder(targetFolderPath);
                continue;
            }

            await _graphApiClient.EnsureFolderExistsAsync(targetContext.Drive.Id, targetFolderPath, cancellationToken);
            _report.RecordFolderEnsured();
        }

        foreach (DriveItemInfo file in inventory.Files)
        {
            var targetFilePath = PathUtilities.CombineRelativePath(targetRoot, file.RelativePath);

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

                using var temporaryFile = TemporaryFile.Create(Path.GetExtension(file.Name));
                await _graphApiClient.DownloadFileAsync(
                    sourceContext.Drive.Id,
                    file.Id,
                    temporaryFile.Path,
                    cancellationToken);

                UploadResult? uploadResult = await _graphApiClient.UploadFileAsync(
                    targetContext.Drive.Id,
                    targetFilePath,
                    temporaryFile.Path,
                    _migrationOptions.OverwriteExistingFiles,
                    cancellationToken);

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
                        schemaValidation,
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

    private async Task ProcessMetadataAsync(
        ResolvedLibraryContext sourceContext,
        ResolvedLibraryContext targetContext,
        DriveItemInfo file,
        SchemaValidationResult schemaValidation,
        UploadResult uploadResult,

        CancellationToken cancellationToken)
    {
        var sourceFields = await _graphApiClient.GetListItemFieldsAsync(sourceContext.Drive.Id, file.Id, cancellationToken);
        var selectedMetadata = _fieldSelector.SelectUpdatableFields(
            sourceFields,
            schemaValidation!.WritableColumns,
            _migrationOptions.MetadataFieldMappings);

        if (selectedMetadata.Fields.Count > 0)
        {
            var targetListItemId = await _graphApiClient.GetListItemIdForDriveItemAsync(
                targetContext.Drive.Id,
                uploadResult.Item.Id,
                cancellationToken);

            await _graphApiClient.UpdateListItemFieldsAsync(
                targetContext.Site.Id,
                targetContext.List.Id,
                targetListItemId,
                selectedMetadata.Fields,
                cancellationToken);

            _report.RecordMetadataUpdated();
        }
    }

    private async Task<LibraryInventory> EnumerateAsync(string driveId, string sourceRoot, CancellationToken cancellationToken)
    {
        var root = await _graphApiClient.GetDriveItemByPathAsync(driveId, sourceRoot, cancellationToken);
        if (!root.IsFolder)
        {
            throw new InvalidOperationException("The configured SharePoint source root must be a folder.");
        }

        var folders = new List<DriveItemInfo>();
        var files = new List<DriveItemInfo>();
        var pendingFolders = new Queue<DriveItemInfo>();
        pendingFolders.Enqueue(new DriveItemInfo(root.Id, root.Name, string.Empty, true, root.WebUrl, root.Size));

        while (pendingFolders.Count > 0)
        {
            var currentFolder = pendingFolders.Dequeue();
            var children = await _graphApiClient.ListChildrenAsync(driveId, currentFolder.Id, currentFolder.RelativePath, cancellationToken);

            foreach (var child in children)
            {
                if (child.IsFolder)
                {
                    folders.Add(child);
                    pendingFolders.Enqueue(child);
                }
                else
                {
                    files.Add(child);
                }
            }
        }

        return new LibraryInventory(folders, files);
    }

    private sealed record LibraryInventory(IReadOnlyList<DriveItemInfo> Folders, IReadOnlyList<DriveItemInfo> Files);
}
