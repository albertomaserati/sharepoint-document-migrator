using SharePointDocumentMigrator.Graph;

namespace SharePointDocumentMigrator.Migration;

public sealed class MigrationReport(bool dryRun)
{
    private readonly List<string> _plannedFiles = [];
    private readonly List<string> _plannedFolders = [];
    private readonly List<string> _migratedFiles = [];
    private readonly List<string> _skippedFiles = [];
    private readonly List<FileFailure> _failures = [];

    public bool DryRun { get; } = dryRun;

    // Counts
    public int PlannedFiles => _plannedFiles.Count;
    public int PlannedFolders => _plannedFolders.Count;
    public int MigratedFiles => _migratedFiles.Count;
    public int SkippedFiles => _skippedFiles.Count;
    public int FailedFiles => _failures.Count;
    public int MetadataUpdated { get; private set; }
    public int FoldersEnsured { get; private set; }

    public void RecordDryRunFile(string relativePath) => _plannedFiles.Add(relativePath);

    public void RecordDryRunFolder(string relativePath) => _plannedFolders.Add(relativePath);

    public void RecordMigratedFile(string relativePath) => _migratedFiles.Add(relativePath);

    public void RecordSkippedFile(string relativePath) => _skippedFiles.Add(relativePath);

    public void RecordFolderEnsured() => FoldersEnsured++;

    public void RecordMetadataUpdated() => MetadataUpdated++;

    public void RecordFailure(string relativePath, string errorMessage) =>
        _failures.Add(new FileFailure(relativePath, errorMessage));

    public string PrintReport() => DryRun
        ? $"Dry run completed. Planned files: {PlannedFiles}, planned folders: {PlannedFolders}."
        : $"Migration completed. Migrated files: {MigratedFiles}, skipped: {SkippedFiles}, failures: {FailedFiles}, updated metadata: {MetadataUpdated}, folders created: {FoldersEnsured}.";
}
