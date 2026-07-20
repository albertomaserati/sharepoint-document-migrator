using SharePointDocumentMigrator.Migration;

namespace SharePointDocumentMigrator.Tests;

public sealed class MigrationReportTests
{
    [Fact]
    public void NewReport_AllCountersAreZero()
    {
        var report = new MigrationReport(dryRun: false);

        Assert.Equal(0, report.PlannedFiles);
        Assert.Equal(0, report.PlannedFolders);
        Assert.Equal(0, report.MigratedFiles);
        Assert.Equal(0, report.SkippedFiles);
        Assert.Equal(0, report.FailedFiles);
        Assert.Equal(0, report.MetadataUpdated);
        Assert.Equal(0, report.FoldersEnsured);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DryRun_Property_ReflectsConstructorArgument(bool dryRun)
    {
        var report = new MigrationReport(dryRun);

        Assert.Equal(dryRun, report.DryRun);
    }

    [Fact]
    public void RecordDryRunFile_IncreasesPlannedFiles()
    {
        var report = new MigrationReport(dryRun: true);

        report.RecordDryRunFile("folder/a.pdf");
        report.RecordDryRunFile("folder/b.pdf");

        Assert.Equal(2, report.PlannedFiles);
    }

    [Fact]
    public void RecordDryRunFolder_IncreasesPlannedFolders()
    {
        var report = new MigrationReport(dryRun: true);

        report.RecordDryRunFolder("folder/sub");

        Assert.Equal(1, report.PlannedFolders);
    }

    [Fact]
    public void RecordMigratedFile_IncreasesMigratedFiles()
    {
        var report = new MigrationReport(dryRun: false);

        report.RecordMigratedFile("a/b.pdf");

        Assert.Equal(1, report.MigratedFiles);
    }

    [Fact]
    public void RecordSkippedFile_IncreasesSkippedFiles()
    {
        var report = new MigrationReport(dryRun: false);

        report.RecordSkippedFile("a/b.pdf");

        Assert.Equal(1, report.SkippedFiles);
    }

    [Fact]
    public void RecordFolderEnsured_IncreasesFoldersEnsured()
    {
        var report = new MigrationReport(dryRun: false);

        report.RecordFolderEnsured();
        report.RecordFolderEnsured();
        report.RecordFolderEnsured();

        Assert.Equal(3, report.FoldersEnsured);
    }

    [Fact]
    public void RecordMetadataUpdated_IncreasesMetadataUpdated()
    {
        var report = new MigrationReport(dryRun: false);

        report.RecordMetadataUpdated();
        report.RecordMetadataUpdated();

        Assert.Equal(2, report.MetadataUpdated);
    }

    [Fact]
    public void RecordFailure_IncreasesFailedFiles()
    {
        var report = new MigrationReport(dryRun: false);

        report.RecordFailure("a/b.pdf", "Upload failed");
        report.RecordFailure("a/c.pdf", "Timeout");

        Assert.Equal(2, report.FailedFiles);
    }

    [Fact]
    public void PrintReport_DryRun_ContainsDryRunCompletedAndCounts()
    {
        var report = new MigrationReport(dryRun: true);
        report.RecordDryRunFile("f1.pdf");
        report.RecordDryRunFile("f2.pdf");
        report.RecordDryRunFolder("folder1");

        var text = report.PrintReport();

        Assert.Contains("Dry run completed", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", text);  // planned files
        Assert.Contains("1", text);  // planned folders
    }

    [Fact]
    public void PrintReport_RealRun_ContainsMigrationCompletedAndCounts()
    {
        var report = new MigrationReport(dryRun: false);
        report.RecordMigratedFile("f1.pdf");
        report.RecordMigratedFile("f2.pdf");
        report.RecordSkippedFile("f3.pdf");
        report.RecordFailure("f4.pdf", "err");
        report.RecordMetadataUpdated();
        report.RecordFolderEnsured();
        report.RecordFolderEnsured();

        var text = report.PrintReport();

        Assert.Contains("Migration completed", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", text);  // migrated / folders
        Assert.Contains("1", text);  // skipped / failures / metadata
    }

    [Fact]
    public void PrintReport_DryRun_DoesNotContainMigrationCompletedText()
    {
        var report = new MigrationReport(dryRun: true);

        Assert.DoesNotContain("Migration completed", report.PrintReport(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PrintReport_RealRun_DoesNotContainDryRunText()
    {
        var report = new MigrationReport(dryRun: false);

        Assert.DoesNotContain("Dry run", report.PrintReport(), StringComparison.OrdinalIgnoreCase);
    }
}
