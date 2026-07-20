using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace SharePointDocumentMigrator.Migration;

/// <summary>
/// Project-specific migration transformations.
/// Swap this implementation (or register a different one in DI) to change
/// how file paths and metadata values are mapped from source to target.
/// </summary>
public sealed partial class MigrationCustomizations : IMigrationCustomizations
{
    [GeneratedRegex(@"_V\d+(\.[^.]+)$", RegexOptions.Compiled)]
    private static partial Regex VersionRegex();

    private static readonly Regex VersionSuffixRegex = VersionRegex();

    /// <inheritdoc/>
    public string TransformTargetPath(string sourceRelativePath) =>
        VersionSuffixRegex.Replace(sourceRelativePath, "$1");

    /// <inheritdoc/>
    public JsonNode? TransformFieldValue(string fieldName, JsonNode? value)
    {
        var res = value?.DeepClone();

        if (fieldName == "Reference_x0020_process")
        {
            string raw = string.Empty;

            if (value is JsonArray arr && arr.Count > 0)
                raw = arr[0]?.GetValue<string>() ?? string.Empty;

            res = JsonValue.Create(MapEnteEmittenteToReferenceProcess(raw));
        }
        else if (fieldName == "Document_x0020_type")
        {
            res = JsonValue.Create(value?.GetValue<string>() ?? string.Empty);
        }

        return res;
    }

    private static string MapEnteEmittenteToReferenceProcess(string enteEmittente) =>
        enteEmittente switch
        {
            "ACQ" => "Purchasing",
            "AFC" => "Finance",
            "AMM" => "Administration",
            "CDO" => "Innovation",
            "COMM" or "Communication" => "Communication",
            "Datore di Lavoro" or "Dir" or "PRES" => "Executive Management",
            "DPO" => "DPO - Privacy",
            "Ente Esterno" => "Third-party",
            "H&S" or "RSPP" => "Health&Safety",
            "Candidati;" or "HR" => "Human Resources",
            "ICT" => "ICT",
            "Legal" => "Legal",
            "MKT" => "Strategic Marketing",
            "BM" or "OPR" or "OPS - Prod"
                or "OPS-delivery" or "OPS-production" or "PM" => "Operations",
            "P&GS" => "Facility",
            "QM" => "Quality Management",
            "RISMS" or "RSGSI" => "Information Security",
            _ => string.Empty
        };
}
