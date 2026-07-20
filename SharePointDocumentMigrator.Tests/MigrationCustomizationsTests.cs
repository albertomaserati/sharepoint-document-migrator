using SharePointDocumentMigrator.Migration;
using System.Text.Json.Nodes;

namespace SharePointDocumentMigrator.Tests;

public sealed class MigrationCustomizationsTests
{
    private readonly MigrationCustomizations _sut = new();

    // ──────────────────────────── TransformTargetPath ────────────────────────────

    [Theory]
    [InlineData("TEO01-I025-IT_V04.pdf",           "TEO01-I025-IT.pdf")]
    [InlineData("TEO01-I025-IT_V10.docx",          "TEO01-I025-IT.docx")]
    [InlineData("TEO01-I025-IT_V4.pdf",            "TEO01-I025-IT.pdf")]
    [InlineData("folder/TEO01-I025-IT_V04.pdf",    "folder/TEO01-I025-IT.pdf")]
    [InlineData("doc_V1_thing_V2.txt",             "doc_V1_thing.txt")]
    public void TransformTargetPath_WithVersionSuffix_StripsSuffix(string input, string expected)
    {
        Assert.Equal(expected, _sut.TransformTargetPath(input));
    }

    [Theory]
    [InlineData("CERT ISO 27001_BindingFuture.pdf")]
    [InlineData("noversion.pdf")]
    [InlineData("folder/document.docx")]
    [InlineData("TEO01_V2_description.pdf")]   // _V2 not at end
    public void TransformTargetPath_WithoutVersionSuffix_ReturnsUnchanged(string input)
    {
        Assert.Equal(input, _sut.TransformTargetPath(input));
    }

    // ──────────────────── TransformFieldValue — Reference_x0020_process ──────────

    [Theory]
    [InlineData("ACQ",              "Purchasing")]
    [InlineData("AFC",              "Finance")]
    [InlineData("AMM",              "Administration")]
    [InlineData("CDO",              "Innovation")]
    [InlineData("COMM",             "Communication")]
    [InlineData("Communication",    "Communication")]
    [InlineData("Datore di Lavoro", "Executive Management")]
    [InlineData("Dir",              "Executive Management")]
    [InlineData("PRES",             "Executive Management")]
    [InlineData("DPO",              "DPO - Privacy")]
    [InlineData("Ente Esterno",     "Third-party")]
    [InlineData("H&S",              "Health&Safety")]
    [InlineData("RSPP",             "Health&Safety")]
    [InlineData("Candidati;",       "Human Resources")]
    [InlineData("HR",               "Human Resources")]
    [InlineData("ICT",              "ICT")]
    [InlineData("Legal",            "Legal")]
    [InlineData("MKT",              "Strategic Marketing")]
    [InlineData("BM",               "Operations")]
    [InlineData("OPR",              "Operations")]
    [InlineData("OPS - Prod",       "Operations")]
    [InlineData("OPS-delivery",     "Operations")]
    [InlineData("OPS-production",   "Operations")]
    [InlineData("PM",               "Operations")]
    [InlineData("P&GS",             "Facility")]
    [InlineData("QM",               "Quality Management")]
    [InlineData("RISMS",            "Information Security")]
    [InlineData("RSGSI",            "Information Security")]
    [InlineData("UNKNOWN_ORG",      "")]
    public void TransformFieldValue_ReferenceProcess_MapsAllKnownValues(string sourceValue, string expected)
    {
        var value = new JsonArray { JsonValue.Create(sourceValue) };

        var result = _sut.TransformFieldValue("Reference_x0020_process", value);

        Assert.Equal(expected, result?.GetValue<string>());
    }

    [Fact]
    public void TransformFieldValue_ReferenceProcess_WithEmptyArray_ReturnsEmptyString()
    {
        var result = _sut.TransformFieldValue("Reference_x0020_process", new JsonArray());

        Assert.Equal(string.Empty, result?.GetValue<string>());
    }

    [Fact]
    public void TransformFieldValue_ReferenceProcess_WithNull_ReturnsEmptyString()
    {
        // null value → raw = string.Empty → MapEnteEmittenteToReferenceProcess("") → ""
        var result = _sut.TransformFieldValue("Reference_x0020_process", null);

        Assert.Equal(string.Empty, result?.GetValue<string>());
    }

    // ──────────────────── TransformFieldValue — Document_x0020_type ──────────────

    [Fact]
    public void TransformFieldValue_DocumentType_PassesThroughStringValue()
    {
        var result = _sut.TransformFieldValue("Document_x0020_type", JsonValue.Create("Generali"));

        Assert.Equal("Generali", result?.GetValue<string>());
    }

    [Fact]
    public void TransformFieldValue_DocumentType_WithNull_ReturnsEmptyString()
    {
        var result = _sut.TransformFieldValue("Document_x0020_type", null);

        Assert.Equal(string.Empty, result?.GetValue<string>());
    }

    // ──────────────────── TransformFieldValue — other fields ─────────────────────

    [Fact]
    public void TransformFieldValue_OtherStringField_ReturnsDeepClone()
    {
        var original = JsonValue.Create("hello");

        var result = _sut.TransformFieldValue("Title", original);

        Assert.Equal("hello", result?.GetValue<string>());
        Assert.False(ReferenceEquals(original, result), "Should be a deep clone, not the same instance.");
    }

    [Fact]
    public void TransformFieldValue_OtherField_WithNull_ReturnsNull()
    {
        var result = _sut.TransformFieldValue("Title", null);

        Assert.Null(result);
    }
}
