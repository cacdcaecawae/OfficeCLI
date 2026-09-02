// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>Controls whether proven Office-compatible diagnostics remain blocking.</summary>
public enum ValidationProfile
{
    Strict,
    OfficeCompatible,
}

/// <summary>A validation message with its profile-specific severity and classification.</summary>
public sealed record ValidationDiagnostic(
    string Type,
    string Description,
    string? Path,
    string? Part,
    string Severity,
    string Classification,
    string Reason);

/// <summary>The complete validation verdict for one profile.</summary>
public sealed record ValidationReport(
    ValidationProfile Profile,
    IReadOnlyList<ValidationDiagnostic> Diagnostics)
{
    public int ErrorCount => Diagnostics.Count(d => d.Severity == "error");
    public int WarningCount => Diagnostics.Count(d => d.Severity == "warning");
    public bool Success => ErrorCount == 0;
}

/// <summary>Parses validation profiles and evaluates strict SDK diagnostics.</summary>
public static class ValidationProfiles
{
    public const string StrictName = "strict";
    public const string OfficeCompatibleName = "office-compatible";

    /// <summary>Parse a CLI profile name, defaulting to strict validation.</summary>
    public static ValidationProfile Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)
            || string.Equals(value, StrictName, StringComparison.OrdinalIgnoreCase))
            return ValidationProfile.Strict;
        if (string.Equals(value, OfficeCompatibleName, StringComparison.OrdinalIgnoreCase))
            return ValidationProfile.OfficeCompatible;

        throw new CliException(
            $"Unknown validation profile '{value}'. Expected '{StrictName}' or '{OfficeCompatibleName}'.")
        {
            Code = "invalid_value",
            ValidValues = [StrictName, OfficeCompatibleName],
        };
    }

    /// <summary>Return the stable CLI spelling for a profile.</summary>
    public static string ToCliName(ValidationProfile profile) => profile switch
    {
        ValidationProfile.Strict => StrictName,
        ValidationProfile.OfficeCompatible => OfficeCompatibleName,
        _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, null),
    };

    /// <summary>Apply a validation profile without removing any diagnostic.</summary>
    public static ValidationReport Evaluate(
        IReadOnlyList<ValidationError> strictErrors,
        ValidationProfile profile)
    {
        var diagnostics = new List<ValidationDiagnostic>(strictErrors.Count);
        foreach (var error in strictErrors)
        {
            var isCompatibilityWarning = profile == ValidationProfile.OfficeCompatible
                && error.OfficeCompatibilityReason != null;
            var classification = error.OfficeCompatibilityReason != null
                ? "office-compatible-element-order"
                : ClassifyBlockingError(error.ErrorType);
            var reason = error.OfficeCompatibilityReason
                ?? BlockingReason(classification);
            diagnostics.Add(new ValidationDiagnostic(
                error.ErrorType,
                error.Description,
                error.Path,
                error.Part,
                isCompatibilityWarning ? "warning" : "error",
                classification,
                reason));
        }
        return new ValidationReport(profile, diagnostics);
    }

    private static string ClassifyBlockingError(string errorType) => errorType switch
    {
        "MalformedXml" or "OrphanedReference" or "PackageStructure" => "document-corruption",
        "ValidatorException" or "ValidatorNullReference" => "validator-failure",
        _ => "schema-validation-error",
    };

    private static string BlockingReason(string classification) => classification switch
    {
        "document-corruption" =>
            "The package contains malformed XML, a missing relationship, or another structural defect and remains invalid in every profile.",
        "validator-failure" =>
            "The SDK validator could not complete reliably, so the document cannot receive a passing verdict.",
        _ =>
            "This diagnostic is not one of the proven Office-compatible element-order cases and remains blocking in every profile.",
    };
}
