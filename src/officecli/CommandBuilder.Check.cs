// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildValidateCommand(Option<bool> jsonOption)
    {
        var validateFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var profileOption = new Option<string>("--profile")
        {
            Description = "Validation profile: strict or office-compatible (default: strict)",
            DefaultValueFactory = _ => ValidationProfiles.StrictName,
        };
        var validateCommand = new Command("validate", "Validate document against OpenXML schema using a strict or Office-compatible profile");
        validateCommand.Add(validateFileArg);
        validateCommand.Add(profileOption);
        validateCommand.Add(jsonOption);
        validateCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(validateFileArg)!;
            var profile = ValidationProfiles.Parse(result.GetValue(profileOption));

            if (TryResident(file.FullName, req =>
            {
                req.Command = "validate";
                req.Json = json;
                req.Args["profile"] = ValidationProfiles.ToCliName(profile);
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var report = ValidationProfiles.Evaluate(handler.Validate(), profile);
            PrintValidationReport(report, json);
            return report.Success ? 0 : 1;
        }, json); });

        return validateCommand;
    }

    internal static void PrintValidationReport(ValidationReport report, bool json)
    {
        if (json)
        {
            Console.WriteLine(OutputFormatter.WrapEnvelope(
                FormatValidationReport(report),
                success: report.Success));
            return;
        }

        if (report.Diagnostics.Count == 0)
        {
            Console.WriteLine("Validation passed: no errors found.");
            return;
        }

        if (report.Profile == ValidationProfile.Strict)
        {
            Console.Error.WriteLine($"Found {report.ErrorCount} validation error(s):");
            foreach (var diagnostic in report.Diagnostics)
            {
                Console.Error.WriteLine($"  [{diagnostic.Type}] {diagnostic.Description}");
                if (diagnostic.Path != null) Console.Error.WriteLine($"    Path: {diagnostic.Path}");
                if (diagnostic.Part != null) Console.Error.WriteLine($"    Part: {diagnostic.Part}");
            }
            return;
        }

        if (report.Success)
        {
            Console.WriteLine(
                $"Validation passed under the {ValidationProfiles.OfficeCompatibleName} profile with {report.WarningCount} compatibility warning(s).");
            Console.Error.WriteLine($"Found {report.WarningCount} Office-compatible warning(s):");
        }
        else
        {
            Console.Error.WriteLine(
                $"Found {report.ErrorCount} validation error(s) and {report.WarningCount} Office-compatible warning(s):");
        }

        foreach (var diagnostic in report.Diagnostics)
        {
            Console.Error.WriteLine(
                $"  [{diagnostic.Severity.ToUpperInvariant()} {diagnostic.Type}] {diagnostic.Description}");
            if (diagnostic.Path != null) Console.Error.WriteLine($"    Path: {diagnostic.Path}");
            if (diagnostic.Part != null) Console.Error.WriteLine($"    Part: {diagnostic.Part}");
            Console.Error.WriteLine($"    Classification: {diagnostic.Classification}");
            Console.Error.WriteLine($"    Reason: {diagnostic.Reason}");
        }
    }
}
