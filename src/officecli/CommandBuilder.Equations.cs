// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildConvertEquationsCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Source DOCX (always read-only)" };
        var outOption = new Option<FileInfo?>("--out") { Description = "New DOCX containing native Word equations; never overwrites an existing file" };
        var dryRunOption = new Option<bool>("--dry-run") { Description = "Read equations and report conversion support without writing a document" };
        var preserveOption = new Option<bool>("--preserve-unsupported") { Description = "Keep unsupported equations as original OLE objects and report warnings; malformed equations still fail" };
        var command = new Command("convert-equations", "Read MathType equations and export native Word OMML in a new DOCX");
        command.Add(fileArg);
        command.Add(outOption);
        command.Add(dryRunOption);
        command.Add(preserveOption);
        command.Add(jsonOption);
        command.SetAction(result =>
        {
            bool json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                var output = result.GetValue(outOption);
                bool dryRun = result.GetValue(dryRunOption);
                if (dryRun == (output != null))
                    throw new CliException("Specify exactly one of --dry-run or --out <new.docx>.") { Code = "invalid_argument" };
                var report = MathTypeConverter.Convert(result.GetValue(fileArg)!.FullName, output?.FullName, result.GetValue(preserveOption));
                bool success = report["success"]!.GetValue<bool>();
                if (json) Console.WriteLine(report.ToJsonString(new JsonSerializerOptions { WriteIndented = true, TypeInfoResolver = AppJsonContext.Default }));
                else
                {
                    var data = report["data"]!;
                    Console.WriteLine($"Equations: {data["equations"]}; convertible: {data["convertible"]}; converted: {data["converted"]}; unsupported: {data["unsupported"]}; invalid: {data["invalid"]}");
                    foreach (var item in data["results"]!.AsArray())
                        Console.WriteLine($"  {item!["part"]} OLE[{item["objectIndex"]}] {item["status"]}: {item["text"] ?? item["message"]}");
                    foreach (var warning in report["warnings"]!.AsArray()) Console.WriteLine($"Warning: {warning!["message"]}");
                    if (success && output != null) Console.WriteLine($"Written: {output.FullName}");
                    if (!success) Console.WriteLine("Conversion failed; no output was written.");
                }
                return success ? 0 : 1;
            }, json);
        });
        return command;
    }
}
