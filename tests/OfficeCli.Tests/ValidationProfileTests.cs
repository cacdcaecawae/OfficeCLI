// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeCli.Tests;

public sealed class ValidationProfileTests
{
    [Fact]
    public async Task CompatibilityFixtureHasLayeredVerdictsAndVisibleDiagnostics()
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);

        var defaultStrict = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--json");
        Assert.Equal(1, defaultStrict.ExitCode);
        using (var json = JsonDocument.Parse(defaultStrict.StdOut))
        {
            var root = json.RootElement;
            Assert.False(root.GetProperty("success").GetBoolean());
            var data = root.GetProperty("data");
            Assert.Equal("strict", data.GetProperty("profile").GetString());
            Assert.Equal(4, data.GetProperty("errorCount").GetInt32());
            Assert.Equal(0, data.GetProperty("warningCount").GetInt32());
            Assert.Equal(4, data.GetProperty("errors").GetArrayLength());
            Assert.Equal(4, data.GetProperty("diagnostics").GetArrayLength());
        }

        var explicitStrict = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "strict", "--json");
        Assert.Equal(1, explicitStrict.ExitCode);

        var invalidProfile = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "lenient", "--json");
        Assert.Equal(1, invalidProfile.ExitCode);
        using (var json = JsonDocument.Parse(invalidProfile.StdOut))
        {
            var root = json.RootElement;
            Assert.False(root.GetProperty("success").GetBoolean());
            Assert.Equal("invalid_value", root.GetProperty("error").GetProperty("code").GetString());
        }

        var compatible = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "office-compatible", "--json");
        Assert.Equal(0, compatible.ExitCode);
        using var compatibleJson = JsonDocument.Parse(compatible.StdOut);
        var compatibleRoot = compatibleJson.RootElement;
        Assert.True(compatibleRoot.GetProperty("success").GetBoolean());
        var compatibleData = compatibleRoot.GetProperty("data");
        Assert.Equal("office-compatible", compatibleData.GetProperty("profile").GetString());
        Assert.Equal(0, compatibleData.GetProperty("errorCount").GetInt32());
        Assert.Equal(4, compatibleData.GetProperty("warningCount").GetInt32());
        Assert.Equal(0, compatibleData.GetProperty("errors").GetArrayLength());
        Assert.Equal(4, compatibleData.GetProperty("warnings").GetArrayLength());
        Assert.Equal(4, compatibleData.GetProperty("diagnostics").GetArrayLength());

        var diagnostics = compatibleData.GetProperty("diagnostics").EnumerateArray().ToList();
        Assert.All(diagnostics, diagnostic =>
        {
            Assert.Equal("warning", diagnostic.GetProperty("severity").GetString());
            Assert.Equal(
                "office-compatible-element-order",
                diagnostic.GetProperty("classification").GetString());
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.GetProperty("reason").GetString()));
        });
        var reasons = diagnostics
            .Select(diagnostic => diagnostic.GetProperty("reason").GetString() ?? "")
            .ToList();
        Assert.Contains(reasons, reason => reason.Contains("m:scr", StringComparison.Ordinal));
        Assert.Contains(reasons, reason => reason.Contains("m:limLoc", StringComparison.Ordinal));
        Assert.Contains(reasons, reason => reason.Contains("m:plcHide", StringComparison.Ordinal));
        Assert.Contains(reasons, reason => reason.Contains("w:uiPriority", StringComparison.Ordinal));
    }

    [Fact]
    public async Task IllegalFormulaStructureRemainsBlockingInCompatibleProfile()
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateIllegalFormulaFixture(workspace.DirectoryPath);

        var result = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "office-compatible", "--json");

        Assert.Equal(1, result.ExitCode);
        using var json = JsonDocument.Parse(result.StdOut);
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        var data = root.GetProperty("data");
        Assert.True(data.GetProperty("errorCount").GetInt32() >= 1);
        Assert.Equal(4, data.GetProperty("warningCount").GetInt32());
        Assert.Contains(
            data.GetProperty("diagnostics").EnumerateArray(),
            diagnostic =>
                diagnostic.GetProperty("severity").GetString() == "error"
                && diagnostic.GetProperty("classification").GetString() == "schema-validation-error"
                && diagnostic.GetProperty("description").GetString()!.Contains("math:limLoc", StringComparison.Ordinal));
    }

    [Fact]
    public async Task DanglingRelationshipRemainsBlockingInCompatibleProfile()
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateDanglingRelationshipFixture(workspace.DirectoryPath);

        var result = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "office-compatible", "--json");

        Assert.Equal(1, result.ExitCode);
        using var json = JsonDocument.Parse(result.StdOut);
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        var diagnostics = root.GetProperty("data").GetProperty("diagnostics");
        Assert.Contains(
            diagnostics.EnumerateArray(),
            diagnostic =>
                diagnostic.GetProperty("severity").GetString() == "error"
                && diagnostic.GetProperty("classification").GetString() == "document-corruption"
                && diagnostic.GetProperty("type").GetString() == "OrphanedReference");
    }

    [Fact]
    public async Task HtmlPreviewAndAdjacentTextEditPreserveOmml()
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);
        var before = DocxFixtureFactory.ReadNormalizedOmml(fixture);
        var stats = DocxFixtureFactory.ReadFormulaStats(fixture);
        Assert.Equal(new FormulaStats(3, 1, 2, 1, 2), stats);

        var previewPath = Path.Combine(workspace.DirectoryPath, "preview.html");
        var preview = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "view", fixture, "html", "--out", previewPath);
        Assert.Equal(0, preview.ExitCode);
        var html = await File.ReadAllTextAsync(previewPath);
        var formulaMatches = Regex.Matches(
            html,
            "<span class=\"katex-formula\" data-formula=\"([^\"]+)\"",
            RegexOptions.CultureInvariant);
        Assert.Equal(3, formulaMatches.Count);
        Assert.All(
            formulaMatches.Cast<Match>(),
            match => Assert.False(string.IsNullOrWhiteSpace(match.Groups[1].Value)));

        var edit = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "set", fixture, "/body/p[1]/r[1]", "--prop", "text=Changed");
        Assert.Equal(0, edit.ExitCode);

        var after = DocxFixtureFactory.ReadNormalizedOmml(fixture);
        Assert.Equal(before, after);
        var validate = await CliTestDriver.RunAsync(
            workspace.DirectoryPath,
            "validate", fixture, "--profile", "office-compatible", "--json");
        Assert.Equal(0, validate.ExitCode);
    }
}
