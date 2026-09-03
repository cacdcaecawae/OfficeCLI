// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using static OfficeCli.Tests.MathTypeTestData;

namespace OfficeCli.Tests;

public class MathTypeConversionTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void DryRunReadsMathematicsAndLeavesSourceByteIdentical()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx");
        CreateDocument(source, suite: true);
        byte[] before = File.ReadAllBytes(source);
        var report = MathTypeConverter.Convert(source, null);
        Assert.True(report["success"]!.GetValue<bool>());
        Assert.Equal(6, report["data"]!["convertible"]!.GetValue<int>());
        Assert.Equal(1, report["data"]!["existingNativeEquations"]!.GetValue<int>());
        Assert.Equal(1, report["data"]!["nonEquationObjects"]!.GetValue<int>());
        Assert.Equal(0, report["data"]!["converted"]!.GetValue<int>());
        Assert.False(report["data"]!["fullyNative"]!.GetValue<bool>());
        Assert.All(report["data"]!["results"]!.AsArray(), r => Assert.NotEmpty(r!["latex"]!.GetValue<string>()));
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Single(Directory.GetFiles(dir.Path));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XmlContentTypeFindsEquationsWithoutAnXmlFileExtension(bool mixedCaseOverride)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, suite: true);
        using (var package = System.IO.Packaging.Package.Open(source, FileMode.Open, FileAccess.ReadWrite))
        {
            var oldUri = new Uri("/word/header1.xml", UriKind.Relative);
            var newUri = new Uri("/word/header.native", UriKind.Relative);
            var old = package.GetPart(oldUri);
            var renamed = package.CreatePart(newUri, old.ContentType);
            using (var input = old.GetStream(FileMode.Open, FileAccess.Read))
            using (var target = renamed.GetStream(FileMode.Create, FileAccess.Write)) input.CopyTo(target);
            foreach (var relationship in old.GetRelationships())
                renamed.CreateRelationship(relationship.TargetUri, relationship.TargetMode, relationship.RelationshipType, relationship.Id);
            var main = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            var headerRef = Assert.Single(main.GetRelationships(), r => r.RelationshipType.EndsWith("/header", StringComparison.Ordinal));
            main.DeleteRelationship(headerRef.Id);
            main.CreateRelationship(new Uri("header.native", UriKind.Relative), headerRef.TargetMode, headerRef.RelationshipType, headerRef.Id);
            package.DeletePart(oldUri);
        }
        if (mixedCaseOverride)
        {
            ChangeXmlPart(source, "[Content_Types].xml", xml => xml.Root!.Add(
                new XElement(xml.Root.Name.Namespace + "Override", new XAttribute("PartName", "/WORD/HEADER.NATIVE"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"))));
        }
        using (var original = WordprocessingDocument.Open(source, false))
            Assert.Empty(new OpenXmlValidator().Validate(original));
        var report = MathTypeConverter.Convert(source, output);
        Assert.True(report["success"]!.GetValue<bool>());
        Assert.Equal(6, report["data"]!["converted"]!.GetValue<int>());
        Assert.Equal("h", Assert.Single(Xml(output, "word/header.native").Descendants(M + "oMath")).Value);
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
    }

    [Theory]
    [InlineData("xml", false)]
    [InlineData("xml", true)]
    [InlineData("native", false)]
    [InlineData("native", true)]
    public async Task CliMainPartCasingFollowsOpcAndPreservesEntryNames(string extension, bool uppercaseEntry)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, suite: true);
        string mainEntry = ChangeMainPartCasing(source, extension, uppercaseEntry);
        if (Environment.GetEnvironmentVariable("OFFICECLI_MATHTYPE_ARTIFACTS") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            File.Copy(source, Path.Combine(artifactDirectory, $"opc-{extension}-{uppercaseEntry}-synthetic.docx"), overwrite: false);
        }
        using (var doc = WordprocessingDocument.Open(source, false))
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        byte[] original = File.ReadAllBytes(source);
        var before = Entries(source);

        var inspect = await Cli("convert-equations", source, "--dry-run", "--json");
        Assert.True(inspect.Exit == 0, inspect.Out + inspect.Error);
        var inspected = JsonNode.Parse(inspect.Out)!;
        Assert.True(inspected["success"]!.GetValue<bool>());
        Assert.Equal(6, inspected["data"]!["convertible"]!.GetValue<int>());
        Assert.Equal(0, inspected["data"]!["converted"]!.GetValue<int>());
        Assert.Equal(1, inspected["data"]!["existingNativeEquations"]!.GetValue<int>());
        Assert.Contains(inspected["data"]!["results"]!.AsArray(), r => r!["part"]!.GetValue<string>() == "/" + mainEntry);
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.False(File.Exists(output));

        var export = await Cli("convert-equations", source, "--out", output, "--json");
        Assert.True(export.Exit == 0, export.Out + export.Error);
        var exported = JsonNode.Parse(export.Out)!;
        Assert.True(exported["success"]!.GetValue<bool>());
        Assert.True(exported["data"]!["fullyNative"]!.GetValue<bool>());
        Assert.Equal(6, exported["data"]!["converted"]!.GetValue<int>());
        Assert.Equal(4, Xml(output, mainEntry).Descendants(M + "oMath").Count());
        Assert.Equal(original, File.ReadAllBytes(source));

        var after = Entries(output);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        string[] changed = [mainEntry, "word/header1.xml", "word/footer1.xml", "word/footnotes.xml"];
        foreach (string name in before.Keys.Except(changed)) Assert.Equal(before[name], after[name]);
        var validation = await Cli("validate", output, "--json");
        Assert.True(validation.Exit == 0, validation.Out + validation.Error);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("native")]
    public void MainPartCasingCannotBypassDocumentRootValidation(string extension)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("broken.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        ChangeMainXml(source, xml => xml.Root!.ReplaceWith(new XElement("notAWordDocument")));
        ChangeMainPartCasing(source, extension, uppercaseEntry: false);
        byte[] original = File.ReadAllBytes(source);
        foreach (string? destination in new[] { null, output })
        {
            var exception = Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, destination, true));
            Assert.Equal("corrupt_file", exception.Code);
            Assert.Equal("The main part must contain a Word document and one body.", exception.Message);
            Assert.False(File.Exists(output));
            Assert.Equal(original, File.ReadAllBytes(source));
        }
    }

    [Fact]
    public void CaseOnlyDuplicatePartsRemainAmbiguousAndRejected()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("ambiguous.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        byte[] main = Entries(source)["word/document.xml"];
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
        using (var stream = archive.CreateEntry("WORD/DOCUMENT.XML").Open()) stream.Write(main);
        byte[] original = File.ReadAllBytes(source);
        var exception = Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, output, true));
        Assert.Equal("corrupt_file", exception.Code);
        Assert.Equal("Package has duplicate entry names.", exception.Message);
        Assert.False(File.Exists(output));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData("equation.dsmt4", true)]
    [InlineData("equation.axmath", false)]
    public void ProducerCasingCannotHideAnEmbeddedEquation(string producer, bool supported)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx");
        CreateDocument(source, progId: producer);
        var report = MathTypeConverter.Convert(source, null);
        Assert.Equal(1, report["data"]!["equations"]!.GetValue<int>());
        Assert.Equal(supported, report["success"]!.GetValue<bool>());
        Assert.Equal(supported ? 0 : 1, report["data"]!["unsupported"]!.GetValue<int>());
        Assert.False(report["data"]!["fullyNative"]!.GetValue<bool>());
    }

    [Fact]
    public void ExportPreservesMixedTextTablesStoriesAndUntouchedPackageBytes()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, suite: true);
        var before = Entries(source);
        byte[] hash = SHA256.HashData(File.ReadAllBytes(source));
        var report = MathTypeConverter.Convert(source, output);
        Assert.True(report["success"]!.GetValue<bool>());
        Assert.True(report["data"]!["fullyNative"]!.GetValue<bool>());
        Assert.Equal(6, report["data"]!["converted"]!.GetValue<int>());
        Assert.Equal(hash, SHA256.HashData(File.ReadAllBytes(source)));
        var after = Entries(output);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        string[] changed = ["word/document.xml", "word/header1.xml", "word/footer1.xml", "word/footnotes.xml"];
        foreach (string part in before.Keys.Except(changed)) Assert.Equal(before[part], after[part]);
        var xml = Xml(output);
        var first = xml.Descendants(W + "p").First();
        Assert.Equal(["Before ", " after"], first.Descendants(W + "t").Select(t => t.Value));
        Assert.Equal(["r", "oMath", "r"], first.Elements().Select(e => e.Name.LocalName));
        Assert.All(first.Elements(W + "r"), r => Assert.NotNull(r.Element(W + "rPr")?.Element(W + "b")));
        Assert.Single(xml.Descendants(M + "oMathPara"));
        Assert.Single(xml.Descendants(W + "tbl").Single().Descendants(M + "oMath"));
        Assert.Equal("Table ", xml.Descendants(W + "tbl").Single().Descendants(W + "t").Single().Value);
        Assert.Equal(4, xml.Descendants(M + "oMath").Count());
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
    }

    [Fact]
    public void EditingAdjacentBodyTextAndSavingPreservesNormalizedOmml()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, suite: true);
        Assert.True(MathTypeConverter.Convert(source, output)["success"]!.GetValue<bool>());
        string[] before = NormalizedOmml(output);
        using (var handler = new WordHandler(output, editable: true))
        {
            handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["text"] = "Changed text " });
            handler.Save();
        }
        Assert.Equal(before, NormalizedOmml(output));
        Assert.Contains("Changed text ", Xml(output).Descendants(W + "t").Select(t => t.Value));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
    }

    [Fact]
    public void SharedVmlShapeDefinitionsSurviveConversion()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, suite: true);
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var definition = new XElement(v + "shapetype", new XAttribute("id", "sharedPreview"), new XAttribute("coordsize", "21600,21600"));
        ChangeMainXml(source, xml =>
        {
            xml.Descendants(W + "object").First().AddFirst(definition);
            foreach (var shape in xml.Descendants(v + "shape")) shape.SetAttributeValue("type", "#sharedPreview");
        });
        var report = MathTypeConverter.Convert(source, output);
        Assert.True(report["success"]!.GetValue<bool>());
        var after = Xml(output);
        var retained = Assert.Single(after.Descendants(v + "shapetype"));
        Assert.Equal(W + "pict", retained.Parent!.Name);
        Assert.Equal("sharedPreview", (string?)retained.Attribute("id"));
        Assert.Equal("#sharedPreview", (string?)Assert.Single(after.Descendants(W + "object")).Element(v + "shape")!.Attribute("type"));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
    }

    [Fact]
    public async Task CliRootEquationStreamDeterminesBothDryRunAndExportedMath()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("nested-storage.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        ReplaceEquationPayload(source, NestedEquations());
        byte[] original = File.ReadAllBytes(source);
        using (var doc = WordprocessingDocument.Open(source, false))
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        var inspect = await Cli("convert-equations", source, "--dry-run", "--json");
        Assert.True(inspect.Exit == 0, inspect.Out + inspect.Error);
        var result = JsonNode.Parse(inspect.Out)!;
        Assert.True(result["success"]!.GetValue<bool>());
        Assert.Equal("x", result["data"]!["results"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("x", XElement.Parse(result["data"]!["results"]![0]!["omml"]!.GetValue<string>()).Value);
        Assert.Single(Directory.GetFiles(dir.Path));

        var export = await Cli("convert-equations", source, "--out", output, "--json");
        Assert.True(export.Exit == 0, export.Out + export.Error);
        Assert.True(JsonNode.Parse(export.Out)!["data"]!["fullyNative"]!.GetValue<bool>());
        Assert.Equal("x", Assert.Single(Xml(output).Descendants(M + "oMath")).Value);
        Assert.Equal(original, File.ReadAllBytes(source));
        var before = Entries(source);
        var after = Entries(output);
        foreach (string part in before.Keys.Where(p => p != "word/document.xml"))
            Assert.Equal(before[part], after[part]);
        var validation = await Cli("validate", output, "--json");
        Assert.True(validation.Exit == 0, validation.Out + validation.Error);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("cycle")]
    public async Task CliInvalidRootEquationStorageFailsEvenWithPreservation(string invalid)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("invalid-storage.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        var payload = NestedEquations(rootName: invalid == "missing" ? "RootData" : "Equation Native");
        if (invalid == "duplicate") MathTypeStorageTests.SetPointer(payload, 1, 72, 2);
        if (invalid == "cycle") MathTypeStorageTests.SetPointer(payload, 3, 72, 3);
        ReplaceEquationPayload(source, payload);
        byte[] original = File.ReadAllBytes(source);
        foreach (bool preserve in new[] { false, true })
        foreach (bool dryRun in new[] { false, true })
        {
            string[] destination = dryRun ? ["--dry-run"] : ["--out", output];
            string[] preservation = preserve ? ["--preserve-unsupported"] : [];
            var run = await Cli(["convert-equations", source, .. destination, .. preservation, "--json"]);
            Assert.True(run.Exit == 1, run.Out + run.Error);
            var report = JsonNode.Parse(run.Out)!;
            Assert.False(report["success"]!.GetValue<bool>());
            Assert.Equal(1, report["data"]!["invalid"]!.GetValue<int>());
            Assert.Equal(0, report["data"]!["unsupported"]!.GetValue<int>());
            Assert.Equal("invalid_equation_ole", Assert.Single(report["errors"]!.AsArray())!["code"]!.GetValue<string>());
            Assert.False(File.Exists(output));
            Assert.Equal(original, File.ReadAllBytes(source));
        }
    }

    [Theory]
    [InlineData("floating", "unsupported_floating_equation")]
    [InlineData("textbox", "unsupported_equation_container")]
    [InlineData("extraShape", "unsupported_equation_container")]
    [InlineData("wrongShapeId", "invalid_equation_container")]
    public void NonInlineOrAmbiguousPreviewContentIsNotSilentlyRemoved(string modification, string code)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        XNamespace v = "urn:schemas-microsoft-com:vml";
        ChangeMainXml(source, xml =>
        {
            var shape = xml.Descendants(v + "shape").Single();
            switch (modification)
            {
                case "floating": shape.SetAttributeValue("style", "position:absolute;width:80pt"); break;
                case "textbox": shape.Add(new XElement(v + "textbox", new XElement(W + "txbxContent", new XElement(W + "p")))); break;
                case "extraShape": shape.AddAfterSelf(new XElement(v + "shape", new XAttribute("id", "other"))); break;
                case "wrongShapeId": shape.SetAttributeValue("id", "different"); break;
            }
        });
        var report = MathTypeConverter.Convert(source, output);
        Assert.False(report["success"]!.GetValue<bool>());
        Assert.Equal(code, report["errors"]![0]!["code"]!.GetValue<string>());
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void DisplayEquationBesideOtherRunContentStaysInline()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("native.docx");
        CreateDocument(source, Raw(false, Line(Fraction)));
        ChangeMainXml(source, xml =>
        {
            xml.Descendants(W + "t").Remove();
            xml.Descendants(W + "r").First().Add(new XElement(W + "br"));
        });
        Assert.True(MathTypeConverter.Convert(source, output)["success"]!.GetValue<bool>());
        var xml = Xml(output);
        Assert.Empty(xml.Descendants(M + "oMathPara"));
        Assert.Single(xml.Descendants(M + "oMath"));
        Assert.Single(xml.Descendants(W + "br"));
    }

    [Fact]
    public void UnsupportedEquationsAbortUnlessExplicitlyPreserved()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("unknown.docx"), output = dir.File("native.docx");
        CreateDocument(source, progId: "Equation.AxMath", suite: true);
        byte[] original = File.ReadAllBytes(source);
        var rejected = MathTypeConverter.Convert(source, output);
        Assert.False(rejected["success"]!.GetValue<bool>());
        Assert.False(File.Exists(output));
        Assert.NotEmpty(rejected["errors"]!.AsArray());
        var preserved = MathTypeConverter.Convert(source, output, preserveUnsupported: true);
        Assert.True(preserved["success"]!.GetValue<bool>());
        Assert.False(preserved["data"]!["fullyNative"]!.GetValue<bool>());
        Assert.Equal(0, preserved["data"]!["converted"]!.GetValue<int>());
        Assert.All(preserved["data"]!["results"]!.AsArray(), r => Assert.Equal("preserved", r!["status"]!.GetValue<string>()));
        Assert.Equal(original, File.ReadAllBytes(source));
        var before = Entries(source);
        foreach (var (part, bytes) in Entries(output)) Assert.Equal(before[part], bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MalformedMtefFailsEvenWhenUnsupportedEquationsMayBePreserved(bool preserve)
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("malformed.docx"), output = dir.File("native.docx");
        CreateDocument(source, Equation(Fraction)[..^3]);
        byte[] original = File.ReadAllBytes(source);
        var report = MathTypeConverter.Convert(source, output, preserve);
        Assert.False(report["success"]!.GetValue<bool>());
        Assert.Equal(1, report["data"]!["invalid"]!.GetValue<int>());
        Assert.False(File.Exists(output));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    [Fact]
    public void MissingRelationshipTargetsAreNotRepaired()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("broken.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
            archive.Entries.First(e => e.FullName.StartsWith("word/embeddings/", StringComparison.Ordinal)).Delete();
        byte[] before = File.ReadAllBytes(source);
        Assert.Equal("corrupt_file", Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, output, true)).Code);
        Assert.False(File.Exists(output));
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Fact]
    public void MalformedXmlIsRejectedWithoutOutput()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("broken.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            archive.GetEntry("word/document.xml")!.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open());
            writer.Write("<broken>");
        }
        Assert.Throws<XmlException>(() => MathTypeConverter.Convert(source, output, true));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void WellFormedXmlWithoutAWordBodyIsNotADocx()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("broken.docx"), output = dir.File("native.docx");
        CreateDocument(source);
        ChangeMainXml(source, xml => xml.Root!.ReplaceWith(new XElement("notAWordDocument")));
        Assert.Equal("corrupt_file", Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, output, true)).Code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void InputAndExistingOutputCannotBeOverwritten()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx"), output = dir.File("existing.docx");
        CreateDocument(source);
        File.WriteAllText(output, "sentinel");
        byte[] before = File.ReadAllBytes(source);
        Assert.Equal("invalid_value", Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, source)).Code);
        Assert.Equal("file_exists", Assert.Throws<CliException>(() => MathTypeConverter.Convert(source, output)).Code);
        Assert.Equal("sentinel", File.ReadAllText(output));
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Fact]
    public async Task CliDryRunExportAndHtmlPreviewAreUsable()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source with spaces.docx"), output = dir.File("native.docx"), html = dir.File("preview.html");
        CreateDocument(source, suite: true);
        var inspect = await Cli("convert-equations", source, "--dry-run", "--json");
        Assert.Equal(0, inspect.Exit);
        Assert.True(JsonNode.Parse(inspect.Out)!["data"]!["dryRun"]!.GetValue<bool>());
        var export = await Cli("convert-equations", source, "--out", output, "--json");
        Assert.Equal(0, export.Exit);
        var json = JsonNode.Parse(export.Out)!;
        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(6, json["data"]!["converted"]!.GetValue<int>());
        Assert.NotEmpty(json["warnings"]!.AsArray());
        var validation = await Cli("validate", output, "--json");
        Assert.True(validation.Exit == 0, validation.Out + validation.Error);
        var preview = await Cli("view", output, "html", "-o", html);
        Assert.True(preview.Exit == 0, preview.Out + preview.Error);
        string rendered = File.ReadAllText(html);
        Assert.Contains("data-formula=", rendered);
        Assert.DoesNotContain("data-formula=\"\"", rendered);
        var read = await Cli("view", output, "text");
        Assert.Equal(0, read.Exit);
        Assert.Contains("Before", read.Out);
        Assert.Contains("x", read.Out);
        Assert.Contains("Table", read.Out);

        // Optional reviewer artifacts are generated fixtures, never user files.
        if (Environment.GetEnvironmentVariable("OFFICECLI_MATHTYPE_ARTIFACTS") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            File.Copy(source, Path.Combine(artifactDirectory, "mathtype-synthetic.docx"), overwrite: false);
            File.Copy(output, Path.Combine(artifactDirectory, "native-synthetic.docx"), overwrite: false);
            File.Copy(html, Path.Combine(artifactDirectory, "native-preview.html"), overwrite: false);
            File.WriteAllText(Path.Combine(artifactDirectory, "conversion.json"), export.Out);
        }
    }

    [Fact]
    public async Task CliSummationLimitsSurviveExportPreviewAndAdjacentTextEdit()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("summation-source.docx"), output = dir.File("summation-native.docx"), html = dir.File("summation-preview.html");
        CreateDocument(source, Equation(NestedSums));
        byte[] original = File.ReadAllBytes(source);
        var inspect = await Cli("convert-equations", source, "--dry-run", "--json");
        Assert.Equal(0, inspect.Exit);
        var inspected = JsonNode.Parse(inspect.Out)!;
        Assert.True(inspected["success"]!.GetValue<bool>());
        var math = XElement.Parse(inspected["data"]!["results"]![0]!["omml"]!.GetValue<string>());
        Assert.Equal(["a=2", "b=3"], math.Descendants(M + "nary").Select(n => n.Element(M + "sub")!.Value));
        Assert.Equal(["A", "B"], math.Descendants(M + "nary").Select(n => n.Element(M + "sup")!.Value));
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.Single(Directory.GetFiles(dir.Path));

        var export = await Cli("convert-equations", source, "--out", output, "--json");
        Assert.Equal(0, export.Exit);
        Assert.True(JsonNode.Parse(export.Out)!["data"]!["fullyNative"]!.GetValue<bool>());
        var sums = Xml(output).Descendants(M + "nary").ToArray();
        Assert.Equal(["a=2", "b=3"], sums.Select(n => n.Element(M + "sub")!.Value));
        Assert.Equal(["A", "B"], sums.Select(n => n.Element(M + "sup")!.Value));
        string[] before = NormalizedOmml(output);
        using (var handler = new WordHandler(output, editable: true))
        {
            handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["text"] = "Changed text " });
            handler.Save();
        }
        Assert.Equal(before, NormalizedOmml(output));
        Assert.Equal(original, File.ReadAllBytes(source));
        var validation = await Cli("validate", output, "--json");
        Assert.True(validation.Exit == 0, validation.Out + validation.Error);
        var preview = await Cli("view", output, "html", "-o", html);
        Assert.True(preview.Exit == 0, preview.Out + preview.Error);
        string rendered = File.ReadAllText(html);
        Assert.Contains("data-formula=", rendered);
        Assert.DoesNotContain("data-formula=\"\"", rendered);
        Assert.Contains("Changed text", rendered);

        if (Environment.GetEnvironmentVariable("OFFICECLI_MATHTYPE_ARTIFACTS") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            File.Copy(source, Path.Combine(artifactDirectory, "summation-source-synthetic.docx"), overwrite: false);
            File.Copy(output, Path.Combine(artifactDirectory, "summation-native-synthetic.docx"), overwrite: false);
            File.Copy(html, Path.Combine(artifactDirectory, "summation-native-preview.html"), overwrite: false);
        }
    }

    [Fact]
    public async Task CliFailuresHaveJsonAndNonzeroExitCodes()
    {
        using var dir = new MathTypeTestDirectory();
        string source = dir.File("source.docx");
        CreateDocument(source, Equation(Fraction)[..^1]);
        var broken = await Cli("convert-equations", source, "--out", dir.File("native.docx"), "--preserve-unsupported", "--json");
        Assert.Equal(1, broken.Exit);
        Assert.False(JsonNode.Parse(broken.Out)!["success"]!.GetValue<bool>());
        Assert.False(File.Exists(dir.File("native.docx")));
        foreach (string[] options in new[] { new[] { "--json" }, new[] { "--dry-run", "--out", dir.File("unused.docx"), "--json" } })
        {
            var result = await Cli(["convert-equations", source, .. options]);
            Assert.Equal(1, result.Exit);
            Assert.Equal("invalid_argument", JsonNode.Parse(result.Out)!["error"]!["code"]!.GetValue<string>());
        }
    }

    private static void ReplaceEquationPayload(string path, byte[] payload)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        using var data = new MemoryStream(payload);
        doc.MainDocumentPart!.EmbeddedObjectParts.Single().FeedData(data);
    }

    private static string ChangeMainPartCasing(string path, string extension, bool uppercaseEntry)
    {
        string partName = "word/document." + extension;
        string entryName = uppercaseEntry ? partName.ToUpperInvariant() : partName;
        string target = "/" + (uppercaseEntry ? partName : partName.ToUpperInvariant());
        ChangeXmlPart(path, "_rels/.rels", xml => xml.Root!.Elements()
            .Single(e => ((string?)e.Attribute("Type"))?.EndsWith("/officeDocument", StringComparison.Ordinal) == true)
            .SetAttributeValue("Target", target));
        if (extension != "xml")
        {
            // A Default keeps this fixture readable by the bundled Packaging
            // version, independently of its separate content-type Override bug.
            ChangeXmlPart(path, "[Content_Types].xml", xml => xml.Root!.Add(
                new XElement(xml.Root.Name.Namespace + "Default", new XAttribute("Extension", extension),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))));
        }
        RenameEntry(path, "word/document.xml", entryName);
        RenameEntry(path, "word/_rels/document.xml.rels", "word/_rels/document." + extension + ".rels");
        return entryName;
    }

    private static void RenameEntry(string path, string previousName, string name)
    {
        if (previousName == name) return;
        byte[] bytes = Entries(path)[previousName];
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(previousName)!.Delete();
        using var stream = archive.CreateEntry(name).Open();
        stream.Write(bytes);
    }

    private static void ChangeMainXml(string path, Action<XDocument> change) =>
        ChangeXmlPart(path, "word/document.xml", change);

    private static void ChangeXmlPart(string path, string partName, Action<XDocument> change)
    {
        var xml = Xml(path, partName);
        change(xml);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(partName)!.Delete();
        using var stream = archive.CreateEntry(partName).Open();
        xml.Save(stream, SaveOptions.DisableFormatting);
    }

    private static string[] NormalizedOmml(string path)
    {
        XElement Normalize(XElement element) => new(element.Name,
            element.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString()),
            element.Nodes().Select(n => n is XElement e ? (object)Normalize(e)
                : n is XText t && (element.Name == M + "t" || !string.IsNullOrWhiteSpace(t.Value)) ? new XText(t.Value) : null));
        return Entries(path).Where(p => p.Key.StartsWith("word/") && p.Key.EndsWith(".xml")).OrderBy(p => p.Key)
            .SelectMany(p =>
            {
                using var stream = new MemoryStream(p.Value);
                return XDocument.Load(stream).Descendants(M + "oMath").ToArray();
            })
            .Select(m => Normalize(m).ToString(SaveOptions.DisableFormatting)).ToArray();
    }

    private static async Task<(int Exit, string Out, string Error)> Cli(params string[] arguments)
    {
        string? binary = Environment.GetEnvironmentVariable("OFFICECLI_TEST_BINARY");
        var start = new ProcessStartInfo(binary ?? Environment.GetEnvironmentVariable("OFFICECLI_TEST_DOTNET") ?? "dotnet")
        {
            RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true,
        };
        if (binary == null) start.ArgumentList.Add(typeof(MathTypeReader).Assembly.Location);
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        start.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        using var process = Process.Start(start)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        try { await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(45)); }
        catch (TimeoutException) { process.Kill(entireProcessTree: true); throw; }
        return (process.ExitCode, await stdout, await stderr);
    }
}
