// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace OfficeCli.Tests;

public sealed class MarkupCompatibilityValidationTests
{
    private static readonly XNamespace M =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Mc =
        "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Extension = "urn:officecli-tests:extension";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReorderedParentsRetainInheritedIgnorableContent(bool extensionElement)
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);
        AddInheritedExtensions(fixture, extensionElement);

        // The control changes only the four known orders in a complete in-memory
        // document. All inherited declarations and extension content stay intact.
        Assert.Empty(ValidateCanonicalDocument(fixture));
        await AssertLayeredProfilesAsync(workspace, fixture);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReorderedParentsRetainScopedIgnorableNamespaces(bool extensionElement)
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);
        AddInheritedExtensions(fixture, extensionElement);
        XNamespace scoped = "urn:officecli-tests:scoped-extension";
        foreach (var partName in new[] { "word/document.xml", "word/styles.xml" })
        {
            EditPart(fixture, partName, root =>
            {
                foreach (var parent in CompatibilityParents(root))
                {
                    parent.SetAttributeValue(XNamespace.Xmlns + "outer", Extension.NamespaceName);
                    parent.SetAttributeValue(XNamespace.Xmlns + "vendor", scoped.NamespaceName);
                    parent.SetAttributeValue(Mc + "Ignorable", "vendor");
                    parent.SetAttributeValue(
                        Mc + (extensionElement ? "PreserveElements" : "PreserveAttributes"), "vendor:*");
                    if (extensionElement)
                        parent.Add(new XElement(scoped + "metadata"));
                    else
                        parent.SetAttributeValue(scoped + "metadata", "nested");
                }
            });
        }

        // The same prefix names different namespaces at two scopes; both
        // remain ignorable. Flattening the MC strings would lose the outer one.
        Assert.Empty(ValidateCanonicalDocument(fixture));
        await AssertLayeredProfilesAsync(workspace, fixture);
    }

    [Theory]
    [InlineData("scr", "duplicate-before")]
    [InlineData("scr", "duplicate-after")]
    [InlineData("limLoc", "duplicate-before")]
    [InlineData("limLoc", "duplicate-after")]
    [InlineData("plcHide", "duplicate-before")]
    [InlineData("plcHide", "duplicate-after")]
    [InlineData("uiPriority", "duplicate-before")]
    [InlineData("uiPriority", "duplicate-after")]
    [InlineData("scr", "illegal-child")]
    public async Task IgnorableExtensionsCannotHideInvalidChildren(string propertyName, string corruption)
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateInvalidCompatibilityFixture(
            workspace.DirectoryPath, propertyName, corruption);
        AddInheritedExtensions(fixture, extensionElement: true);

        Assert.NotEmpty(ValidateCanonicalDocument(fixture));
        await AssertBlockingProfilesAsync(workspace, fixture, compatibleWarnings: 3);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("illegal-child")]
    [InlineData("invalid-value")]
    public async Task InheritedProcessContentCannotHideInvalidFormulaProperties(string corruption)
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);
        AddInheritedExtensions(fixture, extensionElement: true);
        EditPart(fixture, "word/document.xml", root =>
        {
            root.SetAttributeValue(Mc + "ProcessContent", "vendor:metadata");
            var parent = root.Descendants(M + "rPr").Single();
            var child = corruption switch
            {
                "duplicate" => new XElement(parent.Element(M + "scr")!),
                "illegal-child" => new XElement(M + "limLoc", new XAttribute(M + "val", "undOvr")),
                "invalid-value" => new XElement(M + "aln", new XAttribute(M + "val", "not-a-boolean")),
                _ => throw new ArgumentException("Unknown fixture corruption.", nameof(corruption)),
            };
            parent.Element(Extension + "metadata")!.Add(child);
        });

        // The extension wrapper is ignorable, but its children must be checked
        // under the inherited ProcessContent rule after the known order is fixed.
        Assert.NotEmpty(ValidateCanonicalDocument(fixture));
        await AssertBlockingProfilesAsync(workspace, fixture, compatibleWarnings: 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NonIgnorableExtensionsRemainBlocking(bool extensionElement)
    {
        using var workspace = new TestWorkspace();
        var fixture = DocxFixtureFactory.CreateCompatibilityFixture(workspace.DirectoryPath);
        AddInheritedExtensions(fixture, extensionElement);
        foreach (var partName in new[] { "word/document.xml", "word/styles.xml" })
            EditPart(fixture, partName, root => root.Attribute(Mc + "Ignorable")!.Remove());

        Assert.NotEmpty(ValidateCanonicalDocument(fixture));
        await AssertBlockingProfilesAsync(workspace, fixture, compatibleWarnings: 0);
    }

    private static async Task AssertLayeredProfilesAsync(TestWorkspace workspace, string fixture)
    {
        var original = await File.ReadAllBytesAsync(fixture);
        foreach (var profile in new[] { "strict", "office-compatible" })
        {
            var result = await CliTestDriver.RunAsync(
                workspace.DirectoryPath, "validate", fixture, "--profile", profile, "--json");
            var compatible = profile == "office-compatible";
            Assert.Equal(compatible ? 0 : 1, result.ExitCode);
            using var json = JsonDocument.Parse(result.StdOut);
            Assert.Equal(compatible, json.RootElement.GetProperty("success").GetBoolean());
            var data = json.RootElement.GetProperty("data");
            Assert.Equal(compatible ? 0 : 4, data.GetProperty("errorCount").GetInt32());
            Assert.Equal(compatible ? 4 : 0, data.GetProperty("warningCount").GetInt32());
            Assert.Equal(4, data.GetProperty("diagnostics").GetArrayLength());
            if (compatible)
            {
                Assert.All(data.GetProperty("warnings").EnumerateArray(), diagnostic =>
                {
                    Assert.Equal("office-compatible-element-order", diagnostic.GetProperty("classification").GetString());
                    Assert.False(string.IsNullOrWhiteSpace(diagnostic.GetProperty("reason").GetString()));
                });
            }
            Assert.Equal(original, await File.ReadAllBytesAsync(fixture));
        }
    }

    private static async Task AssertBlockingProfilesAsync(
        TestWorkspace workspace, string fixture, int compatibleWarnings)
    {
        var original = await File.ReadAllBytesAsync(fixture);
        foreach (var profile in new[] { "strict", "office-compatible" })
        {
            var result = await CliTestDriver.RunAsync(
                workspace.DirectoryPath, "validate", fixture, "--profile", profile, "--json");
            Assert.Equal(1, result.ExitCode);
            using var json = JsonDocument.Parse(result.StdOut);
            Assert.False(json.RootElement.GetProperty("success").GetBoolean());
            var data = json.RootElement.GetProperty("data");
            Assert.True(data.GetProperty("errorCount").GetInt32() >= 1);
            Assert.Equal(
                profile == "strict" ? 0 : compatibleWarnings, data.GetProperty("warningCount").GetInt32());
            Assert.Contains(data.GetProperty("diagnostics").EnumerateArray(), diagnostic =>
                diagnostic.GetProperty("severity").GetString() == "error"
                && diagnostic.GetProperty("classification").GetString() == "schema-validation-error");
            Assert.Equal(original, await File.ReadAllBytesAsync(fixture));
        }
    }

    private static void AddInheritedExtensions(string path, bool extensionElement)
    {
        foreach (var partName in new[] { "word/document.xml", "word/styles.xml" })
        {
            EditPart(path, partName, root =>
            {
                root.SetAttributeValue(XNamespace.Xmlns + "mc", Mc.NamespaceName);
                root.SetAttributeValue(XNamespace.Xmlns + "vendor", Extension.NamespaceName);
                root.SetAttributeValue(Mc + "Ignorable", "vendor");
                foreach (var parent in CompatibilityParents(root))
                {
                    if (extensionElement)
                        parent.Add(new XElement(Extension + "metadata"));
                    else
                        parent.SetAttributeValue(Extension + "metadata", "synthetic");
                }
            });
        }
    }

    private static IEnumerable<XElement> CompatibilityParents(XElement root) =>
        root.Descendants().Where(element =>
            element.Name == M + "rPr" || element.Name == M + "naryPr"
            || element.Name == M + "mPr" || element.Name == W + "style");

    private static void EditPart(string path, string partName, Action<XElement> edit)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry(partName)!;
        XDocument document;
        using (var stream = entry.Open())
            document = XDocument.Load(stream);
        edit(document.Root!);
        entry.Delete();
        using var output = archive.CreateEntry(partName).Open();
        document.Save(output);
    }

    private static IReadOnlyList<string> ValidateCanonicalDocument(string path)
    {
        using var stream = new MemoryStream(File.ReadAllBytes(path));
        using var document = WordprocessingDocument.Open(stream, false);
        var roots = new OpenXmlElement[]
        {
            document.MainDocumentPart!.Document!,
            document.MainDocumentPart.StyleDefinitionsPart!.Styles!,
        };
        foreach (var root in roots)
        {
            foreach (var (ns, childName, previousName) in new[]
                     {
                         (M.NamespaceName, "scr", "sty"),
                         (M.NamespaceName, "limLoc", "grow"),
                         (M.NamespaceName, "plcHide", "mcs"),
                         (W.NamespaceName, "uiPriority", "qFormat"),
                     })
            {
                var child = root.Descendants().FirstOrDefault(element =>
                    element.NamespaceUri == ns && element.LocalName == childName
                    && element.PreviousSibling() is { } previous
                    && previous.NamespaceUri == ns && previous.LocalName == previousName);
                if (child == null) continue;
                var previous = child.PreviousSibling()!;
                child.Remove();
                previous.InsertBeforeSelf(child);
            }
        }
        return new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document).Select(error => error.Description).ToList();
    }
}
