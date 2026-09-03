// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Tests;

// Entirely synthetic equations and document text; no user content is embedded here.
internal static class MathTypeTestData
{
    internal static byte[] Join(params byte[][] items) => items.SelectMany(i => i).ToArray();
    internal static byte[] Line(params byte[][] items) => Join([1, 0], Join(items), [0]);
    internal static byte[] NullLine => [1, 1];
    internal static byte[] Character(char value, int typeface = 3) => [2, 0, (byte)(typeface + 128), (byte)value, (byte)(value >> 8)];
    internal static byte[] Text(string value) => Join(value.Select(c => Character(c)).ToArray());
    internal static byte[] Template(int selector, int variation, params byte[][] items) => Join(
        [3, 0, (byte)selector], variation >= 128 ? [(byte)((variation & 127) | 128), (byte)(variation >> 8)] : [(byte)variation],
        [0], Join(items), [0]);
    internal static byte[] Raw(bool inline, params byte[][] records) => Join(
        [5, 1, 0, 7, 0], Encoding.ASCII.GetBytes("OfficeCLI-test\0"), [(byte)(inline ? 1 : 0)], Join(records), [0]);
    internal static byte[] Equation(params byte[][] items) => Raw(true, Line(items));
    internal static byte[] Power => Join(Character('x'), Template(28, 0, NullLine, Line(Character('2', 8))));
    internal static byte[] Fraction => Template(11, 0, Line(Power), Line(Text("y+1")));

    internal static byte[] Ole(byte[] mtef)
    {
        var native = new byte[28 + mtef.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(native, 28);
        BinaryPrimitives.WriteUInt32LittleEndian(native.AsSpan(2), 0x00020000);
        BinaryPrimitives.WriteUInt32LittleEndian(native.AsSpan(8), (uint)mtef.Length);
        mtef.CopyTo(native, 28);
        return CompoundFile.WriteSingleStream("Equation Native", native);
    }

    internal static EmbeddedObject AddObject(OpenXmlPart part, byte[] data, string progId, int index)
    {
        var embedded = part.AddNewPart<EmbeddedObjectPart>("application/vnd.openxmlformats-officedocument.oleObject", "ole" + index);
        using (var stream = new MemoryStream(data)) embedded.FeedData(stream);
        return new EmbeddedObject($"""
            <w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <v:shape id="shape{index}" style="width:80pt;height:24pt" o:ole=""/>
              <o:OLEObject Type="Embed" ProgID="{progId}" ShapeID="shape{index}" DrawAspect="Content" ObjectID="_{1000 + index}" r:id="ole{index}"/>
            </w:object>
            """);
    }

    internal static void CreateDocument(string path, byte[]? equation = null, string progId = "Equation.DSMT4", bool suite = false)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var body = new Body();
        main.Document = new Document(body);
        var ole = Ole(equation ?? Equation(Fraction));
        body.AddChild(new Paragraph(new Run(new RunProperties(new Bold()), new Text("Before ") { Space = SpaceProcessingModeValues.Preserve },
            AddObject(main, ole, progId, 1), new Text(" after") { Space = SpaceProcessingModeValues.Preserve })));
        if (!suite) return;
        body.Append(new Paragraph(new Run(AddObject(main, Ole(Raw(false, Line(Fraction))), progId, 2))));
        body.Append(new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "4800" }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve },
            AddObject(main, Ole(Equation(Power)), progId, 3)))))));
        body.Append(new Paragraph(new M.OfficeMath(new M.Run(new M.Text("z")))));
        body.Append(new Paragraph(new Run(AddObject(main, [1, 2, 3, 4], "Visio.Drawing.15", 4))));
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(AddObject(header, Ole(Equation(Character('h'))), progId, 5))));
        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Paragraph(new Run(AddObject(footer, Ole(Equation(Character('f'))), progId, 6))));
        var footnotes = main.AddNewPart<FootnotesPart>();
        footnotes.Footnotes = new Footnotes(new Footnote(new Paragraph(new Run(AddObject(footnotes, Ole(Equation(Character('n'))), progId, 7)))) { Id = 1 });
        body.Append(new Paragraph(new Run(new FootnoteReference { Id = 1 })));
        body.Append(new SectionProperties(new HeaderReference { Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default },
            new FooterReference { Id = main.GetIdOfPart(footer), Type = HeaderFooterValues.Default }));
    }

    internal static Dictionary<string, byte[]> Entries(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return archive.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var stream = e.Open();
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            return bytes.ToArray();
        });
    }

    internal static XDocument Xml(string path, string part = "word/document.xml")
    {
        using var stream = new MemoryStream(Entries(path)[part]);
        return XDocument.Load(stream);
    }
}

internal sealed class MathTypeTestDirectory : IDisposable
{
    internal string Path { get; } = Directory.CreateTempSubdirectory("officecli-mathtype-").FullName;
    internal string File(string name) => System.IO.Path.Combine(Path, name);
    public void Dispose() => Directory.Delete(Path, recursive: true);
}
