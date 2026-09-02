// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace OfficeCli.Tests;

internal sealed record FormulaStats(
    int OMath,
    int OMathPara,
    int InlineOMath,
    int TableOMath,
    int MixedTextMathParagraphs);

internal static class DocxFixtureFactory
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string CreateCompatibilityFixture(string directory) =>
        CreateFixture(directory, "office-compatible.docx", false, false);

    public static string CreateDanglingRelationshipFixture(string directory) =>
        CreateFixture(directory, "dangling-relationship.docx", true, false);

    public static string CreateIllegalFormulaFixture(string directory) =>
        CreateFixture(directory, "illegal-formula.docx", false, true);

    public static IReadOnlyList<string> ReadNormalizedOmml(string path)
    {
        var document = ReadPart(path, "word/document.xml");
        return document
            .Descendants(M + "oMath")
            .Select(NormalizeElement)
            .Select(element => element.ToString(SaveOptions.DisableFormatting))
            .ToList();
    }

    public static FormulaStats ReadFormulaStats(string path)
    {
        var document = ReadPart(path, "word/document.xml");
        var formulas = document.Descendants(M + "oMath").ToList();
        var paragraphs = document.Descendants(W + "p").ToList();
        return new FormulaStats(
            formulas.Count,
            document.Descendants(M + "oMathPara").Count(),
            formulas.Count(formula => !formula.Ancestors(M + "oMathPara").Any()),
            document.Descendants(W + "tbl").SelectMany(table => table.Descendants(M + "oMath")).Count(),
            paragraphs.Count(paragraph =>
                paragraph.Descendants(W + "t").Any()
                && paragraph.Descendants(M + "oMath").Any()));
    }

    private static string CreateFixture(
        string directory,
        string fileName,
        bool danglingHeader,
        bool illegalFormula)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, fileName);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteEntry(archive, "[Content_Types].xml", ContentTypesXml);
        WriteEntry(archive, "_rels/.rels", RootRelationshipsXml);
        WriteEntry(archive, "word/document.xml", DocumentXml(danglingHeader, illegalFormula));
        WriteEntry(archive, "word/_rels/document.xml.rels", DocumentRelationshipsXml);
        WriteEntry(archive, "word/styles.xml", StylesXml);
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string name, string contents)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.SmallestSize);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(contents);
    }

    private static XDocument ReadPart(string path, string partName)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(partName)
            ?? throw new InvalidDataException($"Missing fixture part {partName}.");
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
    }

    private static XElement NormalizeElement(XElement source)
    {
        var normalized = new XElement(source.Name);
        foreach (var attribute in source.Attributes()
                     .Where(attribute => !attribute.IsNamespaceDeclaration)
                     .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                     .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal))
        {
            normalized.Add(new XAttribute(attribute.Name, attribute.Value));
        }
        foreach (var node in source.Nodes())
        {
            switch (node)
            {
                case XElement child:
                    normalized.Add(NormalizeElement(child));
                    break;
                case XText text when !string.IsNullOrWhiteSpace(text.Value):
                    normalized.Add(new XText(text.Value));
                    break;
            }
        }
        return normalized;
    }

    private static string DocumentXml(bool danglingHeader, bool illegalFormula)
    {
        var headerReference = danglingHeader
            ? "<w:headerReference w:type=\"default\" r:id=\"rIdMissing\"/>"
            : "";
        var invalidMathChild = illegalFormula
            ? "<m:limLoc m:val=\"undOvr\"/>"
            : "";
        return $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:body>
                <w:p>
                  <w:r><w:t>Before</w:t></w:r>
                  <m:oMath>
                    <m:r>
                      <m:rPr><m:sty m:val="p"/><m:scr m:val="roman"/></m:rPr>
                      <m:t>x</m:t>
                    </m:r>
                    {{invalidMathChild}}
                  </m:oMath>
                  <w:r><w:t xml:space="preserve"> after</w:t></w:r>
                </w:p>
                <w:p>
                  <m:oMathPara>
                    <m:oMath>
                      <m:nary>
                        <m:naryPr>
                          <m:chr m:val="∑"/>
                          <m:grow m:val="1"/>
                          <m:limLoc m:val="undOvr"/>
                        </m:naryPr>
                        <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
                        <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
                        <m:e><m:r><m:t>x</m:t></m:r></m:e>
                      </m:nary>
                    </m:oMath>
                  </m:oMathPara>
                </w:p>
                <w:tbl>
                  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
                  <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                      <w:p>
                        <w:r><w:t>Cell</w:t></w:r>
                        <m:oMath>
                          <m:m>
                            <m:mPr>
                              <m:mcs>
                                <m:mc>
                                  <m:mcPr>
                                    <m:count m:val="1"/>
                                    <m:mcJc m:val="center"/>
                                  </m:mcPr>
                                </m:mc>
                              </m:mcs>
                              <m:plcHide m:val="1"/>
                            </m:mPr>
                            <m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e></m:mr>
                          </m:m>
                        </m:oMath>
                      </w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
                <w:sectPr>
                  {{headerReference}}
                  <w:pgSz w:w="12240" w:h="15840"/>
                  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                           w:header="720" w:footer="720" w:gutter="0"/>
                </w:sectPr>
              </w:body>
            </w:document>
            """;
    }

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>
        """;

    private const string RootRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                        Target="word/document.xml"/>
        </Relationships>
        """;

    private const string DocumentRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdStyles"
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                        Target="styles.xml"/>
        </Relationships>
        """;

    private const string StylesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
            <w:name w:val="Normal"/>
            <w:qFormat/>
            <w:uiPriority w:val="1"/>
          </w:style>
        </w:styles>
        """;
}
