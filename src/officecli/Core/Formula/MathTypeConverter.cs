// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.IO.Packaging;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;

namespace OfficeCli.Core;

/// <summary>
/// Converts embedded equations in a new DOCX. Input is held read-only throughout;
/// untouched package entries, including original embedded payloads, retain their bytes.
/// No OLE code is executed and external relationships are never downloaded.
/// </summary>
internal static class MathTypeConverter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private const string OleRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    private const long MaxXmlBytes = 32 * 1024 * 1024;

    internal static JsonObject Convert(string input, string? output, bool preserveUnsupported = false)
    {
        input = Path.GetFullPath(input);
        if (!File.Exists(input)) throw new CliException($"File not found: {input}") { Code = "file_not_found" };
        if (!Path.GetExtension(input).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            throw new CliException("Equation conversion requires a .docx file.") { Code = "unsupported_type" };
        if (output != null)
        {
            output = Path.GetFullPath(output);
            if (!Path.GetExtension(output).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                throw new CliException("The output must be a .docx file.") { Code = "unsupported_type" };
            if (string.Equals(input, output, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                throw new CliException("The output must differ from the input; in-place conversion is not supported.") { Code = "invalid_value" };
            if (File.Exists(output))
                throw new CliException($"Output already exists: {output}") { Code = "file_exists" };
            if (!Directory.Exists(Path.GetDirectoryName(output)))
                throw new DirectoryNotFoundException("The output directory does not exist.");
        }

        using var inputStream = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);
        GuardArchive(archive);
        if (output != null && archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
            throw new CliException("Converting a digitally signed package requires removing/reapplying its signature explicitly.") { Code = "signed_document" };
        using var packageStream = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var package = Package.Open(packageStream, FileMode.Open, FileAccess.Read);
        string mainPart = CheckRelationships(package);
        var mainEntry = archive.Entries.FirstOrDefault(e => e.FullName.Equals(mainPart, StringComparison.OrdinalIgnoreCase));
        if (mainEntry == null)
            throw new CliException("The main document part is missing from the archive.") { Code = "corrupt_file" };
        var xmlPartNames = package.GetParts()
            .Where(p => p.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
                || p.ContentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase)
                || p.ContentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase))
            .Select(p => p.Uri.OriginalString.TrimStart('/')).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var records = new JsonArray();
        var warnings = new JsonArray();
        var errors = new JsonArray();
        var xmlParts = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        var conversions = new Dictionary<XElement, MathTypeEquation>();
        int existingNative = 0, otherObjects = 0, unsupported = 0, invalid = 0;
        foreach (var entry in archive.Entries.Where(e => xmlPartNames.Contains(e.FullName) || e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var xml = ReadXml(entry);
            if (ReferenceEquals(entry, mainEntry) && (xml.Root?.Name != W + "document" || xml.Root.Elements(W + "body").Count() != 1))
                throw new CliException("The main part must contain a Word document and one body.") { Code = "corrupt_file" };
            existingNative += xml.Descendants(M + "oMath").Count();
            var objects = xml.Descendants(O + "OLEObject").ToList();
            int index = 0;
            foreach (var ole in objects)
            {
                index++;
                string progId = (string?)ole.Attribute("ProgID") ?? "";
                if (!progId.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase)) { otherObjects++; continue; }
                string? relationshipId = (string?)ole.Attribute(R + "id");
                var record = new JsonObject
                {
                    ["part"] = "/" + entry.FullName, ["objectIndex"] = index,
                    ["progId"] = progId, ["relationshipId"] = relationshipId,
                };
                records.Add((JsonNode)record);
                try
                {
                    if (!progId.Equals("Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
                        && !progId.Equals("Equation.3", StringComparison.OrdinalIgnoreCase))
                        throw new MathTypeException("unsupported_equation_producer", $"Producer {progId} is not supported.");
                    var obj = ole.Parent;
                    if (obj?.Name != W + "object" || obj.Parent?.Name != W + "r" || obj.Parent.Parent?.Name != W + "p"
                        || obj.Elements(O + "OLEObject").Count() != 1)
                        throw new MathTypeException("unsupported_equation_container", "Only inline OLE objects in paragraph runs are supported (including tables, headers and footnotes).");
                    if (obj.Elements().Any(e => e.Name != V + "shape" && e.Name != V + "shapetype" && e != ole)
                        || obj.Elements(V + "shape").Count() > 1
                        || obj.Descendants().Any(e => e.Name.Namespace == W || e.Name.Namespace == M
                            || (e.Name == V + "textpath" && !string.IsNullOrEmpty((string?)e.Attribute("string")))))
                        throw new MathTypeException("unsupported_equation_container", "The equation object also contains non-preview content.");
                    if (obj.Element(V + "shape") is { } shape && (string?)shape.Attribute("id") != (string?)ole.Attribute("ShapeID"))
                        throw new MathTypeException("invalid_equation_container", "The equation preview does not match its OLE ShapeID.");
                    if (obj.Descendants().Attributes("style").Any(a => a.Value.Split(';').Any(s =>
                        s.Replace(" ", "", StringComparison.Ordinal).Equals("position:absolute", StringComparison.OrdinalIgnoreCase))))
                        throw new MathTypeException("unsupported_floating_equation", "Floating OLE equations cannot be moved into the text flow implicitly.");
                    if ((string?)ole.Attribute("Type") != "Embed")
                        throw new MathTypeException("unsupported_linked_equation", "Linked equations are not downloaded or converted.");
                    var partUri = PackUriHelper.CreatePartUri(new Uri("/" + entry.FullName, UriKind.Relative));
                    var part = package.GetPart(partUri);
                    if (string.IsNullOrEmpty(relationshipId) || !part.RelationshipExists(relationshipId))
                        throw new MathTypeException("invalid_equation_relationship", "The equation relationship is missing.");
                    var relationship = part.GetRelationship(relationshipId);
                    if (relationship.TargetMode != TargetMode.Internal || relationship.RelationshipType != OleRelationship)
                        throw new MathTypeException("invalid_equation_relationship", "Expected an internal OLE relationship.");
                    var payloadPart = package.GetPart(PackUriHelper.ResolvePartUri(partUri, relationship.TargetUri));
                    using var payload = payloadPart.GetStream(FileMode.Open, FileAccess.Read);
                    var bytes = ReadBounded(payload, MathTypeReader.MaxOleBytes);
                    var equation = MathTypeReader.ReadOle(bytes);
                    conversions.Add(obj, equation);
                    xmlParts[entry.FullName] = xml;
                    record["status"] = output == null ? "convertible" : "pending";
                    record["latex"] = FormulaParser.ToLatex(equation.Math);
                    record["text"] = FormulaParser.ToReadableText(equation.Math);
                    record["omml"] = equation.Math.OuterXml;
                }
                catch (MathTypeException ex)
                {
                    bool isUnsupported = ex.Code.StartsWith("unsupported_", StringComparison.Ordinal);
                    if (isUnsupported) unsupported++; else invalid++;
                    record["status"] = isUnsupported ? "unsupported" : "invalid";
                    record["code"] = ex.Code;
                    record["message"] = ex.Message;
                    record["byteOffset"] = ex.Offset;
                    var diagnostic = new JsonObject
                    {
                        ["part"] = "/" + entry.FullName, ["objectIndex"] = index,
                        ["code"] = ex.Code, ["message"] = ex.Message, ["byteOffset"] = ex.Offset,
                    };
                    if (isUnsupported && preserveUnsupported) warnings.Add((JsonNode)diagnostic);
                    else errors.Add((JsonNode)diagnostic);
                }
            }
        }

        bool success = errors.Count == 0;
        if (conversions.Count > 0)
            warnings.Insert(0, new JsonObject
            {
                ["code"] = "math_layout_normalized",
                ["message"] = "Converted equations use Word's native math font, sizing, spacing and alignment. MathType-specific typography and nudges are not preserved; mathematical structure, Unicode characters, bold/italic and RGB run colors are retained.",
            });
        int converted = 0;
        if (success && output != null)
        {
            foreach (var group in conversions.Keys.GroupBy(obj => obj.Parent!).ToList())
                ReplaceRun(group.Key, conversions);
            WriteOutput(archive, xmlParts, output);
            converted = conversions.Count;
            foreach (var record in records.OfType<JsonObject>())
            {
                if (record["status"]?.GetValue<string>() == "pending") record["status"] = "converted";
                else if (record["status"]?.GetValue<string>() == "unsupported" && preserveUnsupported) record["status"] = "preserved";
            }
        }
        else if (output != null)
        {
            foreach (var record in records.OfType<JsonObject>().Where(r => r["status"]?.GetValue<string>() == "pending"))
                record["status"] = "convertible";
        }
        return new JsonObject
        {
            ["success"] = success,
            ["data"] = new JsonObject
            {
                ["dryRun"] = output == null, ["output"] = success ? output : null,
                ["equations"] = records.Count, ["convertible"] = conversions.Count, ["converted"] = converted,
                ["unsupported"] = unsupported, ["invalid"] = invalid,
                ["existingNativeEquations"] = existingNative, ["nonEquationObjects"] = otherObjects,
                ["fullyNative"] = success && unsupported == 0 && invalid == 0 && (output != null || records.Count == 0),
                ["results"] = records,
            },
            ["warnings"] = warnings, ["errors"] = errors,
        };
    }

    private static void ReplaceRun(XElement run, Dictionary<XElement, MathTypeEquation> conversions)
    {
        var paragraph = run.Parent!;
        bool equationOnly = paragraph.Elements().All(e => e.Name == W + "pPr" || e.Name == W + "r")
            && paragraph.Elements(W + "r").SelectMany(r => r.Elements()).All(e => e.Name == W + "rPr"
                || e.Name == W + "object" || (e.Name == W + "t" && string.IsNullOrWhiteSpace(e.Value)))
            && paragraph.Descendants(W + "object").Count() == 1
            && !paragraph.Descendants(M + "oMath").Any();
        var replacement = new List<XElement>();
        XElement NewRun() => new(run.Name, run.Attributes(), run.Element(W + "rPr") is { } pr ? new XElement(pr) : null);
        var current = NewRun();
        foreach (var child in run.Elements().Where(e => e.Name != W + "rPr"))
        {
            if (conversions.TryGetValue(child, out var equation))
            {
                // A preview can define a VML shape type shared by other pictures
                // or OLE objects. Keep its definition at the original position.
                var shapeTypes = child.Elements(V + "shapetype").Select(e => new XElement(e)).ToList();
                if (shapeTypes.Count > 0) current.Add(new XElement(W + "pict", shapeTypes));
                if (current.Elements().Any(e => e.Name != W + "rPr")) replacement.Add(current);
                var math = XElement.Parse(equation.Math.OuterXml, LoadOptions.PreserveWhitespace);
                replacement.Add(!equation.Inline && equationOnly ? new XElement(M + "oMathPara", math) : math);
                current = NewRun();
            }
            else current.Add(new XElement(child));
        }
        if (current.Elements().Any(e => e.Name != W + "rPr")) replacement.Add(current);
        run.ReplaceWith(replacement);
    }

    private static void GuardArchive(ZipArchive archive)
    {
        if (archive.Entries.Count > DocumentLimits.MaxZipEntries)
            throw new CliException("Package has too many entries.") { Code = "decompression_bomb" };
        long uncompressed = 0, compressed = 0;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries)
        {
            if (!names.Add(entry.FullName)) throw new CliException("Package has duplicate entry names.") { Code = "corrupt_file" };
            uncompressed = checked(uncompressed + entry.Length);
            compressed = checked(compressed + entry.CompressedLength);
        }
        if (uncompressed > DocumentLimits.MaxUncompressedBytes || uncompressed / Math.Max(1, compressed) > DocumentLimits.MaxCompressionRatio)
            throw new CliException("Package exceeds the decompression limits.") { Code = "decompression_bomb" };
        if (archive.GetEntry("[Content_Types].xml") == null)
            throw new CliException("Not an OPC document package.") { Code = "corrupt_file" };
    }

    private static string CheckRelationships(Package package)
    {
        void Check(IEnumerable<PackageRelationship> relationships)
        {
            foreach (var rel in relationships.Where(r => r.TargetMode == TargetMode.Internal))
                if (!package.PartExists(PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri)))
                    throw new CliException($"Missing relationship target: {rel.SourceUri}, {rel.Id}") { Code = "corrupt_file" };
        }
        Check(package.GetRelationships());
        foreach (var part in package.GetParts())
            if (!PackUriHelper.IsRelationshipPartUri(part.Uri)) Check(part.GetRelationships());
        var main = package.GetRelationshipsByType(R.NamespaceName + "/officeDocument").SingleOrDefault();
        if (main == null || main.TargetMode != TargetMode.Internal
            || package.GetPart(PackUriHelper.ResolvePartUri(main.SourceUri, main.TargetUri)).ContentType
                != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
            throw new CliException("Expected a transitional, non-macro DOCX document.") { Code = "unsupported_type" };
        return PackUriHelper.ResolvePartUri(main.SourceUri, main.TargetUri).OriginalString.TrimStart('/');
    }

    private static XDocument ReadXml(ZipArchiveEntry entry)
    {
        if (entry.Length > MaxXmlBytes) throw new CliException($"XML part exceeds 32 MiB: {entry.FullName}") { Code = "equation_limit" };
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxXmlBytes };
        using (var scanStream = entry.Open())
        using (var scan = XmlReader.Create(scanStream, settings))
        {
            long elements = 0;
            while (scan.Read())
            {
                if (scan.Depth > DocumentLimits.MaxRecursionDepth || (scan.NodeType == XmlNodeType.Element && ++elements > DocumentLimits.MaxDomElements))
                    throw new CliException($"XML part exceeds structural limits: {entry.FullName}") { Code = "equation_limit" };
            }
        }
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadBounded(Stream stream, int limit)
    {
        using var result = new MemoryStream();
        var buffer = new byte[8192];
        for (int count = stream.Read(buffer); count > 0; count = stream.Read(buffer))
        {
            if (result.Length + count > limit) throw new MathTypeException("equation_limit", "OLE equation exceeds the 8 MiB limit.");
            result.Write(buffer, 0, count);
        }
        return result.ToArray();
    }

    private static void WriteOutput(ZipArchive source, Dictionary<string, XDocument> changed, string output)
    {
        string temporary = Path.Combine(Path.GetDirectoryName(output)!, $".officecli-equations-{Guid.NewGuid():N}.tmp");
        try
        {
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Create))
            {
                foreach (var entry in source.Entries)
                {
                    var target = archive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    target.LastWriteTime = entry.LastWriteTime;
                    target.ExternalAttributes = entry.ExternalAttributes;
                    using var stream = target.Open();
                    if (changed.TryGetValue(entry.FullName, out var xml)) xml.Save(stream, SaveOptions.DisableFormatting);
                    else { using var original = entry.Open(); original.CopyTo(stream); }
                }
            }
            File.Move(temporary, output, overwrite: false);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
