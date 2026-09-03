// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Validation;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Core;

internal sealed record MathTypeEquation(M.OfficeMath Math, bool Inline);

internal sealed class MathTypeException(string code, string message, int offset = 0)
    : Exception(message)
{
    public string Code { get; } = code;
    public int Offset { get; } = offset;
}

/// <summary>
/// Reads MTEF v5 without activating OLE servers. Unknown mathematical records,
/// characters and template variants fail closed; typography uses Word's math layout.
/// Format reference: https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later
/// </summary>
internal sealed partial class MathTypeReader
{
    internal const int MaxOleBytes = 8 * 1024 * 1024;
    private const int MaxRecords = 100_000;
    private const int MaxDepth = 64;
    private readonly byte[] _data;
    private int _position;
    private int _records;
    private string? _color;
    private readonly List<string> _colors = [];
    private readonly List<(int Font, int Style)> _fontStyles = [];
    private readonly List<(int Encoding, string Name)> _fonts = [];
    private readonly Dictionary<int, int> _styles = new()
    {
        [1] = 0, [2] = 0, [3] = 2, [4] = 2, [5] = 0, [6] = 0,
        [7] = 1, [8] = 0, [9] = 0, [10] = 0, [11] = 0, [12] = 0,
    };
    private int _encodingCount = 4;

    private MathTypeReader(byte[] data) => _data = data;

    internal static MathTypeEquation ReadOle(byte[] ole)
    {
        if (ole.Length > MaxOleBytes)
            throw new MathTypeException("equation_limit", "OLE equation exceeds the 8 MiB limit.");
        // Bound allocations in the shared CFB reader before following untrusted FATs.
        if (ole.Length < 512 || BinaryPrimitives.ReadUInt16LittleEndian(ole.AsSpan(28)) != 0xfffe
            || BinaryPrimitives.ReadUInt16LittleEndian(ole.AsSpan(30)) is not (9 or 12)
            || BinaryPrimitives.ReadUInt16LittleEndian(ole.AsSpan(32)) != 6
            || BinaryPrimitives.ReadUInt32LittleEndian(ole.AsSpan(56)) != 4096
            || BinaryPrimitives.ReadUInt32LittleEndian(ole.AsSpan(44)) > ole.Length / 512)
            throw new MathTypeException("invalid_equation_ole", "Invalid or unsupported CFB storage header.");
        var native = CompoundFile.ReadStream(ole, "Equation Native")
            ?? throw new MathTypeException("invalid_equation_ole", "Missing or invalid Equation Native stream.");
        if (native.Length < 29 || BinaryPrimitives.ReadUInt16LittleEndian(native) != 28
            || BinaryPrimitives.ReadUInt32LittleEndian(native.AsSpan(2)) != 0x00020000
            || BinaryPrimitives.ReadUInt32LittleEndian(native.AsSpan(8)) != native.Length - 28)
            throw new MathTypeException("invalid_equation_header", "Invalid Equation Native header or payload length.");
        return ReadMtef(native[28..]);
    }

    internal static MathTypeEquation ReadMtef(byte[] data)
    {
        if (data.Length > MaxOleBytes)
            throw new MathTypeException("equation_limit", "MTEF equation exceeds the 8 MiB limit.");
        var reader = new MathTypeReader(data);
        int version = reader.Byte();
        if (version != 5)
            throw new MathTypeException("unsupported_mtef_version", $"MTEF version {version} is not supported; expected version 5.");
        if (reader.Byte() > 1 || reader.Byte() > 1)
            throw reader.Invalid("Invalid MTEF platform or product.");
        reader.Byte(); // producer version
        reader.Byte(); // producer minor version
        reader.CString(); // producer application key, not an instruction
        int options = reader.Byte();
        if ((options & ~1) != 0) throw reader.Unsupported("Unknown equation options.");
        var nodes = reader.ReadList(0);
        if (reader._position != data.Length)
            throw reader.Invalid("Data follows the final MTEF END record.");
        if (nodes.Count == 0)
            throw new MathTypeException("unsupported_empty_equation", "Empty MathType object has no formula to convert.", reader._position);
        if (nodes.Count != 1 || nodes[0].Kind is not (1 or 4))
            throw reader.Invalid("An equation must contain one top-level line or pile.");
        var xml = Element("oMath", reader.RenderSequence(nodes, 0));
        if (!xml.Descendants(MathNs + "t").Any(t => t.Value.Any(c => !char.IsWhiteSpace(c) && char.GetUnicodeCategory(c) != UnicodeCategory.Format)))
            throw new MathTypeException("unsupported_empty_equation", "The equation contains no visible mathematical content.", reader._position);
        var math = new M.OfficeMath(xml.ToString(SaveOptions.DisableFormatting));
        var error = new OpenXmlValidator().Validate(math).FirstOrDefault();
        if (error != null)
            throw reader.Invalid($"Converted OMML is invalid: {error.Description}");
        return new MathTypeEquation(math, (options & 1) != 0);
    }

    private sealed class Node(int kind)
    {
        internal int Kind { get; } = kind;
        internal List<Node> Children { get; set; } = [];
        internal int Selector { get; set; }
        internal int Variation { get; set; }
        internal int Options { get; set; }
        internal int Code { get; set; }
        internal int Style { get; set; }
        internal string? Color { get; set; }
        internal int Rows { get; set; }
        internal int Columns { get; set; }
        internal int Alignment { get; set; }
        internal bool Null { get; set; }
    }

    private List<Node> ReadList(int depth)
    {
        Depth(depth);
        var nodes = new List<Node>();
        while (true)
        {
            if (++_records > MaxRecords) throw Invalid("Too many MTEF records.");
            int type = Byte();
            if (type == 0) return nodes;
            var node = ReadRecord(type, depth + 1);
            if (node != null) nodes.Add(node);
        }
    }

    private Node? ReadRecord(int type, int depth)
    {
        Depth(depth);
        var node = new Node(type);
        switch (type)
        {
            case 1:
            {
                int options = Options(0x0f);
                if ((options & 4) != 0) UInt16(); // custom line spacing
                if ((options & 2) != 0) ReadRulerBody();
                node.Null = (options & 1) != 0;
                if (!node.Null) node.Children = ReadList(depth);
                return node;
            }
            case 2:
            {
                int options = Options(0x3f);
                int typeface = Signed();
                if ((options & 0x14) == 0x14) throw Invalid("Conflicting character encodings.");
                int code = (options & 0x20) == 0 ? UInt16() : -1;
                if ((options & 4) != 0) Byte();
                if ((options & 0x10) != 0) UInt16();
                if (code < 0) throw Unsupported("Font-only character has no Unicode/MTCode value.");
                node.Code = code;
                if (typeface < 0)
                {
                    int index = -typeface - 1;
                    if (index >= _fontStyles.Count) throw Invalid("Undefined explicit font style.");
                    node.Style = _fontStyles[index].Style;
                }
                else if (_styles.TryGetValue(typeface, out int style)) node.Style = style;
                else if (typeface is 22 or 23 or 24) node.Style = 0;
                else throw Unsupported($"Unknown typeface {typeface}.");
                node.Color = _color;
                if ((options & 1) != 0)
                {
                    node.Children = ReadList(depth);
                    if (node.Children.Any(n => n.Kind != 6)) throw Invalid("Non-embellishment in a character embellishment list.");
                }
                return node;
            }
            case 3:
                Options(8);
                node.Selector = Byte();
                int variation = Byte();
                node.Variation = (variation & 0x80) == 0 ? variation : (variation & 0x7f) | (Byte() << 8);
                node.Options = Byte();
                node.Children = ReadList(depth);
                return node;
            case 4:
            {
                int options = Options(0x0a);
                node.Alignment = Byte();
                int vertical = Byte();
                if (node.Alignment is < 1 or > 5 || vertical > 4) throw Invalid("Invalid pile alignment.");
                if ((options & 2) != 0) ReadRulerBody();
                node.Children = ReadList(depth);
                if (node.Children.Count == 0 || node.Children.Any(n => n.Kind != 1))
                    throw Invalid("A pile must contain lines.");
                return node;
            }
            case 5:
                Options(8);
                int valign = Byte();
                node.Alignment = Byte();
                int vjust = Byte();
                node.Rows = Byte();
                node.Columns = Byte();
                if (node.Rows == 0 || node.Columns == 0) throw Invalid("Invalid matrix dimensions.");
                if (valign > 4 || vjust > 4 || node.Alignment > 5)
                    throw Unsupported("Unknown matrix alignment.");
                for (int i = 0; i < (node.Rows + 4) / 4 + (node.Columns + 4) / 4; i++)
                    if (Byte() != 0) throw Unsupported("Matrix partition lines cannot be represented losslessly by this converter.");
                node.Children = ReadList(depth);
                if (node.Children.Count != node.Rows * node.Columns || node.Children.Any(n => n.Kind is not (1 or 4)))
                    throw Invalid("Matrix cell count or structure does not match its dimensions.");
                return node;
            case 6:
                Options(8);
                node.Code = Byte();
                return node;
            case 7:
                ReadRulerBody();
                return null;
            case 8:
                int font = Unsigned();
                int fontStyle = Style();
                if (font < 1 || font > _fonts.Count) throw Invalid("Undefined font.");
                _fontStyles.Add((font, fontStyle));
                return null;
            case 9:
                int size = Byte();
                if (size == 101) UInt16();
                else if (size == 100) { Byte(); UInt16(); }
                else Byte();
                return null;
            case >= 10 and <= 14:
                return null;
            case 15:
                int color = Unsigned();
                if (color > _colors.Count) throw Invalid("Undefined color.");
                _color = color == 0 ? null : _colors[color - 1];
                return null;
            case 16:
                int colorOptions = Byte();
                if ((colorOptions & ~7) != 0) throw Unsupported("Unknown color options.");
                int[] components = new int[(colorOptions & 1) == 0 ? 3 : 4];
                for (int i = 0; i < components.Length; i++)
                {
                    components[i] = UInt16();
                    if (components[i] > 1000) throw Invalid("Invalid color component.");
                }
                if ((colorOptions & 4) != 0) CString();
                if (components.Length == 4) throw Unsupported("CMYK equation colors require an explicit color conversion.");
                _colors.Add(string.Concat(components.Select(c => ((int)Math.Round(c * 255.0 / 1000)).ToString("X2"))));
                return null;
            case 17:
                int encoding = Unsigned();
                if (encoding < 1 || encoding > _encodingCount) throw Invalid("Undefined font encoding.");
                _fonts.Add((encoding, CString()));
                return null;
            case 18:
                if (Byte() != 0) throw Unsupported("Unknown equation preference options.");
                Dimensions();
                Dimensions();
                int count = Byte();
                for (int i = 1; i <= count; i++)
                {
                    int definition = Unsigned();
                    if (definition > _fonts.Count) throw Invalid("Undefined preference font.");
                    if (definition != 0) _styles[i] = Style();
                }
                return null;
            case 19:
                CString();
                _encodingCount++;
                return null;
            default:
                throw Unsupported($"Unsupported MTEF record {type}.");
        }
    }

    private int Options(int allowed)
    {
        int options = Byte();
        if ((options & ~allowed) != 0) throw Unsupported("Unknown MTEF record options.");
        if ((options & 8) != 0)
        {
            int x = Byte(), y = Byte();
            if (x == 128 && y == 128) { UInt16(); UInt16(); }
        }
        return options;
    }

    private void ReadRulerBody()
    {
        int count = Byte();
        for (int i = 0; i < count; i++)
        {
            if (Byte() > 4) throw Invalid("Invalid tab-stop type.");
            UInt16();
        }
    }

    private void Dimensions()
    {
        int count = Byte();
        int pending = -1;
        int Nibble()
        {
            if (pending >= 0) { int value = pending; pending = -1; return value; }
            int valueByte = Byte();
            pending = valueByte & 15;
            return valueByte >> 4;
        }
        for (int i = 0; i < count; i++)
        {
            if (Nibble() > 4) throw Invalid("Invalid dimension unit.");
            int digits = 0;
            for (int value = Nibble(); value != 15; value = Nibble())
                if (value > 11 || ++digits > 32) throw Invalid("Invalid dimension value.");
        }
        if (pending is not (-1 or 0)) throw Invalid("Invalid dimension padding.");
    }

    private int Style()
    {
        int style = Byte();
        if (style > 3) throw Unsupported("Unknown font style bits.");
        return style;
    }

    private int Byte()
    {
        if (_position >= _data.Length) throw Invalid("Truncated MTEF record.");
        return _data[_position++];
    }

    private int UInt16() => Byte() | (Byte() << 8);
    private int Unsigned() { int value = Byte(); return value == 255 ? UInt16() : value; }
    private int Signed() { int value = Byte(); return value == 255 ? UInt16() - 32768 : value - 128; }

    private string CString()
    {
        var bytes = new List<byte>();
        for (int value = Byte(); value != 0; value = Byte())
        {
            if (bytes.Count >= 1024) throw Invalid("MTEF string exceeds its length limit.");
            bytes.Add((byte)value);
        }
        return Encoding.Latin1.GetString(bytes.ToArray());
    }

    private void Depth(int depth)
    {
        if (depth > MaxDepth) throw Invalid("MTEF nesting exceeds 64 levels.");
        DocumentLimits.EnsureDepth(depth);
    }

    private MathTypeException Invalid(string message) => new("invalid_mtef", message, _position);
    private MathTypeException Unsupported(string message) => new("unsupported_mtef", message, _position);
}
