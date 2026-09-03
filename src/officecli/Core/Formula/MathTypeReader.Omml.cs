// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Xml;
using System.Xml.Linq;

namespace OfficeCli.Core;

internal sealed partial class MathTypeReader
{
    private static readonly XNamespace MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement Element(string name, params object?[] content) => new(MathNs + name, content);
    private static XElement Value(string name, string value) => Element(name, new XAttribute(MathNs + "val", value));
    private static XElement Run(string text, int style = 0, string? color = null) => Element("r",
        Element("rPr", Value("sty", style switch { 1 => "b", 2 => "i", 3 => "bi", _ => "p" })),
        color == null ? null : new XElement(WordNs + "rPr", new XElement(WordNs + "color", new XAttribute(WordNs + "val", color))),
        Element("t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    private List<XElement> RenderSequence(List<Node> nodes, int depth)
    {
        Depth(depth);
        var output = new List<XElement>();
        for (int i = 0; i < nodes.Count; i++)
        {
            var node = nodes[i];
            if (node.Kind == 1)
            {
                output.AddRange(RenderSequence(node.Children, depth + 1));
                continue;
            }
            if (node.Kind == 3 && node.Selector is >= 27 and <= 29)
            {
                if (node.Options != 0) throw Unsupported("Unknown script options.");
                Variants(node, 1);
                Slots(node, 2);
                if (node.Selector == 27) RequireAbsent(node.Children[1]);
                if (node.Selector == 28) RequireAbsent(node.Children[0]);
                XElement basis;
                if ((node.Variation & 1) != 0)
                {
                    if (++i >= nodes.Count || nodes[i].Kind is not (2 or 3 or 5))
                        throw Unsupported("Prescript has no unambiguous base.");
                    basis = Render(nodes[i], depth + 1);
                    output.Add(Element("sPre", Slot("sub", node.Children[0], depth),
                        Slot("sup", node.Children[1], depth), Element("e", basis)));
                    continue;
                }
                if (output.Count == 0) throw Invalid("Script has no preceding base.");
                basis = output[^1];
                output.RemoveAt(output.Count - 1);
                if (node.Selector == 27)
                {
                    output.Add(Element("sSub", Element("e", basis), Slot("sub", node.Children[0], depth)));
                }
                else if (node.Selector == 28)
                {
                    output.Add(Element("sSup", Element("e", basis), Slot("sup", node.Children[1], depth)));
                }
                else output.Add(Element("sSubSup", Element("e", basis),
                    Slot("sub", node.Children[0], depth), Slot("sup", node.Children[1], depth)));
                continue;
            }
            output.Add(Render(node, depth + 1));
        }
        return output;
    }

    private XElement Render(Node node, int depth)
    {
        Depth(depth);
        switch (node.Kind)
        {
            case 2:
                var character = Run(Unicode(node.Code), node.Style, node.Color);
                foreach (var embellishment in node.Children)
                    character = Embellish(character, embellishment.Code);
                return character;
            case 3:
                return Template(node, depth + 1);
            case 4:
                if (node.Alignment > 3)
                    throw Unsupported("Relational/decimal pile alignment is not supported.");
                return Element("eqArr", node.Children.Select(n => Slot("e", n, depth)).ToArray());
            case 5:
                var matrix = Element("m", Element("mPr", Element("mcs", Element("mc", Element("mcPr",
                    Value("count", node.Columns.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                    Value("mcJc", node.Alignment switch { 1 => "left", 3 => "right", _ => "center" }))))));
                for (int row = 0; row < node.Rows; row++)
                    matrix.Add(Element("mr", node.Children.Skip(row * node.Columns).Take(node.Columns)
                        .Select(n => Slot("e", n, depth)).ToArray()));
                return matrix;
            default:
                throw Invalid($"Record {node.Kind} cannot appear in a mathematical expression.");
        }
    }

    private XElement Template(Node node, int depth)
    {
        Depth(depth);
        if (node.Selector is >= 0 and <= 9) return Fence(node, depth);
        if (node.Options != 0 && !(node.Selector == 15 && node.Options == 1))
            throw Unsupported($"Unsupported template options for selector {node.Selector}.");
        switch (node.Selector)
        {
            case 10:
                Variants(node, 1);
                Slots(node, 2);
                if (node.Variation == 0) RequireAbsent(node.Children[1]);
                return Element("rad", Element("radPr", Value("degHide", node.Variation == 0 ? "1" : "0")),
                    Slot("deg", node.Children[1], depth), Slot("e", node.Children[0], depth));
            case 11:
                Variants(node, 7);
                Slots(node, 2);
                if ((node.Variation & 4) != 0 && (node.Variation & 2) == 0)
                    throw Unsupported("Baseline fraction without a slash.");
                return Element("f", Element("fPr", Value("type", (node.Variation & 2) == 0 ? "bar"
                    : (node.Variation & 4) == 0 ? "skw" : "lin")),
                    Slot("num", node.Children[0], depth), Slot("den", node.Children[1], depth));
            case 12:
            case 13:
                Variants(node, 0);
                Slots(node, 1);
                return Element("bar", Element("barPr", Value("pos", node.Selector == 12 ? "bot" : "top")),
                    Slot("e", node.Children[0], depth));
            case >= 15 and <= 22:
                return Nary(node, depth);
            case 23:
                Variants(node, 0x43);
                Slots(node, 3);
                var expression = RenderSequence(node.Children[0].Children, depth);
                if (!node.Children[1].Null)
                    expression = [Element("limLow", Element("e", expression), Slot("lim", node.Children[1], depth))];
                if (!node.Children[2].Null)
                    expression = [Element("limUpp", Element("e", expression), Slot("lim", node.Children[2], depth))];
                return Element("box", Element("e", expression));
            case 24:
            case 25:
                Variants(node, 1);
                Slots(node, 2, 1);
                string brace = Glyph(node.Children[2]);
                if (brace is not ("⏞" or "⏟" or "⎴" or "⎵"))
                    throw Unsupported("Unrecognized horizontal brace character.");
                bool top = (node.Variation & 1) != 0;
                var group = Element("groupChr", Element("groupChrPr", Value("chr", brace),
                    Value("pos", top ? "top" : "bot"), Value("vertJc", top ? "bot" : "top")),
                    Slot("e", node.Children[0], depth));
                return Element(top ? "limUpp" : "limLow", Element("e", group), Slot("lim", node.Children[1], depth));
            case 31:
                Variants(node, 0x0f);
                Slots(node, 1);
                if ((node.Variation & 4) != 0) throw Unsupported("Under-vectors are not supported.");
                string arrow = (node.Variation & 3, (node.Variation & 8) != 0) switch
                {
                    (1, false) => "⃖", (2, false) => "⃗", (3, false) => "⃡",
                    (1, true) => "⃐", (2, true) => "⃑",
                    _ => throw Unsupported("Unsupported vector direction."),
                };
                return Element("acc", Element("accPr", Value("chr", arrow)), Slot("e", node.Children[0], depth));
            case 32:
            case 33:
            case 34:
                Variants(node, 0);
                Slots(node, 1);
                return Element("acc", Element("accPr", Value("chr", node.Selector switch
                { 32 => "̃", 33 => "̂", _ => "⌢" })), Slot("e", node.Children[0], depth));
            default:
                throw Unsupported($"Unsupported MTEF template {node.Selector}, variation {node.Variation}.");
        }
    }

    private XElement Fence(Node node, int depth)
    {
        if (node.Options > 2) throw Unsupported("Unknown fence alignment.");
        Variants(node, node.Selector == 9 ? 0x33 : 3);
        bool left = node.Selector == 9 || (node.Variation & 1) != 0;
        bool right = node.Selector == 9 || (node.Variation & 2) != 0;
        Slots(node, 1, (left ? 1 : 0) + (right ? 1 : 0));
        int index = 1;
        string opening = left ? Glyph(node.Children[index++]) : "";
        string closing = right ? Glyph(node.Children[index]) : "";
        const string fences = "()[]{}|‖∥⌊⌋⌈⌉⟨⟩〈〉⟦⟧<>";
        if ((left && !fences.Contains(opening, StringComparison.Ordinal))
            || (right && !fences.Contains(closing, StringComparison.Ordinal)))
            throw Unsupported("Unrecognized fence character.");
        return Element("d", Element("dPr", Value("begChr", opening), Value("endChr", closing), Value("grow", "1")),
            Slot("e", node.Children[0], depth));
    }

    private XElement Nary(Node node, int depth)
    {
        // MathType 7 also stores limit-presence bits at 0x10/0x20. The explicit
        // upper/lower slots carry the content; neither encoding may discard one.
        Variants(node, node.Selector == 15 ? 0x17f : 0x73);
        Slots(node, 3, 1);
        string glyph = Glyph(node.Children[3]);
        string accepted = node.Selector switch
        {
            15 => "∫∬∭∮∯∰∱∲∳", 16 => "∑", 17 => "∏", 18 => "∐", 19 => "⋃∪", 20 => "⋂∩",
            _ => "∫∬∭∮∑∏∐⋃⋂⋁⋀⨁⨂⨀",
        };
        if (!accepted.Contains(glyph, StringComparison.Ordinal)) throw Unsupported("Unrecognized large-operator character.");
        return Element("nary", Element("naryPr", Value("chr", glyph),
                Value("limLoc", (node.Variation & 0x40) != 0 ? "undOvr" : "subSup"),
                Value("grow", node.Selector != 15 || (node.Variation & 0x100) != 0 ? "1" : "0"),
                Value("subHide", node.Children[2].Null ? "1" : "0"),
                Value("supHide", node.Children[1].Null ? "1" : "0")),
            Slot("sub", node.Children[2], depth), Slot("sup", node.Children[1], depth), Slot("e", node.Children[0], depth));
    }

    private XElement Embellish(XElement basis, int code)
    {
        if (code is 5 or 6 or 18)
            return Element("sSup", Element("e", basis), Element("sup", Run(code switch { 5 => "′", 6 => "″", _ => "‴" })));
        if (code is 17 or 29)
            return Element("bar", Element("barPr", Value("pos", code == 17 ? "top" : "bot")), Element("e", basis));
        string accent = code switch
        {
            2 => "̇", 3 => "̈", 4 => "⃛", 8 => "̃", 9 => "̂", 11 => "⃗", 12 => "⃖", 13 => "⃡",
            14 => "⃑", 15 => "⃐", 19 => "⌢", 20 => "⌣", 24 => "⃜",
            _ => throw Unsupported($"Unsupported MTEF embellishment {code}."),
        };
        return Element("acc", Element("accPr", Value("chr", accent)), Element("e", basis));
    }

    private XElement Slot(string name, Node node, int depth)
    {
        if (node.Kind == 4) return Element(name, Render(node, depth + 1));
        if (node.Kind != 1) throw Invalid("Expected a template slot.");
        return Element(name, RenderSequence(node.Children, depth + 1));
    }

    private void Slots(Node node, int lines, int characters = 0)
    {
        if (node.Children.Count != lines + characters || node.Children.Take(lines).Any(n => n.Kind != 1)
            || node.Children.Skip(lines).Any(n => n.Kind != 2))
            throw Invalid($"Template {node.Selector} has an invalid slot/glyph structure.");
    }

    private void Variants(Node node, int allowed)
    {
        if ((node.Variation & ~allowed) != 0)
            throw Unsupported($"Unsupported variation {node.Variation} of template {node.Selector}.");
    }

    private void RequireAbsent(Node node)
    {
        if (node.Children.Count != 0) throw Invalid("An absent template slot contains data.");
    }

    private string Glyph(Node node)
    {
        if (node.Kind != 2 || node.Children.Count != 0) throw Unsupported("Decorated template glyph is not supported.");
        return Unicode(node.Code);
    }

    private string Unicode(int code)
    {
        // MathType's published MathML translator maps these virtual spaces to
        // Unicode. Other private MTCode values must never become blank glyphs.
        string? space = code switch
        {
            0xef01 => "\u200b", 0xef02 => "\u2009", 0xef03 => "\u205f",
            0xef04 => "\u2004", 0xef05 => "\u2003", 0xef06 => "\u2003\u2003",
            0xef08 => "\u200a", _ => null,
        };
        if (space != null) return space;
        if (code is >= 0xe000 and <= 0xf8ff)
            throw Unsupported($"Private MTCode U+{code:X4} has no verified Unicode mapping.");
        if (code < 0x20 || !XmlConvert.IsXmlChar((char)code) || char.IsSurrogate((char)code))
            throw Unsupported($"Unsupported character U+{code:X4}.");
        return ((char)code).ToString();
    }
}
