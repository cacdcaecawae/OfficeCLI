// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using Xunit;
using static OfficeCli.Tests.MathTypeTestData;

namespace OfficeCli.Tests;

public class MathTypeReaderTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static XElement Read(params byte[][] records)
    {
        var equation = MathTypeReader.ReadMtef(Equation(records));
        Assert.Empty(new OpenXmlValidator().Validate(equation.Math));
        Assert.False(string.IsNullOrWhiteSpace(FormulaParser.ToLatex(equation.Math)));
        return XElement.Parse(equation.Math.OuterXml);
    }

    [Fact]
    public void FractionAndScriptsRetainTheirOperandOrder()
    {
        var xml = Read(Fraction);
        var fraction = Assert.Single(xml.Elements(M + "f"));
        var power = Assert.Single(fraction.Element(M + "num")!.Elements(M + "sSup"));
        Assert.Equal("x", power.Element(M + "e")!.Value);
        Assert.Equal("2", power.Element(M + "sup")!.Value);
        Assert.Equal("y+1", fraction.Element(M + "den")!.Value);
    }

    [Theory]
    [InlineData(27, "sSub")]
    [InlineData(28, "sSup")]
    [InlineData(29, "sSubSup")]
    public void PostfixScriptAttachesToLastBaseOnly(int selector, string name)
    {
        var xml = Read(Text("ax"), Template(selector, 0,
            selector == 28 ? NullLine : Line(Character('i')),
            selector == 27 ? NullLine : Line(Character('2'))));
        Assert.Equal("a", xml.Elements().First().Value);
        Assert.Equal("x", Assert.Single(xml.Elements(M + name)).Element(M + "e")!.Value);
    }

    [Fact]
    public void PrescriptUsesTheFollowingBase()
    {
        var xml = Read(Template(29, 1, Line(Character('i')), Line(Character('j'))), Character('x'));
        var script = Assert.Single(xml.Elements(M + "sPre"));
        Assert.Equal("x", script.Element(M + "e")!.Value);
    }

    [Fact]
    public void ScriptsRejectUnknownOptionsAndContentInAbsentSlots()
    {
        var script = Template(28, 0, NullLine, Line(Character('2')));
        script[4] = 1;
        Assert.Equal("unsupported_mtef", Assert.Throws<MathTypeException>(() => Read(Character('x'), script)).Code);
        Assert.Equal("invalid_mtef", Assert.Throws<MathTypeException>(() => Read(
            Template(28, 1, Line(Character('i')), Line(Character('2'))), Character('x'))).Code);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    public void RootDegreeIsExplicit(int variation, bool hidden)
    {
        var xml = Read(Template(10, variation, Line(Text("x+1")), hidden ? NullLine : Line(Character('3'))));
        var root = Assert.Single(xml.Elements(M + "rad"));
        Assert.Equal(hidden ? "1" : "0", root.Element(M + "radPr")!.Element(M + "degHide")!.Attribute(M + "val")!.Value);
        Assert.Equal("x+1", root.Element(M + "e")!.Value);
        Assert.Equal(hidden ? "" : "3", root.Element(M + "deg")!.Value);
    }

    [Theory]
    [InlineData(0, "bar")]
    [InlineData(1, "bar")]
    [InlineData(2, "skw")]
    [InlineData(6, "lin")]
    public void FractionStyleIsNative(int variation, string type)
    {
        var xml = Read(Template(11, variation, Line(Character('a')), Line(Character('b'))));
        Assert.Equal(type, xml.Descendants(M + "type").Single().Attribute(M + "val")!.Value);
    }

    [Theory]
    [InlineData(0, '⟨', '⟩')]
    [InlineData(1, '(', ')')]
    [InlineData(2, '{', '}')]
    [InlineData(3, '[', ']')]
    [InlineData(4, '|', '|')]
    [InlineData(5, '‖', '‖')]
    [InlineData(6, '⌊', '⌋')]
    [InlineData(7, '⌈', '⌉')]
    [InlineData(8, '⟦', '⟧')]
    public void FencesRetainActualGlyphs(int selector, char left, char right)
    {
        var xml = Read(Template(selector, 3, Line(Fraction), Character(left), Character(right)));
        Assert.Equal(left.ToString(), xml.Descendants(M + "begChr").Single().Attribute(M + "val")!.Value);
        Assert.Equal(right.ToString(), xml.Descendants(M + "endChr").Single().Attribute(M + "val")!.Value);
    }

    [Theory]
    [InlineData(15, '∫', 1)]
    [InlineData(16, '∑', 0x43)]
    [InlineData(16, '∑', 0x70)]
    [InlineData(17, '∏', 0x43)]
    [InlineData(18, '∐', 0x43)]
    [InlineData(19, '⋃', 0x43)]
    [InlineData(20, '⋂', 0x43)]
    [InlineData(21, '⨂', 0)]
    [InlineData(22, '⨁', 0x40)]
    public void LargeOperatorsRetainBodyAndLimits(int selector, char glyph, int variation)
    {
        var xml = Read(Template(selector, variation, Line(Character('x')), Line(Text("i=0")), Line(Character('n')), Character(glyph)));
        var nary = Assert.Single(xml.Elements(M + "nary"));
        Assert.Equal("x", nary.Element(M + "e")!.Value);
        Assert.Equal("n", nary.Element(M + "sup")!.Value);
        Assert.Equal("i=0", nary.Element(M + "sub")!.Value);
    }

    [Fact]
    public void NestedMathType7SummationsKeepLowerAndUpperLimits()
    {
        var xml = Read(NestedSums);
        var sums = xml.Descendants(M + "nary").ToArray();
        Assert.Equal(2, sums.Length);
        Assert.Equal(["a=2", "b=3"], sums.Select(n => n.Element(M + "sub")!.Value));
        Assert.Equal(["A", "B"], sums.Select(n => n.Element(M + "sup")!.Value));
        Assert.Equal("p+q", sums[1].Element(M + "e")!.Value);
    }

    [Theory]
    [InlineData(15, '∫')]
    [InlineData(16, '∑')]
    [InlineData(17, '∏')]
    [InlineData(18, '∐')]
    [InlineData(19, '⋃')]
    [InlineData(20, '⋂')]
    [InlineData(21, '⨂')]
    [InlineData(22, '⨁')]
    public void LargeOperatorLimitVisibilityFollowsTheCorrespondingSlot(int selector, char glyph)
    {
        for (int limits = 0; limits < 4; limits++)
        foreach (int position in new[] { 0, 0x40 })
        {
            bool hasLower = (limits & 1) != 0, hasUpper = (limits & 2) != 0;
            int variation = (limits << 4) | position | (selector == 15 ? 1 : 0);
            var xml = Read(Template(selector, variation, Line(Character('x')),
                hasLower ? Line(Text("a")) : NullLine,
                hasUpper ? Line(Text("b")) : NullLine, Character(glyph)));
            var nary = Assert.Single(xml.Elements(M + "nary"));
            Assert.Equal(hasLower ? "a" : "", nary.Element(M + "sub")!.Value);
            Assert.Equal(hasUpper ? "b" : "", nary.Element(M + "sup")!.Value);
            var properties = nary.Element(M + "naryPr")!;
            Assert.Equal(hasLower ? "0" : "1", properties.Element(M + "subHide")!.Attribute(M + "val")!.Value);
            Assert.Equal(hasUpper ? "0" : "1", properties.Element(M + "supHide")!.Attribute(M + "val")!.Value);
            Assert.Equal(position == 0 ? "subSup" : "undOvr", properties.Element(M + "limLoc")!.Attribute(M + "val")!.Value);
        }
    }

    [Theory]
    [InlineData(15, '∫', 0x31)]
    [InlineData(16, '∑', 0x70)]
    public void LimitValuesDoNotDetermineTheirPositions(int selector, char glyph, int variation)
    {
        var xml = Read(Template(selector, variation, Line(Character('x')), Line(Text("9")), Line(Text("1")), Character(glyph)));
        var nary = Assert.Single(xml.Elements(M + "nary"));
        Assert.Equal("9", nary.Element(M + "sub")!.Value);
        Assert.Equal("1", nary.Element(M + "sup")!.Value);
    }

    [Fact]
    public void MalformedLargeOperatorsAndUnknownVariationsRemainRejected()
    {
        Assert.Equal("invalid_mtef", Assert.Throws<MathTypeException>(() => Read(
            Template(16, 0x70, Line(Character('x')), Line(Text("a")), Character('∑')))).Code);
        Assert.Equal("invalid_mtef", Assert.Throws<MathTypeException>(() => Read(
            Template(16, 0x70, Line(Character('x')), Character('a'), Line(Text("b")), Character('∑')))).Code);
        Assert.Equal("unsupported_mtef", Assert.Throws<MathTypeException>(() => Read(
            Template(16, 0x100, Line(Character('x')), Line(Text("a")), Line(Text("b")), Character('∑')))).Code);
    }

    [Fact]
    public void MatricesRetainEveryCellInRowMajorOrder()
    {
        var matrix = Join([5, 0, 1, 0, 1, 2, 2, 0, 0], Line(Character('a')), Line(Character('b')), Line(Character('c')), Line(Character('d')), [0]);
        var xml = Read(matrix);
        Assert.Equal(["a", "b", "c", "d"], xml.Descendants(M + "mr").SelectMany(r => r.Elements(M + "e")).Select(e => e.Value));
    }

    [Fact]
    public void LineRulerPayloadDoesNotConsumeMathematicalContent()
    {
        var equation = MathTypeReader.ReadMtef(Raw(true, Join([1, 2, 1, 0, 100, 0], Text("xyz"), [0])));
        Assert.Equal("xyz", equation.Math.InnerText);
    }

    [Fact]
    public void MatrixCellsMayContainAnExplicitMultilinePile()
    {
        var pile = Join([4, 0, 2, 1], Line(Character('a')), Line(Character('b')), [0]);
        var matrix = Join([5, 0, 1, 0, 1, 1, 2, 0, 0], Line(Character('x')), pile, [0]);
        var xml = Read(matrix);
        Assert.Equal(2, xml.Descendants(M + "mr").Single().Elements(M + "e").Count());
        Assert.Equal(["a", "b"], xml.Descendants(M + "eqArr").Single().Elements(M + "e").Select(e => e.Value));
    }

    [Theory]
    [InlineData(0xef01, "\u200b")]
    [InlineData(0xef02, "\u2009")]
    [InlineData(0xef03, "\u205f")]
    [InlineData(0xef04, "\u2004")]
    [InlineData(0xef05, "\u2003")]
    [InlineData(0xef06, "\u2003\u2003")]
    [InlineData(0xef08, "\u200a")]
    public void VerifiedVirtualSpacesUseUnicodeWithoutDroppingAdjacentVariables(int code, string space)
    {
        var xml = Read(Character('x'), Character((char)code), Character('y'));
        Assert.Equal("x" + space + "y", xml.Value);
    }

    [Fact]
    public void EmptyOrInvisibleMathTypeObjectsAreReportedAsUnconvertible()
    {
        Assert.Equal("unsupported_empty_equation", Assert.Throws<MathTypeException>(() => MathTypeReader.ReadMtef(Raw(true))).Code);
        Assert.Equal("unsupported_empty_equation", Assert.Throws<MathTypeException>(() => Read(Character('\uef01'))).Code);
    }

    [Theory]
    [InlineData(2)] [InlineData(3)] [InlineData(4)] [InlineData(5)] [InlineData(6)]
    [InlineData(8)] [InlineData(9)] [InlineData(11)] [InlineData(12)] [InlineData(13)]
    [InlineData(14)] [InlineData(15)] [InlineData(17)] [InlineData(18)] [InlineData(19)]
    [InlineData(20)] [InlineData(24)] [InlineData(29)]
    public void SupportedEmbellishmentsAreNativeStructures(int code)
    {
        var xml = Read([2, 1, 131, (byte)'x', 0, 6, 0, (byte)code, 0]);
        Assert.Contains(xml.Descendants(M + "t"), e => e.Value == "x");
        Assert.DoesNotContain(xml.Elements(), e => e.Name == M + "r");
    }

    [Fact]
    public void RgbColorAndExplicitBoldItalicAreRetained()
    {
        var mtef = Raw(true, Join([17, 1], System.Text.Encoding.ASCII.GetBytes("Cambria Math\0")), [8, 1, 3],
            [16, 0, 232, 3, 0, 0, 0, 0], [15, 1], Line(Character('x', -1)));
        var xml = XElement.Parse(MathTypeReader.ReadMtef(mtef).Math.OuterXml);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal("bi", xml.Descendants(M + "sty").Single().Attribute(M + "val")!.Value);
        Assert.Equal("FF0000", xml.Descendants(w + "color").Single().Attribute(w + "val")!.Value);
    }

    [Fact]
    public void OleNativeHeaderLengthsAreChecked()
    {
        Assert.NotNull(MathTypeReader.ReadOle(Ole(Equation(Fraction))));
        var native = new byte[64];
        BinaryPrimitives.WriteUInt16LittleEndian(native, 28);
        BinaryPrimitives.WriteUInt32LittleEndian(native.AsSpan(2), 0x00020000);
        BinaryPrimitives.WriteUInt32LittleEndian(native.AsSpan(8), 10000);
        Assert.Equal("invalid_equation_header", Assert.Throws<MathTypeException>(() => MathTypeReader.ReadOle(CompoundFile.WriteSingleStream("Equation Native", native))).Code);
        var invalidCfb = Ole(Equation(Character('x')));
        BinaryPrimitives.WriteUInt32LittleEndian(invalidCfb.AsSpan(56), uint.MaxValue);
        Assert.Equal("invalid_equation_ole", Assert.Throws<MathTypeException>(() => MathTypeReader.ReadOle(invalidCfb)).Code);
    }

    [Fact]
    public void TruncationNeverReturnsAPartialEquation()
    {
        var data = Equation(Fraction);
        for (int length = 0; length < data.Length; length++)
            Assert.Throws<MathTypeException>(() => MathTypeReader.ReadMtef(data[..length]));
        Assert.Throws<MathTypeException>(() => MathTypeReader.ReadMtef(Join(data, [0])));
    }

    [Fact]
    public void UnknownRecordsAndCharactersAreNotSilentlyDropped()
    {
        Assert.Throws<MathTypeException>(() => Read(Character('\ue001')));
        Assert.Throws<MathTypeException>(() => Read([100, 1, 42]));
        Assert.Throws<MathTypeException>(() => Read(Template(99, 0, Line(Character('x')))));
        Assert.Throws<MathTypeException>(() => Read(Template(11, 0x10, Line(Character('x')), Line(Character('y')))));
        Assert.Throws<MathTypeException>(() => Read(Template(28, 0, NullLine, Line(Character('2')))));
        Assert.Throws<MathTypeException>(() => Read(Template(11, 0, Line(Character('x')))));
    }

    [Fact]
    public void NestingIsBounded()
    {
        byte[] nested = Character('x');
        for (int i = 0; i < 100; i++) nested = Template(11, 0, Line(nested), Line(Character('y')));
        Assert.Throws<MathTypeException>(() => Read(nested));
    }
}
