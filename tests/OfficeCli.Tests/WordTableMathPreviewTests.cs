// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Handlers;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Tests;

public class WordTableMathPreviewTests
{
    [Fact]
    public void InlineEquationsRetainSurroundingTextAndEveryEquationInOrder()
    {
        string html = CellHtml(new Run(new Text("Before ")), Math("x"),
            new Run(new Text(" between ")), Math("y"), new Run(new Text(" after")));
        Assert.Equal(2, Regex.Matches(html, "data-formula=").Count);
        Assert.Contains("Before ", html);
        Assert.Contains(" between ", html);
        Assert.Contains(" after", html);
        Assert.True(html.IndexOf("Before ", StringComparison.Ordinal) < html.IndexOf("data-formula=\"x\"", StringComparison.Ordinal));
        Assert.True(html.IndexOf("data-formula=\"x\"", StringComparison.Ordinal) < html.IndexOf(" between ", StringComparison.Ordinal));
        Assert.True(html.IndexOf(" between ", StringComparison.Ordinal) < html.IndexOf("data-formula=\"y\"", StringComparison.Ordinal));
        Assert.True(html.IndexOf("data-formula=\"y\"", StringComparison.Ordinal) < html.IndexOf(" after", StringComparison.Ordinal));
        Assert.DoesNotContain("data-display", html);
    }

    [Fact]
    public void InlineOnlyEquationIsNotPromotedToDisplayMath()
    {
        string html = CellHtml(Math("x"));
        Assert.Contains("data-formula=\"x\"", html);
        Assert.DoesNotContain("data-display", html);
        Assert.DoesNotContain("&nbsp;", html);
    }

    [Fact]
    public void ExclusiveMathParagraphRemainsADisplayBlock()
    {
        string html = CellHtml(new M.Paragraph(Math("x")));
        Assert.Contains("class=\"equation\"", html);
        Assert.Contains("data-display=\"true\"", html);
        Assert.Single(Regex.Matches(html, "data-formula="));
    }

    [Fact]
    public void DisplayMathWithOtherContentDoesNotReplaceTheCellParagraph()
    {
        string html = CellHtml(new Run(new Text("Before ")), new M.Paragraph(Math("x")), new Run(new Text(" after")));
        Assert.Contains("Before ", html);
        Assert.Contains(" after", html);
        Assert.Contains("data-formula=\"x\"", html);
        Assert.Contains("data-display=\"true\"", html);
    }

    private static M.OfficeMath Math(string text) => new(new M.Run(new M.Text(text)));

    private static string CellHtml(params OpenXmlElement[] content)
    {
        string path = Path.Combine(Path.GetTempPath(), $"officecli-table-math-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "4800" }),
                    new TableRow(new TableCell(new Paragraph(content))))));
                Assert.Empty(new OpenXmlValidator().Validate(doc));
            }
            using var handler = new WordHandler(path, editable: false);
            var cell = Regex.Match(handler.ViewAsHtml(), "<td\\b[^>]*>(.*?)</td>", RegexOptions.Singleline);
            Assert.True(cell.Success);
            return cell.Groups[1].Value;
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
