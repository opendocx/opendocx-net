using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OpenDocx;
using OpenXmlPowerTools;
using Normalizer = OpenDocx.Normalizer;
using Xunit;

namespace OpenDocxTemplater.Tests
{
    public partial class ImageParagraphTests
    {
        private static XElement PlacementWhitespace(string kind) => new XElement(W.r,
            kind == "tab" ? new XElement(W.tab) :
            kind == "line-break" ? new XElement(W.br) :
            kind == "carriage-return" ? new XElement(W.cr) :
            kind == "page-break" || kind == "column-break" ? new XElement(W.br, new XAttribute(W.type, kind.Split('-')[0])) :
            kind == "rendered-break" ? new XElement(W.lastRenderedPageBreak) :
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), PlacementText(kind)));

        private static string PlacementText(string kind) => kind switch
        {
            "space" => " ", "text-tab" => "\t", "nbsp" => "\u00a0", "narrow-nbsp" => "\u202f",
            "em-space" => "\u2003", "text-lf" => "\n", "text-cr" => "\r", _ => ""
        };

        public static IEnumerable<object[]> ContentPlacementCases()
        {
            foreach (var controls in new[] { false, true })
                foreach (var before in new[] { false, true })
                    foreach (var kind in new[] { "space", "tab", "text-tab", "line-break", "carriage-return", "page-break",
                        "column-break", "nbsp", "narrow-nbsp", "em-space", "text-lf", "text-cr", "rendered-break", "none" })
                        yield return new object[] { controls, before, kind };
        }

        [Theory]
        [MemberData(nameof(ContentPlacementCases))]
        public void ContentPlacementTreatsAllSurroundingWhitespaceAsSignificant(bool controls, bool before, string kind)
        {
            var runLevel = kind != "none" && kind != "rendered-break";
            foreach (var anchor in new[] { false, true })
            {
                var properties = new XElement(W.pPr,
                    new XElement(W.tabs, new XElement(W.tab, new XAttribute(W.val, "right"), new XAttribute(W.pos, 7200))),
                    new XElement(W.ind, new XAttribute(W.left, 720)));
                var whitespace = PlacementWhitespace(kind);
                var field = Field("Value", controls);
                var source = Document(new XElement(W.p, properties, before ? whitespace : field,
                    before ? field : whitespace, anchor ? new XElement(W.r, Drawing(true)) : null));
                var analysis = Normalizer.AnalyzeTemplate(source, contentControls: controls);
                Assert.Equal(runLevel, JsonNode.Parse(analysis.ExtractedFields)[0][0] is JsonArray);
                if (anchor && !runLevel)
                {
                    Assert.Equal("FloatingImageOnBlockField", Assert.Single(analysis.Diagnostics).Code);
                    Assert.Throws<FieldParseException>(() => Normalizer.ExtractFieldsOnly(source, contentControls: controls));
                    continue;
                }
                Assert.Empty(analysis.Diagnostics);
                Assert.Equal(analysis.ExtractedFields, Normalizer.ExtractFieldsOnly(source, contentControls: controls));
                var fields = Templater.ParseFieldsToDict(analysis.ExtractedFields);
                var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, fields);
                Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
                foreach (var value in new[] { null, "", "Replacement" })
                {
                    var data = new XElement("_odx");
                    if (value != null) data.Add(new XElement(fields["1"].Atom, value));
                    var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
                    Assert.False(assembled.HasErrors, assembled.Error);
                    var paragraph = Assert.Single(Root(assembled.Bytes).Element(W.body).Elements(W.p));
                    Assert.True(XNode.DeepEquals(properties, WithoutNamespaceDeclarations(paragraph.Element(W.pPr))));
                    // The fixture's XML serialization normalizes literal CR to LF.
                    // Authored w:cr is tested separately and must remain an element.
                    var text = PlacementText(kind).Replace('\r', '\n');
                    Assert.Equal(before ? text + (value ?? "") : (value ?? "") + text,
                        string.Concat(paragraph.Descendants(W.t).Select(t => t.Value)));
                    Assert.Equal(kind == "tab" ? 1 : 0, paragraph.Elements(W.r).Elements(W.tab).Count());
                    Assert.Equal(kind.EndsWith("-break") && kind != "rendered-break" ? 1 : 0, paragraph.Descendants(W.br).Count());
                    Assert.Equal(kind == "carriage-return" ? 1 : 0, paragraph.Descendants(W.cr).Count());
                    Assert.Equal(anchor ? 1 : 0, paragraph.Descendants(WP.anchor).Count());
                }
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void TabsAndSplitFieldRunsKeepTheirBoundaries(bool anchorBefore)
        {
            var source = Document(new XElement(W.p,
                anchorBefore ? new XElement(W.r, Drawing(true)) : null,
                new XElement(W.r, new XElement(W.tab), new XElement(W.t, "{[Val")),
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "ue]}"), new XElement(W.tab)),
                !anchorBefore ? new XElement(W.r, Drawing(true)) : null));
            var normalized = Normalizer.NormalizeTemplate(source, false);
            Assert.IsType<JsonArray>(JsonNode.Parse(normalized.ExtractedFields)[0][0]);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            Assert.Equal("Value", Assert.Single(fields).Value.Expression);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var result = Assembler.AssembleDocument(compiled.Bytes, new XElement("_odx", new XElement(fields["1"].Atom, "Replacement")));
            Assert.False(result.HasErrors, result.Error);
            var paragraph = Assert.Single(Root(result.Bytes).Element(W.body).Elements(W.p));
            Assert.Equal("Replacement", string.Concat(paragraph.Descendants(W.t).Select(t => t.Value)));
            var tokens = paragraph.Elements(W.r).SelectMany(r => r.Elements()).Where(e => e.Name == W.tab || e.Name == W.t).ToList();
            Assert.Equal(new[] { W.tab, W.t, W.tab }, tokens.Select(e => e.Name));
            Assert.Single(paragraph.Descendants(WP.anchor));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void InlineConditionalAndEmptyBlockContentRetainTheirParagraphs(bool condition)
        {
            // Reduced from the consultant's production check: no spaces outside
            // the inline pair, and no whitespace beside the empty content field.
            var properties = new XElement(W.pPr,
                new XElement(W.spacing, new XAttribute(W.after, 160)));
            var source = Document(
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Before"))),
                new XElement(W.p, new XElement(properties), Field("if Ready"),
                    new XElement(W.r, new XElement(W.t, "Conditional text")), Field("endif")),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Middle"))),
                new XElement(W.p, new XElement(properties), Field("BlankFormula")),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "After"))));
            var normalized = Normalizer.NormalizeTemplate(source, false);
            Assert.Empty(normalized.Diagnostics);
            var extracted = JsonNode.Parse(normalized.ExtractedFields)[0].AsArray();
            Assert.Equal(2, extracted.Count);
            Assert.Equal(2, Assert.IsType<JsonArray>(extracted[0]).Count);
            Assert.IsType<JsonObject>(extracted[1]);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var conditional = fields.Values.Single(f => f.Type == FieldType.If);
            var content = fields.Values.Single(f => f.Type == FieldType.Content);
            var data = new XElement("_odx",
                new XElement(conditional.Atom + "b", condition ? "true" : "false"),
                new XElement(content.Atom, ""));
            var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
            Assert.False(assembled.HasErrors, assembled.Error);
            var paragraphs = Root(assembled.Bytes).Element(W.body).Elements(W.p).ToList();
            Assert.Equal(new[] { "Before", condition ? "Conditional text" : "", "Middle", "", "After" },
                paragraphs.Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
            foreach (var index in new[] { 1, 3 })
                Assert.True(XNode.DeepEquals(properties, WithoutNamespaceDeclarations(paragraphs[index].Element(W.pPr))));
        }

        [Theory]
        [InlineData("ifValue")]
        [InlineData("elseifValue")]
        [InlineData("elseifTakingTitle == \"Unequal\"")]
        [InlineData("elseValue")]
        [InlineData("endifValue")]
        [InlineData("listValue")]
        [InlineData("endlistValue")]
        [InlineData("IF Ready")]
        [InlineData("If")]
        [InlineData("if_Ready")]
        public void KeywordLikeContentIsStillContent(string expression)
        {
            foreach (var controls in new[] { false, true })
            {
                var source = Document(new XElement(W.p, MarkerWhitespace("tab"), Field(expression, controls)));
                var normalized = Normalizer.NormalizeTemplate(source, false, contentControls: controls);
                Assert.IsType<JsonArray>(JsonNode.Parse(normalized.ExtractedFields)[0][0]);
                var field = Assert.Single(Templater.ParseFieldsToDict(normalized.ExtractedFields)).Value;
                Assert.Equal(FieldType.Content, field.Type);
                Assert.Equal(expression, field.Expression);
            }
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void ControlKeywordsAndShorthandRetainMatchingAndExpressions(bool shorthand, bool controls)
        {
            var markers = shorthand ? new[] { "? Ready", ":? Other", ": comment", "/? comment", "# Items", "/# comment" }
                : new[] { "if Ready", "elseif Other", "else comment", "endif comment", "list Items", "endlist comment" };
            var content = new List<XElement>();
            for (var i = 0; i < markers.Length; i++)
            {
                content.Add(new XElement(W.p, MarkerWhitespace("tab"), Field(markers[i], controls), MarkerWhitespace("nbsp")));
                if (i == 0 || i == 1 || i == 2 || i == 4)
                    content.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Branch" + i))));
            }
            var normalized = Normalizer.NormalizeTemplate(Document(content.ToArray()), false, contentControls: controls);
            Assert.All(JsonNode.Parse(normalized.ExtractedFields)[0].AsArray(), f => Assert.IsType<JsonObject>(f));
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            Assert.Equal(new[] { FieldType.If, FieldType.ElseIf, FieldType.Else, FieldType.EndIf, FieldType.List, FieldType.EndList },
                Enumerable.Range(1, 6).Select(i => fields[i.ToString()].Type));
            Assert.Equal("Ready", fields["1"].Expression);
            Assert.Equal("Other", fields["2"].Expression);
            Assert.Equal("Items", fields["5"].Expression);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var list = fields["5"].Atom;
            var data = new XElement("_odx", new XElement(fields["1"].Atom + "b", "false"),
                new XElement(fields["2"].Atom + "b", "true"), new XElement(list, new XElement(list + "i")));
            var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
            Assert.False(assembled.HasErrors, assembled.Error);
            Assert.Equal("Branch1Branch4", Root(assembled.Bytes).Value);
        }
    }
}
