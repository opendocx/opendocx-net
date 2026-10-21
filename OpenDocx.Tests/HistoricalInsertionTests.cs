using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Xml.Linq;
using OpenDocx;
using OpenXmlPowerTools;
using Normalizer = OpenDocx.Normalizer;
using Xunit;

namespace OpenDocxTemplater.Tests
{
    public partial class ImageParagraphTests
    {
        private static XElement InsertionTable(XElement cellContent) => new XElement(W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W.w + "w", 2400))),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.tcPr), new XElement(W.p, cellContent))));

        [Theory]
        [InlineData("empty-text")]
        [InlineData("empty-document")]
        [InlineData("empty-assembled-child")]
        [InlineData("blank-paragraph")]
        public async Task EmptyDocumentInsertionReplacesHostParagraphUnlikeEmptyText(string value)
        {
            var properties = new XElement(W.pPr, new XElement(W.ind, new XAttribute(W.left, 720)));
            var normalized = Normalizer.NormalizeTemplate(Document(
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Before"))),
                new XElement(W.p, properties, Field("Child")),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "After")))), false);
            Assert.Empty(normalized.Diagnostics);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var child = value == "blank-paragraph" ? Document(new XElement(W.p)) : Document();
            if (value == "empty-assembled-child")
            {
                var childSource = Document(new XElement(W.p, Field("if Include")),
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "Optional child"))),
                    new XElement(W.p, Field("endif")));
                var childNormalized = Normalizer.NormalizeTemplate(childSource, false);
                Assert.Empty(childNormalized.Diagnostics);
                var childFields = Templater.ParseFieldsToDict(childNormalized.ExtractedFields);
                var childCompiled = Templater.CompileTemplate(childNormalized.NormalizedTemplate, childFields);
                Assert.False(childCompiled.HasErrors, string.Join("\n", childCompiled.Errors));
                var childResult = Assembler.AssembleDocument(childCompiled.Bytes,
                    new XElement("_odx", new XElement(childFields["1"].Atom + "b", "false")));
                Assert.False(childResult.HasErrors, childResult.Error);
                Assert.Empty(Root(childResult.Bytes).Element(W.body).Elements(W.p));
                child = childResult.Bytes;
            }
            var data = new XElement("_odx", new XElement(fields["1"].Atom,
                value == "empty-text" ? "" : "oxpt://DocumentAssembler/insert/child"));
            var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(),
                value == "empty-text" ? null : new[] { new IndirectSource("child", child) });
            Assert.False(result.HasErrors, result.Error);
            var body = Root(result.Bytes).Element(W.body);
            var paragraphs = body.Elements(W.p).ToList();
            Assert.Empty(body.Elements(W.tbl));
            var expectBlankParagraph = value == "empty-text" || value == "blank-paragraph";
            Assert.Equal(expectBlankParagraph ? new[] { "Before", "", "After" } : new[] { "Before", "After" },
                paragraphs.Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
            if (value == "empty-text")
                Assert.True(XNode.DeepEquals(properties, WithoutNamespaceDeclarations(paragraphs[1].Element(W.pPr))));
            else
                Assert.Empty(body.Descendants(W.ind));
        }

        public static IEnumerable<object[]> HistoricalInsertionCases()
        {
            foreach (var provisioning in new[] { "stored", "text", "controls" })
                foreach (var value in new[] { "missing", "empty", "scalar", "blocks", "assembled-child" })
                    yield return new object[] { provisioning, value };
        }

        public static IEnumerable<object[]> TrailingWhitespaceInsertionCases()
        {
            foreach (var whitespace in new[] { "space", "tab", "text-tab" })
                foreach (var anchor in new[] { false, true })
                    foreach (var child in new[] { "empty", "blank-paragraph", "blocks" })
                        yield return new object[] { whitespace, anchor, child };
        }

        [Theory]
        [MemberData(nameof(TrailingWhitespaceInsertionCases))]
        public async Task TrailingWhitespaceInsertionMatchesExistingRunLevelAssembly(string whitespace, bool anchor, string childKind)
        {
            // Characterize today's insertion behavior; future paragraph joining is
            // deliberately outside this provisioning fix.
            var properties = new XElement(W.pPr,
                new XElement(W.tabs, new XElement(W.tab, new XAttribute(W.val, "right"), new XAttribute(W.pos, 7200))),
                new XElement(W.ind, new XAttribute(W.left, 720)));
            byte[] Parent(XElement field) => Document(
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Before"))),
                new XElement(W.p, new XElement(properties), field, PlacementWhitespace(whitespace),
                    anchor ? new XElement(W.r, Drawing(true)) : null),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "After"))));
            var normalized = Normalizer.NormalizeTemplate(Parent(Field("Child")), false);
            Assert.Empty(normalized.Diagnostics);
            Assert.IsType<JsonArray>(JsonNode.Parse(normalized.ExtractedFields)[0][0]);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var child = childKind == "empty" ? Document()
                : childKind == "blank-paragraph" ? Document(new XElement(W.p))
                : Document(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Inserted paragraph"))),
                    InsertionTable(new XElement(W.r, new XElement(W.t, "Inserted cell"))));
            var insert = new[] { new IndirectSource("child", child) };
            XElement Data(string atom) => new XElement("_odx", new XElement(atom, "oxpt://DocumentAssembler/insert/child"));
            var existing = await Assembler.AssembleDocumentAsync(Parent(HistoricalContent()), Data("C1").ToString(), insert);
            var candidate = await Assembler.AssembleDocumentAsync(compiled.Bytes, Data(fields["1"].Atom).ToString(), insert);
            var expectedText = new List<string> { "Before" };
            if (childKind == "blocks") expectedText.AddRange(new[] { "Inserted paragraph", "Inserted cell" });
            if (childKind == "blank-paragraph") expectedText.Add("");
            expectedText.Add(PlacementText(whitespace));
            expectedText.Add("After");
            foreach (var result in new[] { existing, candidate })
            {
                Assert.False(result.HasErrors, result.Error);
                var body = Root(result.Bytes).Element(W.body);
                var blocks = body.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).ToList();
                Assert.Equal(expectedText, blocks.Select(b => string.Concat(b.Descendants(W.t).Select(t => t.Value))));
                Assert.Equal(childKind == "blocks" ? 1 : 0, body.Descendants(W.tbl).Count());
                var host = blocks[blocks.Count - 2];
                Assert.Equal(W.p, host.Name);
                Assert.True(XNode.DeepEquals(properties, WithoutNamespaceDeclarations(host.Element(W.pPr))));
                Assert.Equal(whitespace == "tab" ? 1 : 0, host.Elements(W.r).Elements(W.tab).Count());
                Assert.Equal(anchor ? 1 : 0, body.Descendants(WP.anchor).Count());
                if (anchor)
                    Assert.True(XNode.DeepEquals(WithoutNamespaceDeclarations(Drawing(true).Element(WP.anchor).Element(WP.positionH)),
                        WithoutNamespaceDeclarations(Assert.Single(host.Descendants(WP.anchor)).Element(WP.positionH))));
                Assert.Empty(body.Descendants(W.sdt));
            }
        }

        [Theory]
        [MemberData(nameof(HistoricalInsertionCases))]
        public async Task HistoricalPageBreakSurvivesContentAndDocumentInsertion(string provisioning, string value)
        {
            // Reduced from historical case 04 and followup/insert-child. Source paths
            // must preserve the same assembly contract as the stored run-level metadata.
            var paragraph = HistoricalParagraph("page-break", provisioning == "stored"
                ? HistoricalContent() : Field("Value", provisioning == "controls"));
            var compiled = Document(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Before"))),
                paragraph, new XElement(W.p, new XElement(W.r, new XElement(W.t, "After"))));
            var atom = "C1";
            if (provisioning != "stored")
            {
                var normalized = Normalizer.NormalizeTemplate(compiled, false, contentControls: provisioning == "controls");
                Assert.Empty(normalized.Diagnostics);
                var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
                var result = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
                Assert.False(result.HasErrors, string.Join("\n", result.Errors));
                compiled = result.Bytes;
                atom = fields["1"].Atom;
            }

            var insertion = value == "blocks" || value == "assembled-child";
            var data = new XElement("_odx");
            if (value != "missing") data.Add(new XElement(atom,
                insertion ? "oxpt://DocumentAssembler/insert/insert" : value == "scalar" ? "Replacement" : ""));
            byte[] insert;
            if (value == "assembled-child")
            {
                var child = Document(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Certificate"))),
                    InsertionTable(HistoricalContent("C2")),
                    new XElement(W.p, HistoricalContent()));
                var assembledChild = Assembler.AssembleDocument(child,
                    new XElement("_odx", new XElement("C1", "Child ending"), new XElement("C2", "Child cell")));
                Assert.False(assembledChild.HasErrors, assembledChild.Error);
                insert = assembledChild.Bytes;
            }
            else insert = Document(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Inserted paragraph"))),
                InsertionTable(new XElement(W.r, new XElement(W.t, "Inserted cell"))));

            var assembled = await Assembler.AssembleDocumentAsync(compiled, data.ToString(),
                insertion ? new[] { new IndirectSource("insert", insert) } : null);
            Assert.False(assembled.HasErrors, assembled.Error);
            var root = Root(assembled.Bytes);
            var body = root.Element(W.body);
            var blocks = body.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).ToList();
            var expectedText = value == "blocks" ? new[] { "Before", "", "Inserted paragraph", "Inserted cell", "After" }
                : value == "assembled-child" ? new[] { "Before", "", "Certificate", "Child cell", "Child ending", "After" }
                : new[] { "Before", value == "scalar" ? "Replacement" : "", "After" };
            Assert.Equal(expectedText, blocks.Select(b => string.Concat(b.Descendants(W.t).Select(t => t.Value))).ToArray());
            var breakParagraph = blocks[1];
            Assert.Equal(W.p, breakParagraph.Name);
            Assert.True(XNode.DeepEquals(paragraph.Element(W.pPr), WithoutNamespaceDeclarations(breakParagraph.Element(W.pPr))));
            var pageBreak = Assert.Single(root.Descendants(W.br));
            Assert.Equal("page", (string)pageBreak.Attribute(W.type));
            Assert.Same(breakParagraph, pageBreak.Ancestors(W.p).Single());
            Assert.Equal(value == "assembled-child" ? 6 : value == "blocks" ? 5 : 3, root.Descendants(W.p).Count());
            Assert.Equal(insertion ? 1 : 0, root.Descendants(W.tbl).Count());
            if (insertion)
            {
                Assert.Equal(W.tbl, blocks[3].Name);
                Assert.Single(blocks[3].Elements(W.tr));
                Assert.Single(blocks[3].Descendants(W.tc));
            }
            Assert.Empty(root.Descendants(W.sdt));
        }
    }
}
