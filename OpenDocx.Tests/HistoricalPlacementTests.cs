using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenDocx;
using OpenXmlPowerTools;
using Normalizer = OpenDocx.Normalizer;
using Xunit;

namespace OpenDocxTemplater.Tests
{
    public partial class ImageParagraphTests
    {
        // Reduced structures from the stored historical oxpt corpus; no customer content.
        // See HistoricalPlacementTests.md for test contracts, coverage, and limitations.
        private static XElement HistoricalContent(string atom = "C1") => new XElement(W.sdt,
            new XElement(W.sdtContent, new XElement(W.r,
                new XElement(W.t, "<Content Select=\"./" + atom + "[1]\" Optional=\"true\"/>"))));

        private static XElement WithoutNamespaceDeclarations(XElement element)
        {
            var copy = new XElement(element);
            copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            return copy;
        }

        private static XElement HistoricalParagraph(string placement, XElement field)
        {
            switch (placement)
            {
                case "header-tabs-anchor":
                    return new XElement(W.p,
                        new XElement(W.pPr,
                            new XElement(W.shd, new XAttribute(W.val, "clear"), new XAttribute(W.color, "auto"), new XAttribute(W.fill, "FFFFFF")),
                            new XElement(W.ind, new XAttribute(W.right, -720)),
                            new XElement(W.rPr, new XElement(W.i), new XElement(W.sz, new XAttribute(W.val, 22)))),
                        Enumerable.Range(0, 8).Select(_ => new XElement(W.r, new XElement(W.tab))),
                        field, new XElement(W.r, Drawing(true)));
                case "page-break":
                    return new XElement(W.p,
                        new XElement(W.pPr, new XElement(W.keepNext), new XElement(W.keepLines),
                            new XElement(W.tabs, new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos, 720))),
                            new XElement(W.jc, new XAttribute(W.val, "center"))),
                        new XElement(W.r, new XElement(W.br, new XAttribute(W.type, "page"))), field);
                case "leading-space":
                    return new XElement(W.p,
                        new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, "TextHeading2")), new XElement(W.keepNext)),
                        new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " ")), field);
                default:
                    throw new ArgumentOutOfRangeException(nameof(placement));
            }
        }

        private static byte[] HistoricalDocument(string placement, XElement field)
        {
            var paragraph = HistoricalParagraph(placement, field);
            if (placement != "header-tabs-anchor") return Document(paragraph);
            using var stream = new MemoryStream();
            stream.Write(Document(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body")))));
            using (var document = WordprocessingDocument.Open(stream, true))
            {
                var main = document.MainDocumentPart;
                var header = main.AddNewPart<HeaderPart>();
                header.AddPart(main.ImageParts.Single(), "image");
                header.PutXDocument(new XDocument(new XElement(W.hdr, paragraph)));
                main.GetXDocument().Root.Element(W.body).Add(new XElement(W.sectPr,
                    new XElement(W.headerReference, new XAttribute(W.type, "default"), new XAttribute(R + "id", main.GetIdOfPart(header)))));
                main.PutXDocument();
            }
            return stream.ToArray();
        }

        public static IEnumerable<object[]> HistoricalContentCases()
        {
            foreach (var placement in new[] { "header-tabs-anchor", "page-break", "leading-space" })
                foreach (var value in new[] { null, "", "Replacement" })
                    yield return new object[] { placement, value };
        }

        [Theory]
        [MemberData(nameof(HistoricalContentCases))]
        public async Task HistoricalCompiledContentPreservesItsParagraph(string placement, string value)
        {
            var compiled = HistoricalDocument(placement, HistoricalContent());
            var data = new XElement("_odx");
            if (value != null) data.Add(new XElement("C1", value));
            var sync = Assembler.AssembleDocument(compiled, data);
            var asyncResult = await Assembler.AssembleDocumentAsync(compiled, data.ToString(), null);
            foreach (var result in new[] { sync, asyncResult })
            {
                Assert.False(result.HasErrors, result.Error);
                AssertHistoricalContentResult(result.Bytes, placement, value);
            }
        }

        private static void AssertHistoricalContentResult(byte[] bytes, string placement, string value)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            OpenXmlPart part = placement == "header-tabs-anchor"
                ? document.MainDocumentPart.HeaderParts.Single() : document.MainDocumentPart;
            var root = part.GetXDocument().Root;
            var paragraph = Assert.Single(root.Descendants(W.p));
            var expected = HistoricalParagraph(placement, new XElement(W.r));
            Assert.True(XNode.DeepEquals(expected.Element(W.pPr), WithoutNamespaceDeclarations(paragraph.Element(W.pPr))), paragraph.ToString());
            Assert.Equal((placement == "leading-space" ? " " : "") + (value ?? ""),
                string.Concat(paragraph.Descendants(W.t).Select(t => t.Value)));
            // Paragraph tab-stop definitions are distinct from authored inline tab characters.
            Assert.Equal(placement == "header-tabs-anchor" ? 8 : 0, paragraph.Descendants(W.r).Elements(W.tab).Count());
            Assert.Equal(placement == "page-break" ? 1 : 0, paragraph.Descendants(W.br).Count());
            if (placement == "page-break")
                Assert.Equal("page", (string)Assert.Single(paragraph.Descendants(W.br)).Attribute(W.type));
            Assert.Equal(placement == "header-tabs-anchor" ? 1 : 0, paragraph.Descendants(WP.anchor).Count());
            if (placement == "header-tabs-anchor")
            {
                var embed = (string)Assert.Single(paragraph.Descendants(A + "blip")).Attribute(R + "embed");
                var image = Assert.IsType<ImagePart>(part.GetPartById(embed));
                using var imageStream = image.GetStream();
                Assert.True(imageStream.Length > 0);
                Assert.Equal("Body", Root(bytes).Element(W.body).Element(W.p).Value);
            }
        }

        [Theory]
        [InlineData("header-tabs-anchor")]
        [InlineData("page-break")]
        [InlineData("leading-space")]
        public async Task ProvisionedWhitespaceContentMatchesHistoricalAssembly(string placement)
        {
            foreach (var controls in new[] { false, true })
            {
                var source = HistoricalDocument(placement, Field("Value", controls));
                var normalized = Normalizer.NormalizeTemplate(source, false, contentControls: controls);
                Assert.Empty(normalized.Diagnostics);
                Assert.Equal(normalized.ExtractedFields, Normalizer.ExtractFieldsOnly(source, contentControls: controls));
                var part = JsonNode.Parse(normalized.ExtractedFields).AsArray().Single(p => p.AsArray().Count > 0);
                Assert.IsType<JsonArray>(Assert.Single(part.AsArray()));
                var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
                var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
                Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
                foreach (var value in new[] { null, "", "Replacement" })
                {
                    var data = new XElement("_odx");
                    if (value != null) data.Add(new XElement(fields["1"].Atom, value));
                    var assembled = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
                    Assert.False(assembled.HasErrors, assembled.Error);
                    AssertHistoricalContentResult(assembled.Bytes, placement, value);
                }
            }
        }

        private static XElement MarkerWhitespace(string kind) => new XElement(W.r,
            kind == "tab" ? new XElement(W.tab) :
            kind == "break" ? new XElement(W.br) :
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), kind == "nbsp" ? "\u00a0" : " "));

        public static IEnumerable<object[]> WhitespaceMarkerCases()
        {
            foreach (var controls in new[] { false, true })
                foreach (var list in new[] { false, true })
                    foreach (var whitespace in new[] { "tab", "space", "break", "nbsp" })
                        foreach (var before in new[] { false, true })
                            yield return new object[] { controls, list, whitespace, before };
        }

        [Theory]
        [MemberData(nameof(WhitespaceMarkerCases))]
        public async Task WhitespaceBesidePairedMarkersKeepsCrossParagraphMatching(bool controls, bool list, string whitespace, bool before)
        {
            // Tab cases reproduce historical cases 06/07; other rows protect the agreed compatibility rule.
            var marker = Field(list ? "list Items" : "if Ready", controls);
            var spacing = MarkerWhitespace(whitespace);
            var properties = new XElement(W.pPr, new XElement(W.jc, new XAttribute(W.val, "center")));
            var source = Document(new XElement(W.p, before ? spacing : marker, before ? marker : spacing),
                new XElement(W.p, properties, new XElement(W.r, new XElement(W.t, "Body"))),
                new XElement(W.p, Field(list ? "endlist" : "endif", controls)));
            var normalized = Normalizer.NormalizeTemplate(source, false, contentControls: controls);
            Assert.Empty(normalized.Diagnostics);
            var extracted = JsonNode.Parse(normalized.ExtractedFields)[0].AsArray();
            Assert.Equal(2, extracted.Count);
            Assert.All(extracted, field => Assert.IsType<JsonObject>(field));
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var atom = fields["1"].Atom;
            foreach (var count in list ? new[] { 0, 1, 2 } : new[] { 0, 1 })
            {
                var data = new XElement("_odx", list
                    ? new XElement(atom, Enumerable.Range(0, count).Select(_ => new XElement(atom + "i")))
                    : new XElement(atom + "b", count == 1 ? "true" : "false"));
                var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
                Assert.False(result.HasErrors, result.Error);
                var root = Root(result.Bytes);
                Assert.Equal(string.Concat(Enumerable.Repeat("Body", count)), root.Value);
                Assert.Empty(root.Descendants(W.r).Elements(W.tab));
                Assert.Empty(root.Descendants(W.br));
                var bodyParagraphs = root.Descendants(W.p).Where(p => p.Value == "Body").ToList();
                Assert.Equal(count, bodyParagraphs.Count);
                Assert.All(bodyParagraphs, p => Assert.True(XNode.DeepEquals(properties, WithoutNamespaceDeclarations(p.Element(W.pPr)))));
            }
        }
    }
}
