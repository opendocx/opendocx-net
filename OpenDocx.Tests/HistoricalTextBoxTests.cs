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
        private static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        private static XElement FloatingTextBox(string representation, XElement content)
        {
            var drawing = Drawing(true);
            drawing.Descendants(WP.docPr).Single().SetAttributeValue("name", "Synthetic text box");
            drawing.Descendants(A + "graphicData").Single().ReplaceWith(
                new XElement(A + "graphicData", new XAttribute("uri", Wps.NamespaceName),
                    new XElement(Wps + "wsp",
                        new XElement(Wps + "cNvSpPr", new XAttribute("txBox", 1)),
                        new XElement(Wps + "spPr",
                            new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                new XElement(A + "ext", new XAttribute("cx", 914400), new XAttribute("cy", 914400))),
                            new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))),
                        new XElement(Wps + "txbx", new XElement(content)),
                        new XElement(Wps + "bodyPr", new XElement(A + "noAutofit")))));
            var pict = new XElement(W.pict,
                new XElement(Vml + "shapetype", new XAttribute("id", "_x0000_t202"),
                    new XAttribute("coordsize", "21600,21600"), new XAttribute("path", "m,l,21600r21600,l21600,xe")),
                new XElement(Vml + "shape", new XAttribute("id", "SyntheticTextBox"), new XAttribute("type", "#_x0000_t202"),
                    new XAttribute("style", "position:absolute;margin-left:0;margin-top:0;width:72pt;height:72pt"),
                    new XElement(Vml + "textbox", new XElement(content))));
            if (representation == "drawing") return drawing;
            if (representation == "vml") return pict;
            return new XElement(Mc + "AlternateContent", new XAttribute(XNamespace.Xmlns + "wps", Wps.NamespaceName),
                new XElement(Mc + "Choice", new XAttribute("Requires", "wps"), drawing),
                new XElement(Mc + "Fallback", pict));
        }

        private static XElement TextBoxParagraph(bool nestedField, string representation, XElement field, XElement boxField)
        {
            var box = new XElement(W.txbxContent,
                new XElement(W.p, new XElement(W.pPr, new XElement(W.jc, new XAttribute(W.val, "right"))),
                    new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), "Box label: ")),
                    nestedField ? boxField : new XElement(W.r, new XElement(W.t, "Fixed text"))));
            return new XElement(W.p,
                new XElement(W.pPr, new XElement(W.keepNext),
                    new XElement(W.tabs, new XElement(W.tab, new XAttribute(W.val, "right"), new XAttribute(W.pos, 7200))),
                    new XElement(W.ind, new XAttribute(W.right, -720))),
                nestedField ? new[] { MarkerWhitespace("space") }
                    : Enumerable.Range(0, 8).Select(_ => MarkerWhitespace("tab")),
                field, nestedField ? MarkerWhitespace("tab") : null,
                new XElement(W.r, FloatingTextBox(representation, box)));
        }

        public static IEnumerable<object[]> HistoricalTextBoxCases()
        {
            foreach (var nestedField in new[] { false, true })
                foreach (var representation in new[] { "drawing", "vml", "alternate" })
                    foreach (var value in new[] { null, "", "Replacement" })
                        yield return new object[] { nestedField, representation, value };
        }

        [Theory]
        [MemberData(nameof(HistoricalTextBoxCases))]
        public async Task HistoricalCompiledTextBoxesSurviveOuterContentReplacement(bool nestedField, string representation, string value)
        {
            // AlternateContent reproduces followup/float-09 and float-10/11. Standalone
            // DrawingML/VML rows protect both representations independently.
            var original = TextBoxParagraph(nestedField, representation, HistoricalContent(), HistoricalContent("C2"));
            var compiled = Document(original);
            var data = new XElement("_odx", new XElement("C2", "Inside"));
            if (value != null) data.Add(new XElement("C1", value));
            var sync = Assembler.AssembleDocument(compiled, data);
            var asyncResult = await Assembler.AssembleDocumentAsync(compiled, data.ToString(), null);
            foreach (var result in new[] { sync, asyncResult })
            {
                Assert.False(result.HasErrors, result.Error);
                AssertTextBoxOutput(result.Bytes, original, nestedField, representation, value);
            }
        }

        private static void AssertTextBoxOutput(byte[] bytes, XElement original, bool nestedField, string representation, string value)
        {
            var paragraph = Assert.Single(Root(bytes).Element(W.body).Elements(W.p));
            Assert.True(XNode.DeepEquals(original.Element(W.pPr), WithoutNamespaceDeclarations(paragraph.Element(W.pPr))));
            Assert.Equal((nestedField ? " " : "") + (value ?? ""), string.Concat(
                paragraph.Descendants(W.t).Where(t => !t.Ancestors(W.txbxContent).Any()).Select(t => t.Value)));
            Assert.Equal(nestedField ? 1 : 8, paragraph.Elements(W.r).Elements(W.tab).Count());
            Assert.Equal(representation == "vml" ? 0 : 1, paragraph.Descendants(WP.anchor).Count());
            Assert.Equal(representation == "drawing" ? 0 : 1, paragraph.Descendants(Vml + "shape").Count());
            var boxes = paragraph.Descendants(W.txbxContent).ToList();
            Assert.Equal(representation == "alternate" ? 2 : 1, boxes.Count);
            Assert.All(boxes, box =>
            {
                var boxParagraph = Assert.Single(box.Elements(W.p));
                Assert.Equal("Box label: " + (nestedField ? "Inside" : "Fixed text"), boxParagraph.Value);
                Assert.Equal("right", (string)boxParagraph.Element(W.pPr).Element(W.jc).Attribute(W.val));
                Assert.Empty(box.Descendants(W.sdt));
            });
            // Check positioning/shape properties as well as existence. Field contents
            // are intentionally variable and have their own assertions above.
            XElement WithoutBoxContents(XElement p)
            {
                var copy = WithoutNamespaceDeclarations(p);
                foreach (var box in copy.Descendants(W.txbxContent)) box.RemoveNodes();
                return copy;
            }
            var originalObject = original.Elements(W.r).Last();
            var assembledObject = paragraph.Elements(W.r).Single(r => r.Descendants(W.txbxContent).Any());
            Assert.True(XNode.DeepEquals(WithoutBoxContents(originalObject), WithoutBoxContents(assembledObject)));
        }

        [Theory]
        [MemberData(nameof(HistoricalTextBoxCases))]
        public async Task ProvisionedTextBoxesMatchHistoricalAssembly(bool nestedField, string representation, string value)
        {
            foreach (var controls in new[] { false, true })
            {
                var original = TextBoxParagraph(nestedField, representation, Field("Value", controls), Field("BoxValue", controls));
                var source = Document(original);
                var normalized = Normalizer.NormalizeTemplate(source, false, contentControls: controls);
                Assert.Empty(normalized.Diagnostics);
                Assert.Equal(normalized.ExtractedFields, Normalizer.ExtractFieldsOnly(source, contentControls: controls));
                Assert.IsType<JsonArray>(JsonNode.Parse(normalized.ExtractedFields)[0][0]);
                var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
                Assert.Equal(nestedField ? (representation == "alternate" ? 3 : 2) : 1, fields.Count);
                var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
                Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
                var data = new XElement("_odx");
                if (value != null) data.Add(new XElement(fields["1"].Atom, value));
                if (nestedField) data.Add(new XElement(fields["2"].Atom, "Inside"));
                var assembled = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
                Assert.False(assembled.HasErrors, assembled.Error);
                AssertTextBoxOutput(assembled.Bytes, original, nestedField, representation, value);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void NestedBoxFieldsDoNotHideUnsupportedOuterAnchors(bool controls)
        {
            var paragraph = TextBoxParagraph(true, "alternate", Field("Value", controls), Field("BoxValue", controls));
            paragraph.Elements(W.r).Where(r => r.Elements(W.tab).Any() || r.Elements(W.t).Any()).Remove();
            // The text-delimited field is itself a run; put it back before the object.
            if (!controls) paragraph.Elements(W.r).Single().AddBeforeSelf(Field("Value"));
            var source = Document(paragraph);
            var analysis = Normalizer.AnalyzeTemplate(source, contentControls: controls);
            var diagnostic = Assert.Single(analysis.Diagnostics);
            Assert.Equal("Value", diagnostic.FieldText);
            Assert.Equal("FloatingImageOnBlockField", diagnostic.Code);
            Assert.Throws<FieldParseException>(() => Normalizer.NormalizeTemplate(source, false, contentControls: controls));
        }
    }
}
