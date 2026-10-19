using System;
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
    public class ImageParagraphTests
    {
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static XElement Field(string text, bool controls = false) => controls
            ? new XElement(W.sdt, new XElement(W.sdtPr),
                new XElement(W.sdtContent, new XElement(W.r, new XElement(W.t, "[" + text + "]"))))
            : new XElement(W.r, new XElement(W.t, "{[" + text + "]}"));

        private static XElement Drawing(bool floating)
        {
            var picture = new XElement(A + "graphic",
                new XElement(A + "graphicData", new XAttribute("uri", Pic.NamespaceName),
                    new XElement(Pic + "pic",
                        new XElement(Pic + "nvPicPr", new XElement(Pic + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "Picture")), new XElement(Pic + "cNvPicPr")),
                        new XElement(Pic + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", "image")), new XElement(A + "stretch", new XElement(A + "fillRect"))),
                        new XElement(Pic + "spPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                            new XElement(A + "ext", new XAttribute("cx", 914400), new XAttribute("cy", 914400))),
                            new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))))));
            var container = new XElement(floating ? WP.anchor : WP.inline,
                new XElement(WP.extent, new XAttribute("cx", 914400), new XAttribute("cy", 914400)),
                new XElement(WP.docPr, new XAttribute("id", 1), new XAttribute("name", "Picture")), picture);
            if (floating)
            {
                container.Add(new XAttribute("simplePos", 0), new XAttribute("relativeHeight", 0),
                    new XAttribute("behindDoc", 0), new XAttribute("locked", 0), new XAttribute("layoutInCell", 1), new XAttribute("allowOverlap", 1));
                container.AddFirst(new XElement(WP.simplePos, new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(WP.positionH, new XAttribute("relativeFrom", "column"), new XElement(WP.posOffset, 0)),
                    new XElement(WP.positionV, new XAttribute("relativeFrom", "paragraph"), new XElement(WP.posOffset, 0)));
                container.Element(WP.extent).AddAfterSelf(new XElement(WP.wrapNone));
            }
            return new XElement(W.drawing, container);
        }

        private static byte[] Document(params XElement[] content)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var image = main.AddImagePart(ImagePartType.Png, "image");
                using var png = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2tVUAAAAASUVORK5CYII="));
                image.FeedData(png);
                main.PutXDocument(new XDocument(new XElement(W.document, new XElement(W.body, content))));
            }
            return stream.ToArray();
        }

        private static XElement Root(byte[] bytes)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            return new XElement(document.MainDocumentPart.GetXDocument().Root);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void InlineImagesCountOnEitherSideInBothFieldModes(bool before, bool controls)
        {
            var image = new XElement(W.r, Drawing(false));
            var field = Field("Value", controls);
            var bytes = Document(new XElement(W.p, before ? image : field, before ? field : image));
            var analysis = Normalizer.AnalyzeTemplate(bytes, contentControls: controls);
            Assert.Empty(analysis.Diagnostics);
            Assert.Equal(analysis.ExtractedFields, Normalizer.ExtractFieldsOnly(bytes, contentControls: controls));
            var part = JsonNode.Parse(analysis.ExtractedFields)[0].AsArray();
            Assert.IsType<JsonArray>(Assert.Single(part));
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, Templater.ParseFieldsToDict(analysis.ExtractedFields));
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            Assert.Single(Root(compiled.Bytes).Descendants(WP.inline));
        }

        [Theory]
        [InlineData("if Ready", false)]
        [InlineData("if Ready", true)]
        [InlineData("endif", false)]
        [InlineData("elseif Other", true)]
        [InlineData("else", false)]
        [InlineData("list Items", true)]
        [InlineData("endlist", false)]
        [InlineData("Value", true)]
        public void FloatingImagesDiagnoseAllBlockFieldKindsRegardlessOfPosition(string text, bool before)
        {
            foreach (var controls in new[] { false, true })
            {
                var image = new XElement(W.r, Drawing(true));
                var field = Field(text, controls);
                var bytes = Document(new XElement(W.p, before ? image : field, before ? field : image));
                var analysis = Normalizer.AnalyzeTemplate(bytes, contentControls: controls);
                var diagnostic = Assert.Single(analysis.Diagnostics);
                Assert.Equal("FloatingImageOnBlockField", diagnostic.Code);
                Assert.Equal("1", diagnostic.FieldId);
                Assert.Equal(text, diagnostic.FieldText);
                Assert.Equal("/word/document.xml", diagnostic.PartUri);
                var part = JsonNode.Parse(analysis.ExtractedFields)[0].AsArray();
                Assert.IsType<JsonObject>(Assert.Single(part));
                Assert.Single(Root(analysis.NormalizedTemplate).Descendants(WP.anchor));
                var error = Assert.Throws<FieldParseException>(() => Normalizer.NormalizeTemplate(bytes, false, contentControls: controls));
                Assert.Contains("floating object (such as a picture, text box, or shape)", error.Message);
                Assert.Contains(text, error.Message);
                var extractionError = Assert.Throws<FieldParseException>(() => Normalizer.ExtractFieldsOnly(bytes, contentControls: controls));
                Assert.Equal(error.Message, extractionError.Message);
            }
        }

        [Fact]
        public void AnchorsInMixedRunsDoNotShiftSplitFieldBoundaries()
        {
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(true),
                new XElement(W.t, "{[if Ready]}"))),
                new XElement(W.p, Field("Value")), new XElement(W.p, Field("endif")));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Single(analysis.Diagnostics);
            Assert.Equal(3, Templater.ParseFieldsToDict(analysis.ExtractedFields).Count);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FloatingImagesAreAllowedWithRealParagraphContent(bool inline)
        {
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(true)), Field("Value"),
                inline ? new XElement(W.r, Drawing(false)) : new XElement(W.r, new XElement(W.t, " text"))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Empty(analysis.Diagnostics);
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, Templater.ParseFieldsToDict(analysis.ExtractedFields));
            Assert.False(compiled.HasErrors);
            Assert.Single(Root(compiled.Bytes).Descendants(WP.anchor));
        }

        [Fact]
        public void BlockContentControlWithAnAnchorIsDiagnosed()
        {
            var bytes = Document(new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent,
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "[Value]")), new XElement(W.r, Drawing(true))))));
            Assert.Single(Normalizer.AnalyzeTemplate(bytes, contentControls: true).Diagnostics);
        }

        [Fact]
        public void AFieldControlWrappingATableCellAlsoDiagnosesAnchors()
        {
            var bytes = Document(new XElement(W.tbl, new XElement(W.tr,
                new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent,
                    new XElement(W.tc, new XElement(W.p,
                        new XElement(W.r, new XElement(W.t, "[Value]")), new XElement(W.r, Drawing(true)))))))));
            Assert.Single(Normalizer.AnalyzeTemplate(bytes, contentControls: true).Diagnostics);
        }

        [Fact]
        public void CompilationDropsMarkerAnchorsWithoutCreatingImageParagraphs()
        {
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(true)), Field("if Ready")),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body"))),
                new XElement(W.p, Field("endif"), new XElement(W.r, Drawing(true))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Equal(2, analysis.Diagnostics.Count);
            var fields = Templater.ParseFieldsToDict(analysis.ExtractedFields);
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            Assert.Empty(Root(compiled.Bytes).Descendants(WP.anchor));
            var assembled = Assembler.AssembleDocument(compiled.Bytes,
                new XElement("_odx", new XElement(fields["1"].Atom + "b", "true")));
            Assert.False(assembled.HasErrors, assembled.Error);
            Assert.Empty(Root(assembled.Bytes).Descendants(WP.anchor));
            Assert.Contains("Body", Root(assembled.Bytes).Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void InlineContentImagesSurviveEmptyAndNonemptyValues(bool hasValue)
        {
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(false)), Field("Value")));
            var normalized = Normalizer.NormalizeTemplate(bytes, false);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors);
            var data = new XElement("_odx");
            if (hasValue) data.Add(new XElement(fields["1"].Atom, "Replacement"));
            var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
            Assert.False(assembled.HasErrors, assembled.Error);
            Assert.Single(Root(assembled.Bytes).Descendants(WP.inline));
            Assert.Equal(hasValue, Root(assembled.Bytes).Value.Contains("Replacement"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Replacement")]
        public async Task ImageFreeContentKeepsExistingParagraphBehavior(string value)
        {
            var field = Field("Value");
            field.AddFirst(new XElement(W.rPr, new XElement(W.b)));
            var bytes = Document(new XElement(W.p, new XElement(W.pPr,
                new XElement(W.jc, new XAttribute(W.val, "center"))), field),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "After"))));
            var normalized = Normalizer.NormalizeTemplate(bytes, false);
            Assert.Empty(normalized.Diagnostics);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var compiledRoot = Root(compiled.Bytes);
            Assert.DoesNotContain("Conditional", compiledRoot.Value);
            Assert.Empty(compiledRoot.Descendants(W.sdtContent).Elements(W.p));
            var data = new XElement("_odx");
            if (value != null) data.Add(new XElement(fields["1"].Atom, value));
            var sync = Assembler.AssembleDocument(compiled.Bytes, data);
            var asyncResult = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
            foreach (var result in new[] { sync, asyncResult })
            {
                Assert.False(result.HasErrors, result.Error);
                var paragraphs = Root(result.Bytes).Element(W.body).Elements(W.p).ToList();
                Assert.Equal(2, paragraphs.Count);
                Assert.Equal("After", paragraphs.Last().Value);
                Assert.Equal(value ?? "", paragraphs.First().Value);
                Assert.Equal("center", (string)paragraphs.First().Element(W.pPr).Element(W.jc).Attribute(W.val));
                Assert.Single(paragraphs.First().Descendants(W.b));
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task CompilationDropsAnchoredBlockContentImages(bool hasValue)
        {
            var analysis = Normalizer.AnalyzeTemplate(Document(
                new XElement(W.p, Field("Value"), new XElement(W.r, Drawing(true)))));
            Assert.Single(analysis.Diagnostics);
            var fields = Templater.ParseFieldsToDict(analysis.ExtractedFields);
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            Assert.Empty(Root(compiled.Bytes).Descendants(WP.anchor));
            var data = new XElement("_odx");
            if (hasValue) data.Add(new XElement(fields["1"].Atom, "Replacement"));
            var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
            Assert.False(result.HasErrors, result.Error);
            Assert.Empty(Root(result.Bytes).Descendants(WP.anchor));
            Assert.Single(Root(result.Bytes).Element(W.body).Elements(W.p));
            Assert.Equal(hasValue ? "Replacement" : "", Root(result.Bytes).Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task LegacyCompiledRunContentKeepsItsAnchorsWhenOtherContentIsPresent(bool inline)
        {
            var control = new XElement(W.sdt, new XElement(W.sdtContent, new XElement(W.r,
                new XElement(W.t, "<Content Select=\"./C1[1]\" Optional=\"true\"/>"))));
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(true)), control,
                inline ? new XElement(W.r, Drawing(false)) : new XElement(W.r, new XElement(W.t, "Text"))));
            var result = await Assembler.AssembleDocumentAsync(bytes, "<_odx/>", null);
            Assert.False(result.HasErrors, result.Error);
            Assert.Single(Root(result.Bytes).Descendants(WP.anchor));
            Assert.Single(Root(result.Bytes).Element(W.body).Elements(W.p));
        }

        [Fact]
        public async Task ImageFreeContentInAListKeepsExistingEmptyParagraphs()
        {
            var bytes = Document(new XElement(W.p, Field("list Items")),
                new XElement(W.p, Field("Value")), new XElement(W.p, Field("endlist")));
            var normalized = Normalizer.NormalizeTemplate(bytes, false);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var listAtom = fields["1"].Atom;
            var data = new XElement("_odx", new XElement(listAtom,
                new XElement(listAtom + "i"),
                new XElement(listAtom + "i", new XElement(fields["2"].Atom, "Second"))));
            var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), null);
            Assert.False(result.HasErrors, result.Error);
            Assert.Equal("Second", Root(result.Bytes).Value);
            Assert.Equal(2, Root(result.Bytes).Element(W.body).Elements(W.p).Count());
        }

        [Fact]
        public async Task BlockContentStillComposesInsertedDocuments()
        {
            var normalized = Normalizer.NormalizeTemplate(Document(new XElement(W.p, Field("Subdocument"))), false);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var data = new XElement("_odx", new XElement(fields["1"].Atom, "oxpt://DocumentAssembler/insert/insert"));
            var source = new IndirectSource("insert", Document(new XElement(W.p,
                new XElement(W.r, new XElement(W.t, "Inserted")))));
            var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, data.ToString(), new[] { source });
            Assert.False(result.HasErrors, result.Error);
            Assert.Equal("Inserted", Root(result.Bytes).Value);
        }

        [Fact]
        public void TableCellsRetainTheirBoundaryWhileAnchorsAreDiagnosed()
        {
            var bytes = Document(new XElement(W.tbl, new XElement(W.tr,
                new XElement(W.tc, new XElement(W.p, Field("Value"), new XElement(W.r, Drawing(true)))),
                new XElement(W.tc, new XElement(W.p, Field("Other"))))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Single(analysis.Diagnostics);
            var part = JsonNode.Parse(analysis.ExtractedFields)[0].AsArray();
            Assert.Equal(2, part.Count);
            Assert.All(part, cell => Assert.IsType<JsonArray>(cell));
        }

        [Theory]
        [InlineData("position:absolute", true)]
        [InlineData("position:relative", true)]
        [InlineData("width:10pt;height:10pt", false)]
        public void LegacyVmlPicturesDistinguishFloatingFromInline(string style, bool floating)
        {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            var bytes = Document(new XElement(W.p, Field("Value"), new XElement(W.r,
                new XElement(W.pict, new XElement(v + "shape", new XAttribute("style", style),
                    new XElement(v + "imagedata", new XAttribute(R + "id", "image")))))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Equal(floating ? 1 : 0, analysis.Diagnostics.Count);
            var part = JsonNode.Parse(analysis.ExtractedFields)[0].AsArray();
            Assert.Equal(floating, part[0] is JsonObject);
        }

        [Theory]
        [InlineData("position:absolute", false, true)]
        [InlineData("position:relative", false, true)]
        [InlineData("position:absolute", true, true)]
        [InlineData("position:relative", true, true)]
        [InlineData("width:10pt;height:10pt", false, false)]
        [InlineData("width:10pt;height:10pt", true, false)]
        public void LegacyVmlGroupsSupplyThePicturePosition(string style, bool nested, bool floating)
        {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            foreach (var controls in new[] { false, true })
            {
                var shape = new XElement(v + "shape", new XAttribute("style", "width:10pt;height:10pt"),
                    new XElement(v + "imagedata", new XAttribute(R + "id", "image")));
                var group = new XElement(v + "group", new XAttribute("style", style),
                    nested ? new XElement(v + "group", new XAttribute("style", "width:10pt;height:10pt"), shape) : shape);
                var bytes = Document(new XElement(W.p, Field("Value", controls),
                    new XElement(W.r, new XElement(W.pict, group))));
                var analysis = Normalizer.AnalyzeTemplate(bytes, contentControls: controls);
                Assert.Equal(floating ? 1 : 0, analysis.Diagnostics.Count);
                Assert.Equal(floating, JsonNode.Parse(analysis.ExtractedFields)[0][0] is JsonObject);
                if (floating)
                    Assert.Throws<FieldParseException>(() => Normalizer.NormalizeTemplate(bytes, false, contentControls: controls));
                var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate,
                    Templater.ParseFieldsToDict(analysis.ExtractedFields));
                Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
                Assert.Equal(floating ? 0 : 1, Root(compiled.Bytes).Descendants(W.pict).Count());
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void VmlWithBothFloatingAndInlineShapesStillCountsAsContent(bool controls)
        {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XElement Shape(string style) => new XElement(v + "shape", new XAttribute("style", style),
                new XElement(v + "imagedata", new XAttribute(R + "id", "image")));
            var bytes = Document(new XElement(W.p, Field("Value", controls), new XElement(W.r,
                new XElement(W.pict, new XElement(v + "group", new XAttribute("style", "width:10pt;height:10pt"),
                    Shape("position:absolute"), Shape("width:10pt;height:10pt"))))));
            var analysis = Normalizer.AnalyzeTemplate(bytes, contentControls: controls);
            Assert.Empty(analysis.Diagnostics);
            Assert.IsType<JsonArray>(JsonNode.Parse(analysis.ExtractedFields)[0][0]);
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate,
                Templater.ParseFieldsToDict(analysis.ExtractedFields));
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            Assert.Equal(2, Root(compiled.Bytes).Descendants(v + "shape").Count());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MixedVmlOnABlockFieldControlStillDiagnosesItsFloatingShape(bool alternateContent)
        {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            XElement Shape(string style) => new XElement(v + "shape", new XAttribute("style", style),
                new XElement(v + "imagedata", new XAttribute(R + "id", "image")));
            var picture = new XElement(W.pict, Shape("position:absolute"), Shape("width:10pt;height:10pt"));
            var bytes = Document(new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent,
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "[Value]")), new XElement(W.r,
                    alternateContent ? new XElement(mc + "AlternateContent",
                        new XElement(mc + "Choice", new XAttribute("Requires", "v"), picture),
                        new XElement(mc + "Fallback", new XElement(picture))) : picture)))));
            var analysis = Normalizer.AnalyzeTemplate(bytes, contentControls: true);
            var diagnostic = Assert.Single(analysis.Diagnostics);
            Assert.Equal("FloatingImageOnBlockField", diagnostic.Code);
            Assert.Equal("1", diagnostic.FieldId);
            Assert.Equal("/word/document.xml", diagnostic.PartUri);
            Assert.Throws<FieldParseException>(() => Normalizer.NormalizeTemplate(bytes, false, contentControls: true));
            Assert.Throws<FieldParseException>(() => Normalizer.ExtractFieldsOnly(bytes, contentControls: true));
        }

        [Fact]
        public void OtherOpaqueRunContentIsNotIgnoredAlongsideAnAnchor()
        {
            var bytes = Document(new XElement(W.p, Field("Value"), new XElement(W.r, Drawing(true)),
                new XElement(W.r, new XElement(W.footnoteReference, new XAttribute(W.id, 1)))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Empty(analysis.Diagnostics);
            Assert.IsType<JsonArray>(JsonNode.Parse(analysis.ExtractedFields)[0][0]);
        }

        [Fact]
        public void AlternateContentAnchorsUseTheSelectedDrawingBranch()
        {
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            var bytes = Document(new XElement(W.p, Field("Value"), new XElement(W.r,
                new XElement(mc + "AlternateContent",
                    new XElement(mc + "Choice", new XAttribute("Requires", "wp"), Drawing(true)),
                    new XElement(mc + "Fallback", Drawing(true))))));
            Assert.Single(Normalizer.AnalyzeTemplate(bytes).Diagnostics);
        }

        [Fact]
        public void FieldCodeOnlyRunsDoNotChangeBlockClassification()
        {
            var bytes = Document(new XElement(W.p,
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
                Field("Value"), new XElement(W.r, Drawing(true))));
            var analysis = Normalizer.AnalyzeTemplate(bytes);
            Assert.Single(analysis.Diagnostics);
            var compiled = Templater.CompileTemplate(analysis.NormalizedTemplate, Templater.ParseFieldsToDict(analysis.ExtractedFields));
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            Assert.Empty(Root(compiled.Bytes).Descendants(WP.anchor));
        }

        [Fact]
        public void ImageFreeNativeFieldsKeepTheirExistingNonTextHandling()
        {
            var bytes = Document(new XElement(W.p, Field("if Ready", true), new XElement(W.r,
                    new XElement(W.footnoteReference, new XAttribute(W.id, 1)))),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body"))),
                new XElement(W.p, Field("endif", true)));
            var normalized = Normalizer.NormalizeTemplate(bytes, false, contentControls: true);
            Assert.Empty(normalized.Diagnostics);
            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var result = Assembler.AssembleDocument(compiled.Bytes, new XElement("_odx",
                new XElement(fields["1"].Atom + "b", "true")));
            Assert.False(result.HasErrors, result.Error);
            Assert.Equal("Body", Root(result.Bytes).Value);
        }

        [Fact]
        public async Task EmptyBlockContentKeepsTheRequiredTableCellParagraph()
        {
            var bytes = Document(new XElement(W.tbl, new XElement(W.tr,
                new XElement(W.tc, new XElement(W.p, Field("Value"))))));
            var normalized = Normalizer.NormalizeTemplate(bytes, false);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, Templater.ParseFieldsToDict(normalized.ExtractedFields));
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));
            var result = await Assembler.AssembleDocumentAsync(compiled.Bytes, "<_odx/>", null);
            Assert.False(result.HasErrors, result.Error);
            Assert.Single(Assert.Single(Root(result.Bytes).Descendants(W.tc)).Elements(W.p));
            Assert.Empty(Root(result.Bytes).Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task LegacyPairedMarkerParagraphsLoseTheirAnchorsInAssembly(bool condition)
        {
            XElement Metadata(string xml) => new XElement(W.sdt, new XElement(W.sdtContent,
                new XElement(W.r, new XElement(W.t, xml))));
            var bytes = Document(new XElement(W.p, new XElement(W.r, Drawing(true)),
                    Metadata("<Conditional Select=\"Ready\" Match=\"true\"/>")),
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body"))),
                new XElement(W.p, Metadata("<EndConditional/>"), new XElement(W.r, Drawing(true))));
            var result = await Assembler.AssembleDocumentAsync(bytes,
                new XElement("_odx", new XElement("Ready", condition ? "true" : "false")).ToString(), null);
            Assert.False(result.HasErrors, result.Error);
            Assert.Empty(Root(result.Bytes).Descendants(WP.anchor));
            Assert.Equal(condition ? "Body" : "", Root(result.Bytes).Value);
        }

        [Fact]
        public void DiagnosticsIncludeFieldsInHeadersAndAccumulateAcrossParts()
        {
            var bytes = Document(new XElement(W.p, Field("Value"), new XElement(W.r, Drawing(true))));
            using var stream = new MemoryStream();
            stream.Write(bytes);
            using (var document = WordprocessingDocument.Open(stream, true))
            {
                var header = document.MainDocumentPart.AddNewPart<HeaderPart>();
                header.PutXDocument(new XDocument(new XElement(W.hdr,
                    new XElement(W.p, new XElement(W.r, Drawing(true)), Field("HeaderValue")))));
                document.MainDocumentPart.GetXDocument().Root.Element(W.body).Add(new XElement(W.sectPr,
                    new XElement(W.headerReference, new XAttribute(W.type, "default"),
                        new XAttribute(R + "id", document.MainDocumentPart.GetIdOfPart(header)))));
                document.MainDocumentPart.PutXDocument();
            }
            var analysis = Normalizer.AnalyzeTemplate(stream.ToArray());
            Assert.Equal(2, analysis.Diagnostics.Count);
            Assert.Contains(analysis.Diagnostics, d => d.FieldText == "HeaderValue" && d.PartUri.Contains("header"));
            Assert.Equal(2, analysis.Diagnostics.Select(d => d.FieldId).Distinct().Count());
        }
    }
}
