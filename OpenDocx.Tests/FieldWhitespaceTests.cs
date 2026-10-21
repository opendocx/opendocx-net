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
        public static IEnumerable<object[]> NonBreakingFieldBoundaries()
        {
            foreach (var space in new[] { null, "default", "preserve" })
                foreach (var boundary in new[] { "expression", "keyword", "separator", "internal" })
                    yield return new object[] { space, boundary };
        }

        [Theory]
        [MemberData(nameof(NonBreakingFieldBoundaries))]
        public void NonBreakingFieldSeparatorsSurviveExtractionAndReplacement(string space, string boundary)
        {
            XElement Run(string text) => new XElement(W.r, new XElement(W.t,
                space == null ? null : new XAttribute(XNamespace.Xml + "space", space), text));

            const string field = "elseif\u00A0TakingTitle == \"Unequal\"";
            var parts = boundary switch
            {
                "expression" => new[] { "{[elseif", "\u00A0TakingTitle == \"Unequal\"]}" },
                "keyword" => new[] { "{[elseif\u00A0", "TakingTitle == \"Unequal\"]}" },
                "separator" => new[] { "{[elseif", "\u00A0", "TakingTitle == \"Unequal\"]}" },
                _ => new[] { "{[" + field + "]}" }
            };
            var runs = new List<XElement> { Run("Start\u00A0{[if Ready]}First") };
            runs.AddRange(parts.Take(parts.Length - 1).Select(Run));
            // Shared runs and the final content field exercise replacement offsets
            // on both sides of the split NBSP-containing control field.
            runs.Add(Run(parts.Last() + "Second{[else]}Third{[endif]}\u00A0End{[Value]}\u00A0Tail"));
            var source = Document(new XElement(W.p, runs));
            var normalized = Normalizer.NormalizeTemplate(source, false);
            Assert.Empty(normalized.Diagnostics);
            Assert.Equal(normalized.ExtractedFields, Normalizer.ExtractFieldsOnly(source));
            var extracted = JsonNode.Parse(normalized.ExtractedFields)[0][0].AsArray();
            Assert.Equal(new[] { "if Ready", field, "else", "endif", "Value" },
                extracted.Select(f => f["content"].GetValue<string>()));

            var fields = Templater.ParseFieldsToDict(normalized.ExtractedFields);
            Assert.Equal(new[] { FieldType.If, FieldType.ElseIf, FieldType.Else, FieldType.EndIf, FieldType.Content },
                Enumerable.Range(1, 5).Select(i => fields[i.ToString()].Type));
            Assert.Equal("TakingTitle == \"Unequal\"", fields["2"].Expression);
            var compiled = Templater.CompileTemplate(normalized.NormalizedTemplate, fields);
            Assert.False(compiled.HasErrors, string.Join("\n", compiled.Errors));

            foreach (var scenario in new[] { (true, true, "First"), (false, true, "Second"), (false, false, "Third") })
            {
                var data = new XElement("_odx",
                    new XElement(fields["1"].Atom + "b", scenario.Item1 ? "true" : "false"),
                    new XElement(fields["2"].Atom + "b", scenario.Item2 ? "true" : "false"),
                    new XElement(fields["5"].Atom, "Replacement"));
                var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
                Assert.False(assembled.HasErrors, assembled.Error);
                var paragraph = Assert.Single(Root(assembled.Bytes).Element(W.body).Elements(W.p));
                Assert.Equal("Start\u00A0" + scenario.Item3 + "\u00A0EndReplacement\u00A0Tail",
                    string.Concat(paragraph.Descendants(W.t).Select(t => t.Value)));
            }
        }
    }
}
