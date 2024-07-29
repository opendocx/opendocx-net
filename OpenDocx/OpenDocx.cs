using Newtonsoft.Json.Linq;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class PrepareTemplateOptions
    {
        public bool GenerateFlatPreview { get; set; }
        public bool GenerateLogicTree { get; set; }
        public bool RemoveCustomProperties { get; set; }
        public List<string> KeepPropertyNames { get; set; }

        public PrepareTemplateOptions(bool removeCustomProperties = true, List<string> keepPropertyNames = null, bool generateFlatPreview = false, bool generateLogicTree = true)
        {
            GenerateFlatPreview = generateFlatPreview;
            GenerateLogicTree = generateLogicTree;
            RemoveCustomProperties = removeCustomProperties;
            KeepPropertyNames = keepPropertyNames;
        }

        public static readonly PrepareTemplateOptions Default = new();
    }

    public class PrepareTemplateResult
    {
        public byte[] OXPTTemplateBytes { get; internal set; }
        public Dictionary<string, ParsedField> Fields { get; internal set; }
        public List<FieldLogicNode> LogicTree { get; internal set; }
        public byte[] FlatPreviewBytes { get; internal set; }
        public Dictionary<string, string> FlatPreviewFields { get; internal set; }

        internal PrepareTemplateResult(byte[] oxptTemplateBytes, Dictionary<string, ParsedField> fields, IEnumerable<FieldLogicNode> logicTree = null, byte[] flatPreviewBytes = null, Dictionary<string, string> flatPreviewFields = null)
        {
            OXPTTemplateBytes = oxptTemplateBytes;
            Fields = fields;
            if (logicTree != null) LogicTree = logicTree.ToList();
            FlatPreviewBytes = flatPreviewBytes;
            FlatPreviewFields = flatPreviewFields;
        }

    }

    public class IndirectSource
    {
        public string ID { get; set; }
        public byte[] Bytes { get; set; }
        public bool KeepSections { get; set; }

        public IndirectSource(string id, byte[] bytes, bool keepSections = false)
        {
            ID = id;
            Bytes = bytes;
            KeepSections = keepSections;
        }
    }

    public static class OpenDocx
    {
        public static JsonSerializerOptions DefaultJsonOptions = new()
        {
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters =
                    {
                        new JsonStringEnumConverter()
                    },
        };

        public static PrepareTemplateResult PrepareTemplate(byte[] templateBytes, PrepareTemplateOptions options = null)
        {
            if (options == null) options = PrepareTemplateOptions.Default;

            var normalizeResult = FieldExtractor.NormalizeTemplate(templateBytes, options.RemoveCustomProperties, options.KeepPropertyNames);
            var fieldList = JsonNode.Parse(normalizeResult.ExtractedFields);
            var ast = FieldParser.ParseContentArray(fieldList as JsonArray);
            // create a map from field ID to nodes in the AST, which before would have been saved in fieldDictPath = templatePath + 'obj.fields.json'
            var fieldDict = new Dictionary<string, ParsedField>();
            var atoms = new FieldExprNamer();
            FieldParser.BuildFieldDictionary(ast, fieldDict, atoms); // this also atomizes expressions in fields
            // note: it ALSO mutates ast, adding atom annotations for expressions

            var compileResult = Templater.CompileTemplate(normalizeResult.NormalizedTemplate, fieldDict);
            if (compileResult.HasErrors)
            {
                throw new Exception("CompileTemplate failed:\n" + string.Join('\n', compileResult.Errors));
            }
            var result = new PrepareTemplateResult(compileResult.Bytes, fieldDict);
            if (options.GenerateLogicTree)
            {
                result.LogicTree = LogicTree.BuildLogicTree(ast);
            }
            if (options.GenerateFlatPreview)
            {
                var previewResult = TemplateTransformer.TransformTemplate(
                    normalizeResult.NormalizedTemplate,
                    TemplateFormat.PreviewDocx,
                    null); // field map is ignored when output = TemplateFormat.PreviewDocx
                if (!previewResult.HasErrors)
                {
                    var previewFields = new Dictionary<string, string>();
                    foreach (var fieldId in fieldDict.Keys)
                    {
                        previewFields[fieldId] = fieldDict[fieldId].Text;
                    }
                    result.FlatPreviewBytes = previewResult.Bytes;
                    result.FlatPreviewFields = previewFields;
                }
                else
                {
                    Console.WriteLine("Preview failed to generate:\n" + string.Join('\n', previewResult.Errors));
                }
            }
            return result;
        }
    }
}
