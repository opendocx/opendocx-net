using System;
using System.Collections.Generic;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OpenDocx
{
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

    public class PrepareTemplateOptions
    {
        public bool GenerateFlatPreview { get; set; }
        public bool GenerateLogicTree { get; set; }
        public bool GenerateLegacyLogicModule { get; set; }
        public bool RemoveCustomProperties { get; set; }
        public List<string> KeepPropertyNames { get; set; }

        public PrepareTemplateOptions(bool removeCustomProperties = true, List<string> keepPropertyNames = null, bool generateFlatPreview = false, bool generateLogicTree = true, bool generateLegacyModule = false)
        {
            GenerateFlatPreview = generateFlatPreview;
            GenerateLogicTree = generateLogicTree;
            GenerateLegacyLogicModule = generateLegacyModule;
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
        public string LegacyLogicModule { get; internal set; }
        public byte[] FlatPreviewBytes { get; internal set; }
        public Dictionary<string, string> FlatPreviewFields { get; internal set; }

        internal PrepareTemplateResult(byte[] oxptTemplateBytes, Dictionary<string, ParsedField> fields)
        {
            OXPTTemplateBytes = oxptTemplateBytes;
            Fields = fields;
        }
    }

    public static class OpenDocx
    {
        public static readonly JsonSerializerOptions DefaultJsonOptions = new()
        {
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters = { new JsonStringEnumConverter() },
            NumberHandling = JsonNumberHandling.AllowReadingFromString,
        };

        public class NormalizeDocxResult
        {
            public string UnparsedFieldJson { get; internal set; }
            public byte[] NormalizedBytes { get; internal set; }
            public byte[] FlatPreviewBytes { get; internal set; }
            public string PreviewError { get; internal set; }

            internal NormalizeDocxResult(string unparsedFieldJson, byte[] normalizedBytes,
                byte[] flatPreviewBytes = null, string previewError = null)
            {
                UnparsedFieldJson = unparsedFieldJson;
                NormalizedBytes = normalizedBytes;
                FlatPreviewBytes = flatPreviewBytes;
                PreviewError = previewError;
            }
        }

        public class TransformDocxResult
        {
            public byte[] Bytes { get; }
            public string Error { get; }

            internal TransformDocxResult(byte[] bytes, string error = null)
            {
                Bytes = bytes;
                Error = error;
            }
        }

        // work performed by 'prepare template: stage 1' lambda:
        public static NormalizeDocxResult NormalizeDocx(byte[] docxBytes, PrepareTemplateOptions options)
        {
            var normalizeResult = FieldExtractor.NormalizeTemplate(docxBytes,
                options.RemoveCustomProperties, options.KeepPropertyNames);
            byte[] previewBytes = null;
            string previewError = null;

            if (options.GenerateFlatPreview)
            {
                var previewResult = TemplateTransformer.TransformTemplate(
                    normalizeResult.NormalizedTemplate,
                    TemplateFormat.PreviewDocx,
                    null); // field map is ignored when output = TemplateFormat.PreviewDocx
                if (!previewResult.HasErrors)
                    previewBytes = previewResult.Bytes;
                else
                    previewError = "Preview failed to generate:\n" + string.Join('\n', previewResult.Errors);
            }
            return new NormalizeDocxResult(normalizeResult.ExtractedFields,
                normalizeResult.NormalizedTemplate, previewBytes, previewError);
        }

        // work to be performed by 'prepare template: stage 2' lambda, whether you use this implementation
        // or something else (for example a NodeJS implementation that actually parses field expressions):
        public static ParseFieldsResult ParseDocx(string extractedFields, bool generateLegacyLogicModule = false)
        {
            return Templater.ParseFields(extractedFields, generateLegacyLogicModule);
        }

        // work performed by 'prepare template: stage 3' lambda:
        public static TransformDocxResult TransformDocx(byte[] normalizedBytes,
            Dictionary<string, ParsedField> fieldDictionary)
        {
            string errorMessage = null;
            var compileResult = Templater.CompileTemplate(normalizedBytes, fieldDictionary);
            if (compileResult.HasErrors)
                errorMessage = string.Join('\n', compileResult.Errors);
            return new TransformDocxResult(compileResult.Bytes, errorMessage);
        }

        public static PrepareTemplateResult PrepareTemplate(byte[] templateBytes, PrepareTemplateOptions options = null)
        {
            if (options == null) options = PrepareTemplateOptions.Default;

            var result1 = NormalizeDocx(templateBytes, options);
            var result2 = ParseDocx(result1.UnparsedFieldJson, options.GenerateLegacyLogicModule);
            var result3 = TransformDocx(result1.NormalizedBytes, result2.ParsedFields);

            if (!string.IsNullOrEmpty(result3.Error))
            {
                throw new Exception("TransformDocx failed:\n" + result3.Error);
            }
            var result = new PrepareTemplateResult(result3.Bytes, result2.ParsedFields);
            if (options.GenerateLogicTree || options.GenerateLegacyLogicModule)
            {
                result.LogicTree = result2.LogicTree;
                if (options.GenerateLegacyLogicModule)
                {
                    result.LegacyLogicModule = result2.LegacyLogicModule;
                }
            }
            if (options.GenerateFlatPreview)
            {
                if (string.IsNullOrEmpty(result1.PreviewError))
                {
                    var previewFields = new Dictionary<string, string>();
                    foreach (var fieldId in result2.ParsedFields.Keys)
                    {
                        previewFields[fieldId] = result2.ParsedFields[fieldId].Text;
                    }
                    result.FlatPreviewBytes = result1.FlatPreviewBytes;
                    result.FlatPreviewFields = previewFields;
                }
                else
                {
                    Console.WriteLine("Preview failed to generate:\n" + result1.PreviewError);
                }
            }
            return result;
        }
    }
}
