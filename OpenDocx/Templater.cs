/***************************************************************************

Copyright (c) Lowell Stewart 2018-2025.
Licensed under the Mozilla Public License. See LICENSE file in the project root for full license information.

Published at https://github.com/opendocx/opendocx-net
Developer: Lowell Stewart
Email: lowell@opendocx.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using Newtonsoft.Json;
using OpenXmlPowerTools;

namespace OpenDocx
{
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

    public static class Templater
    {
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

            var compileResult = CompileTemplate(normalizeResult.NormalizedTemplate, fieldDict);
            if (compileResult.HasErrors)
            {
                throw new Exception("CompileTemplate failed:\n" + string.Join('\n', compileResult.Errors));
            }
            var result = new PrepareTemplateResult(compileResult.Bytes, fieldDict);
            if (options.GenerateLogicTree || options.GenerateLegacyLogicModule)
            {
                result.LogicTree = LogicTree.BuildLogicTree(ast);
                if (options.GenerateLegacyLogicModule)
                {
                    result.LegacyLogicModule = GetLegacyTemplateModule(result.LogicTree);
                }
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

        public static void AddFieldsToDict(JsonArray jsonArray, Dictionary<string, string> fieldDict)
        {
            // Loop through each item
            foreach (JsonNode item in jsonArray)
            {
                if (item is JsonArray)
                {
                    AddFieldsToDict(item as JsonArray, fieldDict);
                } else {
                    var content = item["content"]?.ToString();
                    var fieldId = item["id"]?.ToString();
                    fieldDict[fieldId] = content;
                }
            }
        }

        public static CompileResult CompileTemplate(string originalTemplateFile, string preProcessedTemplateFile, string parsedFieldInfoFile)
        {
            string json = File.ReadAllText(parsedFieldInfoFile);
            var xm = JsonConvert.DeserializeObject<FieldTransformIndex>(json);
            // translate xm into a simple Dictionary<string, string> so we can use basic TemplateTransformer
            // instead of the former custom implementation
            var fieldMap = new FieldReplacementIndex();
            foreach (var fieldId in xm.Keys)
            {
                fieldMap[fieldId] = new FieldReplacement(xm[fieldId]);
            }
            string destinationTemplatePath = originalTemplateFile + "gen.docx";
            var errors = TemplateTransformer.TransformTemplate(preProcessedTemplateFile,
                destinationTemplatePath, TemplateFormat.ObjectDocx, fieldMap);
            return new CompileResult(destinationTemplatePath, errors);
        }

        public static TemplateTransformResult CompileTemplate(byte[] preProcessedTemplate, Dictionary<string, ParsedField> fieldDict)
        {
            // translate fieldDict into a simple Dictionary<string, FieldReplacement> so we can use TemplateTransformer
            var fieldMap = new FieldReplacementIndex();
            foreach (var fieldId in fieldDict.Keys)
            {
                fieldMap[fieldId] = new FieldReplacement(fieldDict[fieldId]);
            }
            return TemplateTransformer.TransformTemplate(preProcessedTemplate,
                TemplateFormat.ObjectDocx, fieldMap);
        }

        public static string GetLegacyTemplateModule(List<FieldLogicNode> logicTree)
        {
            var result = new StringBuilder();
            result.Append(@"'use strict';
exports.version='2.0.2';
exports.evaluate=function(cx,cl,h)
{
h.beginObject('_odx',cx,cl);
");
            logicTree.ForEach(node => node.SerializeToLegacyModule(result));
            result.Append(@"h.endObject()
}");
            return result.ToString();
        }
    }

}
