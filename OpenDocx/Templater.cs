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
    public class ParseFieldsResult
    {
        public Dictionary<string, ParsedField> ParsedFields { get; internal set; }
        public List<FieldLogicNode> LogicTree { get; internal set; }
        public string LegacyLogicModule { get; internal set; }

        internal ParseFieldsResult(Dictionary<string, ParsedField> parsedFields, List<FieldLogicNode> logicTree,
            string legacyLogicModule = null)
        {
            ParsedFields = parsedFields;
            LogicTree = logicTree;
            LegacyLogicModule = legacyLogicModule;
        }
    }

    public static class Templater
    {
        public static ParseFieldsResult ParseFields(string extractedFields, bool generateLegacyLogicModule = false)
        {
            var fieldList = JsonNode.Parse(extractedFields);
            // this "dumb" implementation creates a simple template AST, but it does not bother parsing
            // the expressions inside individual fields. A complete implementation would do that, so it
            // could check for errors inside fields!
            var ast = FieldParser.ParseContentArray(fieldList as JsonArray);
            // create a map from field ID (for DOCX, basially, field number) to nodes in the AST
            var fieldDict = new Dictionary<string, ParsedField>();
            var atoms = new FieldExprNamer();
            FieldParser.BuildFieldDictionary(ast, fieldDict, atoms); // this also atomizes expressions in fields
            // note: it ALSO mutates ast, adding atom annotations for expressions
            var logicTree = LogicTree.BuildLogicTree(ast);
            string legacyLogicModule = null;
            if (generateLegacyLogicModule)
            {
                legacyLogicModule = GetLegacyTemplateModule(logicTree);
            }
            return new ParseFieldsResult(fieldDict, logicTree, legacyLogicModule);
        }

        public static void AddFieldsToDict(JsonArray jsonArray, Dictionary<string, string> fieldDict)
        {
            // Loop through each item
            foreach (JsonNode item in jsonArray)
            {
                if (item is JsonArray)
                {
                    AddFieldsToDict(item as JsonArray, fieldDict);
                }
                else
                {
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
