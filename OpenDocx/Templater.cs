/***************************************************************************

Copyright (c) Lowell Stewart 2018-2022.
Licensed under the Mozilla Public License. See LICENSE file in the project root for full license information.

Published at https://github.com/opendocx/opendocx-net
Developer: Lowell Stewart
Email: lowell@opendocx.com

***************************************************************************/

using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using OpenXmlPowerTools;

namespace OpenDocx
{
    public static class Templater
    {
        public static CompileResult CompileTemplate(string originalTemplateFile, string preProcessedTemplateFile, string parsedFieldInfoFile)
        {
            string json = File.ReadAllText(parsedFieldInfoFile);
            var xm = JsonConvert.DeserializeObject<FieldTransformIndex>(json);
            // translate xm into a simple Dictionary<string, string> so we can use basic TemplateTransformer
            // instead of the former custom implementation
            var fieldMap = new FieldReplacementIndex();
            foreach (var fieldId in xm.Keys) {
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
    }
}
