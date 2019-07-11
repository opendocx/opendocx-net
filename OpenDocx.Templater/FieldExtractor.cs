﻿using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Xml.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;

namespace OpenDocx
{
    public class FieldExtractor
    {
        public async Task<object> ExtractFieldsAsync(dynamic input)
        {
            var templateFile = (string)input.templateFile;
            return ExtractFields(templateFile);
        }

        public static FieldExtractResult ExtractFields(string templateFileName)
        {
            string newTemplateFileName = templateFileName + "obj.docx";
            string outputFile = templateFileName + "obj.json";
            WmlDocument templateDoc = new WmlDocument(templateFileName); // just reads the template's bytes into memory (that's all), read-only
            WmlDocument preprocessedTemplate = null;
            byte[] byteArray = templateDoc.DocumentByteArray;
            var fieldAccumulator = new FieldAccumulator();
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, byteArray.Length); // copy template file (binary) into memory -- I guess so the template/file handle isn't held/locked?
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true)) // read & parse that byte array into OXML document (also in memory)
                {
                    ExtractAllTemplateFields(wordDoc, fieldAccumulator);
                }
                preprocessedTemplate = new WmlDocument(newTemplateFileName, mem.ToArray());
            }
            // save the output (even in the case of error, since error messages are in the file)
            preprocessedTemplate.Save();

            using (StreamWriter sw = File.CreateText(outputFile))
            {
                sw.Write('[');
                bool first = true;
                foreach (var field in fieldAccumulator)
                {
                    if (first)
                        first = false;
                    else
                        sw.Write(',');
                    sw.Write(field.JsonSerialize());
                }
                sw.Write(']');
                sw.Close();
            }

            return new FieldExtractResult(newTemplateFileName, outputFile);
        }

        private static void ExtractAllTemplateFields(WordprocessingDocument wordDoc, FieldAccumulator fieldAccumulator)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                throw new FieldParseException("Invalid template - contains tracked revisions");

            foreach (var part in wordDoc.ContentParts())
            {
                ExtractFieldsFromPart(part, fieldAccumulator);
            }
        }

        private static void ExtractFieldsFromPart(OpenXmlPart part, FieldAccumulator fieldAccumulator)
        {
            XDocument xDoc = part.GetXDocument();
            var xDocRoot = (XElement)IdentifyAndTransformFields(xDoc.Root, fieldAccumulator);
            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.PutXDocument();
        }

        private static int CountSubstring(string substring, string source)
        {
            int count = 0, n = 0;
            if (!string.IsNullOrEmpty(substring))
            {
                while ((n = source.IndexOf(substring, n, StringComparison.Ordinal)) != -1)
                {
                    n += substring.Length;
                    ++count;
                }
            }
            return count;
        }

        private static object IdentifyAndTransformFields(XNode node, FieldAccumulator fieldAccumulator)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.sdt)
                {
                    var alias = (string)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    if (alias == null || alias == "")
                    {
                        var ccContents = element
                            .DescendantsTrimmed(W.txbxContent)
                            .Where(e => e.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate()
                            .CleanUpInvalidCharacters();
                        if (FieldRecognizer.IsField(ccContents, out ccContents))
                        {
                            var fieldId = fieldAccumulator.Count.ToString();
                            var newCC = new XElement(element.Name, element.Attributes(), element.Nodes());
                            var props = newCC.Elements(W.sdtPr).FirstOrDefault();
                            if (props == null)
                            {
                                props = new XElement(W.sdtPr);
                                newCC.Add(props);
                            }
                            var tagElem = props.Elements(W.tag).FirstOrDefault();
                            if (tagElem == null)
                            {
                                tagElem = new XElement(W.tag);
                                props.Add(tagElem);
                            }
                            tagElem.SetAttributeValue(W.val, fieldId);
                            fieldAccumulator.Add(new FieldInfo(ccContents, fieldId));
                            return newCC;
                        }
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => IdentifyAndTransformFields(n, fieldAccumulator)));
                    }
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => IdentifyAndTransformFields(n, fieldAccumulator)));
                }
                if (element.Name == W.p)
                {
                    var paraContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == W.t)
                        .Select(t => (string)t)
                        .StringConcatenate()
                        .Trim();
                    int occurances = CountSubstring(FieldRecognizer.EmbedBegin, paraContents);
                    if (occurances == 1
                        && paraContents.StartsWith(FieldRecognizer.EmbedBegin)
                        && paraContents.EndsWith(FieldRecognizer.EmbedEnd))
                    {
                        var content = paraContents
                            .Substring(FieldRecognizer.EmbedBegin.Length,
                                       paraContents.Length - FieldRecognizer.EmbedBegin.Length - FieldRecognizer.EmbedEnd.Length)
                            .Trim();
                        if (FieldRecognizer.IsField(content, out content))
                        {
                            var fieldId = fieldAccumulator.Count.ToString();
                            fieldAccumulator.Add(new FieldInfo(content, fieldId));
                            var ppr = element.Elements(W.pPr).FirstOrDefault();
                            var rpr = (ppr != null) ? ppr.Elements(W.rPr).FirstOrDefault() : null;
                            XElement r = new XElement(W.r, rpr,
                                new XElement(W.t, FieldRecognizer.FieldBegin + content + FieldRecognizer.FieldEnd));
                            return new XElement(element.Name,
                                element.Attributes(),
                                element.Elements(W.pPr),
                                CCTWrap(fieldId, r)
                            );
                        }
                        // else fall through to (slower) case
                    }
                    if (paraContents.Contains(FieldRecognizer.EmbedBegin))
                    {
                        var runReplacementInfo = new List<XElement>();
                        var placeholderText = Guid.NewGuid().ToString();
                        var r = new Regex(
                                Regex.Escape(FieldRecognizer.EmbedBegin)
                                + "\\s*"
                                + Regex.Escape(FieldRecognizer.FieldBegin)
                                + ".*?"
                                + Regex.Escape(FieldRecognizer.FieldEnd)
                                + "\\s*"
                                + Regex.Escape(FieldRecognizer.EmbedEnd));
                        OpenXmlRegex.Replace(new[] { element }, r, placeholderText, (para, match) =>
                        {
                            var matchString = match.Value.Trim();
                            var content = matchString.Substring(
                                    FieldRecognizer.EmbedBegin.Length,
                                    matchString.Length - FieldRecognizer.EmbedBegin.Length - FieldRecognizer.EmbedEnd.Length
                                ).CleanUpInvalidCharacters();
                            if (FieldRecognizer.IsField(content, out content))
                            {
                                runReplacementInfo.Add(CCWrap(new XElement(W.r, new XElement(W.t,
                                    FieldRecognizer.FieldBegin + content + FieldRecognizer.FieldEnd))));
                                return true;
                            }
                            return false;
                        }, false);

                        var newPara = new XElement(element);
                        foreach (var elem in runReplacementInfo)
                        {
                            var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == placeholderText
                                                                                             && rn.Parent.Name != Templater.OD.Content);
                            if (runToReplace == null)
                                throw new InvalidOperationException("Internal error");
                            else
                            {
                                //elem.Add(runToReplace); // does this work? what does it do?
                                runToReplace.ReplaceWith(elem);
                            }
                        }
                        var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                        return IdentifyAndTransformFields(coalescedParagraph, fieldAccumulator);
                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => IdentifyAndTransformFields(n, fieldAccumulator)));
            }
            return node;
        }

        static XElement CCWrap(params object[] content) => new XElement(W.sdt, new XElement(W.sdtContent, content));
        static XElement CCTWrap(string tag, params object[] content) =>
            new XElement(W.sdt,
                new XElement(W.sdtPr,
                    new XElement(W.tag, new XAttribute(W.val, tag))
                ),
                new XElement(W.sdtContent, content)
            );
        static XElement PWrap(params object[] content) => new XElement(W.p, content);
    }

    public static class StringFixerUpper
    {
        public static string CleanUpInvalidCharacters(this string fieldText)
        {
            return fieldText.Trim()
                            .Replace('“', '"') // replace curly quotes with straight ones
                            .Replace('”', '"')
                            .Replace("\u200b", null) // remove zero-width spaces -- potentially inserted via Macro or Word add-in for purposes of making the template look better
                            .Replace("\u200c", null);
        }
    }
}
