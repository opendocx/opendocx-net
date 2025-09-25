/***************************************************************************

Copyright (c) Lowell Stewart 2019-2023.
Licensed under the Mozilla Public License. See LICENSE file in the project root for full license information.

Published at https://github.com/opendocx/opendocx
Developer: Lowell Stewart
Email: lowell@opendocx.com

Uses a Recursive Pure Functional Transform (RPFT) approach to process a DOCX file and extract "field" metadata.
"Fields" may be either in regular text runs (delimited by special characters) or in content controls,
or any mixture thereof.

In the process, fields are all normalized so they are uniformly contained in content controls.
The process produces generic JSON metadata about all fields thus located, which includes depth indicators
so matching begin/end fields can be detected/enforced.

General RPFT approach was adapted from the Open XML Power Tools project. Those parts may contain...
  Portions Copyright (c) Microsoft. All rights reserved.
  Portions Copyright (c) Eric White Inc. All rights reserved.

***************************************************************************/
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomProperties;

namespace OpenDocx
{
    public class Normalizer
    {
        // public static FieldExtractResult ExtractFields(string templateFileName,
        //     bool removeCustomProperties = true, IEnumerable<string> keepPropertyNames = null,
        //     string fieldDelimiters = null)
        // {
        //     string newTemplateFileName = templateFileName + "obj.docx";
        //     string outputFile = templateFileName + "obj.json";
        //     WmlDocument templateDoc = new WmlDocument(templateFileName); // just reads the template's bytes into memory (that's all), read-only

        //     var result = NormalizeTemplate(templateDoc.DocumentByteArray, removeCustomProperties, keepPropertyNames, fieldDelimiters);
        //     // save the output (even in the case of error, since error messages are in the file)
        //     var preprocessedTemplate = new WmlDocument(newTemplateFileName, result.NormalizedTemplate);
        //     preprocessedTemplate.Save();
        //     // also save extracted fields
        //     File.WriteAllText(outputFile, result.ExtractedFields);
        //     return new FieldExtractResult(newTemplateFileName, outputFile);
        // }

        public static NormalizeResult NormalizeTemplate(byte[] templateBytes, bool removeCustomProperties = true,
            IEnumerable<string> keepPropertyNames = null, string fieldDelimiters = null)
        {
            var fieldAccumulator = new FieldAccumulator();
            var recognizer = FieldRecognizer.Default;
            OpenSettings openSettings = new OpenSettings();
            if (!string.IsNullOrWhiteSpace(fieldDelimiters))
            {
                recognizer = new FieldRecognizer(fieldDelimiters, null);
                // commented out, because this causes corruption in some templates??
                // openSettings.MarkupCompatibilityProcessSettings =
                //     new MarkupCompatibilityProcessSettings(
                //         MarkupCompatibilityProcessMode.ProcessAllParts, 
                //         DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            }
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateBytes, 0, templateBytes.Length); // copy template bytes into memory stream
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true, openSettings)) // read & parse that byte array into OXML document (also in memory)
                {
                    // first, remove all the task panes / web extension parts from the template (if there are any)
                    wordDoc.DeleteParts<WebExTaskpanesPart>(wordDoc.GetPartsOfType<WebExTaskpanesPart>());
                    // next, extract all fields (and thus logic) from the template's content parts
                    // (also normalizes field structure, replacing plain text fields with content controls)
                    ExtractAllTemplateFields(wordDoc, recognizer, fieldAccumulator, false,
                        removeCustomProperties, keepPropertyNames);
                }
                using (var sw = new StringWriter())
                {
                    fieldAccumulator.JsonSerialize(sw);
                    return new NormalizeResult(mem.ToArray(), sw.ToString());
                }
            }
        }

        public static string ExtractFieldsOnly(byte[] docxBytes, string fieldDelimiters = null)
        {
            var fieldAccumulator = new FieldAccumulator();
            var recognizer = FieldRecognizer.Default;
            OpenSettings openSettings = new OpenSettings();
            if (!string.IsNullOrWhiteSpace(fieldDelimiters))
            {
                recognizer = new FieldRecognizer(fieldDelimiters, null);
            }
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(docxBytes, 0, docxBytes.Length); // copy template bytes into memory stream
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true, openSettings)) // read & parse those bytes into OXML document (also in memory)
                {
                    // next, extract all fields (and thus logic) from the template's content parts
                    ExtractAllTemplateFields(wordDoc, recognizer, fieldAccumulator, false, false, null);
                }
            }
            using (var sw = new StringWriter())
            {
                fieldAccumulator.JsonSerialize(sw);
                return sw.ToString();
            }
        }

        private static void ExtractAllTemplateFields(WordprocessingDocument wordDoc, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, bool readFieldComments, bool removeCustomProperties = true,
            IEnumerable<string> keepPropertyNames = null)
        {
            if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                throw new FieldParseException("Invalid template - contains tracked revisions");

            CommentReader comments = null;
            if (readFieldComments)
            {
                comments = new CommentReader(wordDoc);
            }

            // extract fields from each part of the document
            foreach (var part in wordDoc.ContentParts())
            {
                ExtractFieldsFromPart(part, recognizer, fieldAccumulator, comments);

                if (removeCustomProperties)
                {
                    // remove document variables and custom properties
                    // (in case they have any sensitive information that should not carry over to assembled documents!)
                    MainDocumentPart main = part as MainDocumentPart;
                    if (main != null)
                    {
                        var docVariables = main.DocumentSettingsPart.Settings.Descendants<DocumentVariables>();
                        foreach (DocumentVariables docVars in docVariables.ToList())
                        {
                            foreach (DocumentVariable docVar in docVars.ToList())
                            {
                                if (keepPropertyNames == null || !Enumerable.Contains<string>(keepPropertyNames, docVar.Name))
                                {
                                    docVar.Remove();
                                    //docVar.Name = "Id";
                                    //docVar.Val.Value = "123";
                                }
                            }
                        }
                    }
                }
            }
            if (removeCustomProperties)
            {
                // remove custom properties if there are any (custom properties are the new/non-legacy version of document variables)
                var custom = wordDoc.CustomFilePropertiesPart;
                if (custom != null)
                {
                    foreach (CustomDocumentProperty prop in custom.Properties.ToList())
                    {
                        if (keepPropertyNames == null || !Enumerable.Contains<string>(keepPropertyNames, prop.Name))
                        {
                            prop.Remove();
                            // string propName = prop.Name;
                            // string value = prop.VTLPWSTR.InnerText;
                        }
                    }
                }
            }
        }

        private static void ExtractFieldsFromPart(OpenXmlPart part, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            XDocument xDoc = part.GetXDocument();
            fieldAccumulator.BeginBlock();
            var xDocRoot = (XElement)IdentifyAndNormalizeFields(xDoc.Root, recognizer, fieldAccumulator, comments);
            fieldAccumulator.EndBlock();
            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.PutXDocument();
        }

        private static object IdentifyAndNormalizeFields(XNode node, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.sdt)
                {
                    throw new Exception("Content controls not currently supported in templates - please contact Support");
                }
                if (element.Name == W.p)
                {
                    fieldAccumulator.BeginBlock();
                    var transformedPara = ProcessParagraphContent(element, recognizer, fieldAccumulator, comments);
                    fieldAccumulator.EndBlock();
                    return transformedPara;
                }
                else if (element.Name == W.lastRenderedPageBreak)
                {
                    // documents assembled from templates will almost always change pagination, so remove Word's pagination hints
                    // (also because they're not handled cleanly by OXPT)
                    return null;
                }
                else if (element.Name == W.r)
                {
                    return ProcessTextRun(element, recognizer, fieldAccumulator, comments);
                }
                // For all other elements, just process children
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => IdentifyAndNormalizeFields(n, recognizer, fieldAccumulator, comments)));
            }
            return node;
        }

        private static XElement ProcessParagraphContent(XElement element, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            var paraContents = ExtractTextContent(element);
            if (!paraContents.Contains(recognizer.CombinedBegin))
            {
                return ProcessContentNoFields(element, recognizer, fieldAccumulator, comments);
            }

            // PRE-SPLIT RUNS so field boundaries align with run boundaries
            var paragraphWithSplitRuns = PreSplitRunsForFields(element, recognizer);
            
            // Now we can process with simple run replacement
            return ProcessParagraphWithAlignedRuns(paragraphWithSplitRuns, recognizer, fieldAccumulator, comments);
        }

        private static XElement PreSplitRunsForFields(XElement paragraph, FieldRecognizer recognizer)
        {
            // Get all field match positions
            var paraContents = ExtractTextContent(paragraph);
            var matches = recognizer.Regex.Matches(paraContents);
    
            if (matches.Count == 0) return paragraph;
    
            // Collect all positions where we need to split runs
            var splitPositions = new HashSet<int>();
            foreach (Match match in matches.Cast<Match>())
            {
                splitPositions.Add(match.Index);           // Start of field
                splitPositions.Add(match.Index + match.Length); // End of field
            }
    
            // Split runs at these positions
            return SplitRunsAtPositions(paragraph, splitPositions.OrderBy(x => x).ToList());
        }

        private static XElement SplitRunsAtPositions(XElement paragraph, List<int> splitPositions)
        {
            if (splitPositions.Count == 0) return paragraph;
    
            var runs = paragraph.DescendantsTrimmed(W.txbxContent)
                .Where(d => d.Name == W.r && (d.Parent == null || d.Parent.Name != W.del))
                .ToList();
    
            var newRuns = new List<XElement>();
            int currentPosition = 0;
    
            foreach (var run in runs)
            {
                var runText = UnicodeMapper.RunToString(run);
                if (string.IsNullOrEmpty(runText))
                {
                    newRuns.Add(run);
                    continue;
                }
    
                var runProperties = run.Elements(W.rPr).FirstOrDefault();
                int runStart = currentPosition;
                int runEnd = currentPosition + runText.Length;
    
                // Find split positions within this run
                var splitsInRun = splitPositions
                    .Where(pos => pos > runStart && pos < runEnd)
                    .Select(pos => pos - runStart)
                    .ToList();

                if (splitsInRun.Count == 0)
                {
                    // No splits needed in this run
                    newRuns.Add(run);
                }
                else
                {
                    // Split this run
                    int lastSplit = 0;
                    foreach (var splitPos in splitsInRun)
                    {
                        if (splitPos > lastSplit)
                        {
                            var textSegment = runText[lastSplit..splitPos];
                            newRuns.Add(new XElement(W.r,
                                runProperties != null ? new XElement(runProperties) : null,
                                CreateTextElement(textSegment)));
                        }
                        lastSplit = splitPos;
                    }
                    // Add the final segment
                    if (lastSplit < runText.Length)
                    {
                        var textSegment = runText[lastSplit..];
                        newRuns.Add(new XElement(W.r,
                            runProperties != null ? new XElement(runProperties) : null,
                            CreateTextElement(textSegment)));
                    }
                }
                currentPosition = runEnd;
            }
            // Rebuild paragraph with split runs
            return ReplaceParagraphRuns(paragraph, runs, newRuns);
        }

        private static XElement CreateTextElement(string textContent)
        {
            // Check if the text starts or ends with significant whitespace
            bool needsXmlSpace = textContent.Length > 0
                && (char.IsWhiteSpace(textContent[0]) || char.IsWhiteSpace(textContent[^1]));
            XAttribute xmlSpace = needsXmlSpace ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
            return new XElement(W.t, xmlSpace, textContent);
        }

        private static XElement ReplaceParagraphRuns(XElement paragraph, List<XElement> originalRuns, List<XElement> newRuns)
        {
            if (originalRuns.Count == 0 || newRuns.Count == 0)
                return paragraph;

            // Simple case: if all original runs are direct children of the paragraph
            if (originalRuns.All(r => r.Parent == paragraph))
            {
                var newParagraph = new XElement(paragraph.Name, paragraph.Attributes());

                bool runsReplaced = false;
                foreach (var node in paragraph.Nodes())
                {
                    if (node is XElement element && element.Name == W.r && originalRuns.Contains(element))
                    {
                        // Replace all original runs with new runs (but only on the first match)
                        if (!runsReplaced)
                        {
                            foreach (var newRun in newRuns)
                            {
                                newParagraph.Add(newRun);
                            }
                            runsReplaced = true;
                        }
                        // Skip this original run
                    }
                    else
                    {
                        // Keep other nodes as-is
                        newParagraph.Add(node);
                    }
                }

                return newParagraph;
            }
            else
            {
                // Complex case with nested runs - fall back to original paragraph
                // This shouldn't happen often with the DescendantsTrimmed approach
                return paragraph;
            }
        }

        private static XElement ProcessParagraphWithAlignedRuns(XElement paragraph, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            var paraContents = ExtractTextContent(paragraph);
            var matches = recognizer.Regex.Matches(paraContents);
            
            // FIRST: Register all fields in forward order for the accumulator
            var fieldReplacements = new List<(Match match, string fieldId, XElement contentControl)>();
            
            foreach (Match match in matches.Cast<Match>().OrderBy(m => m.Index)) // Forward order
            {
                var matchString = match.Value.Trim().Replace("\u0001", "");
                var fieldContent = matchString
                    .Substring(recognizer.EmbedBeginLength, matchString.Length - recognizer.EmbedDelimLength)
                    .CleanUpInvalidCharacters();
                    
                if (recognizer.IsField(fieldContent, out fieldContent))
                {
                    // Register field in forward order
                    var fieldId = fieldAccumulator.AddField(fieldContent);
                    var contentControl = CCTWrap(fieldId, new XElement(W.r, new XElement(W.t, '[' + fieldContent + ']')));
                    
                    fieldReplacements.Add((match, fieldId, contentControl));
                }
            }

            if (fieldReplacements.Count == 0)
            {
                return paragraph;
            }
            
            // check if single-field paragraphs have any non-whitespace text besides that field
            if (fieldReplacements.Count == 1)
            {
                var fld = fieldReplacements[0].match;
                var nonFieldText = paraContents[..fld.Index] + paraContents[(fld.Index + fld.Length)..];
                if (!string.IsNullOrWhiteSpace(nonFieldText))
                {
                    fieldAccumulator.RegisterNonFieldContentInBlock();
                }
            }

            // Apply XML replacements in reverse order to avoid position shifts
            var newParagraph = new XElement(paragraph);
            foreach (var (match, fieldId, contentControl) in fieldReplacements.OrderByDescending(fr => fr.match.Index))
            {
                ReplaceMatchWithContentControl(newParagraph, match, contentControl);
            }
            return WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newParagraph);
        }

        private static void ReplaceMatchWithContentControl(XElement paragraph, Match match, XElement contentControl)
        {
            var runs = paragraph.DescendantsTrimmed(W.txbxContent)
                .Where(d => d.Name == W.r && (d.Parent == null || d.Parent.Name != W.del))
                .ToList();
    
            int currentPosition = 0;
            var runsToReplace = new List<XElement>();
    
            foreach (var run in runs)
            {
                var runText = UnicodeMapper.RunToString(run);
                int runStart = currentPosition;
                int runEnd = currentPosition + runText.Length;
        
                // Check if this run overlaps with the match
                if (runStart < match.Index + match.Length && runEnd > match.Index)
                {
                    runsToReplace.Add(run);
                }
        
                currentPosition = runEnd;
        
                if (currentPosition > match.Index + match.Length)
                    break;
            }
    
            if (runsToReplace.Count > 0)
            {
                // Preserve formatting from first run
                var firstRun = runsToReplace.First();
                var runProperties = firstRun.Elements(W.rPr).FirstOrDefault();
        
                if (runProperties != null)
                {
                    var contentRun = contentControl.Element(W.sdtContent)?.Element(W.r);
                    contentRun?.AddFirst(new XElement(runProperties));
                }
        
                // Replace first run with content control, remove others
                firstRun.ReplaceWith(contentControl);
                foreach (var runToRemove in runsToReplace.Skip(1))
                {
                    runToRemove.Remove();
                }
            }
        }
        
        private static XElement ProcessContentNoFields(XElement element, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            var transformedChildren = new List<object>();
            foreach (var child in element.Nodes())
            {
                if (child is XElement childElement)
                {
                    if (childElement.Name == W.lastRenderedPageBreak)
                    {
                        continue; // Skip this element
                    }
                    else if (childElement.Name == W.r)
                    {
                        transformedChildren.Add(
                            ProcessTextRun(childElement, recognizer, fieldAccumulator, comments)
                        );
                    }
                    else
                    {
                        // Recursively process - might(?) find nested paragraphs with fields
                        transformedChildren.Add(
                            IdentifyAndNormalizeFields(child, recognizer, fieldAccumulator, comments)
                        );
                    }
                }
                else
                {
                    transformedChildren.Add(child);
                }
            }
            return new XElement(element.Name, element.Attributes(), transformedChildren);
        }

        private static string ExtractTextContent(XElement element)
        {
            if (element.Name == W.r)
            {
                return UnicodeMapper.RunToString(element);
            }

            // Use the same pattern as OpenXmlRegex for consistent text extraction
            return element
                .DescendantsTrimmed(W.txbxContent)
                .Where(d => d.Name == W.r && (d.Parent == null || d.Parent.Name != W.del))
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
        }

        private static XElement ProcessTextRun(XElement runElement, FieldRecognizer recognizer,
            FieldAccumulator fieldAccumulator, CommentReader comments)
        {
            if (runElement.Elements(W.t).Any(t => !string.IsNullOrWhiteSpace((string)t)))
            {
                // apparently, spaces and non-text outside of a field will (in at least some cases?)
                // get ignored in assembly, even for block-level if's, so we only note non-spaces.
                fieldAccumulator.RegisterNonFieldContentInBlock();
            }
            // Always process children recursively to maintain pure functional approach
            return new XElement(runElement.Name,
                runElement.Attributes(),
                runElement.Nodes().Select(n => IdentifyAndNormalizeFields(n, recognizer, fieldAccumulator, comments)));
        }

        private class TextRunInfo
        {
            public XElement Run { get; set; }
            public XElement TextElement { get; set; }
            public string Text { get; set; }
            public int StartPosition { get; set; }
            public int Length { get; set; }
            public XElement RunProperties { get; set; }
        }

        private class FieldReplacement
        {
            public System.Text.RegularExpressions.Match Match { get; set; }
            public XElement ContentControl { get; set; }
            public List<XElement> AffectedRuns { get; set; }
            public string FieldContent { get; set; }
        }

        static XElement CCTWrap(string tag, params object[] content) =>
            new XElement(W.sdt,
                new XElement(W.sdtPr,
                    new XElement(W.tag, new XAttribute(W.val, tag))
                ),
                new XElement(W.sdtContent, content)
            );
    }

    // public static class StringFixerUpper
    // {
    //     public static string CleanUpInvalidCharacters(this string fieldText)
    //     {
    //         return fieldText.Trim()
    //                         .Replace('“', '"') // replace curly quotes with straight ones
    //                         .Replace('”', '"')
    //                         .Replace("\u200b", null) // remove zero-width spaces -- potentially inserted via Macro or Word add-in for purposes of allowing word wrap
    //                         .Replace("\u200c", null);
    //     }
    // }
}
