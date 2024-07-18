using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenDocx
{
    public static class FieldParser
    {
        private static Regex _ifRE = new(@"^(?:if\b|\?)\s*(.*)$");
        private static Regex _elseifRE = new(@"^(?:elseif\b|:\?)\s*(.*)$");
        private static Regex _elseRE = new(@"^(?:else\b|:)(.*)?$");
        private static Regex _endifRE = new(@"^(?:endif\b|\/\?)(?:.*)$");
        private static Regex _listRE = new(@"^(?:list\b|#)\s*(.*)$");
        private static Regex _endlistRE = new(@"^(?:endlist\b|\/#)(.*)$");
        // private static Regex _curlyquotes = /[“”]/g
        // private static Regex _zws = /[\u{200B}\u{200C}]/gu

        public static List<ParsedField> ParseContentArray(JsonArray contentArray)
        {
            // contentArray is an array of objects with field text and field IDs (as extracted from a DOCX template)
            // sub-arrays indicate hierarchical blocks of content (doc parts, paragraphs, table cells, etc.)
            var astBody = new List<ParsedField>();
            uint i = 0;
            while (i < contentArray.Count) // we use a 'while' because contentArray gets shorter as we go!
            {
                var parsedContentItem = ParseContentItem(i, contentArray);
                if (parsedContentItem.Count == 1)
                {
                    var parsedContent = parsedContentItem[0];
                    if (parsedContent.Type == FieldType.EndList || parsedContent.Type == FieldType.EndIf || parsedContent.Type == FieldType.Else || parsedContent.Type == FieldType.ElseIf)
                    {
                        // Field i's EndList/EndIf/Else/ElseIf has no matching List/If
                        var errMsg = string.Format("The {0} in field {1} has no matching {2}",
                            parsedContent.Type.ToString(),
                            parsedContent.Number.ToString(),
                            parsedContent.Type == FieldType.EndList ? "List" : "If");
                        throw new FieldParseException(errMsg);
                    }
                }
                astBody.AddRange(parsedContentItem); // js was: Array.prototype.push.apply(astBody, parsedContentItem)
                i++;
            }
            return astBody;
        }

        private static List<ParsedField> ParseContentItem(uint idx, JsonArray contentArray)
        {
            var contentItem = contentArray[(int)idx];
            var parsedItems = new List<ParsedField>();
            if (contentItem is JsonArray)
            {
                // if there's a sub-array, that item must be its own valid sequence of fields
                // with appropriately matched ifs/endifs and/or lists/endlists
                var parsedBlockContent = ParseContentArray(contentItem as JsonArray);
                parsedItems.AddRange(parsedBlockContent); // js was: Array.prototype.push.apply(parsedItems, parsedBlockContent)
            }
            else
            {
                var parsedContent = ParseField(contentArray, idx);
                if (parsedContent != null)
                {
                    parsedItems.Add(parsedContent);
                }
            }
            return parsedItems;
        }

        private static ParsedField ParseField(JsonArray contentArray, uint idx)
        {
            var contentArrayItem = contentArray[(int)idx];
            var content = contentArrayItem["content"].ToString();
            var fieldId = contentArrayItem["id"].ToString();
            // parse the field
            ParsedField node;
            var match = _ifRE.Match(content);
            if (match.Success)
            {
                node = CreateNode(FieldType.If, match.Groups[1].Value, fieldId);
                node.ContentArray = ParseContentUntilMatch(contentArray, idx + 1, FieldType.EndIf, node.Number);
                return node;
            }
            match = _elseifRE.Match(content);
            if (match.Success)
            {
                node = CreateNode(FieldType.ElseIf, match.Groups[1].Value, fieldId);
                node.ContentArray = new List<ParsedField>();
                return node;
            }
            if (_elseRE.IsMatch(content))
            {
                return CreateNode(FieldType.Else, null, fieldId, new List<ParsedField>());
            }
            if (_endifRE.IsMatch(content))
            {
                return CreateNode(FieldType.EndIf, null, fieldId);
            }
            match = _listRE.Match(content);
            if (match.Success)
            {
                node = CreateNode(FieldType.List, match.Groups[1].Value, fieldId);
                node.ContentArray = ParseContentUntilMatch(contentArray, idx + 1, FieldType.EndList, node.Number);
                return node;
            }
            if (_endlistRE.IsMatch(content))
            {
                return CreateNode(FieldType.EndList, null, fieldId);
            }
            // else
            return CreateNode(FieldType.Content, content.Trim(), fieldId);
        }

        public static ParsedField ParseFieldContent(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
                return CreateNode(FieldType.Content, "");
            // parse the content
            var match = _ifRE.Match(content);
            if (match.Success)
                return CreateNode(FieldType.If, match.Groups[1].Value);
            match = _elseifRE.Match(content);
            if (match.Success)
                return CreateNode(FieldType.ElseIf, match.Groups[1].Value);
            match = _elseRE.Match(content);
            if (match.Success)
            {
                var result = CreateNode(FieldType.Else);
                var comment = match.Groups[1].Value;
                if (!string.IsNullOrWhiteSpace(comment))
                    result.Comment = comment;
                return result;
            }
            match = _endifRE.Match(content);
            if (match.Success)
            {
                var result = CreateNode(FieldType.EndIf);
                var comment = match.Groups[1].Value;
                if (!string.IsNullOrWhiteSpace(comment))
                    result.Comment = comment;
                return result;
            }
            match = _listRE.Match(content);
            if (match.Success)
                return CreateNode(FieldType.List, match.Groups[1].Value);
            match = _endlistRE.Match(content);
            if (match.Success)
            {
                var result = CreateNode(FieldType.EndList);
                var comment = match.Groups[1].Value;
                if (!string.IsNullOrWhiteSpace(comment))
                    result.Comment = comment;
                return result;
            }
            // else
            return CreateNode(FieldType.Content, content.Trim());
        }

        private static List<ParsedField> ParseContentUntilMatch(JsonArray contentArray, uint startIdx, FieldType targetType, uint originId)
        {
            // parses WITHIN THE SAME CONTENT ARRAY (block) until it finds a field of the given targetType
            // returns a content array
            var idx = startIdx;
            var result = new List<ParsedField>();
            var parentContent = result;
            var elseEncountered = false;
            while (true)
            {
                if (idx >= contentArray.Count)
                {
                    // Field idx's List/If has no matching EndList/EndIf
                    var errMsg = string.Format("The {0} in field {1} has no matching {2}",
                        targetType == FieldType.EndList ? "List" : "If",
                        originId.ToString(),
                        targetType.ToString());
                    throw new FieldParseException(errMsg);
                }
                var parsedContent = ParseContentItem(idx, contentArray);
                ParsedField parsedContent0 = null;
                if (parsedContent.Count == 1)
                {
                    parsedContent0 = parsedContent[0];
                }
                idx++;
                if (parsedContent0 != null && parsedContent0.Type == targetType)
                {
                    if (parsedContent0.Type == FieldType.EndList)
                    {
                        // always insert a punctuation placeholder at the tail-end of every list) for DOCX templates.
                        // See "puncElem" in OpenDocx.Templater\Templater.cs
                        InjectListPunctuationNode(parentContent);
                    }
                    parentContent.Add(parsedContent0);
                    break;
                }
                foreach (var pc in parsedContent) {
                    if (pc != null)
                    {
                        parentContent.Add(pc);
                    }
                }
                if (parsedContent0 != null)
                {
                    string errMsg = null;
                    switch (parsedContent0.Type)
                    {
                        case FieldType.ElseIf:
                        case FieldType.Else:
                            if (targetType == FieldType.EndIf)
                            {
                                if (elseEncountered)
                                {
                                    // js: Encountered [field Y's|an] [Else/ElseIf] when expecting an EndIf (following [field X's|an] Else)
                                    // or: "Encountered an {0} in field {1} (after the Else in field {2}) when expecting an EndIf"
                                    errMsg = string.Format("The Else in field {0} needs a matching EndIf prior to the {1} in field {2}",
                                        originId.ToString(),
                                        parsedContent0.Type.ToString(),
                                        parsedContent0.Number.ToString());
                                }
                                if (parsedContent0.Type == FieldType.Else)
                                    elseEncountered = true;
                                if (errMsg == null)
                                    parentContent = parsedContent0.ContentArray;
                            }
                            else if (targetType == FieldType.EndList)
                            {
                                // js: Encountered [field Y's|an] [Else|ElseIf] when expecting [the end of field X's List|an EndList]
                                errMsg = string.Format("The List in field {0} needs a matching EndList prior to the {1} in field {2}",
                                    originId.ToString(),
                                    parsedContent0.Type.ToString(),
                                    parsedContent0.Number.ToString());
                            }
                            break;
                        case FieldType.EndIf:
                        case FieldType.EndList:
                            // Field X's EndIf/EndList has no matching If/List
                            errMsg = string.Format("Unexpected {0} in field {1} (could not locate the matching {2})",
                                parsedContent0.Type.ToString(),
                                parsedContent0.Number.ToString(),
                                parsedContent0.Type == FieldType.EndList ? "List" : "If");
                            break;
                    }
                    if (errMsg != null)
                        throw new FieldParseException(errMsg);
                }
            }
            // remove (consume) all parsed items from the contentArray before returning
            // js was: contentArray.splice(startIdx, idx - startIdx)
            while (idx > startIdx)
            {
                contentArray.RemoveAt((int)idx - 1);
                idx--;
            }
            return result;
        }

        private static void InjectListPunctuationNode(List<ParsedField> contentArray)
        {
            // synthesize list punctuation node
            var puncNode = CreateNode(FieldType.Content, "_punc");
            // field number == null because there is not (yet) a corresponding field in the template
            contentArray.Add(puncNode);
        }

        private static ParsedField CreateNode(FieldType type, string expr = null, string id = null, List<ParsedField> contentArray = null)
        {
            var newNode = new ParsedField();
            newNode.Type = type;
            if (expr != null) newNode.Expression = expr;
            if (id != null) newNode.Number = uint.Parse(id);
            if (contentArray != null) newNode.ContentArray = contentArray;
            return newNode;
        }

        public static void BuildFieldDictionary(List<ParsedField> astBody, Dictionary<string, ParsedField> fieldDict, FieldExprNamer atoms, ParsedField parent = null)
        {
            foreach (var obj in astBody)
            {
                if (obj.ContentArray != null)
                {
                    BuildFieldDictionary(obj.ContentArray, fieldDict, atoms, obj);
                }
                if (obj.Number > 0)
                {
                    var fieldObj = new ParsedField();
                    fieldObj.Type = obj.Type;
                    if (!string.IsNullOrEmpty(obj.Expression))
                    {
                        fieldObj.Expression = obj.Expression;
                        fieldObj.Atom = atoms.GetFieldAtom(obj);
                        // also cross-pollinate atomizedExpr across to ast (for later use)
                        obj.Atom = fieldObj.Atom;
                    }
                    else
                    {
                        fieldObj.ParentNumber = parent.Number;
                        // EndList fields are also stored with the atomized expression of their matching List field,
                        // because this is (or at least, used to be?) needed to make list punctuation work
                        if (obj.Type == FieldType.EndList)
                        {
                            fieldObj.Atom = atoms.GetFieldAtom(parent);
                        }
                    }
                    fieldDict[obj.Number.ToString()] = fieldObj;
                }
            }
        }
    }
}
