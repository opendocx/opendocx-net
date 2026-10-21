using System.Text.RegularExpressions;

namespace OpenDocx
{
    // Shared shallow recognition only. Expression parsing and pair matching stay
    // in FieldParser. Keep its historical ordering, boundaries and case sensitivity.
    internal static class FieldKindRecognizer
    {
        private static readonly (FieldType Kind, Regex Pattern)[] Patterns =
        {
            (FieldType.If, new Regex(@"^(?:if\b|\?)\s*(.*)$")),
            (FieldType.ElseIf, new Regex(@"^(?:elseif\b|:\?)\s*(.*)$")),
            (FieldType.Else, new Regex(@"^(?:else\b|:)(.*)?$")),
            (FieldType.EndIf, new Regex(@"^(?:endif\b|\/\?)(?:.*)$")),
            (FieldType.List, new Regex(@"^(?:list\b|#)\s*(.*)$")),
            (FieldType.EndList, new Regex(@"^(?:endlist\b|\/#)(.*)$"))
        };

        internal static Match Recognize(string content, out FieldType kind)
        {
            foreach (var pattern in Patterns)
            {
                var match = pattern.Pattern.Match(content);
                if (match.Success)
                {
                    kind = pattern.Kind;
                    return match;
                }
            }
            kind = FieldType.Content;
            return null;
        }
    }
}
