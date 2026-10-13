using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenXmlPowerTools;

namespace OpenDocx
{
    // Keep UnicodeMapper's character positions intact. Only classification ignores
    // floating objects; extraction/replacement still uses the original run text.
    internal static class ParagraphImages
    {
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        private static readonly HashSet<XName> Invisibles = new()
        {
            W.rPr, W.fldChar, W.instrText, W.lastRenderedPageBreak
        };

        internal static bool RunIsVisible(XElement run) =>
            run.Elements(W.t).Any() || !run.Elements().All(e => Invisibles.Contains(e.Name));

        internal static bool IsFloating(XElement child)
        {
            if (child.Name == Mc + "AlternateContent")
            {
                var branch = child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback");
                return branch != null && branch.Elements().Any() && branch.Elements().All(IsFloating);
            }
            if (child.Name == W.drawing)
                return child.Elements(WP.anchor).Any() && !child.Elements(WP.inline).Any();
            if (child.Name == W.pict)
            {
                var shapes = child.Descendants(V + "shape").ToList();
                return shapes.Any() && shapes.All(shape => VmlShapeIsFloating(shape, child));
            }
            return false;
        }

        // Classification ignores an object only when it is entirely floating.
        // Diagnostics must also find anchors in objects with mixed content.
        private static bool ContainsFloating(XElement child)
        {
            if (child.Name == Mc + "AlternateContent")
            {
                var branch = child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback");
                return branch != null && branch.Elements().Any(ContainsFloating);
            }
            if (child.Name == W.drawing)
                return child.Elements(WP.anchor).Any();
            if (child.Name == W.pict)
                return child.Descendants(V + "shape").Any(shape => VmlShapeIsFloating(shape, child));
            return false;
        }

        // A group can supply the floating position for its shapes.
        private static bool VmlShapeIsFloating(XElement shape, XElement picture) => shape.AncestorsAndSelf()
            .TakeWhile(element => element != picture)
            .Where(element => element.Name == V + "shape" || element.Name == V + "group")
            .Any(element => Regex.IsMatch((string)element.Attribute("style") ?? "",
                @"(?:^|;)\s*position\s*:\s*(absolute|relative)\s*(?:;|$)", RegexOptions.IgnoreCase));

        internal static bool HasFloating(XElement paragraph) => paragraph
            .DescendantsTrimmed(W.txbxContent)
            .Where(e => e.Name == W.r && !e.Ancestors(W.del).Any())
            .SelectMany(r => r.Elements()).Any(ContainsFloating);

        internal static bool IsInline(XElement child)
        {
            if (child.Name == Mc + "AlternateContent")
            {
                var branch = child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback");
                return branch != null && branch.Elements().Any(IsInline);
            }
            return (child.Name == W.drawing || child.Name == W.pict) && !IsFloating(child);
        }

        internal static string ClassificationText(XElement run) => !RunIsVisible(run) ? "" : run.Elements()
            .Select(child => IsFloating(child)
                ? new string(' ', UnicodeMapper.RunToString(child).Length)
                : UnicodeMapper.RunToString(child)).StringConcatenate();
    }
}
