using System.Xml.Linq;
using OpenXmlPowerTools;

namespace OpenDocx
{
    internal static class FieldPlacement
    {
        // Use the normalized XML, not character offsets: floating objects occupy
        // UnicodeMapper placeholders but are not leading whitespace. The field is
        // an SDT during extraction and an OD.Content node during transformation.
        internal static bool ContentRequiresRunLevel(XElement paragraph, XElement field)
        {
            return Visit(paragraph);

            bool Visit(XElement element)
            {
                if (ReferenceEquals(element, field))
                    return false;
                if (element.Name == W.pPr || element.Name == W.rPr || element.Name == W.sdtPr ||
                    element.Name == W.txbxContent || element.Name == W.del ||
                    element.Name == W.lastRenderedPageBreak || element.Name == W.fldChar || element.Name == W.instrText)
                    return false;
                if (ParagraphImages.IsFloating(element)) return false;
                if (element.Name == W.t)
                    return element.Value.Length != 0;
                // Authored whitespace on either side belongs to the host paragraph.
                if (element.Name == W.tab) return true;
                if (element.Parent?.Name == W.r)
                    return true; // Explicit breaks, inline objects, symbols, references, etc.
                foreach (var child in element.Elements())
                    if (Visit(child)) return true;
                return false;
            }
        }
    }
}
