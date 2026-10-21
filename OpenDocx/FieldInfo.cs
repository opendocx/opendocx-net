using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OpenXmlPowerTools;

namespace OpenDocx
{
    public class FieldAccumulator
    {
        private Stack<FieldBlock> blocks;
        private int fieldCount;
        internal string PartUri { get; set; }
        internal List<TemplateDiagnostic> Diagnostics { get; } = new();

        internal void CompleteParagraphPlacement(XElement paragraph)
        {
            var block = blocks.Peek();
            // Fields inside text boxes have their own nested groups. They must not
            // hide an unsupported anchor on the outer paragraph's sole field.
            var fields = block.DirectFields.ToList();
            if (fields.Count != 1) return;
            var field = fields[0];
            var control = paragraph.DescendantsTrimmed(W.txbxContent)
                .FirstOrDefault(e => e.Name == W.sdt &&
                    (string)e.Element(W.sdtPr)?.Element(W.tag)?.Attribute(W.val) == field.id);
            if (field.Kind == FieldType.Content && control != null &&
                !control.Elements(W.sdtContent).Elements(W.p).Any() &&
                FieldPlacement.ContentRequiresRunLevel(paragraph, control))
                block.RegisterOtherContent();
            if (!block.HasOtherContent && ParagraphImages.HasFloating(paragraph))
                Diagnostics.Add(new TemplateDiagnostic(field.id, field.content, PartUri));
        }

        internal void CheckBlockControlImages(XElement control, string fieldId, string fieldText)
        {
            if (control.Elements(W.sdtContent)
                .SelectMany(content => content.DescendantsTrimmed(W.txbxContent))
                .Any(e => e.Name == W.p && ParagraphImages.HasFloating(e)))
                Diagnostics.Add(new TemplateDiagnostic(fieldId, fieldText, PartUri));
        }

        public FieldAccumulator()
        {
            blocks = new Stack<FieldBlock>();
            blocks.Push(new FieldBlock());
            fieldCount = 0;
        }

        public string AddField(string contents)
        {
            fieldCount++;
            var fieldId = fieldCount.ToString();
            blocks.Peek().AddField(new FieldInfo(contents, fieldId));
            return fieldId;
        }

        public void BeginBlock()
        {
            blocks.Push(new FieldBlock());
        }

        public void RegisterNonFieldContentInBlock()
        {
            blocks.Peek().RegisterOtherContent();
        }

        public void EndBlock(bool promoteSingleField = false)
        {
            var block = blocks.Pop();
            if (!block.IsEmpty)
            {
                if (promoteSingleField && block.FieldCount == 1 && !block.HasOtherContent)
                    blocks.Peek().AddField(block.GetItem(0));
                else
                    blocks.Peek().AddField(block);
            }
        }

        public void JsonSerialize(TextWriter writer)
        {
            System.Diagnostics.Debug.Assert(blocks.Count == 1);
            blocks.Peek().JsonSerialize(writer);
        }

        private interface ExtractedItem
        {
            void JsonSerialize(TextWriter writer);
        }

        private class FieldBlock : ExtractedItem
        {
            private List<ExtractedItem> list;

            public FieldBlock()
            {
                list = new List<ExtractedItem>();
                HasOtherContent = false;
            }

            public bool IsEmpty { get { return list.Count == 0; } }

            public int FieldCount { get { return list.Count; } }

            public IEnumerable<FieldInfo> DirectFields => list.OfType<FieldInfo>();

            public bool HasOtherContent { get; private set; }

            public ExtractedItem GetItem(int index)
            {
                return list[index];
            }

            public void AddField(ExtractedItem field)
            {
                list.Add(field);
            }
            public void RegisterOtherContent()
            {
                HasOtherContent = true;
            }
            public void JsonSerialize(TextWriter sw)
            {
                sw.Write('[');
                bool first = true;
                foreach (var field in list)
                {
                    if (first)
                        first = false;
                    else
                        sw.Write(',');
                    field.JsonSerialize(sw);
                }
                sw.Write(']');
            }
        }

        private class FieldInfo : ExtractedItem
        {
            public string content { get; }
            public string id { get; }
            public FieldType Kind { get; }
            public FieldInfo(string fieldContent, string fieldId)
            {
                content = fieldContent;
                id = fieldId;
                FieldKindRecognizer.Recognize(fieldContent, out var kind);
                Kind = kind;
            }
            public void JsonSerialize(TextWriter sw)
            {
                sw.Write('{');
                sw.Write("\"content\":\"");
                sw.Write(content.Replace(@"\", @"\\").Replace(@"""", @"\""")
                    .Replace("\r", String.Empty).Replace("\n", @"\\n"));
                sw.Write("\",\"id\":\"");
                sw.Write(id);
                sw.Write("\"");
                sw.Write('}');
            }
        }
    }
}
