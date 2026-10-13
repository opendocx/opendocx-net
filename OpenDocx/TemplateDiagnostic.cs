namespace OpenDocx
{
    public sealed class TemplateDiagnostic
    {
        public string Code { get; }
        public string FieldId { get; }
        public string FieldText { get; }
        public string PartUri { get; }
        public string Message { get; }

        internal TemplateDiagnostic(string fieldId, string fieldText, string partUri)
        {
            Code = "FloatingImageOnBlockField";
            FieldId = fieldId;
            FieldText = fieldText;
            PartUri = partUri;
            Message = "A floating image is anchored to a paragraph containing a block-level field. " +
                "Move the image anchor to a separate content paragraph, or change the image to inline " +
                "and keep any matching fields in the same paragraph.";
        }

        public override string ToString() => $"Field {FieldId} (\"{FieldText}\"): {Message}";
    }
}
