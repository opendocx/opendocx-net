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
            // Retain the existing code for consumers and saved audit reports.
            Code = "FloatingImageOnBlockField";
            FieldId = fieldId;
            FieldText = fieldText;
            PartUri = partUri;
            Message = "A floating object (such as a picture, text box, or shape) is anchored to a paragraph containing a block-level field. " +
                "Move the object's anchor to a separate content paragraph, or change the object to inline if supported " +
                "and keep any matching fields in the same paragraph.";
        }

        public override string ToString() => $"Field {FieldId} (\"{FieldText}\"): {Message}";
    }
}
