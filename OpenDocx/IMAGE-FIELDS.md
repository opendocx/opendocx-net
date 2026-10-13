# Images and template fields

An inline picture is content in the paragraph's text flow. It counts as content
on either side of a field, just as ordinary text does. A field with an inline
picture beside it is run-level. Paired run-level fields must have their matching
fields in the same paragraph.

A floating picture is attached to a paragraph by its anchor. It does not count
as text-flow content. A paragraph containing only a field, whitespace, and
floating pictures has a block-level field.

Floating pictures attached to **any block-level field paragraph** are unsupported
at provisioning time. This applies to `if`, `elseif`, `else`, `endif`, `list`,
`endlist`, and ordinary content fields. Move the anchor to a separate paragraph
containing actual document content, or make the picture inline and use run-level
fields. Merely putting the anchor after the field does not resolve the error.

Floating pictures beside run-level fields are allowed when the paragraph also
has real content (including inline pictures). Images on content paragraphs inside
a block-level conditional or list are also allowed.

## Provisioning and audits

`Normalizer.NormalizeTemplate`, `Normalizer.ExtractFieldsOnly`, and
`OpenDocx.PrepareTemplate` enforce the rule and throw `FieldParseException` with
field-specific messages. Provisioning can report these using the existing
template-error path.

Existing `Normalizer.ExtractFieldsOnly` callers can now receive
`FieldParseException` for unsupported anchors instead of extracted JSON.
`Normalizer.ExtractFields`, `OpenDocx.NormalizeDocx`, and `OpenDocx.PrepareTemplate`
also use strict normalization. The extracted JSON schema and public signatures
are unchanged, but paragraph nesting can change under the clarified image rules.
`ExtractFieldsOnly` uses the same in-memory processing and diagnostics but skips
the final normalized DOCX byte-array copy. It still needs a working document in
memory. `AnalyzeTemplate` and `NormalizeTemplate` continue returning DOCX bytes.
`Templater.ParseFields`, `Templater.ParseFieldsToDict`, and `OpenDocx.ParseDocx`
have unchanged parsing rules; newly extracted nesting can make previously
accepted field pairs fail their existing checks.

Structured analysis diagnostics include `PartUri`. The exception messages from
strict normalization include the field ID, expression, and diagnostic message,
without the document-part URI. The separate legacy `FieldExtractor` API has not
been updated with these image rules.

An audit of an existing DOCX can use:

```csharp
var analysis = Normalizer.AnalyzeTemplate(originalDocxBytes);
foreach (var diagnostic in analysis.Diagnostics)
{
    // Code == "FloatingImageOnBlockField"
    // FieldId, FieldText, PartUri, Message
}
```

This uses the same classification and diagnostic check as provisioning. It
normalizes an in-memory copy and returns normalized bytes and extracted fields
alongside the diagnostics. It does not parse field pairs, compile, change the
source bytes, or write files. Other malformed-template errors can still throw;
an audit must report these as failures rather than claiming a clean result.

For this image-rule migration audit, also run
`Templater.ParseFields(analysis.ExtractedFields, false)` and report any
`FieldParseException` alongside the analysis diagnostics.
This catches matching-field failures when an inline image makes a marker
run-level. Full compilation is not required for this audit.

The classifier inspects DrawingML `wp:inline` and `wp:anchor`, as well as legacy
VML pictures positioned with `position:absolute` or `position:relative` on the
shape or an enclosing VML group.
Unknown non-text run children remain content. No UnicodeMapper sentinel meanings
are changed, so run and field offsets stay intact.

## Assembly and old templates

If a floating picture reaches compilation on a block-field paragraph, its anchor
is discarded with that source paragraph. It is not moved into another paragraph
to preserve it. Provisioning normally prevents such templates from being newly
installed.

All image-specific checks and transformations occur during provisioning. The
assembler has no additional image scan, compatibility pass, or generated
conditional wrappers. Existing behavior for image-free content fields, including
empty values and list punctuation, is retained. OpenXmlPowerTools is unchanged.

Previously compiled templates are assembled using their existing metadata. The
audit identifies original source templates affected by the clarified image rules;
it does not promise identical output after those templates are re-provisioned.

Earlier compiler versions could move an image off its marker paragraph into a
new image paragraph. Such compiled artifacts no longer contain the original
attachment information; they need correction in the original template and
re-provisioning. An old pinned App needs a corrected, newly compiled version and
a pin update. This change does not rewrite historical S3 objects.

The consuming Lambda must use the updated OpenDocx package for upload checks.
Its editing-only `FixTemplateFields` handler uses `AnalyzeTemplate` so
Word add-in formatting remains available while authors repair an invalid anchor.
The existing-template scanner and Issues integration are separate follow-up work.
