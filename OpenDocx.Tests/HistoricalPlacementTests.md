# Field placement regression tests

These synthetic fixtures protect compatibility between provisioning a source
DOCX and assembling equivalent precompiled OXPT field structures. They generate
all documents locally; no customer documents, external storage, or historical
OpenDocx binaries are required. Both paths use the current assembler.

## Behavior under test

- **Content fields:** any surrounding authored whitespace makes a content field
  run-level, including leading or trailing spaces, tabs, breaks, and NBSP.
  Filled, empty, and missing scalar values preserve surrounding content,
  paragraph properties, and supported floating anchors. Paragraph tab stops and
  Word's rendered-page-break hints do not themselves count as authored whitespace.
- **Control markers:** surrounding whitespace retains the existing block-level
  classification for otherwise solitary conditional/list markers, allowing pairs
  in separate paragraphs. Marker paragraphs and their whitespace are removed
  during assembly; body paragraph properties are preserved.
- **Empty text:** a field-only block content paragraph retains its paragraph and
  properties when merging an empty string. A false inline conditional also
  retains its containing paragraph. These tests do not define how a host
  application represents missing values or adds placeholders.
- **DOCX insertion:** a block content field's paragraph is replaced by incoming
  blocks. Zero incoming blocks remove it; one blank child paragraph replaces it
  with the child's paragraph rather than retaining the host's properties.
  For the tested run-level fields beside trailing whitespace, incoming blocks
  precede the surviving host paragraph, whose whitespace, properties, and supported
  anchor remain. This characterizes current behavior, not paragraph splitting or
  joining for arbitrary inline insertion.
- **Text boxes:** DrawingML, VML, and AlternateContent representations preserve
  positioning and independently merged nested fields. Both alternate XML
  representations are processed; this does not imply two visible boxes in Word.
  Nested fields must not hide an unsupported anchor on an outer block-field paragraph.

## Coverage and execution

The fixtures are partial definitions of `ImageParagraphTests`:

| File | Coverage |
| --- | --- |
| [HistoricalPlacementTests.cs](HistoricalPlacementTests.cs) | Headers with tabs and floating pictures, explicit page breaks, styled paragraphs with leading spaces, and cross-paragraph control pairs |
| [HistoricalTextBoxTests.cs](HistoricalTextBoxTests.cs) | Static and field-containing text boxes across modern, legacy, and alternate representations |
| [HistoricalInsertionTests.cs](HistoricalInsertionTests.cs) | Empty text versus empty/blank DOCX insertion, paragraphs and tables, assembled child documents, and whitespace beside insertion fields |
| [FieldPlacementTests.cs](FieldPlacementTests.cs) | Content placement, anchor diagnostics, and paragraph retention for empty content and false inline conditionals |

Coverage includes text-delimited and native content-control source templates.
Companion tests in [FieldWhitespaceTests.cs](FieldWhitespaceTests.cs) cover Unicode
separators and split runs, including NBSP preservation during `elseif` recognition.

Run the placement group from the repository root:

```sh
dotnet test OpenDocx.Tests/OpenDocx.Tests.csproj --filter FullyQualifiedName~ImageParagraphTests
```

This filter includes the companion whitespace tests. Omit it to run the full project.
Assertions compare document structure, text, paragraph properties, relationships,
and object positioning; namespace-declaration placement is ignored where it is
irrelevant to the comparison. These checks do not establish identical behavior
across all historical library versions or certify Word layout, complete run
formatting, numbering, or section behavior.
