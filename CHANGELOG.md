# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.8] - 2026-10-21

### Added
- Document the provisioning/assembly boundary, host-supplied expression evaluation, and a standalone C# example in the repository and NuGet README. The AWS Lambda example includes offline provisioning-to-assembly tests.
- Whitespace regression tests compare 43 additional Unicode and run-boundary cases with Microsoft Word's saved output, retaining the original 107-case comparison. Mapper round trips, regex replacement offsets, and OpenDocx conditional assembly cover NBSP separators with `xml:space` absent, `default`, and `preserve`.
- Historical regression fixtures cover floating pictures in headers, DrawingML/VML text boxes and their nested fields, empty/missing values, and DOCX insertion after a page break. Source-to-assembly checks cover both text-delimited and native content-control templates.
- Placement tests cover leading/trailing whitespace, explicit breaks, NBSP and other Unicode whitespace, split/shared runs, shorthand markers, and keyword-like content expressions. Compatibility tests protect existing cross-paragraph conditional/list matching.
- Empty-insertion regression fixtures distinguish empty text (host paragraph retained), zero-block DOCX insertion (host paragraph removed), and a DOCX containing a blank paragraph (host replaced by the child's paragraph), including a child emptied by a false block conditional.
- Trailing-whitespace insertion fixtures compare newly provisioned fields with existing run-level compiled metadata for spaces/tabs, floating anchors, and empty/blank/nonempty child documents. They protect current insertion behavior without implementing future paragraph joining.

### Fixed
- The bundled `UnicodeMapper` preserves NBSP and other Unicode whitespace at text-element boundaries without `xml:space="preserve"`. Only XML whitespace (space, tab, CR, and LF) is trimmed. Split-run fields retain their source separators, allowing `elseif` fields to be recognized and paired correctly without broadening keyword recognition.
- Content fields with any surrounding authored whitespace, including ordinary trailing spaces/tabs, explicit breaks, and Unicode whitespace, now remain run-level during extraction. Surrounding content and paragraph properties are preserved for filled, empty, and missing scalar values. Control-marker whitespace treatment is unchanged. Truly block-level content fields merging empty text continue to retain their paragraphs; automatic empty-paragraph cleanup is outside this change.
- Content-field anchor diagnostics and compilation use the same placement rule, preserving supported floating objects instead of promoting their fields and discarding their paragraphs' surrounding content. Fields inside a text box no longer hide an unsupported anchor on a block-field paragraph.
- Text-delimited fields inside text boxes are normalized and extracted even when their containing paragraph also has a text-delimited field. Nested paragraph scopes and alternate object representations are retained.

### Changed
- The illustrative Lambda project, its tests, and SAM runtime settings target .NET 10 alongside the library and CLI.
- OpenDocx.NET 1.1.8, its CLI, and its tests target .NET 10. The bundled OpenXmlPowerTools is rebased onto upstream `0d55d75`, retaining this fork's public namespace, assembly identity, and composition changes while adopting upstream test namespaces. Dependencies are aligned with the updated library: DocumentFormat.OpenXml 3.5.1, System.IO.Packaging 10.0.2, and SkiaSharp/native assets 4.153.1.
- Extraction and parsing share field-kind recognition, retaining keyword boundaries, case sensitivity, shorthand syntax, and existing control-marker whitespace behavior. The nested-array JSON format is unchanged; newly extracted text-box fields receive IDs in the normal extraction sequence.
- Floating-object diagnostics now refer to pictures, text boxes, and shapes rather than calling every object an image. The diagnostic code remains `FloatingImageOnBlockField`.

## [1.1.7] - 2026-10-13

### Added
- `Normalizer.AnalyzeTemplate` returns structured `FloatingImageOnBlockField` diagnostics via `NormalizeResult.Diagnostics` for reuse in audits of existing templates, without parsing field pairs or writing files.

### Changed
- Inline images count as paragraph content on either side of a field, in both text-delimited and content-control templates. This can make formerly block-level paired fields run-level, requiring their matching fields to be in the same paragraph. Floating images do not count as paragraph content when determining block/run classification.
- `Normalizer.NormalizeTemplate` rejects floating images anchored to block-field paragraphs, including content fields, with `FieldParseException`. `Normalizer.ExtractFieldsOnly`, `Normalizer.ExtractFields`, `OpenDocx.NormalizeDocx`, and `OpenDocx.PrepareTemplate` inherit this rejection. Structured diagnostics from `AnalyzeTemplate` identify the field, expression, and document part; thrown messages include the field ID, expression, and message.

### Fixed
- Compilation removes anchors with their block-field paragraph instead of moving the image to a separate paragraph. Inline images remain in the text flow.
- Legacy VML pictures inherit floating-position classification from their enclosing VML groups.
- Mixed floating/inline VML pictures still count as paragraph content, while their floating shapes are detected when validating block-level field controls.
- `Normalizer.ExtractFieldsOnly` shares normalization and validation without allocating a discarded copy of the normalized DOCX bytes.

## [1.1.6] - 2026-07-31

### Changed
- Allowed graphics following a block-level conditional or list field in the same paragraph. Compilation preserves the graphic in a separate content paragraph. Graphics preceding the field still prevent block-level classification. This behavior is superseded by the image classification and validation rules in 1.1.7.

## [1.1.5] - 2026-07-22

### Fixed
- Preserved table-cell boundaries when a cell contains only a single field, so paired fields in different cells are correctly rejected instead of being treated as sharing a container.

## [1.1.4] - 2026-06-18

### Changed
- Bumped OpenDocx.NET package version metadata to 1.1.4
- Updated bundled OpenXmlPowerTools with the latest insert-processing changes in `DocumentBuilder`/`DocumentComposer`

### Fixed
- Fixed repeated insert-processing scenarios (where the same insert is handled multiple times), including image relationship rewrite handling in embedded OpenXmlPowerTools

## [1.1.3] - 2026-03-05

### Changed
- Updated bundled OpenXmlPowerTools (OXPT) to a newer version

### Fixed
- Fixed operation of implicit inserts with filename-based templates
- Fixed undesirable blank paragraphs following some inserts (via OXPT update)

## [1.1.2] - 2026-02-06

### Added
- Table cell hierarchy tracking in field normalization to properly detect paired fields across cell boundaries

### Changed
- Added these release notes
- Field parsing now accumulates all errors instead of failing on first error
- Standardized error message format across field parsing: `Field {id} ("{text}"): {message}`
- Made `FieldParser` internal; `Templater` is the public API

### Fixed
- Field pairing errors now properly detected when paired fields are in different table cells
- NullReferenceException when processing orphaned end fields (endif, endlist)

## [1.1.1] - 2026-01-15

### Changed
- Our OpenXmlPowerTools fork is no longer a separate Nuget dependency;
  it is now bundled directly into this package

## [1.1.0] - 2026-01-01

### Added
- Initial stable release with .NET 8.0 support

### Changed
- Migrated to .NET 8.0 target framework
- rebased our OpenXmlPowerTools fork from OpenXmlDev (original developers, but inactive)
  to Codeuctivity's fork (more actively maintained)
- Updated dependencies to latest versions

## [1.0.0] - 2025-12-01

### Added
- Initial release of OpenDocx.NET
- Document assembly and templating capabilities
- OpenXML document processing tools
- Template transformation and validation
- Field extraction and replacement
- Comment management utilities

[1.1.8]: https://github.com/opendocx/opendocx-net/compare/d606f3fa4424c374f58dcb2d7ec599ac2714191d...HEAD
[1.1.7]: https://github.com/opendocx/opendocx-net/compare/df2abf3cb43565ade1afbdf8b2b12787fbde49cf...d606f3fa4424c374f58dcb2d7ec599ac2714191d
[1.1.6]: https://github.com/opendocx/opendocx-net/compare/v1.1.5...df2abf3cb43565ade1afbdf8b2b12787fbde49cf
[1.1.5]: https://github.com/opendocx/opendocx-net/compare/v1.1.4...v1.1.5
[1.1.4]: https://github.com/opendocx/opendocx-net/compare/v1.1.3...v1.1.4
[1.1.3]: https://github.com/opendocx/opendocx-net/compare/v1.1.2...v1.1.3
[1.1.2]: https://github.com/opendocx/opendocx-net/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/opendocx/opendocx-net/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/opendocx/opendocx-net/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/opendocx/opendocx-net/releases/tag/v1.0.0
