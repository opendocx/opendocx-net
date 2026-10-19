# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- Floating-object diagnostics now refer to pictures, text boxes, and shapes rather than calling every object an image. The diagnostic code remains `FloatingImageOnBlockField`; detection and assembly behavior are unchanged.

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

[Unreleased]: https://github.com/opendocx/opendocx-net/compare/d606f3fa4424c374f58dcb2d7ec599ac2714191d...HEAD
[1.1.7]: https://github.com/opendocx/opendocx-net/compare/df2abf3cb43565ade1afbdf8b2b12787fbde49cf...d606f3fa4424c374f58dcb2d7ec599ac2714191d
[1.1.6]: https://github.com/opendocx/opendocx-net/compare/v1.1.5...df2abf3cb43565ade1afbdf8b2b12787fbde49cf
[1.1.5]: https://github.com/opendocx/opendocx-net/compare/v1.1.4...v1.1.5
[1.1.4]: https://github.com/opendocx/opendocx-net/compare/v1.1.3...v1.1.4
[1.1.3]: https://github.com/opendocx/opendocx-net/compare/v1.1.2...v1.1.3
[1.1.2]: https://github.com/opendocx/opendocx-net/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/opendocx/opendocx-net/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/opendocx/opendocx-net/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/opendocx/opendocx-net/releases/tag/v1.0.0
