# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- 

### Changed
- 

### Fixed
- 

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

[Unreleased]: https://github.com/opendocx/opendocx-net/compare/v1.1.2...HEAD
[1.1.2]: https://github.com/opendocx/opendocx-net/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/opendocx/opendocx-net/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/opendocx/opendocx-net/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/opendocx/opendocx-net/releases/tag/v1.0.0
