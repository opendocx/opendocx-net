# OpenDocx.NET

A .NET library for document processing and templating with OpenXML. This package provides a low-level toolkit for transforming and assembling Word documents. It is layered on top of Eric White's excellent Open XML Power Tools, extending it with additional capabilities:

- **Flexible field syntax** - Specify your own field delimiters and syntax for document templates
- **Basic field types** - Supports content fields, nestable conditional logic (`if`/`elseif`/`else`/`endif`), and nestable lists (`list`/`endlist`) without imposing restrictions on field expressions
- **Template transformation** - Convert DOCX files with your custom syntax into Open XML Power Tools-compatible templates by separating them into:
  - **Logic tree (JSON)** - Encapsulates the data structure and logical rules of the template
  - **Optimized DOCX file** - Ready for use with Open XML Power Tools' DocumentAssembler
- **Template optimization** - Utilities for optimizing template logic
- **Multi-template assembly** - Handles assembling and composing complex documents from multiple template sources

## What's NOT Included

OpenDocx.NET is intentionally data-source agnostic. It doesn't know or care where your data is stored or how it's structured—it only cares about the data structure that your template requires.

**This package does NOT include:**
- Data transformation tools for converting your source data into template-compatible formats
- Database connectors or data access layers
- Business logic for data processing or validation

**What you need to provide:**
OpenDocx.NET generates a logic tree (a simple JSON structure) that describes exactly what data structure your template expects. It's your responsibility to transform your source data to match this structure and provide it as XML or JSON. Once you do that, OpenDocx.NET handles the rest—-assembling and composing even the most complex documents.

## Features

- Document assembly and templating
- Field extraction and processing
- Template transformation
- Logic tree generation
- Content control processing
- Document validation

## Installation

Install the package via NuGet Package Manager:

```
Install-Package OpenDocx
```

Or via .NET CLI:

```
dotnet add package OpenDocx
```

## Usage

The OpenDocx library provides various classes for working with Word documents:

- **`FieldExtractor`** - Extract fields from DOCX files and analyze template structure
- **`Templater`** - Process templates and transform field syntax
- **`Assembler`** - Perform document assembly operations using logic trees and data
- **`Composer`** - Compose complex documents from multiple template sources
- **`Validator`** - Validate document structure and template integrity

## Requirements

- .NET 6.0 or later
- DocumentFormat.OpenXml 2.19.0 (compatibility with later versions not yet tested)

## License

This project is licensed under the [Mozilla Public License 2.0](https://www.mozilla.org/en-US/MPL/2.0/).

## Contributing

This package is part of the [OpenDocx.NET](https://github.com/opendocx/opendocx-net) project. Please refer to the main repository for contribution guidelines.