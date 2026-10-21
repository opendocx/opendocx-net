# OpenDocx.NET

**OpenDocx.NET translates Word templates containing text-delimited fields into
OpenXmlPowerTools templates, so applications can assemble documents using their
own expression language and data model.**

Authors work in Word and write fields such as `{[Customer.Name]}`,
`{[if IncludeDetails]}`, and `{[list Items]}`. OpenDocx identifies those fields,
checks their structure, and compiles the DOCX into the form expected by
OpenXmlPowerTools (OXPT). Your application evaluates the expressions; OXPT
handles the document assembly.

OpenDocx is a .NET library for document-generation applications. Its main
integration point is the boundary between a Word template and the application
that supplies its values.

## How it fits together

There are two distinct stages:

1. **Provision a template**, usually when an author uploads or changes it.
   Normalize the source DOCX, extract its fields, parse conditional/list structure,
   and compile an OXPT-compatible DOCX. Save the compiled template together with
   its field dictionary and logic tree.
2. **Assemble a document**, whenever data is available. Evaluate the expressions
   in the appropriate conditional and list scopes, construct the XML data expected
   by the compiled template, and pass both to the assembler.

```text
source.docx with {[fields]}
        |
        | normalize, extract, parse, compile
        v
compiled oxpt.docx + field dictionary / logic tree
        |                          |
        |                 your expression evaluator + application data
        |                          |
        |                          v
        +-------------------- evaluated XML
                                   |
                                   | OXPT document assembly
                                   v
                             assembled.docx
```

The compiled DOCX contains OXPT content controls and selectors. Expressions from
the source template are mapped to generated names such as `C1` or `L4`, called
*atoms* in the API. Those names connect the field metadata, evaluated XML, and
compiled template. Keep all three associated with the same template version.

This separation lets a host application implement an expression language in
.NET, JavaScript, or another environment. The host can validate expressions against
its own schema, retrieve data, apply business rules, and format values before
assembly. OpenDocx parses the template's control structure; it does not execute
the expressions or fetch application data. JSON metadata describes the work for
the host; the assembly APIs shown here consume evaluated XML.

## Authoring templates

Type fields as ordinary text in a DOCX document. For example:

```text
Dear {[Customer.Name]},

{[if IncludeDetails]}
Your reference is {[Reference]}.
{[elseif IncludeSummary]}
A summary will follow.
{[else]}
Please contact us for details.
{[endif]}

{[list Items]}
{[Description]}: {[FormattedAmount]}
{[endlist]}
```

`if` / `elseif` / `else` / `endif` and `list` / `endlist` establish the template's
structure; the remaining field text is an expression for the host to interpret.
Conditionals and lists can nest. Custom text delimiters are also supported.

Placement in Word is significant. Paired fields within a paragraph must match
within that paragraph; block-level markers can enclose multiple paragraphs.
Starting with 1.1.8, any authored whitespace around a content field makes it
run-level. Existing whitespace treatment for control markers is retained.
Floating pictures and text boxes must have supported anchors; normalization reports
unsupported anchors instead of silently losing them during assembly. See the
[changelog](https://github.com/opendocx/opendocx-net/blob/master/CHANGELOG.md) for
release-specific behavior.

## A minimal C# example

Create `greeting.docx` in Word containing one paragraph:

```text
Hello {[Customer.Name]}.
```

The following console-program code provisions that template and assembles
`greeting-assembled.docx`. Its evaluator is deliberately a single explicit lookup;
a real host would supply its own expression evaluator.

```csharp
using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using OpenDocx;

// Provision once for this version of the template.
var source = File.ReadAllBytes("greeting.docx");
var normalized = Normalizer.NormalizeTemplate(source);
var parsed = Templater.ParseFields(normalized.ExtractedFields);
var compiled = Templater.CompileTemplate(
    normalized.NormalizedTemplate, parsed.ParsedFields);
if (compiled.HasErrors)
    throw new InvalidOperationException(string.Join("\n", compiled.Errors));

var jsonOptions = OpenDocx.OpenDocx.DefaultJsonOptions;
File.WriteAllBytes("greeting.oxpt.docx", compiled.Bytes);
File.WriteAllText("greeting.fields.json",
    JsonSerializer.Serialize(parsed.ParsedFields, jsonOptions));
File.WriteAllText("greeting.logic.json",
    JsonSerializer.Serialize(parsed.LogicTree, jsonOptions));

// At assembly time, evaluate the expressions using application data.
// This example has content fields only, so no scope traversal is needed.
string Evaluate(string expression) => expression switch
{
    "Customer.Name" => "Ada Lovelace",
    _ => throw new InvalidOperationException($"Unknown expression: {expression}")
};

var data = new XElement("_odx",
    parsed.ParsedFields.Values
        .Where(field => field.Type == FieldType.Content)
        .GroupBy(field => field.Atom)
        .Select(group => new XElement(group.Key, Evaluate(group.First().Expression))));

var assembled = Assembler.AssembleDocument(compiled.Bytes, data);
if (assembled.HasErrors)
    throw new InvalidOperationException(assembled.Error);
File.WriteAllBytes("greeting-assembled.docx", assembled.Bytes);
```

The output reads `Hello Ada Lovelace.` with the template's formatting. The example
uses returned atoms rather than assuming a field number. `NormalizeTemplate`
rejects unsupported placements, and `ParseFields` throws `FieldParseException`
for structural errors such as unmatched markers. Use `Normalizer.AnalyzeTemplate`
when you need structured normalization diagnostics before deciding whether to
continue provisioning.

For more complex templates, the host traverses `parsed.LogicTree` and supplies
values in the correct scope:

| Field | Evaluated XML convention |
| --- | --- |
| Content with atom `C1` | `<C1>formatted text</C1>` |
| Condition with atom `C2` | `<C2b>true</C2b>` or `<C2b>false</C2b>` |
| List with atom `L3` | `<L3><L3i>item-scoped values</L3i><L3i>…</L3i></L3>` |

Names above are illustrative: use the atoms in the returned metadata. List
expressions are evaluated in the enclosing scope, and their child expressions in
each item scope. Inline-list punctuation has additional generated metadata; a
complete host integration should handle the logic tree rather than simply flatten
all fields into a dictionary of values.

## Library and examples

The package bundles the OpenDocx-maintained OXPT fork. Core APIs include:

- `Normalizer`: normalization, field extraction, and placement diagnostics.
- `Templater`: field parsing, logic metadata, and compilation to OXPT templates.
- `Assembler`: assembly from compiled templates and evaluated XML.
- `Assembler.AssembleDocumentAsync` and `Composer`: assembly/composition involving
  additional DOCX sources.
- `Validator`: Open XML document validation.

The [Lambda example](https://github.com/opendocx/opendocx-net/tree/master/Lambdas/DocxTemplater)
shows separate provisioning and assembly handlers with S3 storage. It illustrates
an integration pattern, including the boundary where a host must evaluate data.
The `odx` command-line utility exposes
normalization, extraction, compilation, assembly, and validation for local use.
Run `dotnet run --project odx -- --help` for its options.

## Install, build, and test

The NuGet package is named `OpenDocx.NET`:

```sh
dotnet add package OpenDocx.NET
```

Version 1.1.8 targets **.NET 10**. Dependencies include DocumentFormat.OpenXml
3.5.1 and SkiaSharp 4.153.1. Word automation is not required for provisioning or
assembly.

Clone with submodules, install the .NET 10 SDK, and run from the repository root:

```sh
git submodule update --init --recursive
dotnet build opendocx.sln -c Release
dotnet test OpenDocx.Tests/OpenDocx.Tests.csproj -c Release
dotnet test Lambdas/test/DocxTemplater.Tests/DocxTemplater.Tests.csproj -c Release
```

The Lambda tests use in-memory S3 storage and require no AWS account. OXPT has its
own tests and build guidance in its submodule. Package publishing instructions
are in [PUBLISHING.md](https://github.com/opendocx/opendocx-net/blob/master/OpenDocx/PUBLISHING.md).

## License

OpenDocx.NET is licensed under the [Mozilla Public License 2.0](https://github.com/opendocx/opendocx-net/blob/master/LICENSE).
It builds on the document-processing capabilities of OpenXmlPowerTools; see the
bundled fork for its license and acknowledgments.

Development of OpenDocx is sponsored by Knackly.
