# OpenDocx.NET Lambda example

This sample shows how to separate template provisioning from document assembly
using OpenDocx.NET and S3. Both handlers target .NET 10.

The example is useful for its API flow and persisted artifacts. Its S3 key
conventions, rename behavior, and deployment configuration are illustrative and
should be adapted to the application hosting it.

## Provision a template

`Functions.ImportDocxTemplate` accepts an S3 event. It reads the source DOCX,
normalizes and extracts fields, parses their structure, and writes:

- `fields.dict.json`: field types, expressions, and generated atom names.
- `logic.json`: the expression/conditional/list structure for a host evaluator.
- `oxpt.docx`: the compiled template for assembly.
- `preview.obj.docx`: an optional field-marked preview.

Keys must contain `/templates/`. For a source named `import.docx`, artifacts go in
the same folder, and the source is copied to `template.docx` then deleted from its
original key. For another `.docx` filename, artifacts go under
`support/<filename-without-extension>/` beside the original. Import errors are
logged; this event handler does not return a provisioning-result object.

The SAM file does not configure S3 upload notifications. A host can invoke the
handler with an S3 event or configure notifications separately. Triggering only
on `import.docx` avoids feeding generated DOCX artifacts back into provisioning.

## Evaluate and assemble

Between handlers, the host reads the field/logic metadata, evaluates expressions,
and constructs the required XML. The sample does not implement that evaluator.
See the [main README](../../README.md) for a standalone example.

`Functions.AssembleDocx` accepts an `AssembleRequest` with:

| Property | Meaning |
| --- | --- |
| `Bucket` | Bucket containing the template and output. |
| `TemplateKey` | Key of the compiled `oxpt.docx`. |
| `Data` | Evaluated XML as a string. |
| `DestinationKey` | Prefix for output and optional component documents. |
| `Sources` | Component IDs; pass an empty list when there are none. |
| `OutputName` | Filename for the final document. |
| `ID` | Optional intermediate key for later composition; provide this or `OutputName`. |

The handler loads the compiled template and any components, assembles the DOCX,
writes the output to S3, and returns an `AssembleResponse`. For composition, the
XML's insert references and `Sources` must agree, and component DOCX files must
already exist under `DestinationKey`.

## Build and test locally

From the repository root, with the .NET 10 SDK and OXPT submodule available:

```sh
dotnet build Lambdas/DocxTemplater/DocxTemplater.csproj -c Release
dotnet test Lambdas/test/DocxTemplater.Tests/DocxTemplater.Tests.csproj -c Release
```

Tests use synthetic DOCX templates and an in-memory S3 client. They provision a
template through the import handler and assemble it through the runtime handler,
including both conditional outcomes. No AWS credentials or cloud resources are
used by the tests.

## Deployment files

`serverless.template` describes two arm64 functions using the managed `dotnet10`
runtime. `default-config.json` contains example deployment-tool defaults; `deploy`
selects an environment-specific config file. No such environment is required for
local tests, and building this project does not deploy it.

Before adapting the sample for an AWS environment, select its bucket, function
prefix, IAM permissions, and invocation mechanism. The inherited SAM policies are
broad example defaults, not a least-privilege deployment prescription. The sample
also leaves application authorization, expression validation/evaluation, artifact
versioning, and import-error reporting to its host.
