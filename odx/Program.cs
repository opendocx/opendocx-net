using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace OpenDocx.CommandLine
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string inFile = null, outFile = null, dataFile = null, sourceId = null,
                helpMode = null, unrecognized = null;
            var sources = new List<IndirectSource>();
            bool nextIsOutput = false, nextIsSource = false, nextIsHelp = false, flagError = false;
            bool verifyMode = false, wantPreview = false, nextIsData = false, keepSections = false;
            bool prepareMode = false, wantNormalize = false, wantFields = false, wantTransform = false;
            int i = 0, srcIdx = -1;
            while (i < args.Length) {
                if (nextIsOutput) {
                    outFile = args[i];
                    nextIsOutput = false;
                } else if (nextIsSource) {
                    var argLower = args[i].ToLowerInvariant();
                    if (argLower == "--keepsections" || argLower == "-k") {
                        keepSections = true;
                    }
                    else if (string.IsNullOrEmpty(sourceId)) {
                        sourceId = args[i];
                    } else {
                        var s = new IndirectSource(sourceId, File.ReadAllBytes(args[i]), keepSections);
                        sources.Add(s);
                        // reset source stuff
                        keepSections = false;
                        sourceId = null;
                        nextIsSource = false;
                    }
                } else if (nextIsData) {
                    dataFile = args[i];
                    nextIsData = false;
                } else if (nextIsHelp) {
                    helpMode = args[i];
                    nextIsHelp = false;
                } else {
                    var argLower = args[i].ToLowerInvariant();
                    if (argLower == "--output" || argLower == "-o")
                    {
                        nextIsOutput = true;
                    }
                    else if (argLower == "--source" || argLower == "-s")
                    {
                        nextIsSource = true;
                        srcIdx++;
                    }
                    else if (argLower == "--prepare" || argLower == "-p")
                    {
                        prepareMode = true;
                    }
                    else if (argLower == "--normalize" || argLower == "-n")
                    {
                        wantNormalize = true;
                    }
                    else if (argLower == "--fields" || argLower == "-f")
                    {
                        wantFields = true;
                    }
                    else if (argLower == "--transform" || argLower == "-t")
                    {
                        wantTransform = true;
                    }
                    else if (argLower == "--preview")
                    {
                        wantPreview = true;
                    }
                    else if (argLower == "--assemble" || argLower == "-a")
                    {
                        nextIsData = true;
                    }
                    else if (argLower == "--verify" || argLower == "-v")
                    {
                        verifyMode = true;
                    }
                    else if (argLower == "--help" || argLower == "-h")
                    {
                        nextIsHelp = true;
                    }
                    else if (string.IsNullOrEmpty(inFile) && !args[i].StartsWith('-'))
                    {
                        inFile = args[i];
                    }
                    else
                    {
                        flagError = true;
                        unrecognized = (unrecognized == null)
                            ? args[i]
                            : unrecognized + " " + args[i];
                    }
                }
                i++;
            }
            if (nextIsHelp) {
                Help(helpMode);
                return;
            }
            if (flagError) {
                Console.WriteLine("\nUnrecognized input: {0}", unrecognized);
                Usage();
                return;
            }
            if (string.IsNullOrWhiteSpace(outFile) && !verifyMode)
            {
                if (prepareMode || wantNormalize || wantFields || wantTransform || wantFields)
                { // infer default output directory for prepare template mode
                    if (inFile.EndsWith(Path.DirectorySeparatorChar + "template.docx"))
                    {
                        outFile = Path.GetDirectoryName(inFile) ?? ".";
                    }
                    else
                    {
                        outFile = (Path.GetDirectoryName(inFile) ?? ".")
                            + Path.DirectorySeparatorChar
                            + Path.GetFileNameWithoutExtension(inFile);
                    }
                }
                else if (!string.IsNullOrEmpty(dataFile))
                {
                    Console.WriteLine("\noutfile not specified!");
                    Usage();
                    return;
                }
            }

            if (verifyMode)
            {
                var validator = new Validator();
                if (Path.GetExtension(inFile).ToLower() == @".docx")
                {
                    // single-DOCX verify
                    var result = validator.ValidateDocument(inFile);
                    Console.WriteLine(result.HasErrors
                        ? "VERIFY failed:\n" + result.ErrorList
                        : "VERIFY succeeded"
                    );
                }
                else
                { // multi-DOCX verify
                    try
                    {
                        DirectoryInfo indir = GetExistingDirectory(inFile);
                        // directory-wide verification
                        foreach (string currentFile in EnumerateDocxFiles(inFile))
                        {
                            string fileName = Path.GetFileName(currentFile);
                            var result = validator.ValidateDocument(currentFile);
                            Console.WriteLine(fileName + ": " + (result.HasErrors
                                ? "VERIFY failed:\n" + result.ErrorList
                                : "VERIFY succeeded"));
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("VERIFY (directory) failed: {0}\n", e.ToString());
                        Usage();
                    }
                }
            }
            else if (!string.IsNullOrEmpty(dataFile))
            { // ASSEMBLE mode!
                // todo: assemble the template, answers and (optionally) sources into the output doc
                var result = await Assembler.AssembleDocumentAsync(File.ReadAllBytes(inFile), File.ReadAllText(dataFile), sources);
                if (!result.HasErrors)
                {
                    File.WriteAllBytes(outFile, result.Bytes);
                    Console.WriteLine("Assembled document written to " + outFile);
                }
                else
                { // HasErrors
                    Console.Error.WriteLine(result.Error);
                }
            }
            else if (!string.IsNullOrEmpty(inFile)) // PREPARE TEMPLATE mode
            {
                if (prepareMode || wantNormalize || wantFields || wantTransform || wantPreview)
                {
                    byte[] docxBytes;
                    byte[] normalizedBytes;
                    string rawFields;
                    string normalizedPath = Path.Combine(outFile, "normalized.obj.docx");
                    string rawFieldPath = Path.Combine(outFile, "fields.obj.json");
                    if (prepareMode || wantNormalize || !File.Exists(normalizedPath))
                    {
                        // read template into memory
                        docxBytes = await File.ReadAllBytesAsync(inFile);
                        Console.WriteLine("Template retrieved successfully; normalizing (step 1A) using NEW algorithm...");
                        var start = DateTime.Now;
                        var normalizeResult = Normalizer.NormalizeTemplate(docxBytes, true, new List<string>() { "UpdateFields", "PlayMacros" });
                        Console.WriteLine($"  Normalization took {(DateTime.Now - start).TotalSeconds:F1} seconds");
                        normalizedBytes = normalizeResult.NormalizedTemplate;
                        rawFields = normalizeResult.ExtractedFields;
                        if (!prepareMode)
                        {
                            Directory.CreateDirectory(outFile);
                            await File.WriteAllBytesAsync(normalizedPath, normalizedBytes);
                            await File.WriteAllTextAsync(rawFieldPath, rawFields);
                        }
                        Console.Write("Template normalized; ");
                    }
                    else
                    {
                        normalizedBytes = await File.ReadAllBytesAsync(normalizedPath);
                        rawFields = await File.ReadAllTextAsync(rawFieldPath);
                        Console.Write("Normalized template retrieved successfully; ");
                    }
                    Dictionary<string, ParsedField> fieldDict = null;
                    string fieldDictPath = Path.Combine(outFile, "fields.dict.json");
                    if (prepareMode || wantFields || (wantTransform && !File.Exists(fieldDictPath)))
                    {
                        Console.WriteLine("building field dictionary (step 1B)...");
                        fieldDict = Templater.ParseFieldsToDict(rawFields);
                        Console.WriteLine("Storing field dictionary...");
                        Directory.CreateDirectory(outFile);
                        await File.WriteAllTextAsync(fieldDictPath,
                            JsonSerializer.Serialize(fieldDict, OpenDocx.DefaultJsonOptions));
                        Console.WriteLine("Field dictionary stored (so other processes can use it): OK1");
                    }
                    else if (wantTransform)
                    {
                        var fieldDictStr = await File.ReadAllTextAsync(fieldDictPath);
                        fieldDict = JsonSerializer.Deserialize<Dictionary<string, ParsedField>>(fieldDictStr, OpenDocx.DefaultJsonOptions);
                        Console.Write("Field dictionary retrieved successfully; ");
                    }
                    if ((prepareMode || wantTransform) && fieldDict != null) {
                        Console.WriteLine("Transforming docx to oxpt (step 1C)...");
                        var compileResult = Templater.CompileTemplate(normalizedBytes, fieldDict);
                        if (compileResult.HasErrors)
                        {
                            throw new Exception("Error while transforming DOCX:\n"
                                + string.Join('\n', compileResult.Errors));
                        }
                        Console.WriteLine("DOCX transformed; storing oxpt.docx...");
                        Directory.CreateDirectory(outFile);
                        await File.WriteAllBytesAsync(Path.Combine(outFile, "oxpt.docx"), normalizedBytes);
                        Console.WriteLine("oxpt.docx stored");
                        Console.WriteLine("Success: OK2");
                    }
                    if (prepareMode || wantPreview)
                    {
                        Console.WriteLine("Now replacing docx fields prior to markdown conversion (step 1D)...");
                        try
                        {
                            var previewResult = TemplateTransformer.TransformTemplate(
                                normalizedBytes,
                                TemplateFormat.PreviewDocx,
                                null); // field map is ignored when output = TemplateFormat.PreviewDocx
                            if (!previewResult.HasErrors)
                            {
                                Console.WriteLine("Preview generated; storing preview.obj.docx...");
                                Directory.CreateDirectory(outFile);
                                await File.WriteAllBytesAsync(Path.Combine(outFile, "preview.obj.docx"), previewResult.Bytes);
                            }
                            else
                            {
                                Console.WriteLine("Preview failed to generate:\n" + string.Join('\n', previewResult.Errors));
                            }
                        }
                        catch (Exception e)
                        {
                            // some other random exception, typically an internal error
                            Console.WriteLine("Preview error: " + e.Message);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("error: unsupported arguments\n");
                Usage();
            }
        }

        private static DirectoryInfo GetExistingDirectory(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (dir.Exists)
                return dir;
            else throw new DirectoryNotFoundException();
        }

        private static DirectoryInfo GetOrCreateDirectory(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
                dir.Create();
            return dir;
        }
        private static IEnumerable<string> EnumerateDocxFiles(string path)
        {
            var fullPath = Path.GetFullPath(path);
            return Directory.EnumerateFiles(fullPath, "*.docx", SearchOption.TopDirectoryOnly)
                .Where(name => !name.Contains(".docxobj."));
        }

        private static IEnumerable<string> EnumerateFieldFiles(string path)
        {
            var fullPath = Path.GetFullPath(path);
            return Directory.EnumerateFiles(fullPath, "*.fields.csv", SearchOption.TopDirectoryOnly);
        }

        static void Usage() {
            Console.WriteLine(@"
usage:  odx.exe [TBA - see help!]
   or   odx.exe --help
");
        }

        static void Help(string command = null) {
            Console.WriteLine(@"
odx.exe performs these operations:

PREPARE TEMPLATE
    A. Normalizes and extracts all the {[delimited fields]} in the template
    B. Validates paired fields and creates a JSON 'field dictionary'
    C. Creates an OpenXmlPowerTools version of the template
    D. Creates a no-fields version of the template for potential conversion to Markdown (if --preview specified)

    --normalize   Performs only step A
    --fields      Performs step A (if not already done) followed by step B
    --transform   Performs steps A and B (if not already done) followed by step C
    --preview     Performs steps A thru C (if not already done) followed by step D
    --prepare     Performs ALL the above steps, in order

    By default, if the input file is named `template.docx`, all output files will be written to the same directory.
    If the template is named anything else, then by default, all output files are created in a subdirectory of the
    template's directory, with the same base name as the template itself.
    
    --output      Specify a different/custom directory where output should be written

    Sample:
        odx.exe template.docx --prepare

ASSEMBLE DOCUMENT
    input: an oxpt.docx format template + matching XML data + options (including output path)
    output: writes the assembled file to output path OR provides one or more error messages

        odx.exe template.docx --assemble data.xml --output assembled.docx
        odx.exe template.docx --assemble data.xml
            --source <SourceId1> --keepSections Source1.docx
            --source <SourceId2> Source2.docx
            --output assembled.docx

VERIFY DOCUMENT
    Attempt to check a .docx file for structural integrity/validity. If a file appears to be valid,
    no output is produced. If problems are detected, it may output some possibly-cryptic messages
    to hopefully narrow in on what is wrong with the file.

        odx.exe infile.docx --verify

HELP
    Dislay the text you're seeing right now.

        odx.exe --help
");
        }
    }
}
