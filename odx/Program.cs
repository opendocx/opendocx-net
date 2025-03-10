using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    if (argLower == "--output" || argLower == "-o") {
                        nextIsOutput = true;
                    } else if (argLower == "--source" || argLower == "-s") {
                        nextIsSource = true;
                        srcIdx++;
                    } else if (argLower == "--preview" || argLower == "-p") {
                        wantPreview = true;
                    } else if (argLower == "--assemble" || argLower == "-a") {
                        nextIsData = true;
                    } else if (argLower == "--verify" || argLower == "-v") {
                        verifyMode = true;
                    } else if (argLower == "--help" || argLower == "-h") {
                        nextIsHelp = true;
                    } else if (string.IsNullOrEmpty(inFile) && !args[i].StartsWith('-')) {
                        inFile = args[i];
                    } else if (string.IsNullOrEmpty(outFile) && !args[i].StartsWith('-')) {
                        outFile = args[i];
                    } else {
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
            if (string.IsNullOrWhiteSpace(outFile) && !verifyMode && !string.IsNullOrEmpty(dataFile)) {
                Console.WriteLine("\noutfile not specified!");
                Usage();
                return;
            }

            if (verifyMode) {
                var validator = new Validator();
                if (Path.GetExtension(inFile).ToLower() == @".docx") {
                    // single-DOCX verify
                    var result = validator.ValidateDocument(inFile);
                    Console.WriteLine(result.HasErrors
                        ? "VERIFY failed:\n" + result.ErrorList
                        : "VERIFY succeeded"
                    );
                } else { // multi-DOCX verify
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
            else if (!string.IsNullOrEmpty(dataFile)) { // ASSEMBLE mode!
                // todo: assemble the template, answers and (optionally) sources into the output doc
                var result = await Assembler.AssembleDocumentAsync(File.ReadAllBytes(inFile), File.ReadAllText(dataFile), sources);
                if (!result.HasErrors) {
                    File.WriteAllBytes(outFile, result.Bytes);
                    Console.WriteLine("Assembled document written to " + outFile);
                } else { // HasErrors
                    Console.Error.WriteLine(result.Error);
                }
            }
            else if (!string.IsNullOrEmpty(inFile)) { // PREPARE TEMPLATE mode
                var o = new PrepareTemplateOptions();
                if (wantPreview) {
                    o.GenerateFlatPreview = true;
                }
                var result = OpenDocx.PrepareTemplate(File.ReadAllBytes(inFile), o);
                // save results: (TODO)
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
    1. Extracts the content of all OpenDocx {[delimited fields]} into JSON.
    2. Creates an OpenXmlPowerTools version of the template
    3. Creates a no-fields version of the template for potential conversion to Markdown (if --preview specified)
    4. Outputs a JSON file with field IDs to use after the Markdown conversion process (if --preview specified)

    All output files will be written to the same directory, and with the same BASE filename, where template.docx is located.

        odx.exe template.docx --preview

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
