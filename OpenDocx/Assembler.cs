/***************************************************************************

Copyright (c) Lowell Stewart 2018-2019.
Licensed under the Mozilla Public License. See LICENSE file in the project root for full license information.

Published at https://github.com/opendocx/opendocx
Developer: Lowell Stewart
Email: lowell@opendocx.com

***************************************************************************/

using System;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using OpenXmlPowerTools;

namespace OpenDocx;

public static class Assembler
{
    public static AssembleResult AssembleDocument(byte[] templateBytes, XElement xmlData)
    {
        WmlDocument templateDoc = new(new OpenXmlPowerToolsDocument(templateBytes));
        WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(templateDoc, xmlData, out bool templateError);
        string errorMessage = null;
        if (templateError)
        {
            errorMessage = "Unexpected template errors - please inspect the assembled document.";
            Console.WriteLine(errorMessage);
        }
        return new AssembleResult(wmlAssembledDoc.DocumentByteArray, errorMessage);
    }

    public static async Task<AssembleResult> AssembleDocAsync(string templateFile, XElement data, string outputFile, List<DocxSource> sources)
    {
        var result = await AssembleDocAsync(File.ReadAllBytes(templateFile), data, sources);
        if (!string.IsNullOrEmpty(outputFile))
        {
            //// save the output (even in the case of error, since error messages are in the file)
            await File.WriteAllBytesAsync(outputFile, result.Bytes);
        }
        return result;
    }

    public static async Task<AssembleResult> AssembleDocAsync(byte[] templateBytes, XElement data, List<DocxSource> sources)
    {
        WmlDocument templateDoc = new(new OpenXmlPowerToolsDocument(templateBytes));
        WmlDocument wmlAssembledDoc = await DocumentComposer.ComposeDocument(templateDoc, data, sources);
        return new AssembleResult(wmlAssembledDoc.DocumentByteArray);
    }

    public static async Task<AssembleResult> AssembleDocumentAsync(byte[] templateBytes, string xmlData, IEnumerable<IndirectSource> sources)
    {
        List<DocxSource> oxptSources = null;
        if (sources != null)
        {
            oxptSources = sources.Select(source => {
                var doc = new WmlDocument(new OpenXmlPowerToolsDocument(source.Bytes));
                var oxptSource = new DocxSource(doc, source.ID);
                if (source.KeepSections) {
                    oxptSource.KeepSections = true;
                }
                return oxptSource;
            }).ToList();
        }
        using var xmlReader = new StringReader(xmlData);
        try
        {
            return await AssembleDocAsync(templateBytes, XElement.Load(xmlReader), oxptSources);
        }
        catch (XmlException e)
        {
            e.Data.Add("xml", xmlData);
            throw;
        }
    }
}
