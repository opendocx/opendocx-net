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
using OpenXmlPowerTools;
using System.Collections.Generic;

namespace OpenDocx
{
    public class Composer
    {
        public static AssembleResult ComposeDocument(List<Source> oxptSources)
        {
            var composedDoc = DocumentBuilder.BuildDocument(oxptSources);
            return new AssembleResult(composedDoc.DocumentByteArray);
        }

        public static AssembleResult ComposeDocument(IEnumerable<IndirectSource> sources)
        {
            // rawSources is an array of objects, where each object has
            // (1) a unique ID, and (2) an array of bytes consisting of its DOCX contents
            if (sources == null || !sources.Any())
            {
                throw new ArgumentException("Invalid sources argument supplied to ComposeDocument");
            }
            var oxptSources = sources.Select(source => {
                var doc = new WmlDocument(new OpenXmlPowerToolsDocument(source.Bytes));
                var oxptSource = string.IsNullOrWhiteSpace(source.ID)
                    ? new Source(doc, true)
                    : new Source(doc, source.ID);
                if (source.KeepSections) {
                    oxptSource.KeepSections = true;
                }
                return oxptSource;
            }).ToList();
            return ComposeDocument(oxptSources);
        }
    }
}
