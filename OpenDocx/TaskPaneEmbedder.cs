/***************************************************************************
This approach was derived from
  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/tree/master

See also...
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document#use-open-xml-to-tag-the-document

Published at https://github.com/opendocx/opendocx
Developer: Lowell Stewart
Email: lowell@opendocx.com

***************************************************************************/

using System;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using We = DocumentFormat.OpenXml.Office2013.WebExtension;
using Wetp = DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using System.Linq;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using System.Xml;

namespace OpenDocx;

public class TaskPaneEmbedder
{
    private const string AutoShowPropertyName = "Office.AutoShowTaskpaneWithDocument";

    private enum EmbeddedAddInKind
    {
        Ours,
        Theirs,
        Unrecognizable
    }

    public static byte[] EmbedTaskPane(byte[] docxBytes, string guid, string addInId, string version, string store,
        string storeType, string dockState, bool visibility, double width, uint row)
    {
        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(docxBytes, 0, docxBytes.Length);
            using (var document = WordprocessingDocument.Open(memoryStream, true))
            {
                // add task panes part if it doesn't exist
                var taskPanesPart = document.WebExTaskpanesPart;
                if (taskPanesPart == null)
                {
                    taskPanesPart = document.AddWebExTaskpanesPart();
                }

                // find web extension child part by guid, or create if it doesn't exist yet
                var webExtensionPart = taskPanesPart.GetPartsOfType<WebExtensionPart>()
                    .FirstOrDefault(p => p.WebExtension.Id == guid);
                if (webExtensionPart != null)
                {
                    // update logic?
                    var webExtension = webExtensionPart.WebExtension;
                    // update store reference? update alternate references? update Office.AutoShowTaskpaneWithDocument?
                }
                else // webExtensionPart == null, so add it
                {
                    webExtensionPart = taskPanesPart.AddNewPart<WebExtensionPart>(); // "rId1");

                    // Generate webExtensionPart Content
                    var webExtension = new We.WebExtension() { Id = guid }; // "{635BF0CD-42CC-4174-B8D2-6D375C9A759E}" };
                    webExtension.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");

                    webExtension.Append(new We.WebExtensionStoreReference()
                    {
                        Id = addInId,
                        Version = version,
                        Store = store,
                        StoreType = storeType
                    });

                    webExtension.Append(new We.WebExtensionReferenceList());

                    // Add the property that makes the taskpane visible.
                    var webExtensionPropertyBag = new We.WebExtensionPropertyBag();
                    var webExtensionProperty = new We.WebExtensionProperty()
                    {
                        Name = "Office.AutoShowTaskpaneWithDocument",
                        Value = "true"
                    };
                    webExtensionPropertyBag.Append(webExtensionProperty);
                    webExtension.Append(webExtensionPropertyBag);

                    webExtension.Append(new We.WebExtensionBindingList());

                    var snapshot = new We.Snapshot();
                    snapshot.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    webExtension.Append(snapshot);

                    webExtensionPart.WebExtension = webExtension;
                }
                var relationshipId = taskPanesPart.GetIdOfPart(webExtensionPart);

                // get (or create) list of task panes from task pane part
                var taskpanes = taskPanesPart.Taskpanes;
                if (taskpanes == null)
                {
                    // Generate taskPanesPart Content
                    taskpanes = new Wetp.Taskpanes();
                    taskpanes.AddNamespaceDeclaration("wetp", "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11");
                }

                // find existing task pane ref, or create if it doesn't exist yet
                // searching for webExtensionTaskpane within existing children of taskpanes
                var webExtensionPartReference = taskpanes
                    .Descendants<Wetp.WebExtensionPartReference>()
                    .FirstOrDefault(r => r.Id == relationshipId);
                if (webExtensionPartReference != null)
                {
                    // update the webExtensionPartReference
                    var webExtensionTaskpane = (WebExtensionTaskpane) webExtensionPartReference.Parent;
                    webExtensionTaskpane.DockState = dockState;
                    webExtensionTaskpane.Visibility = visibility;
                    webExtensionTaskpane.Width = width;
                    webExtensionTaskpane.Row = row;
                }
                else // webExtensionPartReference == null; create the task pane and part reference
                {
                    var webExtensionTaskpane = new Wetp.WebExtensionTaskpane()
                    {
                        DockState = dockState,
                        Visibility = visibility,
                        Width = width,
                        Row = row
                    };
                    webExtensionPartReference = new Wetp.WebExtensionPartReference() { Id = relationshipId };
                    webExtensionPartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                    webExtensionTaskpane.Append(webExtensionPartReference);
                    taskpanes.Append(webExtensionTaskpane);
                }
                if (taskPanesPart.Taskpanes == null)
                {
                    taskPanesPart.Taskpanes = taskpanes;
                }
                // no explicit save -- disposing automatically saves changes to byte stream
            }
            return memoryStream.ToArray(); // and this returns the now-modified byte stream
        }
    }

    public static byte[] EmbedTaskPanes(byte[] docxBytes, IEnumerable<TaskPaneMetadata> taskPanes)
    {
        if (taskPanes == null) throw new ArgumentNullException(nameof(taskPanes));

        var taskPaneList = taskPanes.ToList();
        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(docxBytes, 0, docxBytes.Length);

            using (var document = WordprocessingDocument.Open(memoryStream, true))
            {
                if (taskPaneList.Count == 0)
                {
                    var existingTaskPanesPart = document.WebExTaskpanesPart;
                    if (existingTaskPanesPart != null)
                    {
                        document.DeletePart(existingTaskPanesPart);
                    }
                    return memoryStream.ToArray();
                }

                var taskPanesPart = document.WebExTaskpanesPart ?? document.AddWebExTaskpanesPart();

                foreach (var part in taskPanesPart.GetPartsOfType<WebExtensionPart>().ToList())
                {
                    taskPanesPart.DeletePart(part);
                }

                taskPanesPart.Taskpanes = new Wetp.Taskpanes();
                taskPanesPart.Taskpanes.AddNamespaceDeclaration("wetp", "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11");

                foreach (var taskPane in taskPaneList)
                {
                    var webExtensionPart = taskPanesPart.AddNewPart<WebExtensionPart>();

                    var webExtension = new We.WebExtension() { Id = taskPane.Guid };
                    webExtension.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");

                    webExtension.Append(new We.WebExtensionStoreReference()
                    {
                        Id = taskPane.AddInId,
                        Version = taskPane.Version,
                        Store = taskPane.Store,
                        StoreType = taskPane.StoreType
                    });

                    webExtension.Append(new We.WebExtensionReferenceList());

                    var webExtensionPropertyBag = new We.WebExtensionPropertyBag();
                    var webExtensionProperty = new We.WebExtensionProperty()
                    {
                        Name = "Office.AutoShowTaskpaneWithDocument",
                        Value = XmlConvert.ToString(taskPane.AutoShow)
                    };
                    webExtensionPropertyBag.Append(webExtensionProperty);
                    webExtension.Append(webExtensionPropertyBag);

                    webExtension.Append(new We.WebExtensionBindingList());

                    var snapshot = new We.Snapshot();
                    snapshot.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    webExtension.Append(snapshot);

                    webExtensionPart.WebExtension = webExtension;

                    var relationshipId = taskPanesPart.GetIdOfPart(webExtensionPart);

                    var webExtensionTaskpane = new Wetp.WebExtensionTaskpane()
                    {
                        DockState = taskPane.DockState,
                        Visibility = taskPane.Visibility,
                        Width = taskPane.Width,
                        Row = taskPane.Row
                    };
                    var webExtensionPartReference = new Wetp.WebExtensionPartReference() { Id = relationshipId };
                    webExtensionPartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                    webExtensionTaskpane.Append(webExtensionPartReference);
                    taskPanesPart.Taskpanes.Append(webExtensionTaskpane);
                }
            }
            return memoryStream.ToArray();
        }
    }

    public static byte[] RemoveTaskPane(byte[] docxBytes, string guid)
    {
        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(docxBytes, 0, docxBytes.Length);
            using (var document = WordprocessingDocument.Open(memoryStream, true))
            {
                var taskPanesPart = document.WebExTaskpanesPart;
                if (taskPanesPart != null)
                {
                    var webExtensionPart = taskPanesPart.GetPartsOfType<WebExtensionPart>()
                            .FirstOrDefault(p => p.WebExtension.Id == guid);
                    if (webExtensionPart != null)
                    {
                        var relationshipId = taskPanesPart.GetIdOfPart(webExtensionPart);
                        // find existing task pane ref -- searching for webExtensionTaskpane within existing children of taskpanes
                        var webExtensionPartReference = taskPanesPart.Taskpanes
                            .Descendants<Wetp.WebExtensionPartReference>()
                            .FirstOrDefault(r => r.Id == relationshipId);
                        if (webExtensionPartReference != null)
                        {
                            var parent = webExtensionPartReference.Parent;
                            parent.RemoveAllChildren();
                            parent.Remove();
                            // taskPanesPart.Taskpanes.RemoveChild(parent);
                        }
                        taskPanesPart.DeletePart(webExtensionPart);
                    }
                    if (!taskPanesPart.Taskpanes.HasChildren)
                    {
                        document.DeletePart(taskPanesPart);
                    }
                }
                // no explicit save -- disposing automatically saves changes to byte stream
            }
            return memoryStream.ToArray();
        }
    }

    public static TaskPaneMetadata[] GetTaskPaneInfo(byte[] docxBytes) {
        var result = new List<TaskPaneMetadata>();
        using (MemoryStream memoryStream = new MemoryStream(docxBytes))
        {
            using (var document = WordprocessingDocument.Open(memoryStream, false))
            {
                var taskPanesPart = document.WebExTaskpanesPart;
                if (taskPanesPart != null)
                {
                    foreach (var webExtensionPart in taskPanesPart.GetPartsOfType<WebExtensionPart>()) {
                        var resultItem = new TaskPaneMetadata();
                        var webExtension = webExtensionPart.WebExtension;
                        resultItem.Guid = webExtension.Id;

                        var storeReference = webExtension.WebExtensionStoreReference;
                        resultItem.AddInId = storeReference.Id;
                        resultItem.Version = storeReference.Version;
                        resultItem.Store = storeReference.Store;
                        resultItem.StoreType = storeReference.StoreType;

                        var webExtensionPropertyBag = webExtension.WebExtensionPropertyBag;
                        var autoShowProperty = webExtensionPropertyBag
                            .Descendants<We.WebExtensionProperty>()
                            .FirstOrDefault(p => p.Name == "Office.AutoShowTaskpaneWithDocument");
                        if (autoShowProperty != null) {
                            resultItem.AutoShow = XmlConvert.ToBoolean(autoShowProperty.Value);
                        }
                        // look up task pane for this web extension
                        var relationshipId = taskPanesPart.GetIdOfPart(webExtensionPart);
                        // find existing task pane ref
                        // searching for webExtensionTaskpane within existing children of taskpanes
                        var webExtensionPartReference = taskPanesPart.Taskpanes
                            .Descendants<Wetp.WebExtensionPartReference>()
                            .FirstOrDefault(r => r.Id == relationshipId);
                        if (webExtensionPartReference != null)
                        {
                            var webExtensionTaskpane = (WebExtensionTaskpane) webExtensionPartReference.Parent;
                            resultItem.DockState = webExtensionTaskpane.DockState;
                            resultItem.Visibility = webExtensionTaskpane.Visibility;
                            resultItem.Width = webExtensionTaskpane.Width;
                            resultItem.Row = webExtensionTaskpane.Row;
                        }
                        result.Add(resultItem);
                    }
                }
            }
        }
        return result.ToArray();
    }

    public List<EmbeddedAddInInfo> GetEmbeddedAddIns(Stream docxStream)
    {
        var result = new List<EmbeddedAddInInfo>();
        
        using (var doc = WordprocessingDocument.Open(docxStream, false))
        {
            var webExPart = doc.WebExTaskpanesPart;
            if (webExPart == null) return result;
            
            foreach (var wePart in webExPart.GetPartsOfType<WebExtensionPart>())
            {
                var relationshipId = webExPart.GetIdOfPart(wePart);
                var weInfo = ParseWebExtension(wePart);
                var tpInfo = FindTaskpaneByRelationshipId(webExPart.Taskpanes, relationshipId);
                
                result.Add(new EmbeddedAddInInfo
                {
                    WebExtension = weInfo,
                    Taskpane = tpInfo  // may be null if orphaned
                });
            }
        }
        return result;
    }

    public static List<EmbeddedAddInInfo> GetEmbeddedAddIns(byte[] docxBytes)
    {
        if (docxBytes == null) throw new ArgumentNullException(nameof(docxBytes));

        using (var stream = new MemoryStream(docxBytes, writable: false))
        {
            return new TaskPaneEmbedder().GetEmbeddedAddIns(stream);
        }
    }

    public static byte[] EmbedAddIn(
        byte[] docxBytes,
        string manifestGuid,
        string omexAssetId,
        string version,
        string dockState,
        double width,
        uint row,
        string store = "en-US",
        IEnumerable<string> legacyManifestGuids = null,
        IEnumerable<string> legacyOmexAssetIds = null)
    {
        if (docxBytes == null) throw new ArgumentNullException(nameof(docxBytes));
        if (string.IsNullOrWhiteSpace(manifestGuid)) throw new ArgumentException("Manifest GUID is required.", nameof(manifestGuid));
        if (string.IsNullOrWhiteSpace(omexAssetId)) throw new ArgumentException("OMEX asset ID is required.", nameof(omexAssetId));

        var addInId = omexAssetId;
        var marketplaceAssetId = omexAssetId;
        var storeType = "OMEX";

        var manifestIdentityIds = BuildIdentitySet(manifestGuid, legacyManifestGuids);
        var marketplaceIdentityIds = BuildIdentitySet(omexAssetId, legacyOmexAssetIds);

        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(docxBytes, 0, docxBytes.Length);

            using (var document = WordprocessingDocument.Open(memoryStream, true))
            {
                var webExPart = document.WebExTaskpanesPart ?? document.AddWebExTaskpanesPart();

                if (webExPart.Taskpanes == null)
                {
                    webExPart.Taskpanes = new Wetp.Taskpanes();
                    webExPart.Taskpanes.AddNamespaceDeclaration("wetp", "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11");
                }

                var taskpanes = webExPart.Taskpanes;
                var parts = webExPart.GetPartsOfType<WebExtensionPart>().ToList();
                var candidates = new List<(WebExtensionPart Part, string RelationshipId, EmbeddedAddInInfo Info, EmbeddedAddInKind Kind)>();

                foreach (var part in parts)
                {
                    var relationshipId = webExPart.GetIdOfPart(part);
                    var info = new EmbeddedAddInInfo
                    {
                        WebExtension = ParseWebExtension(part),
                        Taskpane = FindTaskpaneByRelationshipId(taskpanes, relationshipId)
                    };

                    var kind = ClassifyEmbeddedAddIn(info.WebExtension, manifestIdentityIds, marketplaceIdentityIds);
                    candidates.Add((part, relationshipId, info, kind));
                }

                var ours = candidates.Where(c => c.Kind == EmbeddedAddInKind.Ours).ToList();
                var unrecognizable = candidates.Where(c => c.Kind == EmbeddedAddInKind.Unrecognizable).ToList();

                WebExtensionPart selectedPart = null;
                string selectedRelationshipId = null;

                if (ours.Count > 0)
                {
                    selectedPart = ours[0].Part;
                    selectedRelationshipId = ours[0].RelationshipId;
                }
                else
                {
                    var reusable = unrecognizable.FirstOrDefault(c => c.Info.Taskpane != null);
                    if (reusable.Part != null)
                    {
                        selectedPart = reusable.Part;
                        selectedRelationshipId = reusable.RelationshipId;
                    }
                }

                if (selectedPart == null)
                {
                    selectedPart = webExPart.AddNewPart<WebExtensionPart>();
                    selectedRelationshipId = webExPart.GetIdOfPart(selectedPart);
                }

                SetWebExtension(
                    selectedPart,
                    manifestGuid,
                    addInId,
                    version,
                    store,
                    storeType,
                    true,
                    marketplaceAssetId);

                EnsureTaskpane(
                    taskpanes,
                    selectedRelationshipId,
                    dockState,
                    false,
                    width,
                    row);

                foreach (var candidate in candidates)
                {
                    if (candidate.Part == selectedPart)
                    {
                        continue;
                    }

                    if (candidate.Kind == EmbeddedAddInKind.Ours || candidate.Kind == EmbeddedAddInKind.Unrecognizable)
                    {
                        RemoveTaskpaneByRelationshipId(taskpanes, candidate.RelationshipId);
                        webExPart.DeletePart(candidate.Part);
                    }
                }
            }

            return memoryStream.ToArray();
        }
    }

    private static WebExtensionInfo ParseWebExtension(WebExtensionPart wePart)
    {
        if (wePart == null || wePart.WebExtension == null)
        {
            return new WebExtensionInfo();
        }

        var webExtension = wePart.WebExtension;
        var refs = new List<StoreReferenceInfo>();

        var primaryRef = webExtension.WebExtensionStoreReference;
        if (primaryRef != null)
        {
            refs.Add(new StoreReferenceInfo
            {
                Id = primaryRef.Id ?? string.Empty,
                StoreType = primaryRef.StoreType ?? string.Empty,
                Store = primaryRef.Store ?? string.Empty,
                Version = primaryRef.Version ?? string.Empty
            });
        }

        var alternateRefs = webExtension.WebExtensionReferenceList?
            .Elements<We.WebExtensionStoreReference>()
            .ToList();
        if (alternateRefs != null)
        {
            foreach (var alternateRef in alternateRefs)
            {
                refs.Add(new StoreReferenceInfo
                {
                    Id = alternateRef.Id ?? string.Empty,
                    StoreType = alternateRef.StoreType ?? string.Empty,
                    Store = alternateRef.Store ?? string.Empty,
                    Version = alternateRef.Version ?? string.Empty
                });
            }
        }

        var autoShow = false;
        var property = webExtension.WebExtensionPropertyBag?
            .Elements<We.WebExtensionProperty>()
            .FirstOrDefault(p => string.Equals(p.Name, AutoShowPropertyName, StringComparison.OrdinalIgnoreCase));
        if (property != null)
        {
            if (bool.TryParse(property.Value, out var parsedBool))
            {
                autoShow = parsedBool;
            }
            else
            {
                try
                {
                    autoShow = XmlConvert.ToBoolean(property.Value);
                }
                catch
                {
                    autoShow = false;
                }
            }
        }

        return new WebExtensionInfo
        {
            Id = webExtension.Id ?? string.Empty,
            StoreReferences = refs.ToArray(),
            AutoShow = autoShow
        };
    }

    private static TaskPaneInfo FindTaskpaneByRelationshipId(Taskpanes taskpanes, string relationshipId)
    {
        if (taskpanes == null || string.IsNullOrWhiteSpace(relationshipId))
        {
            return null;
        }

        var webExtensionPartReference = taskpanes
            .Descendants<Wetp.WebExtensionPartReference>()
            .FirstOrDefault(r => r.Id == relationshipId);
        if (webExtensionPartReference == null)
        {
            return null;
        }

        var taskpane = webExtensionPartReference.Parent as WebExtensionTaskpane;
        if (taskpane == null)
        {
            return null;
        }

        return new TaskPaneInfo
        {
            PartRelationshipId = relationshipId,
            DockState = taskpane.DockState ?? string.Empty,
            Visibility = taskpane.Visibility ?? false,
            Width = taskpane.Width ?? 0,
            Row = taskpane.Row ?? 0
        };
    }

    private static EmbeddedAddInKind ClassifyEmbeddedAddIn(
        WebExtensionInfo webExtension,
        IReadOnlyCollection<string> manifestIdentityIds,
        IReadOnlyCollection<string> marketplaceIdentityIds)
    {
        if (webExtension == null)
        {
            return EmbeddedAddInKind.Unrecognizable;
        }

        var validReferenceIds = (webExtension.StoreReferences ?? Array.Empty<StoreReferenceInfo>())
            .Where(r => !string.IsNullOrWhiteSpace(r.Id))
            .Select(r => r.Id.Trim())
            .ToList();

        foreach (var referenceId in validReferenceIds)
        {
            if (MatchesAnyId(referenceId, manifestIdentityIds))
            {
                return EmbeddedAddInKind.Ours;
            }

            if (MatchesAnyId(referenceId, marketplaceIdentityIds))
            {
                return EmbeddedAddInKind.Ours;
            }
        }

        if (MatchesAnyId(webExtension.Id, manifestIdentityIds))
        {
            return EmbeddedAddInKind.Ours;
        }

        return validReferenceIds.Count == 0
            ? EmbeddedAddInKind.Unrecognizable
            : EmbeddedAddInKind.Theirs;
    }

    private static HashSet<string> BuildIdentitySet(string currentValue, IEnumerable<string> legacyValues)
    {
        var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrWhiteSpace(currentValue))
        {
            values.Add(currentValue.Trim());
        }

        if (legacyValues != null)
        {
            foreach (var value in legacyValues)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    values.Add(value.Trim());
                }
            }
        }

        return values;
    }

    private static bool MatchesAnyId(string candidate, IReadOnlyCollection<string> ids)
    {
        if (string.IsNullOrWhiteSpace(candidate) || ids == null || ids.Count == 0)
        {
            return false;
        }

        foreach (var id in ids)
        {
            if (IdsMatch(candidate, id))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IdsMatch(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        if (Guid.TryParse(left, out var leftGuid) && Guid.TryParse(right, out var rightGuid))
        {
            return leftGuid == rightGuid;
        }

        return string.Equals(left.Trim(), right.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static void SetWebExtension(
        WebExtensionPart webExtensionPart,
        string manifestGuid,
        string addInId,
        string version,
        string store,
        string storeType,
        bool autoShow,
        string marketplaceAssetId)
    {
        var webExtension = new We.WebExtension() { Id = manifestGuid };
        webExtension.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");

        webExtension.Append(new We.WebExtensionStoreReference()
        {
            Id = addInId,
            Version = version,
            Store = store,
            StoreType = storeType
        });

        var referenceList = new We.WebExtensionReferenceList();
        if (IdsMatch(addInId, manifestGuid) == false)
        {
            referenceList.Append(new We.WebExtensionStoreReference()
            {
                Id = manifestGuid,
                Version = version,
                Store = store,
                StoreType = storeType
            });
        }

        if (!string.IsNullOrWhiteSpace(marketplaceAssetId)
            && !string.Equals(addInId, marketplaceAssetId, StringComparison.OrdinalIgnoreCase))
        {
            referenceList.Append(new We.WebExtensionStoreReference()
            {
                Id = marketplaceAssetId,
                Version = version,
                Store = store,
                StoreType = "OMEX"
            });
        }
        webExtension.Append(referenceList);

        var webExtensionPropertyBag = new We.WebExtensionPropertyBag();
        webExtensionPropertyBag.Append(new We.WebExtensionProperty()
        {
            Name = AutoShowPropertyName,
            Value = XmlConvert.ToString(autoShow)
        });
        webExtension.Append(webExtensionPropertyBag);

        webExtension.Append(new We.WebExtensionBindingList());

        var snapshot = new We.Snapshot();
        snapshot.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        webExtension.Append(snapshot);

        webExtensionPart.WebExtension = webExtension;
    }

    private static void EnsureTaskpane(Taskpanes taskpanes, string relationshipId, string dockState, bool visibility, double width, uint row)
    {
        var existingReference = taskpanes
            .Descendants<Wetp.WebExtensionPartReference>()
            .FirstOrDefault(r => r.Id == relationshipId);

        if (existingReference != null)
        {
            var existingTaskpane = existingReference.Parent as WebExtensionTaskpane;
            if (existingTaskpane != null)
            {
                existingTaskpane.DockState = dockState;
                existingTaskpane.Visibility = visibility;
                existingTaskpane.Width = width;
                existingTaskpane.Row = row;
            }
            return;
        }

        var newTaskpane = new Wetp.WebExtensionTaskpane()
        {
            DockState = dockState,
            Visibility = visibility,
            Width = width,
            Row = row
        };
        var webExtensionPartReference = new Wetp.WebExtensionPartReference() { Id = relationshipId };
        webExtensionPartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        newTaskpane.Append(webExtensionPartReference);
        taskpanes.Append(newTaskpane);
    }

    private static void RemoveTaskpaneByRelationshipId(Taskpanes taskpanes, string relationshipId)
    {
        if (taskpanes == null || string.IsNullOrWhiteSpace(relationshipId))
        {
            return;
        }

        var webExtensionPartReference = taskpanes
            .Descendants<Wetp.WebExtensionPartReference>()
            .FirstOrDefault(r => r.Id == relationshipId);

        if (webExtensionPartReference?.Parent != null)
        {
            webExtensionPartReference.Parent.Remove();
        }
    }

}
