#nullable enable

using System;

namespace OpenDocx;

public class EmbeddedAddInInfo // A paired webextension + taskpane. They always appear together in OOXML.
{
    public required WebExtensionInfo WebExtension { get; set; }
    public TaskPaneInfo? Taskpane { get; set; } // null for webextensions without a taskpane
}

public class WebExtensionInfo
{
    public string Id { get; set; } = string.Empty; // unique id of webextension within this document -- may or may not match add-in's manifest GUID
    public StoreReferenceInfo[] StoreReferences { get; set; } = []; // primary ref first, alternate(s) subsequent; for matching identity.
    public bool AutoShow { get; set; } // Office.AutoShowTaskpaneWithDocument
}

public class StoreReferenceInfo
{
    public string Id { get; set; } = string.Empty; // Marketplace asset ID (OMEX) or manifest GUID (EXCatalog/Registry/FileSystem)
    public string StoreType { get; set; } = string.Empty; // OMEX, EXCatalog, SPCatalog, Registry, FileSystem
    public string Store { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
}

public class TaskPaneInfo
{
    public string PartRelationshipId { get; set; } = string.Empty; // Relationship ID that links this taskpane to its WebExtensionPart (e.g. "rId1")
    public string DockState { get; set; } = string.Empty;
    public bool Visibility { get; set; }
    public double Width { get; set; }
    public uint Row { get; set; }
}
