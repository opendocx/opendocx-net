# DOCX Web Extension & Task Pane Implementation Notes

This document summarizes the design decisions, OOXML structure, and implementation guidance for server-side embedding of Office Add-in task pane metadata in Word documents (.docx). It supports the goal of automatically opening a task pane when a document is opened via Open XML, without conflicting with existing add-in associations.

---

## 1. Background & Constraints

### Add-in Architecture
- Microsoft Word add-in using Office JS API
- Task pane + ribbon with shared runtime
- Documents opened via `context.application.createDocument(base64).open()` for full fidelity (headers, footers, sections, complex formatting)
- `insertFileFromBase64` with `Word.InsertLocation.replace` loses headers/footers/sections—not viable for production documents

### The Problem
- When `createDocument().open()` opens a new Word window, the task pane is closed by default
- Client-side `settings.set("Office.AutoShowTaskpaneWithDocument", true)` on the document returned by `createDocument()` **does not work**—that object is temporary, has no backing store, and does not expose the settings API until after `.open()` is called
- Server-side OOXML injection can embed task pane metadata so the add-in auto-opens when the document is opened

### Past Issues
- Abstract "task pane metadata" mixed properties from different OOXML structures (webextension vs. taskpane)
- Merging logic failed when documents had multiple add-in associations
- No way to designate which add-in should auto-open vs. others
- "Different GUIDs" appearing—document instance IDs vs. manifest GUID vs. reference IDs
- Failed to recognize our add-in when already present, leading to duplicate webextensions and broken task panes

---

## 2. OOXML Structure

### Package Layout
- **WebExTaskpanesPart** (one per document)
  - **WebExtensionParts** (one per add-in)—contains `webextension` element
  - **Taskpanes** container—contains `WebExtensionTaskpane` elements
  - Each `WebExtensionTaskpane` references a WebExtensionPart via relationship ID (`r:id`)

### webextension Element (WebExtensionPart)
- **id** (attribute): Uniquely identifies the add-in instance **in the current document**. NOT the manifest GUID. Office may generate different values when the user saves. Can sometimes match manifest GUID when server-embedded or when refs are damaged—use as fallback for identity matching.
- **reference** (primary): Catalog-specific identity
  - `id`: Marketplace asset ID (OMEX) or manifest GUID (EXCatalog/Registry/FileSystem)
  - `version`, `store`, `storeType`
- **alternateReferences**: Fallback chain if primary reference fails to resolve
- **properties**: Includes `Office.AutoShowTaskpaneWithDocument` (true/false)
- **bindings**: Optional—associations between add-in and document content (ranges, etc.). Webextensions are used for both task panes AND bindings; a webextension with no taskpane may have bindings.

### WebExtensionTaskpane Element
- **webextensionref**: Links to WebExtensionPart via relationship ID
- **dockstate**, **visibility**, **width**, **row**: UI state

### Relationship
- WebExtensionPart ↔ WebExtensionTaskpane is **1:0-or-1** via the relationship reference
- A webextension can have zero or one taskpane pointing to it
- A taskpane references exactly one WebExtensionPart
- WebExtensionPart with no taskpane may be orphaned OR valid (add-in with bindings but no task pane)

---

## 3. Identity & Matching

### Key Insight
- **webextension.id** is document-instance-specific—do NOT use as primary identity
- **reference.id** (primary + alternates) is the primary way to identify the add-in
- **webextension.id** can be used as **fallback** when reference IDs are damaged—it may match manifest GUID when server-embedded or in corrupted documents

### Matching Logic (in order)
1. Check `reference.id` (primary + all alternateReferences) against manifest GUID and Marketplace asset ID
2. If no match, check `webextension.id` against manifest GUID as fallback

### Store Types
| storeType | Meaning | reference.id |
|-----------|---------|--------------|
| OMEX | Microsoft Marketplace (Office.com) | Marketplace asset ID (e.g. WA200008877) |
| EXCatalog | Centralized Deployment (M365 admin) | Manifest GUID or deployment ID |
| SPCatalog | SharePoint app catalog | Varies; can be empty (broken) |
| FileSystem | File share catalog | Manifest GUID |
| Registry | Dev/sideload | Manifest GUID |

### Multi-Region
- **Manifest GUID** is the same regardless of deployment method
- **Marketplace asset ID** (if published) is one per add-in—same across locales; `store` indicates locale (en-US, ru-RU) but the asset ID does not change

---

## 4. Invalid Data & Classification

### Reference Validity
- **Empty reference id**: If any `<we:reference>` has an empty `id` attribute, treat it as invalid and ignore it
- **Valid reference**: Non-empty `id`

### Classification
- **OURS**: Matched via reference IDs or webextension.id fallback
- **UNRECOGNIZABLE**: All references have empty `id`—no one can resolve it
- **THEIRS**: Has at least one valid reference, but not ours—leave alone

### Conservative Approach
- **Leave alone** any webextension that is NOT ours but has at least one reference with non-empty `id`
- Do NOT scrub orphaned/broken webextensions from other add-ins
- Only clean up baggage left behind by our own add-in

---

## 5. Cleanup & Reuse Logic

### Unrecognizable Webextensions (all refs empty)
- **If HAS taskpane**: REUSE—replace its content with ours. Keep relationship IDs, taskpane element. Do not add new parts.
- **If NO taskpane**: REMOVE the webextension entirely

### Scenario: 3 webextensions (2 invalid, 1 ours)
- webextension1: invalid, has taskpane
- webextension2: invalid, has taskpane
- webextension3: **ours** (valid refs)

**Action**: Use webextension3. Remove webextension1 + taskpane and webextension2 + taskpane. Do NOT add a new one.

### Scenario: 1 invalid + 1 ours
- webextension1: invalid, has taskpane
- webextension2: **ours**

**Action**: Use webextension2. Remove webextension1 + taskpane.

### Scenario: 2 invalid only
- webextension1: invalid, has taskpane
- webextension2: invalid, has taskpane

**Action**: Reuse webextension1's slot (replace content with ours). Remove webextension2 + taskpane. Do NOT add new.

### Multiple OURS (duplicates from past bugs)
- Keep one, remove the others

### Order of Operations
- Process in deterministic order (e.g., part order) so "first" is consistent
- When reusing among UNRECOGNIZABLE, pick the first one
- When removing duplicates among OURS, keep the first one

---

## 6. Auto-Show Behavior

### Per-Add-In Limitation
- Microsoft docs: "You can only set one pane **of your add-in** to open automatically with a document"
- The limitation is per add-in, not necessarily per document

### Multiple Add-Ins
- Docs do not explicitly forbid multiple add-ins from having auto-show in the same document
- Task panes can dock to different sides/rows—multiple could theoretically auto-show
- **Do NOT set another add-in's auto-show to false**—we don't modify THEIRS
- Only set ours to `true`
- In current `EmbedAddIn` implementation, taskpane `visibility` is written as `false`; auto-open behavior is driven by `Office.AutoShowTaskpaneWithDocument=true`

---

## 7. Type Model (C#)

### StoreReferenceInfo
```csharp
public class StoreReferenceInfo
{
    public string Id { get; set; }
    public string StoreType { get; set; }
    public string Store { get; set; }
    public string Version { get; set; }
}
```

### WebExtensionInfo
```csharp
public class WebExtensionInfo
{
    /// <summary>webextension.id — document instance ID; use as fallback for matching.</summary>
    public string Id { get; set; }
    
    /// <summary>Primary reference first, then alternates. Any can match for identity.</summary>
    public StoreReferenceInfo[] StoreReferences { get; set; }
    
    /// <summary>Office.AutoShowTaskpaneWithDocument</summary>
    public bool AutoShow { get; set; }
}
```

### TaskPaneInfo
```csharp
public class TaskPaneInfo
{
    public string PartRelationshipId { get; set; }
    public string DockState { get; set; }
    public bool Visibility { get; set; }
    public double Width { get; set; }
    public uint Row { get; set; }
}
```

### EmbeddedAddInInfo
```csharp
public class EmbeddedAddInInfo
{
    public WebExtensionInfo WebExtension { get; set; }
    public TaskPaneInfo Taskpane { get; set; }  // may be null
}
```

---

## 8. Snapshot Helper

Returns all embedded add-ins; `Taskpane` may be null for orphaned webextensions:

```csharp
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
                Taskpane = tpInfo
            });
        }
    }
    
    return result;
}
```

---

## 9. Merge Algorithm (When Adding/Updating Our Add-in)

### Inputs
- Document bytes
- Our add-in metadata (OMEX-first):
    - `manifestGuid` (current)
    - `omexAssetId` (current)
    - `legacyManifestGuids` (optional)
    - `legacyOmexAssetIds` (optional)
    - `version`, `store` (locale), and taskpane layout fields

### Logic
1. Parse document, get WebExTaskpanesPart (or null)
2. Enumerate and classify each webextension: OURS | UNRECOGNIZABLE | THEIRS
3. **Choose slot**:
   - If any OURS → use it (if multiple, keep first, remove others)
   - Else if any UNRECOGNIZABLE with taskpane → reuse first one
   - Else → add new webextension + taskpane
4. **Cleanup**: Remove duplicates (other OURS), remove UNRECOGNIZABLE except reused one
5. **Update**: Set AutoShow on chosen slot (`true`). Do NOT modify THEIRS.
6. **Never overwrite other add-ins**; only add, update, or remove our own

### Current API Surface (implemented)
```csharp
public static List<EmbeddedAddInInfo> GetEmbeddedAddIns(byte[] docxBytes)

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
```

---

## 10. References

- [Automatically open a task pane with a document](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)
- [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) (sample)
- [MS-OWEXML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/a2cd741a-4cca-4b1a-ade4-b2c443972afa)
- [MS-OTASKXML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-otaskxml/652d0608-31b8-4e90-a83a-98d6957b7fed)

---

## 11. Manifest IDs (This Project)

- Default/Knackly: `8eb22e22-73c3-40a5-a8d8-ddae1c07065a`
- Actionstep: `8eb22e22-73c3-40a5-a8d8-ddae1c07068a`
