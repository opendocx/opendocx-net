# `mc:Ignorable` Prefix Error: Current Findings

## What We Observed

Some DOCX templates in production history fail validation with errors like:

- invalid `mc:Ignorable` prefix references
- undeclared extension attributes/elements (`w14:paraId`, `w14:textId`, `w16cid:durableId`, `w14:docId`, etc.)

The failing files show `mc:Ignorable` values that reference prefixes not declared on the part root.

## What We Confirmed

- The issue can exist in a template *before* assembly.
- Corrupted and non-corrupted versions of the same template lineage both exist.
- Validation in `opendocx-net` (via `OpenDocx.Validator`) reports these errors; they are not in the ignore list.

## What We Have **Not** Proven

- We have not proven the exact code path that introduced corruption for the incident file.
- We have not proven a deterministic repro from a clean template using local test inputs.
- We have not proven whether the root cause is always inside OpenDocx/OXPT or sometimes external (for example, client/editor/add-in interactions before upload).

## Operational Recommendation

Treat this as an input-quality gate:

1. Validate uploaded source template DOCX.
2. Validate provisioned/object DOCX.
3. Block template activation when validation returns `mc:Ignorable`/undeclared-extension errors.

This prevents corrupted templates from entering service even while root-cause research continues.
