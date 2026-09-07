---
name: visio-image-rebuilder
description: Create, edit, inspect and explicitly export Microsoft Visio diagrams, or rebuild reference images as native editable VSDX shapes, text and connectors. Requires Windows and Visio for drawing and rendering.
---

# Visio

Use the existing installation and shared scripts. The Markdown guidance and
PowerShell backend are agent-independent; optional UI metadata is not a dependency.
The COM backend requires Windows, PowerShell 7 (`pwsh`) and licensed Visio.
Check the host before drawing. Windows PowerShell 5.1 is rejected before COM
activation because document opening can hang in that host. Report a missing
PowerShell 7 installation and its proposed source before installing anything.
Package inspection and stencil catalog scanning do not launch COM.

## Choose the Mode

- Create: a new diagram from a description, using native shapes and connectors.
- Edit: change only the requested pages or objects in an existing VSDX.
- Rebuild-image: reconstruct a reference as editable geometry, never as an embedded page image.
- Inspect: read and report; no automatic document or configuration changes.
- Export: produce only the formats the user requested from the saved VSDX.

For rebuilds, read [rebuild-guidelines.md](references/rebuild-guidelines.md).
For domain icons, read [icon-strategy.md](references/icon-strategy.md), then search
[stencil-reference.md](references/stencil-reference.md) or the full
[stencil index](references/visio-stencil-index.md) for relevant candidates.
For new diagrams, [design recipes](references/template-library.md) are examples,
not bundled template files or a required layout.

## Input and Canvas

Resolve the reference, target, overwrite permission and requested deliverables.
Ask only when a missing choice changes meaning or output, such as unreadable text.
Keep source images and unrelated files unchanged.

Read the actual reference dimensions with `Get-ReferenceImageDimensions` in
`scripts/visio_runtime.ps1`. Do not copy pixel dimensions from an example.
Derive page height as `PageW * RefH / RefW`; page width is a physical-scale choice.
The scaffold rejects conflicting dimensions. Without an image, Create/Rebuild
accept `PageW/PageH` in inches or `RefW/RefH` for a separate coordinate grid.
Edit preserves the existing page dimensions; with no reference/grid override,
drawing helpers use those inch dimensions with a top-left origin. Direct COM uses
bottom-left page coordinates. For PDF references, render the selected page
first and use its measured raster dimensions.

Analyze panel bounds, whitespace, text sizes, colors, arrow topology and repeated
objects from that reference. Calibrate panel-local coordinates for dense content.
Treat charts in a concept diagram as schematic; never invent quantitative results.

## Native Content and Icons

Use basic shapes for boxes, grids, axes and connectors. Search installed stencils
for domain objects; verify both file/NameU availability and a rendered sample.
A matching name does not prove a matching silhouette or biological meaning.
Inspect master style controls before recoloring children: child shapes can be
alternative filled/outline styles, not separate semantic parts.

Use smooth native curves and grouped editable parts when no appropriate master
exists. Do not force an unrelated stencil or create external SVG by default.
External assets require a user-approved source and license. Never redistribute
Microsoft stencil files. Group and name related modules, preserving editable text.
For movable workflow relationships, use glued connectors, not disconnected lines.

## Execution and Safety

Use `scripts/visio_rebuild_scaffold.ps1` with an explicit operation:

- `-Mode Create`: refuse an existing target; keep supplied template contents.
- `-Mode Rebuild -PageIndex N`: clear only the selected page before drawing.
- `-Mode Edit -PageIndex N`: preserve objects and dimensions; change only the
  requested shapes in the callback. Inspect stable shape names/IDs first and avoid
  global style/master changes that would affect other pages.

`PageIndex` is one-based and defaults to 1. `Rebuild` remains the default mode for
older callers; always specify Edit for local changes. The script accepts a
`-DrawingScript` containing `Draw-VisioPage -Phase` (legacy `Draw-ReferenceFigure`
is also supported); the drawing
file must support `-LoadDrawing` to load definitions without starting a second run.
The callback receives `$script:Page`, `$script:Visio`, measured `RefW/RefH`,
`PageW/PageH`, and `BuildPhase`. Reuse the supplied drawing helpers where suitable.
Assign `-PassThru` results to variables and release them in `finally`; the scaffold
rejects COM objects emitted by the callback to prevent object-output floods.
Rectangle helpers default to a small physical corner radius; explicit `roundPx`
values use the drawing coordinate grid, with 0 selecting square corners.

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
    -Mode Rebuild -VsdxPath $target -ReferenceImagePath $reference -DrawingScript $drawing
```

For complex rebuilds, phases 1/2/3 can represent framework/content/details.
Each invocation redraws the complete state up to that phase, not a persisted
incremental edit. The callback must implement this contract; a phase argument
alone is not evidence of completeness. Simple diagrams do not need three runs.

The scaffold draws to a staging file, closes its COM session, runs checks and
atomically replaces the target only on success. A failed build leaves the original
unchanged. It removes its staging directory on every exit. `-KeepBackup` explicitly
retains the previous target at a unique path; there is no default backup sidecar.
Create/Rebuild never imply permission to overwrite a user file; obtain that choice
before running. Shared callbacks must implement their stated page/object scope.

Use `New-VisioApplication` / `Stop-VisioApplication` for isolated sessions.
Never attach a rebuild to the user's active UI, kill all VISIO processes, save in
a failure handler, or quit a user-owned application. Close only owned documents,
discard only owned unsaved automation state, and release every acquired COM
reference once. Helpers must not emit COM objects unless explicitly requested.

A locked target requires the user to close it or explicitly approve saving or
discarding its unsaved edits. Overwrite permission is not permission to discard UI
edits. Missing licensed software must be reported; installation requires an
approved source and authorization, not an unattended Office deployment.

## Outputs and Verification

Default deliverable: VSDX only. Render a temporary PNG for verification.
For visual review by an agent, explicitly set `-PreviewPath` inside a task-owned
temporary directory, inspect it, then delete it. Without that option the scaffold
removes its internal preview after automated checks. A user-requested PNG/export
is retained and must not overwrite the source image.

Use `visio_quality_gates.ps1` for package integrity, native-content/media checks,
required text/color tokens, page/reference and preview/page ratios, read-only COM
reopen, geometry bounds and owned-process cleanup. `visio_validate.ps1` delegates
to the same implementation. Text, color, ratio and bounds checks apply to
`-PageIndex`; repeat them for every changed page. Package inspection reports all
pages, using relationships and foreground/background ordering, not page filenames.
Optional size limits must be task-specific.

Native-only is the default. If a user explicitly approves retaining or importing
media (for example an existing logo), pass `-AllowMedia` and report mixed content;
it does not prove full native editability. Never delete existing assets just to make
a strict gate pass. For read-only inspection use `visio_page_tools.ps1 -InspectPackage`.

PASS means only the named automated check passed. Skipped tests say SKIPPED.
Phase completeness, overlap, text fit, icon semantics and visual fidelity require
rendered comparison. Inspect the whole page and dense crops; fix mismatches before
delivery. Check reference labels and flow relationships, not only one icon.

Only explicitly requested PNG/SVG/PDF/PPTX files are exported. PNG/SVG/PPTX export
the selected page; PDF exports all foreground pages with their backgrounds. PPTX is a slide containing the
rendered SVG, not decomposed editable PowerPoint shapes.
Report the final VSDX, verified editability, visual approximations and any skipped
checks. Do not claim a remote update without a verified push.
