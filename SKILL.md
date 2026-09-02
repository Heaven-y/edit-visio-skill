---
name: visio-image-rebuilder
description: Create, modify, inspect, export, or rebuild Microsoft Visio diagrams as native editable .vsdx files. Use for flowcharts, architecture, UML sequence diagrams, network or stencil diagrams, process maps, scientific figures, image-to-Visio reconstruction, layout/style changes, and Visio COM automation on Windows.
---

# Visio Automation and Image Rebuilder

This is a portable instruction-and-script bundle for Microsoft Visio. The core
workflow is not tied to Codex: any AI agent or local automation runner that can
read this file and execute the bundled scripts can reuse it. The optional
`agents/openai.yaml` file only supplies Codex UI metadata.

## Core Rule

Recreate the reference as editable Visio content. Do not satisfy a rebuild request by inserting the whole reference image into the page. Embedding the reference image is only allowed as a temporary locked tracing layer if it is removed or hidden before delivery and the final `.vsdx` remains native shapes, text, connectors, and groups.

Treat `.vsdx` as the source of truth. Export PNG, SVG, PDF, or PPTX only after the Visio page has been rebuilt or restyled and checked for native editability.

## Operating Modes

Choose the smallest mode that satisfies the request:

| Mode | Use | Default result |
| --- | --- | --- |
| Create | Draw a new flowchart, architecture, process map, UML sequence, network, or scientific schematic | `.vsdx` |
| Edit | Change an existing document's text, layout, palette, styles, grouping, or connectors | Updated `.vsdx` plus rolling backup |
| Rebuild-image | Convert PNG/JPG/screenshot/scan into native Visio shapes | `.vsdx`; preview only when requested |
| Inspect | Check pages, shape count, text, media, locks, or editability | Report only |
| Export | Render an existing verified `.vsdx` | Only explicitly requested formats |

Use basic shapes for ordinary diagrams. Load Visio Stencils only when the user
requests professional icons or the diagram semantically requires device/service
masters. For domain symbols without a suitable local master, use a small,
licensed SVG/EMF asset or a reusable native-shape helper; do not use Emoji,
generic clip art, or a large raster crop. Read
[references/icon-strategy.md](references/icon-strategy.md) before choosing or
importing an icon. For image reconstruction, automatic contour/OCR extraction
is only a first pass; dense scientific or chart-heavy images require an
agent-assisted native redraw with calibrated panels.

## Request and Output Rules

- Preserve the user's requested output boundary. A `.vsdx`-only request must not create SVG, PDF, PPTX, PNG, JSON, or other sidecars unless needed for an explicitly reported verification step.
- Ask for the output filename only when it cannot be inferred safely. Otherwise use the source stem and keep generated artifacts in the requested directory.
- Keep temporary previews and intermediate JSON under a build or preview directory when possible, not beside the final source file.
- Never use the whole reference image as the final page. A temporary tracing layer must be removed or hidden before delivery.

## Default Output and COM Hygiene

- A rebuild creates only the formats requested by the caller. The default output is the editable `.vsdx`; SVG, PDF, PPTX, and PNG are opt-in.
- A PNG preview is a verification artifact. Generate it only with an explicit preview switch or path, and do not treat it as an implicit deliverable.
- Use one rolling backup beside the target (for example `figure.backup.vsdx`) unless timestamped history is explicitly requested.
- Keep the rolling backup and temporary verification files until acceptance passes; after a successful handoff, remove them unless the user asks to retain recovery history.
- Assign or suppress every COM return value. Drawing helpers must not leak Shape, Page, or Document objects into the PowerShell pipeline.
- Create a dedicated `Visio.Application` instance from an explicit template path. Avoid `Documents.Add('')`, which can open a chooser or hang.
- Save before closing; close only the document created by the script; release COM references; and quit only the dedicated application instance.
- If PowerShell COM becomes unstable or a job is long-running, use the optional Python `pywin32` backend with `EnsureDispatch` in a dedicated process; keep COM references local and release them in reverse order. Read [references/python-com-backend.md](references/python-com-backend.md) before using it.
- After a timeout, inspect the target timestamp and lock state before retrying. Never launch another Visio instance while the previous one may still be active.

## Workflow

1. Inspect inputs.
   - Confirm paths for the target `.vsdx`, reference image, requested output formats, and output directory.
   - Export the current Visio page to PNG before editing only when visual comparison is requested.
   - Inspect the `.vsdx` package for pages, media entries, and shape counts.
   - Back up the target file before any write using the rolling-backup policy above.
   - If Visio is already open, close only the target document or ask before terminating a stuck process.

2. Decode the reference image.
   - Identify page orientation, panel boundaries, module colors, captions, text hierarchy, arrows, dashed lines, and repeated motifs.
   - Build an object inventory: containers, titles, process boxes, icons, charts, graphs, equations, connectors, captions.
   - Calibrate the canvas first, then calibrate each major panel or subpanel with explicit top-left bounds or four corner points.
   - Draw panel internals in panel-local normalized coordinates whenever the figure has dense multi-panel content.
   - Decide whether the task is a full rebuild, color/style transfer, local edit, or export-only job.
   - For dense scientific figures, first create a coarse panel map, then draw panel internals. Do not start with small decorative details.
   - OpenCV contour or OCR extraction may supply a first-pass geometry inventory, but it does not replace semantic annotation for dense scientific figures, charts, or image-like motifs.

3. Prefer Visio automation for native edits.
   - Use COM automation on Windows when Visio is installed.
   - Use `DrawRectangle`, `DrawOval`, `DrawLine`, `Page.Import` only for small vector source assets, and shape cells such as `FillForegnd`, `LineColor`, `LineWeight`, `Char.Size`, `Char.Color`, `Rounding`.
   - Before importing an icon, inspect available masters with `scripts/visio_stencil_catalog.ps1`; prefer a matching built-in master, then a repository asset, then a native helper.
   - Use explicit coordinates and IDs for fragile edits.
   - Keep grouped structure meaningful: major panels, submodules, repeated blocks, legends.
   - For nested modules, use helpers such as `RectRel`, `TextRel`, `LineRel`, and `OvalRel` so child shapes are constrained by their calibrated parent panel.

4. Use package XML edits only for narrow, deterministic changes.
   - XML patching is appropriate for recoloring existing shapes, replacing font tables, or changing known cell values.
   - Preserve Visio XML ordering: shape-level `Cell` nodes should be before `Section`, `Text`, or child `Shapes`.
   - Avoid rebuilding complex geometry by raw XML unless COM automation is unavailable.

5. Export only requested formats from the verified Visio source.
   - Use `scripts/visio_page_tools.ps1` for export-only jobs.
   - Use `scripts/visio_rebuild_scaffold.ps1` with `-ExportFormats` for rebuilds that should immediately create deliverables.
   - Omit `-ExportFormats` and preview arguments for a source-only rebuild.
   - Prefer SVG for vector web/manuscript handoff, PDF for review/print, and PPTX for presentation decks.
   - For PPTX, use PowerPoint COM when available; the generated slide contains the Visio page render, usually inserted from SVG.

6. Verify without overtrusting a single signal.
   - Export at least one preview after editing when possible.
   - Inspect the `.vsdx` package to confirm that no full-size reference PNG/JPG was left in `visio/media`.
   - Check shape count and representative text labels.
   - Check that major panels do not overlap and that child shapes stay within their intended panel bounds.
   - Check every requested output file exists and is non-empty.
   - If Visio automation hangs, stop safely, close the document if possible, and report whether the file was actually modified.

## Implementation Pattern

For full rebuilds, generate a script that:

- Opens the target `.vsdx` with Visio COM.
- Saves the rolling backup before editing; use timestamped history only when requested.
- Clears or duplicates the page depending on user preference.
- Sets page size to match the reference aspect ratio.
- Draws native shapes in top-left reference coordinates converted to Visio coordinates.
- Defines calibrated panel bounds for complex regions and uses panel-local coordinates for their internals.
- Adds reusable helpers for rectangles, text boxes, ovals, lines, arrows, mini charts, graph nodes, and image-like stacks.
- Saves the document and exports only explicitly requested formats.

Start from `scripts/visio_rebuild_scaffold.ps1` when building a full reconstruction script. Copy it into the workspace and customize the `Draw-ReferenceFigure` function rather than editing the skill copy.

For style transfer, generate a script that:

- Reads existing shape IDs, text, approximate geometry, fill, and line colors.
- Maps known modules to target palettes by text and group context.
- Applies fills, borders, text colors, line patterns, and font changes to existing shapes.
- Avoids repositioning unless the user asks for layout changes.
- Exports only after the `.vsdx` has been saved and inspected.

For export-only requests, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\figure.vsdx" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports" `
  -InspectPackage
```

For rebuild plus multi-format export, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_rebuild_scaffold.ps1 `
  -VsdxPath "C:\path\figure.vsdx" `
  -PageW 16 `
  -PageH 9 `
  -RefW 1600 `
  -RefH 900 `
  -PreviewPath "C:\path\exports\figure.png" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports"
```

## Safety Checklist

- Back up before writing outside the workspace.
- Close any open Visio document that locks the target file before direct package edits.
- Never delete or revert unrelated user files.
- If a previous attempt embedded the reference image, restore from backup or remove the image shape before continuing.
- Tell the user clearly whether the final file is native editable shapes or a flat embedded image.
- Tell the user when PPTX export is a rendered slide rather than native PowerPoint shapes.

## Acceptance Criteria

A Visio rebuild is acceptable only when:

- Main panel positions, flow direction, captions, and module hierarchy match the reference at first glance.
- Major panels are aligned to calibrated bounds, with no obvious submodule drift, collision, or cross-panel overlap.
- Text remains editable and uses a consistent academic font, usually Times New Roman.
- Repeated motifs are represented with reusable native shapes rather than pasted raster crops.
- Domain icons use a matching Visio master, a small licensed vector asset, or a native fallback; no Emoji or low-quality clip art is used.
- The final `.vsdx` package has no full-page raster reference image in `visio/media`.
- Only requested export files were produced from the saved `.vsdx` and are non-empty; a source-only rebuild needs no export files.
- A preview export or package inspection was performed, or the final response explicitly states why verification was skipped.

## Useful Resource

Use `scripts/visio_page_tools.ps1` for common inspection, backup, preview export, multi-format export, and package checks. Use `scripts/visio_export_formats.ps1` when a custom task script needs reusable export helpers. Use `scripts/visio_rebuild_scaffold.ps1` as the starting point for native-shape drawing scripts; it includes both global top-left helpers and panel-local calibrated helpers. Use `scripts/visio_stencil_catalog.ps1` to inspect local stencil masters without launching Visio. Read `references/rebuild-guidelines.md` when a task requires a full figure reconstruction or careful one-to-one scientific diagram matching, and read `references/icon-strategy.md` when a diagram contains semantic icons or imported vector assets.
