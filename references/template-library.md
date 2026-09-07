# Visio Design Recipes

These are layout choices, not bundled VSDX templates. Use the user's requested
notation, meaning and existing document style. An installed Visio template can be
passed through `-TemplatePath` when creating a new document.

## Workflow and Swimlanes

Use start/end nodes, process boxes and decision diamonds with labeled branches.
For ownership across teams, arrange steps inside lanes. Prefer installed flowchart
or cross-functional templates when lane/container behavior matters; a rectangle
that looks like a lane does not automatically manage its contained shapes.
Glue connections so they survive moving nodes. Check loops, branch labels and
reading order, including routes that cross lanes.

## Network and Deployment

Group devices and services by site, network or trust boundary. Search installed
stencils for routers, servers, databases and the requested cloud services. Keep
connection direction, protocol and redundancy explicit where provided by the user.
Do not infer a deployment topology from product logos alone.

## Organization and Hierarchy

Use one consistent node style and alignment per level. Distinguish reporting lines
from dotted advisory relationships. Prefer organization-chart masters when their
layout or shape-data features are needed. Keep role/name text editable; grouping
for appearance is not the same as creating an organization-chart data model.

## UML and Sequence

Use installed UML stencils when the task requires that notation. Sequence diagrams
need participant headers, vertical lifelines, ordered messages and activation bars;
class diagrams need compartments and correct association/inheritance endpoints.
Do not substitute a generic arrow if it changes UML meaning. Simple native geometry
is sufficient when only the visual notation, not specialized behavior, is requested.

## Tables and Comparison Matrices

Keep stable row/column tracks, headers and legends. Use actual supplied values for
quantitative cells, and preserve missing values instead of inventing measurements.
Name repeated groups so future edits can target a row, column or category.

## Scientific and Technical Schematics

Arrange inputs, transformations and outputs according to the supplied mechanism or
reference. Domain symbols such as molecules, laboratory equipment or plant organs
must preserve meaning; consult the stencil references and verify silhouettes.
Use [scientific palettes](scientific-color-palettes.md) only as optional starting
points for new diagrams, never to recolor an existing reference automatically.
Actual statistical plots should come from the user's data and plotting workflow.

## Execution and Coordinates

- `-Mode Create` refuses an existing target. It preserves template contents.
- `-Mode Rebuild -PageIndex N` clears only the selected page before drawing.
- `-Mode Edit -PageIndex N` preserves objects and page size; the callback changes
  only explicitly identified shapes, text or connections.
- With an image, use measured `RefW/RefH`; page height follows its aspect ratio.
- Without an image, pass `PageW/PageH` in inches, or `RefW/RefH` for a separate
  coordinate grid. For Edit with no dimensions, helpers use the existing page's
  inch dimensions, with the origin at the top left. Direct Visio COM uses bottom left.

Define `Draw-VisioPage([int]$Phase)` in the task's drawing script and support
`param([switch]$LoadDrawing)`. `Draw-ReferenceFigure` is accepted for older scripts.
Use the shared scaffold; do not duplicate its COM session or staging lifecycle.
Custom callbacks may create additional pages, but must not mutate unrelated pages
or shared masters/styles during a local edit.

Verify the selected page's text, geometry, connectors and rendered appearance.
Package inspection reports every page, while page-specific gates use `-PageIndex`.
Repeat those gates and visual review for every changed page. Keep only requested
deliverables and accurately report any approved non-native assets.
