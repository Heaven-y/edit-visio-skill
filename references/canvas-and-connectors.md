# Canvas, Lines and Final Size

Use this reference for new layouts, reconstruction scaling, fixed paper sizes,
connector behavior, or readability checks. The examples are not a diagram template.

## Three Independent Sizes

| Quantity | Meaning | Controls |
| --- | --- | --- |
| Reference grid | Actual image pixels, or an explicit logical grid without an image | Measured `RefW/RefH` |
| Visio page | Physical page size in the document | `PageWidthMm/PageHeightMm` or `PageW/PageH` in inches |
| Raster preview | Pixels exported from the page | Page inches multiplied by `PreviewDpi` / `PngDpi` |

Do not infer print width from image DPI metadata. A 1200-pixel screenshot does not
determine whether the diagram will be printed at 90 mm, displayed on a slide or
used as a large engineering drawing. Use the requested medium/size; ask when a
fixed dimension would materially affect the result. Choose and report a sensible
working size when none is required. The legacy 16-inch fallback is not a design rule.

For an image, the scaffold reads dimensions itself. Supply one physical dimension
to preserve the ratio automatically:

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
    -Mode Create -VsdxPath $target -DrawingScript $drawing `
    -ReferenceImagePath $reference -PageWidthMm 180
```

With `CanvasFit MatchReference` (default), explicit dimensions must match the
reference ratio. With `CanvasFit Contain`, content is uniformly scaled and centered
on the supplied page, leaving whitespace. `MarginMm` is a minimum on all sides;
the other axis can have more whitespace. No cropping, stretching or automatic
rearrangement occurs. To reflow instead of leaving whitespace, change the layout
under the user's approved redesign scope.

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
    -Mode Create -VsdxPath $target -DrawingScript $drawing `
    -ReferenceImagePath $reference -PageWidthMm 210 -PageHeightMm 297 `
    -CanvasFit Contain -MarginMm 10
```

Edit keeps existing dimensions unless explicitly overridden. Existing drawing
objects are not automatically moved or scaled when page dimensions change.
The helpers target ordinary unscaled drawing pages; engineering drawing scales
and real-world units require explicit Visio ShapeSheet handling and separate QA.

## Positions, Lengths and Typography

Inside the callback, `VX/VY` convert top-left grid positions to Visio inches.
`VL` converts lengths to inches; `VPT` converts grid lengths to points. The latter
two exclude centering offsets. A reference-pixel radius must use `VL`, not `VX`.
`roundPx=0` resets the shape to square corners; a negative value selects the
legacy small physical radius. For new diagrams, select a radius suitable for the
notation instead of copying a screenshot default.

Font and line parameters are points at the document's physical size. To reproduce
a measured reference stroke of two pixels, use `VPT 2` for `linePt`. Font glyph
height in pixels is not the same as font point size; confirm text visually.

If a 180 mm page is later placed at 90 mm, nominal 10 pt text becomes 5 pt and
0.8 pt strokes become 0.4 pt. Work at final size when possible. Increasing PNG DPI
does not repair undersized text or lines.

`MinFontPt`, `MinLinePt` and optional `FinalWidthMm` can catch this reduction.
Thresholds are user/medium-specific, not mandatory Nature limits for all Visio
documents. For example, a dense print figure might use 7-9 pt text and 0.5-1 pt
strokes at final size, while a presentation requires larger labels.

The style gate reads nominal Character rows in text-bearing shapes and LineWeight
where LinePattern is nonzero, recursively through groups. It includes declared
rows that might be unused/hidden; report and inspect such findings, do not silently
rewrite the document. It does not measure actual glyph extents, superscript
reduction, effective layer visibility, clipping, engineering drawing scales or
background-page styles. It cannot certify legibility or absence of overlap.

## Lines and Connections

- `LineTL/LineRel`: fixed axes, dividers and geometry. They do not follow nodes.
- `Connect-VisioShapes`: glued relationship between two 2-D shapes on the same page.
- `Set-VisioLineStyle`: consistent line color, point weight, dash and arrowheads.

```powershell
$edge = Connect-VisioShapes $source $target -FromSide Right -ToSide Left `
    -Routing Orthogonal -LinePt 0.8 -EndArrow 4 -LabelOffsetYPt 8 -PassThru
try {
    $edge.NameU = 'SourceToTarget'
    Set-Text $edge 'request' 9
} finally { Release-VisioComObject $edge }
```

`Auto` attachments glue to the shape pin and let Visio choose a side. Explicit
`Left/Right/Top/Bottom` glue to a midpoint in the shape's local coordinates via
`GlueToPos`; these directions rotate with a rotated shape. Visio creates connection
points on the endpoint shapes. `Routing Auto` keeps the connector's routing default;
`Orthogonal` and `Straight` set ShapeRouteStyle to 1 and 2 respectively.
Do not modify the page-wide route style for one connector. `LabelOffsetXPt` and
`LabelOffsetYPt` shift the native label position without replacing its dynamic
position formula. They default to zero; choose clearance for the actual route and
label length, then inspect it. A fixed offset is not automatic collision avoidance.

No arrow is the helper default. Arrow code 4 is a filled arrow, not a substitute
for a domain notation such as UML inheritance, inhibition or bidirectional flow.
Keep line patterns and arrow meanings consistent. Reserve space for branch labels,
avoid unrelated boxes/text, and inspect routes again after changing node geometry.
Move a representative node, check both endpoints and undo the test movement before
delivery unless the move is intentional.

## Visual Acceptance

Inspect the full page at intended use size and then dense regions. Verify text
fit, glyphs, aligned edges, repeated spacing, arrow semantics, route crossings,
icon silhouettes and color/line distinctions. Grayscale should retain important
meaning. Do not hide a line-text collision with an opaque label patch by default;
adjust label or route geometry. Native routing is useful, not a collision proof.

Rebuilds compare against the reference; new diagrams compare against supplied
relationships and notation; edits also check that unrelated content is unchanged.
For `Contain`, compare the fitted content region, not the whole padded page, with
the reference. Ratio checks alone cannot prove the drawing used the intended map.

## API References

- [Cell.GlueToPos](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.gluetopos)
- [ShapeRouteStyle](https://learn.microsoft.com/en-us/office/client-developer/visio/shaperoutestyle-cell-shape-layout-section)
- [SetRasterExportResolution](https://learn.microsoft.com/en-us/office/vba/api/visio.applicationsettings.setrasterexportresolution)
- [SetRasterExportSize](https://learn.microsoft.com/en-us/office/vba/api/visio.applicationsettings.setrasterexportsize)

Raster settings persist in Visio. The export helper captures and restores both
resolution and size, including after export failure. It refuses a PNG over 100
megapixels to avoid unbounded bitmap allocations; choose a lower DPI or vector
export for such a page. PNG settings do not affect the saved VSDX geometry.
