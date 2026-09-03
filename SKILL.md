---
name: visio-image-rebuilder
version: 2.0.0
description: Scientific figure reconstruction with precise layout, semantic color, and quality gates. Use for flowcharts, architecture, UML sequence diagrams, network diagrams, process maps, scientific figures, and image-to-Visio reconstruction with native editable .vsdx output.
---

# Visio Image Rebuilder v2.0

**Scientific figure reconstruction with precision layout and quality gates.**

This skill specializes in converting reference images (PNG, JPG, PDF) into fully editable Visio `.vsdx` files with native shapes, text, and connectors—never embedded images.

---

## Core Principle

**Recreate the reference as native, editable Visio content.**

Never satisfy a rebuild request by embedding the reference image. The deliverable must be fully editable `.vsdx` where every element—text, shapes, connectors—can be modified independently.

---

## The Six Contracts (New in v2.0)

Before any rebuild, declare these six contracts:

### 1. Canvas Contract
```powershell
$RefW = 1500; $RefH = 750  # Reference dimensions (px)
$PageW = 16; $PageH = 9    # Visio page (inches)
function VX($px) { ($px / $RefW) * $PageW }
function VY($px) { $PageH - ($px / $RefH) * $PageH }
```

### 2. Layout Contract
```powershell
$Regions = @{
    LeftPanel   = @{ RefX=20; RefY=50; RefW=480; RefH=700 }
    CenterPanel = @{ RefX=540; RefY=50; RefW=400; RefH=700 }
    RightPanel  = @{ RefX=980; RefY=50; RefW=500; RefH=700 }
}
```

### 3. Color Contract (Scientific Semantics)
```powershell
$Colors = @{
    DNA_Purple     = "RGB(102,51,204)"   # DNA double helix
    SNP_Blue       = "RGB(51,102,204)"   # SNP encoding
    Gene_Orange    = "RGB(204,153,51)"   # Gene annotation
    Protein_Green  = "RGB(51,153,102)"   # Protein structure
    Cell_Red       = "RGB(204,51,51)"    # Cell/tissue
    Data_Teal      = "RGB(0,153,153)"    # Data matrices
    Model_Purple   = "RGB(153,51,204)"   # ML models
    Result_Gold    = "RGB(204,153,0)"    # Results/predictions
    Neutral_Gray   = "RGB(100,100,100)"  # Neutral elements
}
```

### 4. Icon Contract
- **Priority**: Visio builtin stencil → Scientific library → Custom native shapes
- **Catalog check**: Run `scripts/visio_stencil_catalog.ps1` first
- **Common stencils**: `MEDICAL_M.VSSX` (DNA, cells), `BIOLOGI1_M.VSSX` (organisms)
- **No emoji or clip art**

### 5. Content Inventory
- Count total shapes expected: ~400 for complex figures
- List text labels, connectors, special elements
- Estimate phase breakdown: Framework (60) → Content (250) → Details (100)

### 6. Quality Contract
- ✅ No embedded reference image in final `.vsdx`
- ✅ All text editable (not rasterized)
- ✅ All shapes native Visio primitives or stencil masters
- ✅ Preview PNG matches reference layout
- ✅ File size reasonable (< 2 MB)

---

## Progressive Build Strategy (v2.0)

**Never draw 400+ shapes in one script.** Use three-phase progressive build:

### Phase 1: Framework (~60 shapes, 30s)
- Major region borders
- Panel titles
- Main connection arrows
- **Checkpoint**: Save → Export preview → Verify layout

### Phase 2: Core Content (~250 shapes, 2min)
- Feature visualizations
- Matrix internals
- Icons and symbols
- **Checkpoint**: Save → Export preview → Verify content

### Phase 3: Details (~100 shapes, 1min)
- Model visualizations
- Annotations and labels
- Final polish
- **Checkpoint**: Full validation → Quality audit

---

## Seven Quality Gates (v2.0)

Every rebuild must pass:

1. **INPUT_VALIDATION**: Reference image exists, target path valid
2. **LAYOUT_PLANNING**: Coordinate system calibrated, regions defined
3. **PROGRESSIVE_BUILD**: Three phases complete without COM errors
4. **ALIGNMENT_AUDIT**: Panels aligned, no overlaps
5. **COLOR_AUDIT**: Scientific semantics applied, colorblind-safe
6. **INTEGRITY_VERIFICATION**: Native shapes only, no embedded media
7. **DELIVERY**: All requested formats generated, file size acceptable

**Fail fast:** Stop at first gate failure, report issue, don't proceed.

---

## COM Best Practices (v2.0)

### Pre-computed Coordinates
```powershell
# BAD: Nested loops with COM calls
for ($i = 0; $i -lt 100; $i++) {
    for ($j = 0; $j -lt 50; $j++) {
        $shape = $page.DrawRectangle($i*10, $j*5, $i*10+8, $j*5+3)
    }
}

# GOOD: Pre-compute coordinates, batch draw
$coords = @()
for ($i = 0; $i -lt 100; $i++) {
    for ($j = 0; $j -lt 50; $j++) {
        $coords += @{ X1=$i*10; Y1=$j*5; X2=$i*10+8; Y2=$j*5+3 }
    }
}
foreach ($c in $coords) {
    $null = $page.DrawRectangle($c.X1, $c.Y1, $c.X2, $c.Y2)
}
```

### Proper COM Release
```powershell
try {
    $visio = New-Object -ComObject Visio.Application
    # ... work ...
} finally {
    if ($doc) { $doc.Close($true) }
    if ($visio) { $visio.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
```

### Output Rules
- Default output: editable `.vsdx` only
- Generate rebuild scripts in `$env:TEMP`, delete after success
- No preview PNG unless explicitly requested
- Clean up temporary files after completion

---

## Operating Modes

Choose the smallest mode that satisfies the request:

| Mode | Use | Default result |
| --- | --- | --- |
| Create | Draw a new flowchart, architecture, process map, UML sequence, network, or scientific schematic | `.vsdx` |
| Edit | Change an existing document's text, layout, palette, styles, grouping, or connectors | Updated `.vsdx` (temporary backup on failure) |
| Rebuild-image | Convert PNG/JPG/screenshot/scan into native Visio shapes | `.vsdx`; preview only when requested |
| Inspect | Check pages, shape count, text, media, locks, or editability | Report only |
| Export | Render an existing verified `.vsdx` | Only explicitly requested formats |

---

## Workflow

1. **Inspect inputs** - Confirm paths, inspect package, backup target file
2. **Decode reference** - Identify layout, colors, icons, inventory shapes
3. **Declare Six Contracts** - Canvas, Layout, Color, Icon, Inventory, Quality
4. **Progressive Build** - Three phases with checkpoints after each
5. **Verify Quality** - Run validation script, check shape count
6. **Clean up** - Remove temporary files, keep only final `.vsdx`

---

## Implementation Pattern

### Full Rebuild Template
```powershell
param([string]$VsdxPath, [int]$Phase = 1)

# 1. Declare Six Contracts
$RefW = 1500; $RefH = 750
$PageW = 16; $PageH = 9
function VX($px) { ($px / $RefW) * $PageW }
function VY($px) { $PageH - ($px / $RefH) * $PageH }

$Regions = @{
    LeftPanel = @{ RefX=20; RefY=50; RefW=480; RefH=700 }
}

$Colors = @{
    DNA_Purple = "RGB(102,51,204)"
}

# 2. Open Visio
try {
    $visio = New-Object -ComObject Visio.Application
    $doc = $visio.Documents.Open($VsdxPath)
    $page = $doc.Pages.Item(1)
    
    # Set page size
    $page.PageSheet.CellsSRC(1,1,0).FormulaU = "${PageW} in"
    $page.PageSheet.CellsSRC(1,1,1).FormulaU = "${PageH} in"
    
    # 3. Draw based on phase
    if ($Phase -eq 1) {
        # Framework: ~60 shapes
    } elseif ($Phase -eq 2) {
        # Content: ~250 shapes
    } elseif ($Phase -eq 3) {
        # Details: ~100 shapes
    }
    
    # 4. Save
    $doc.Save()
    Write-Host "Phase $Phase complete"
    
} finally {
    if ($doc) { $doc.Close($true) }
    if ($visio) { $visio.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
    [System.GC]::Collect()
}
```

---

## Safety Checklist

- Back up before writing
- Close any open Visio document that locks the target file
- Never delete unrelated user files
- Tell the user clearly whether the final file is native editable shapes

---

## Acceptance Criteria

A Visio rebuild is acceptable only when:

- Main panel positions, flow direction, captions match the reference
- Major panels aligned to calibrated bounds, no overlaps
- Text remains editable, consistent academic font
- Domain icons use Visio Master or editable native fallback
- Final `.vsdx` has no full-page raster reference image
- Only requested export files produced and non-empty

---

## Useful Resources

- `scripts/visio_page_tools.ps1` - Inspection, backup, export
- `scripts/visio_rebuild_scaffold.ps1` - Native-shape drawing template
- `scripts/visio_stencil_catalog.ps1` - Stencil inspection
- `scripts/visio_validate.ps1` - Package inspection + validation
- `references/icon-strategy.md` - Icon selection
- `references/scientific-color-palettes.md` - Color schemes
