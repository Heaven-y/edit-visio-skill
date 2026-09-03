# Visio Template Library for Scientific Figures

## Overview

Pre-built Visio templates for common scientific figure types, optimized for publication quality and editability.

## Available Templates

### Genomics and Multi-Omics

#### G×E Integration Model (based on fig3-1)
**File**: `templates/genomics/gxe_integration_model.vsdx`

**Use case**: Genome-by-environment interaction modeling, multi-omics integration pipeline

**Components**:
- SNP encoding methods (-1/0/1, 0/1/2, one-hot)
- DNA shape representation with built-in DNA icon
- Functional annotation tracks with semantic shapes
- Gene functional annotation with structured gene models
- Environmental time series panels
- Model comparison blocks (statistical, kernel-based, deep learning)
- Trait prediction outputs with plant icons (corn/wheat)
- Prediction performance scatter plots

**Color scheme**: Multi-omics palette (blue for genotype, purple for DNA shape, teal for annotation, gold for genes)

**Icons used**:
- DNA (MEDICAL_M.VSSX::DNA)
- Corn (HOLIDAYS_M.VSSX::Corn) for plant traits

---

### Flowcharts and Pipelines

#### Analysis Pipeline
**File**: `templates/workflow/analysis_pipeline.vsdx`

**Use case**: Bioinformatics pipeline, data processing workflow

**Components**:
- Input data nodes
- Processing step boxes with rounded corners
- Decision diamonds
- Output result nodes
- Connector arrows with labels

**Color scheme**: Process flow palette (input=teal, process=blue, decision=orange, output=green)

---

### Comparison Matrices

#### Method Comparison Table
**File**: `templates/comparison/method_matrix.vsdx`

**Use case**: Comparing algorithms, methods, or experimental conditions

**Components**:
- Checkerboard matrix cells
- Column headers with semantic colors
- Row labels
- Legend panel

**Color scheme**: Neutral gray borders with categorical fills

---

## Creating a New Template

### 1. Design Principles

- **Editable text**: All labels use Times New Roman or Arial
- **Semantic icons**: Use Visio Masters from built-in stencils, not embedded images
- **Grouped structure**: Group related shapes (e.g., all parts of a gene model)
- **Color consistency**: Follow scientific-color-palettes.md
- **No embedded rasters**: Templates should contain only native shapes and stencil masters

### 2. Template Metadata

Each template should include:
- **Page properties**: Set to standard publication dimensions (e.g., 16×9 for wide figures, 8.5×11 for vertical)
- **Shape data**: Add custom properties to key shapes for easy identification
- **Grouped elements**: Pre-group multi-shape components
- **Named shapes**: Assign meaningful names to major shapes for programmatic access

### 3. Coordinate System

Use reference coordinate calibration:
```powershell
$PageW = 16.0  # inches
$PageH = 9.0
$RefW = 1448.0  # reference pixel width
$RefH = 810.0   # reference pixel height

function VX([double]$x) { $PageW * $x / $RefW }
function VY([double]$y) { $PageH - ($PageH * $y / $RefH) }
```

### 4. Saving Templates

- Save as `.vsdx` (not `.vst` or `.vstx`) for broader compatibility
- Include a `README.txt` inside the template directory with usage instructions
- Test reopening in Visio to verify all icons load correctly

---

## Template Usage Examples

### Python script to load and customize a template:

```python
import win32com.client

visio = win32com.client.Dispatch("Visio.Application")
doc = visio.Documents.Open("templates/genomics/gxe_integration_model.vsdx")
page = doc.Pages.Item(1)

# Find and modify a specific shape
for shape in page.Shapes:
    if shape.Name == "ModelTitle":
        shape.Text = "My Custom Model"

doc.SaveAs("figures/my_figure.vsdx")
doc.Close()
visio.Quit()
```

### PowerShell script example:

```powershell
$visio = New-Object -ComObject Visio.Application
$doc = $visio.Documents.Open("templates/genomics/gxe_integration_model.vsdx")
$page = $doc.Pages.Item(1)

# Modify shapes
foreach ($shape in $page.Shapes) {
    if ($shape.Name -eq "TraitLabel1") {
        $shape.Text = "Grain Yield (GY)"
    }
}

$doc.SaveAs("figures/my_figure.vsdx")
$doc.Close()
$visio.Quit()
```

---

## Contributing New Templates

When adding a new template:

1. Create the `.vsdx` file following the design principles above
2. Document it in this file with use case, components, color scheme, and icons used
3. Add example usage code if the template has complex structure
4. Commit with a descriptive message: `feat(templates): add [type] template for [use case]`

---

## Template Checklist

Before committing a template, verify:

- [ ] No embedded raster images in `visio/media/`
- [ ] All text uses standard fonts (Times New Roman, Arial)
- [ ] Icons are from built-in Visio stencils with documented NameU
- [ ] Colors follow scientific-color-palettes.md
- [ ] Page size is set to standard dimensions
- [ ] File size < 200 KB (pure shape templates should be small)
- [ ] Template opens successfully in Visio 2016+ without errors
- [ ] All major shapes have meaningful names
- [ ] Related shapes are pre-grouped
