# Scientific Figure Color Palettes

Academic and publication-quality color schemes for scientific diagrams, optimized for accessibility and print reproduction.

## General Principles

- **Contrast:** Ensure sufficient contrast between adjacent colors and between colors and white/black text
- **Print-safe:** Test colors in grayscale; critical distinctions should remain visible
- **Colorblind-friendly:** Avoid red-green as the only distinction; use blue-orange or purple-yellow pairs
- **Consistent semantics:** Data types (e.g., genomic, proteomic, environmental) should have consistent color families across figures

## Genomics and Bioinformatics

### DNA/RNA bases
```powershell
$BasePair = @{
    Adenine  = RGBF 76 175 80   # Green
    Cytosine = RGBF 255 152 0   # Orange  
    Guanine  = RGBF 33 150 243  # Blue
    Thymine  = RGBF 244 67 54   # Red
    Uracil   = RGBF 156 39 176  # Purple
}
```

### Genomic features
```powershell
$Genomic = @{
    SNP         = RGBF 56 135 190   # Blue
    Gene        = RGBF 198 147 62   # Gold/Brown
    Regulatory  = RGBF 119 172 84   # Green
    Structural  = RGBF 139 107 181  # Purple
    Epigenetic  = RGBF 240 147 43   # Orange
}
```

### Data types
```powershell
$DataType = @{
    Transcriptomic = RGBF 31 95 184     # Dark Blue
    Proteomic      = RGBF 139 107 181   # Purple
    Metabolomic    = RGBF 119 172 84    # Green
    Phenotypic     = RGBF 240 147 43    # Orange
    Environmental  = RGBF 88 163 173    # Teal
}
```

## Multi-Omics Integration

For figures showing integration across omics layers (like the G×E example):

```powershell
$MultiOmics = @{
    Genotype      = RGBF 56 135 190    # Blue
    DNAShape      = RGBF 139 107 181   # Purple  
    Annotation    = RGBF 88 163 173    # Teal
    GeneFunction  = RGBF 198 147 62    # Gold
    Environment   = RGBF 88 163 173    # Teal
    Phenotype     = RGBF 119 172 84    # Green
}

# Soft backgrounds for matrix panels
$SoftBG = @{
    BlueSoft    = RGBF 232 244 252
    PurpleSoft  = RGBF 243 238 250
    TealSoft    = RGBF 224 242 241
    GoldSoft    = RGBF 238 220 176
    GreenSoft   = RGBF 232 245 233
}
```

## Model Types

Distinguish statistical, kernel, and deep learning models:

```powershell
$ModelType = @{
    Statistical   = RGBF 95 95 95       # Gray - neutral, traditional
    KernelBased   = RGBF 31 95 184      # Dark Blue - mathematical
    DeepLearning  = RGBF 31 95 184      # Dark Blue background
    # For deep learning, use colored layer nodes:
    InputLayer    = RGBF 88 163 173     # Teal
    HiddenLayer   = RGBF 56 135 190     # Blue
    OutputLayer   = RGBF 139 107 181    # Purple
}
```

## Plant Traits

For crop phenotypes (corn, wheat, etc.):

```powershell
$PlantTrait = @{
    Morphological = RGBF 119 172 84    # Green (height, architecture)
    Yield         = RGBF 240 147 43    # Orange (grain, biomass)
    Quality       = RGBF 198 147 62    # Gold (protein, oil)
    Phenology     = RGBF 88 163 173    # Teal (flowering, maturity)
    Stress        = RGBF 244 67 54     # Red (drought, disease)
}
```

## Neutral Palette

General-purpose colors for boxes, containers, and non-semantic elements:

```powershell
$Neutral = @{
    Black      = RGBF 17 17 17
    DarkGray   = RGBF 95 95 95
    Gray       = RGBF 158 158 158
    LightGray  = RGBF 220 220 220
    White      = RGBF 255 255 255
}
```

## Usage Guidelines

1. **Primary distinction by hue:** Use different hue families (blue, green, orange, purple) for major categories
2. **Secondary distinction by saturation/lightness:** Within a category, use lighter tints for sub-types or backgrounds
3. **Lines and borders:** Use darker, more saturated versions of the fill color, or use `$Neutral.DarkGray` for all borders
4. **Text on colored backgrounds:** 
   - White text for dark backgrounds (luminance < 50%)
   - Black text for light backgrounds and tinted panels
5. **Matrices and heatmaps:** Use one primary color at varied opacity/lightness; checkerboard alternates primary with white

## Testing

Before finalizing a figure:
- Print in grayscale and verify all distinctions remain clear
- Use a colorblind simulator (e.g., Coblis) to check deuteranopia and protanopia views
- Ensure all text meets WCAG AA contrast ratio (4.5:1 for body text, 3:1 for large text)
