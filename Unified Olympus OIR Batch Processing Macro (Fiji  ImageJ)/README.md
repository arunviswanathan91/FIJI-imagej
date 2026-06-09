# Unified Olympus OIR Batch Processing Macro (Fiji / ImageJ)

## Overview

Fiji / ImageJ macros for automated batch processing of Olympus OIR microscopy files. Handles channel splitting, Z-stack projection, and composite image generation across nested directory structures with no manual interaction.

---

## Repository Structure

```
.
├── unified fiji.ijm              # Original macro (with upscaling)
├── unified fiji_no_scale.ijm     # Revised macro (native resolution, quality-fixed)
├── README.md
├── input_data/                   # Raw .oir files (user-provided)
└── output_data/                  # Auto-generated outputs
    ├── Individual_Channels/
    └── Composite_Images/
```

---

## Macros

### `unified fiji.ijm`
The original processing macro. Opens each `.oir` file, detects whether it is a Z-stack or channel-only image, splits channels, generates a composite, and saves TIFFs. Applies a 10× upscale before export and uses a hardcoded inch-based scale.

### `unified fiji_no_scale.ijm` *(recommended)*
Revised version with the following changes over the original:

- **No upscaling** — exports at native camera resolution; removes the 10× Bilinear scale that caused pixelation
- **Contrast normalization** — applies `Enhance Contrast` per channel before 16→8-bit RGB conversion to preserve dynamic range
- **LUT-preserving channel export** — duplicates each channel individually from the composite hyperstack so fluorescence colors are retained in saved TIFFs
- **Correct composite rendering** — uses `Flatten` instead of `Stack to RGB`, which renders the composite exactly as displayed including LUT assignments
- **Calibration preserved** — removes the hardcoded `Set Scale` override; Bio-Formats reads the correct µm/pixel calibration from OIR metadata automatically

Both macros mirror the input subfolder structure in the output directories and support arbitrarily nested input folders.

---

## Requirements

- [Fiji](https://fiji.sc) (ImageJ ≥ 1.53 recommended)
- Bio-Formats plugin — included in standard Fiji; handles Olympus OIR import

---

## How to Run

1. Launch Fiji
2. Go to `Plugins → Macros → Edit…` and open the `.ijm` file
3. Click **Run** and choose input/output directories when prompted
4. Monitor progress in the **Log** window

---

## Output

Each run produces:

- `Individual_Channels/` — one TIFF per fluorescence channel per file, with LUT color applied
- `Composite_Images/` — merged composite TIFF with the transmission/brightfield channel excluded

Output folder structure mirrors the input directory tree.
