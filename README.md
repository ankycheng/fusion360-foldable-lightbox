# Foldable Lightbox — Fusion 360 Add-in

Parametric foldable lightbox generator for Fusion 360. One flat sheet with living hinges folds into a **Triangle / Trapezoid / Square** box. Designed for multi-material FDM printing with translucent fills.

![Status](https://img.shields.io/badge/status-beta-yellow) ![Fusion](https://img.shields.io/badge/Fusion%20360-Win%20%7C%20Mac-blue) ![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Three profiles** — Triangle, Trapezoid, Square (Square auto-syncs top/bottom width)
- **Flat sheet + living hinges** — thin-hinge grooves with configurable thickness and groove width
- **Text on front & back** — independent strings, unified size; 16 built-in fonts or any custom font
- **4 text modes**
  - `Flush Recess` — letters recessed into the sheet and filled flush
  - `Through-cut` — letters cut through entirely (ideal for backlit translucent fill)
  - `Emboss` — letters raised above the sheet
  - `Through-cut + Emboss` — both, for raised signage with light transmission
- **Optional end caps** with annular rabbet pockets that flush-fit the sheet
- **Auto-apply appearances** — pick Body + Text colors, assigned by component so islands inside letters (A, O, D, P, B, R, Q) stay sheet-colored
- **Auto-fit text** to panel; optionally auto-size box depth to fit text length
- **UI state persisted** between invocations

## Install

### Option A — Download the ZIP (easiest)

1. Download the latest `FoldableLightbox-vX.Y.Z.zip` from the [Releases page](https://github.com/ankycheng/fusion360-foldable-lightbox/releases)
2. Unzip — you'll get a folder named `FoldableLightbox/`
3. Move that folder into Fusion 360's AddIns directory:
   - **macOS**: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`
   - **Windows**: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\`
4. Restart Fusion 360 (or open `Utilities → ADD-INS → Scripts and Add-Ins`)

### Option B — Clone the repo

```bash
git clone https://github.com/ankycheng/fusion360-foldable-lightbox.git
# then symlink or copy the FoldableLightbox/ folder into the AddIns path above
```

### Activate

1. In Fusion 360: `Utilities → ADD-INS → Scripts and Add-Ins`
2. Click the **Add-Ins** tab, find **Foldable Lightbox**
3. Click **Run** (or check "Run on Startup")
4. A new button appears under `Solid → Create → Foldable Lightbox`

## Usage

Click the **Foldable Lightbox** button on the Solid Create panel. A dialog opens with five groups:

| Group | Notes |
|---|---|
| **Profile** | Triangle / Trapezoid / Square. Square locks top_w = bot_w. |
| **Dimensions** | Top Width, Bottom Width, Profile Height, Depth (axial length). |
| **Sheet & Hinge** | Sheet Thickness, Living Hinge Thickness (`< sheet_t`), Groove Width. |
| **Text** | Front / Back strings, font, size, mode, Auto-fit, Auto-size Depth. |
| **End Caps** | Toggleable. Cap Thickness, Outer Ring Thickness (keeps ≥ 0.8 mm printable), Recess Depth, Clearance, Corner Fillet. |
| **Appearance** | Auto-apply colors — Body (Sheet + End Cap) and Text. |

Click **OK** — a `Lightbox_{Profile}` component appears at origin, containing the flat sheet (with hinges + text) and optional end caps laid out next to it. Export the sheet as one STL and the text as another for multi-material slicing.

## Print notes

- Print the sheet flat. Living hinges need thin walls (0.4 mm default) — match your nozzle/line width.
- For through-cut text: the enclosed inner "islands" of letters like O, A, D are kept as separate sheet-colored bodies. They fuse to the sheet on the first print layer. No support needed.
- End caps: the annular rabbet pocket fits over the sheet's thinned axial end — push-fit with 0.4 mm clearance per side by default.

## Changelog

### v0.2.0 (2026-04-20)
- Fix: polygon offset collapse detection — rejects degenerate / inverted offsets when profile is too small for the clearance
- Fix: `text_autosize_depth` bumps depth up so end-cap recess bands fit (previously could silently break the rabbet fit)

### v0.1.0 (initial)
- Triangle / Trapezoid / Square profiles
- Living hinges, end caps with annular rabbet, text modes, appearance presets

## Known issues

- End-cap fillet classification uses a median split — can misclassify corners on very asymmetric profiles
- Appearance fallback modifies the first `ColorProperty` of a plastic template — may not always be the visible base color in every Fusion library variant
- Back text is ignored if Front text is empty (symmetry bug — to be fixed)

## License

[MIT](LICENSE)
