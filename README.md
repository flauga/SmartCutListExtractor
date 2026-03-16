# SmartCutList — Fusion 360 Add-In

A Fusion 360 Python add-in that automatically generates smart cut lists from selected components and bodies. It classifies parts by geometry, extracts dimensions and materials, and exports production-ready cut lists, DXF flat patterns, STEP files, and procurement summaries.

## Features

- **Automatic Part Classification** — Rule-based classifier identifies 9 part types from geometry, material, and naming conventions
- **Sheet Metal Support** — Detects sheet metal bodies, reads bend thickness, exports DXF flat patterns
- **Fastener Detection** — Identifies content-library fasteners (Insert > Fasteners) with thread spec, head type, and category parsing
- **Sourced Components** — Tracks externally linked/referenced components separately for procurement
- **3D Print Export** — STEP file export for 3D-printed parts
- **Weld Assembly Detection** — Groups multi-body components as weld assemblies with viewport highlighting
- **Wall Thickness Estimation** — Area-weighted detection for hollow sections
- **Interactive Review Palette** — Tabbed HTML UI for reviewing, overriding, and exporting classified parts
- **Multiple Export Formats** — CSV cut lists, JSON data, DXF flat patterns, STEP files, fastener procurement CSV

## Installation

1. Download or clone this repository
2. Copy the `SmartCutList/` folder into your Fusion 360 add-ins directory: (optional)
   - **Windows**: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\`
   - **macOS**: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`
3. In Fusion 360, go to **Utilities > Scripts and Add-Ins > Add-Ins** ( Can do this directly after downloading )
4. Find **SmartCutList** in the list and click **Run**

## Usage

1. **Select Components** — Click the SmartCutList button in the Solid > Make panel. Select one or more components or bodies from your design.
2. **Configure Options** — Choose whether to include sub-components and fasteners.
3. **Review Palette** — A tabbed review window opens with:
   - **Cut List** — All structural parts grouped by type, material, and dimensions
   - **Fasteners** — Content-library fasteners grouped for procurement
   - **Welds** — Multi-body weld assemblies (click to highlight in viewport)
   - **3D Prints** — Parts classified as 3D-printed with STEP export
   - **Sourced** — Externally referenced components
4. **Export** — Export individual formats or use "Export Package" for everything at once.

## Part Classification Types

| Type | Description |
|------|-------------|
| **Sheet Metal** | Fusion 360 sheet metal bodies (API flag) |
| **Hollow Rectangular Channel** | Square/rectangular tubes, channels — hollow with elongated profile |
| **Hollow Circular Cylinder** | Round tubes, pipes — hollow cylindrical sections |
| **Solid Cylinder** | Solid round bar stock |
| **Solid Block** | Solid rectangular stock (may have holes) |
| **3D Printed** | Plastic/resin material bodies for additive manufacturing |
| **Fastener** | Content-library fasteners (screws, bolts, nuts, washers) |
| **Sourced Component** | Externally linked/referenced components (bearings, motors, purchased parts) |
| **Unknown** | Parts that don't match any classification rule |

## Export Formats

| Format | Contents |
|--------|----------|
| **CSV** | Full cut list with part name, type, material, dimensions, quantity |
| **JSON** | Machine-readable export with all classification data |
| **DXF** | Sheet metal flat patterns and profile cross-sections |
| **STEP** | 3D-printed part geometry for slicing/printing services |
| **Fastener CSV** | Procurement list with thread spec, head type, category, quantity |
| **Sourced CSV** | Externally referenced components with dimensions and quantities |

## Project Structure

```
SmartCutList/
├── SmartCutList.py                # Entry point — run(context) / stop(context)
├── SmartCutList.manifest          # Fusion 360 add-in manifest (v1.0.0)
├── commands/
│   ├── __init__.py                # Module init — delegates start() / stop()
│   ├── select_components.py       # Command UI, component selection, body collection
│   ├── feature_extraction.py      # Geometry & material extraction from BRepBodies
│   ├── classifier.py              # Rule-based part classification engine
│   ├── review_palette.py          # HTML palette lifecycle & message handlers
│   ├── export_cutlist.py          # CSV / JSON export
│   ├── export_dxf.py              # DXF flat patterns & STEP export
│   └── settings.py                # Settings persistence & logging
└── resources/
    └── review_palette.html        # Interactive review UI (HTML/CSS/JS)
```

## Classification Pipeline

```
Select Components → Extract Features → Classify → Review Palette → Export
                         │                  │
                    OBB, fill ratio,    Priority-ordered
                    face analysis,      rule chain:
                    wall thickness,     1. Sheet metal (API flag)
                    material name       2. Content-library fastener
                                        3. Sourced component
                                        4. 3D printed (material)
                                        5. Fastener (name keywords)
                                        6. Geometry rules
                                        7. Fallback → Unknown
```

## Requirements

- Autodesk Fusion 360 (Windows or macOS)
- Python 3.x (bundled with Fusion 360)
- No external dependencies — uses only Fusion 360's built-in Python API (`adsk.core`, `adsk.fusion`)

## License

MIT License
