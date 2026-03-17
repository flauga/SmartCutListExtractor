# SmartCutList — Fusion 360 Add-In

A Fusion 360 Python add-in that automatically generates smart cut lists from selected components and bodies. It classifies parts by geometry, extracts dimensions and materials, and exports production-ready cut lists, DXF flat patterns, STEP files, and procurement summaries.

## Features

- **Automatic Part Classification** — Rule-based classifier identifies 9 part types from geometry, material, and naming conventions
- **Sheet Metal Support** — Detects sheet metal bodies, reads bend thickness, exports DXF flat patterns
- **Fastener Detection** — Identifies content-library fasteners (Insert > Fasteners) with thread spec, head type, and category parsing
- **Sourced Components** — Tracks externally linked/referenced components separately for procurement
- **3D Print Export** — STEP file export for 3D-printed parts
- **Hole Detection & Drill Planning** — Detects drilled/cut holes on hollow rectangular channel bodies, groups concentric through-holes, measures edge distances from channel ends, provides context-aware thread hints (tap for blind holes, clearance for through-holes), and captures perpendicular photos with Fusion driven dimension annotations and 6-side reference labelling (A/B ends, C/D/E/F faces)
- **Weld Assembly Detection & Weld Plan Documents** — Auto-detects multi-body weld assemblies, supports interactive step ordering (drag-and-drop reorder), nested sub-assemblies, per-step notes, and generates self-contained HTML weld plan documents with progressive viewport screenshots
- **Wall Thickness Estimation** — Area-weighted detection for hollow sections
- **Interactive Review Palette** — Tabbed HTML UI for reviewing, overriding, and exporting classified parts
- **Multiple Export Formats** — CSV cut lists, JSON data, DXF flat patterns, STEP files, fastener procurement CSV, drill list CSV
- **Export All** — Single-click export of all CSVs, DXFs, STEPs, and weld/drill plan documents to one folder

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
   - **Cut List** — All structural parts grouped by type, material, and dimensions. Includes select all/deselect all, type filtering, and search.
   - **Fasteners** — Content-library fasteners grouped for procurement
   - **Drill** — Holes detected on channel bodies that must be drilled before welding. Shows hole diameter, depth, type (through/blind/counterbored), face location, edge distances, and thread hints. Capture perpendicular photos with dimension annotation overlays showing channel outline, all four side labels (A/B ends, C/D cross-section), and hole position reference lines. Click any image to expand full-screen.
   - **Welds** — Weld assemblies with drag-to-reorder sequencing, per-step notes, nested sub-assemblies, and drill & weld plan document generation (click body name to highlight, Shift+click for cumulative highlight)
   - **3D Prints** — Parts classified as 3D-printed with STEP export
   - **Sourced** — Externally referenced components
4. **Export** — Export individual formats, DXFs only, or use "Export All" for everything at once (all CSVs, DXFs, STEPs, and documents).

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
| **Drill CSV** | Standalone drill list with hole dimensions, positions, thread hints, and face labels |
| **Weld Plan HTML** | Step-by-step drill & weld assembly document with drilling requirements before welding, progressive viewport screenshots, BOM tables, and per-step notes |

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
│   ├── hole_detection.py          # Hole detection on channel bodies for drill planning
│   ├── weld_plan_generator.py     # Drill & weld plan document generation with viewport capture
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

## Design Guidelines for Drill & Weld Planning

SmartCutList's weld planning features work best when you follow these modelling conventions in Fusion 360:

- **Each piece to be welded should be a separate body** in the component browser. For example, if a frame is built from four lengths of box section, model each length as its own body inside a single component (the weld assembly). SmartCutList detects multi-body components as weld assemblies and uses the individual bodies as weld steps.
- **Name your bodies** — Fusion 360 assigns default names like "Body1", "Body2". Giving each body a meaningful name (e.g. "Top_Rail", "Left_Leg") makes the generated weld plan document much easier to follow. SmartCutList will automatically prefix default names with the component name, but custom names are always clearer.
- **Use a single component per weld assembly** — Group all the bodies that will be welded together into one component. Sub-assemblies (a welded assembly inside another welded assembly) are supported via nesting.
- **Assign materials** — SmartCutList reads the material assigned to each body. Correct materials make the BOM table in the weld plan document accurate and useful for workshop prep.
- **Use Hole/Extrude features for drilled holes** — SmartCutList detects holes from cylindrical faces on channel bodies. Use Fusion's Hole feature or Extrude (Cut) with circular profiles so the creating sketch can be overlaid on photos. All holes must be drilled before the channels are welded together.

## Requirements

- Autodesk Fusion 360 (Windows or macOS)
- Python 3.x (bundled with Fusion 360)
- No external dependencies — uses only Fusion 360's built-in Python API (`adsk.core`, `adsk.fusion`)

## Changelog

### 6-Side Labelling, Driven Dimensions & Drill Grouping (v1.2.2)

**Improved: Annotation Sketch — Driven Dimensions**
- Replaced all manual construction lines with **Fusion 360 native sketch dimensions** (`addDistanceDimension`) from End A to each hole centre — these show as proper dimension callouts with arrows and values
- Removed channel outline rectangle, tick marks, leader lines, and construction geometry — only hole circles, crosshair lines, text labels, and driven dimensions are visible
- Added `sketch.arePointsShown = False` to hide construction points from screenshots

**Improved: 6-Side Reference System (A–F)**
- Channels now have **six named reference sides**: A/B (lengthwise ends), C/D (largest cross-section faces), E/F (remaining faces)
- Face labels in hole detection changed from Top/Bottom/Left/Right to C/D/E/F — consistent, orientation-independent identifiers
- Each drill screenshot caption now reads **"Drilling into Side X"** indicating which face the holes are on
- Side labels stay in the **same positions** regardless of which face is being photographed — only the camera rotates

**Improved: Larger, Clearer Labels**
- Text height for A/B labels now scales proportional to **channel length** (`ch_span × 0.08`), not just cross-section — much more visible at screenshot resolution
- Labels placed **further from channel edges** with increased `label_gap` (`ch_span × 0.04`)
- Each photo includes a small **"X (this side)"** indicator showing the camera-facing face

**New: Channel Grouping in Drill Tab**
- Channels with the **same dimensions and material** are grouped together under a shared header (e.g. "Steel — 50.0 × 50.0 × 600.0 mm (3 pieces, same setup)")
- Makes it easy to identify which channels need the **same drill setup** on the bench

**Improved: Hole Type Labels**
- "Through-Hole" renamed to **"Through Hole"** (hole goes through the channel wall)
- "Blind Hole" renamed to **"One Side"** (hole on one face only, does not penetrate through)

**Improved: Image Viewer — Zoom & Pan**
- Lightbox now supports **scroll-to-zoom** (towards cursor) and **click-drag to pan** — works in both the review palette and exported HTML documents
- Close button (×), Escape key, or click-outside-image to dismiss
- Hint bar shows usage instructions

### Annotation & Image Improvements (v1.2.1)

**Improved: Annotation Sketch Side Labels**
- Side A and B markers now correctly placed at the **lengthwise ends** of the channel (left/right in the screenshot), regardless of sketch coordinate orientation
- Channel axis direction is projected into sketch space to determine the correct lengthwise orientation — eliminates mislabelled sides when the sketch X axis doesn't align with the channel length
- Added **Side C** and **Side D** labels on the cross-section faces (top/bottom in the screenshot) so all four visible sides of the channel are clearly identified in the annotation overlay
- Dimension reference lines from End A to each hole are drawn along the correct lengthwise axis

**New: Expandable & Zoomable Images**
- Drill tab images in the review palette are now **click-to-expand** — clicking opens a full-screen lightbox overlay with the image at maximum resolution; press Escape or click outside to close
- Weld plan HTML documents include a **zoom modal** — click any drill or weld step image to view it full-size in an overlay; Escape key to dismiss
- Images show `cursor: zoom-in` on hover to indicate interactivity

### Drill Planning & UI Polish (v1.2.0)

**New: Drill Tab & Hole Detection**
- Dedicated **Drill** tab between Fasteners and Welds for hole drilling requirements on channel bodies
- Detects all cylindrical faces on hollow rectangular channel bodies as holes
- Groups concentric through-holes as a single entry (entry + exit faces merged)
- Context-aware thread hints: blind holes → tap drill match, through-holes → clearance match
- Captures perpendicular photos with temporary annotation sketch overlay showing:
  - Channel outline rectangle (outer dimensions)
  - Hole position circles with crosshairs
  - End A / End B reference labels
  - Dimension reference lines from End A to each hole centre
  - Channel length dimension line
- Hole images displayed 2-per-row (compact layout matching weld step cards)
- Drilling requirements section placed before weld sequence in generated documents

**New: Export All**
- Single-click "Export All" button that generates all CSVs (cut list, drill list, fasteners, sourced), DXF flat patterns, STEP files, and drill & weld plan documents into one folder

**New: DXFs Only Export**
- Dedicated "DXFs Only" export button on the Cut List tab

**Improved: UI Polish**
- Bulk action bar (select all, deselect all, filter, search) now only visible on Cut List tab
- Sheet metal body names in dimension summary now show "ComponentName_Body1" instead of "Body1"
- Weld plan button renamed to "Generate Drill and Weld Plan"
- Row highlighting in Drill and Weld tabs — clicking a row highlights it in the UI and the corresponding entity in the viewport
- Weld step camera resets to isometric view before capture (no longer inherits perpendicular hole camera angle)

### Weld Plan Feature (v1.1.0)

**New: Weld Plan Document Generator**
- Auto-detects weld assemblies (multi-body components) and builds ordered weld step sequences
- Generates self-contained HTML documents with progressive viewport screenshots showing each weld step
- Each step card includes body name, material, dimensions, and user-written notes
- Steps rendered in a compact 2-up grid layout (2 cards per row) with image above info
- Bill of Materials table per assembly
- Table of contents with jump links when multiple assemblies exist
- Print-friendly CSS for clean paper/PDF output

**New: Interactive Weld Sequence Ordering**
- Drag-and-drop step reordering via grab handles in the Welds tab (mouse-based, works reliably in Fusion 360's CEF browser)
- Per-step description fields for weld process notes (weld type, amps, tack-first instructions, etc.)
- Assembly-level notes textarea for general instructions (preheat, material spec, etc.)
- Add bodies to weld assemblies from any group via dropdown
- Create brand-new manual weld assemblies from selected bodies
- Delete assemblies, remove individual steps

**New: Nested Weld Assemblies**
- Sub-assemblies automatically detected from component path hierarchy
- Nested assemblies appear as composite steps in the parent sequence
- Promote (unnest) sub-assemblies back to top-level

**New: Viewport Highlighting**
- Click a body name to highlight it in the Fusion 360 viewport
- Shift+click a body name to cumulatively highlight that body and every body above it in the weld sequence (shows progressive weld state)
- Preview button (eye icon) hides all other bodies and shows cumulative progress up to that step

**Improved: Body Naming**
- Bodies with default Fusion 360 names ("Body1", "Body2", etc.) are automatically prefixed with their parent component name for clarity (e.g. "Frame_Tube_50x50_Body1")
- Applied consistently across the palette UI, generated documents, and JSON export

**Improved: JSON Export**
- JSON export now includes a `weld_plans` array with full step sequences, notes, and nesting information

## License

MIT License
