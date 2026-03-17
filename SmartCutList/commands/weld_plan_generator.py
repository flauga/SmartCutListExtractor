"""
weld_plan_generator.py — Generate step-by-step weld plan documents.

For each weld assembly the generator:
  1. Saves the current visibility state of all bodies in the design
  2. Hides all bodies
  3. Progressively reveals bodies in weld-step order, capturing a viewport
     screenshot at each stage
  4. Restores the original visibility state
  5. Produces a self-contained HTML document with Base64-embedded images

Public API
----------
    generate_weld_plan(weld_plans, output_dir, settings) -> str
    preview_weld_step(weld_plans, weld_id, up_to_step) -> None
"""

from __future__ import annotations

import base64
import logging
import os
import traceback
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import adsk.core
import adsk.fusion

from . import export_dxf
from . import hole_detection

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_IMAGE_WIDTH = 1920
_IMAGE_HEIGHT = 1080


# ---------------------------------------------------------------------------
# Visibility state management
# ---------------------------------------------------------------------------

def _save_visibility_state(
    design: adsk.fusion.Design,
) -> List[Tuple[Any, bool]]:
    """Record isVisible for every BRepBody in the design.

    Returns a list of (body, was_visible) tuples so we can restore exactly.
    Using direct object references is safe within a single operation.
    """
    state: List[Tuple[Any, bool]] = []
    try:
        for comp in design.allComponents:
            for body in comp.bRepBodies:
                try:
                    state.append((body, body.isVisible))
                except Exception:
                    pass
    except Exception:
        logger.exception('Failed to save full visibility state')
    return state


def _restore_visibility_state(state: List[Tuple[Any, bool]]) -> None:
    """Restore previously saved visibility state."""
    for body, was_visible in state:
        try:
            body.isVisible = was_visible
        except Exception:
            pass


def _hide_all_bodies(design: adsk.fusion.Design) -> None:
    """Set isVisible=False on every BRepBody in the design."""
    for comp in design.allComponents:
        for body in comp.bRepBodies:
            try:
                body.isVisible = False
            except Exception:
                pass


def _resolve_step_bodies(
    step: dict,
    all_plans: list,
) -> List[adsk.fusion.BRepBody]:
    """Resolve the BRepBody objects for a single step.

    For sub-assembly steps, recursively resolves all bodies in the child
    weld assembly.
    """
    bodies: List[adsk.fusion.BRepBody] = []

    if step.get('is_sub_assembly'):
        sub_id = step.get('sub_assembly_weld_id', '')
        sub_weld = None
        for w in all_plans:
            if w['weld_id'] == sub_id:
                sub_weld = w
                break
        if sub_weld:
            for sub_step in sub_weld['steps']:
                bodies.extend(_resolve_step_bodies(sub_step, all_plans))
    else:
        token = step.get('body_token', '')
        if token:
            body = export_dxf.resolve_body_token(token)
            if body is not None:
                bodies.append(body)

    return bodies


# ---------------------------------------------------------------------------
# Viewport capture
# ---------------------------------------------------------------------------

def _capture_weld_steps(
    weld: dict,
    all_plans: list,
    output_dir: str,
    image_width: int = _IMAGE_WIDTH,
    image_height: int = _IMAGE_HEIGHT,
) -> List[dict]:
    """Capture progressive viewport screenshots for one weld assembly.

    Returns a list of capture dicts:
        {step_index, image_path, bodies_shown, body_names, description}
    """
    app = adsk.core.Application.get()
    design = adsk.fusion.Design.cast(app.activeProduct)
    viewport = app.activeViewport

    captures: List[dict] = []
    vis_state = _save_visibility_state(design)

    try:
        _hide_all_bodies(design)

        # Reset to isometric view so weld captures start from a clean angle
        # (hole photography may have left the camera in a perpendicular view)
        try:
            camera = viewport.camera
            camera.isFitView = False
            camera.isSmoothTransition = False
            camera.viewOrientation = adsk.core.ViewOrientations.IsoTopRightViewOrientation
            viewport.camera = camera
        except Exception:
            pass

        shown_bodies: List[adsk.fusion.BRepBody] = []
        shown_names: List[str] = []

        for step in weld['steps']:
            # Resolve bodies for this step
            step_bodies = _resolve_step_bodies(step, all_plans)
            if not step_bodies:
                logger.warning(
                    'Step %s in weld %s resolved no bodies; skipping capture',
                    step.get('step_index'), weld.get('weld_id'),
                )
                # Still record the step with no image
                captures.append({
                    'step_index': step['step_index'],
                    'image_path': '',
                    'bodies_shown': len(shown_bodies),
                    'body_names': list(shown_names),
                    'description': step.get('description', ''),
                    'body_name': step.get('body_name', ''),
                    'material_name': step.get('material_name', ''),
                    'dimensions_mm': step.get('dimensions_mm'),
                    'is_sub_assembly': step.get('is_sub_assembly', False),
                })
                shown_names.append(step.get('body_name', ''))
                continue

            # Make newly added bodies visible
            for body in step_bodies:
                try:
                    body.isVisible = True
                    shown_bodies.append(body)
                except Exception:
                    pass
            shown_names.append(step.get('body_name', ''))

            # Frame and capture
            try:
                viewport.refresh()
                viewport.fit()
            except Exception:
                pass

            image_name = '{}_{}.png'.format(
                weld['weld_id'], step['step_index']
            )
            image_path = os.path.join(output_dir, image_name)

            captured = False
            try:
                captured = viewport.saveAsImageFile(
                    image_path, image_width, image_height
                )
            except Exception:
                logger.exception('saveAsImageFile failed for step %s', step['step_index'])

            captures.append({
                'step_index': step['step_index'],
                'image_path': image_path if captured else '',
                'bodies_shown': len(shown_bodies),
                'body_names': list(shown_names),
                'description': step.get('description', ''),
                'body_name': step.get('body_name', ''),
                'material_name': step.get('material_name', ''),
                'dimensions_mm': step.get('dimensions_mm'),
                'is_sub_assembly': step.get('is_sub_assembly', False),
            })

    finally:
        _restore_visibility_state(vis_state)
        try:
            viewport.refresh()
        except Exception:
            pass

    return captures


# ---------------------------------------------------------------------------
# Preview (no screenshots, just visibility)
# ---------------------------------------------------------------------------

def preview_weld_step(
    weld_plans: list,
    weld_id: str,
    up_to_step: int,
) -> None:
    """Show cumulative weld progress in the viewport up to a given step.

    This is a live preview — it modifies body visibility but does NOT
    save/restore state or capture images.  The user can click away from
    the preview and the viewport will remain in that state until they
    interact with the design.
    """
    app = adsk.core.Application.get()
    design = adsk.fusion.Design.cast(app.activeProduct)
    viewport = app.activeViewport

    weld = None
    for w in weld_plans:
        if w['weld_id'] == weld_id:
            weld = w
            break
    if not weld:
        return

    _hide_all_bodies(design)

    for step in weld['steps']:
        if step['step_index'] > up_to_step:
            break
        step_bodies = _resolve_step_bodies(step, weld_plans)
        for body in step_bodies:
            try:
                body.isVisible = True
            except Exception:
                pass

    try:
        viewport.refresh()
        viewport.fit()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Hole photography — perpendicular viewport capture with sketch overlay
# ---------------------------------------------------------------------------

def _compute_up_vector(axis: tuple) -> tuple:
    """Compute a camera up-vector perpendicular to *axis*.

    Uses world-Z as seed unless the axis is near-vertical, then uses world-X.
    """
    ax, ay, az = axis
    # If axis is close to world-Z, seed with world-X
    if abs(az) > 0.9:
        seed = (1.0, 0.0, 0.0)
    else:
        seed = (0.0, 0.0, 1.0)
    # Cross product: seed × axis
    ux = seed[1] * az - seed[2] * ay
    uy = seed[2] * ax - seed[0] * az
    uz = seed[0] * ay - seed[1] * ax
    mag = (ux * ux + uy * uy + uz * uz) ** 0.5
    if mag < 1e-12:
        return (0.0, 1.0, 0.0)
    return (ux / mag, uy / mag, uz / mag)


def _create_annotation_sketch(
    body: adsk.fusion.BRepBody,
    face_holes: list,
    summary: dict,
    face_label: str,
) -> Optional[adsk.fusion.Sketch]:
    """Create a temporary annotation sketch on the channel body for photography.

    Draws:
    - Hole position circles with crosshairs (regular, visible lines)
    - Large "A"/"B" text at lengthwise ends, "C"/"D"/"E"/"F" on cross-section
      and top/bottom sides
    - Fusion sketch distance dimensions from End A to each hole centre
    - No construction lines — only text and driven dimensions visible

    The sketch coordinate system may not align with the camera orientation so
    we project the channel axis into sketch space to find which direction is
    "lengthwise" and place all annotations accordingly.

    Returns the sketch (caller must delete it after capture) or None on failure.
    """
    try:
        app = adsk.core.Application.get()
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design:
            return None

        channel_axis = summary.get('channel_axis')
        bb_min = summary.get('bb_min_cm')
        bb_max = summary.get('bb_max_cm')
        if not channel_axis or not bb_min or not bb_max:
            return None

        hole_axis = face_holes[0].get('axis', (0, 0, 1))
        target_face = None
        plane_type = adsk.core.Plane.classType()

        for face in body.faces:
            try:
                geom = face.geometry
                if geom.objectType != plane_type:
                    continue
                normal = geom.normal
                n = (normal.x, normal.y, normal.z)
                dot = abs(n[0] * hole_axis[0] + n[1] * hole_axis[1] + n[2] * hole_axis[2])
                if dot > 0.99:
                    target_face = face
                    break
            except Exception:
                pass

        if target_face is None:
            return None

        component = body.parentComponent
        if not component:
            return None

        planes = component.constructionPlanes
        plane_input = planes.createInput()
        offset_val = adsk.core.ValueInput.createByReal(0.05)  # 0.5mm offset
        plane_input.setByOffset(target_face, offset_val)
        construction_plane = planes.add(plane_input)

        sketch = component.sketches.add(construction_plane)
        sketch.name = '_SmartCutList_HoleAnnotation_'
        sketch.isVisible = True

        lines = sketch.sketchCurves.sketchLines
        circles = sketch.sketchCurves.sketchCircles
        texts = sketch.sketchTexts
        sk_points = sketch.sketchPoints
        sk_dims = sketch.sketchDimensions

        sketch_transform = sketch.transform
        inv_transform = sketch_transform.copy()
        inv_transform.invert()

        def _w2s(world_cm):
            pt3d = adsk.core.Point3D.create(world_cm[0], world_cm[1], world_cm[2])
            pt3d.transformBy(inv_transform)
            return pt3d

        # --- Project bounding box corners to sketch plane ---
        bb_corners = []
        for x_pick in [bb_min[0], bb_max[0]]:
            for y_pick in [bb_min[1], bb_max[1]]:
                for z_pick in [bb_min[2], bb_max[2]]:
                    bb_corners.append(_w2s((x_pick, y_pick, z_pick)))

        sx_min = min(p.x for p in bb_corners)
        sx_max = max(p.x for p in bb_corners)
        sy_min = min(p.y for p in bb_corners)
        sy_max = max(p.y for p in bb_corners)

        # -----------------------------------------------------------
        # Determine which sketch-space axis is the channel LENGTH.
        # -----------------------------------------------------------
        bb_center_w = (
            (bb_min[0] + bb_max[0]) / 2.0,
            (bb_min[1] + bb_max[1]) / 2.0,
            (bb_min[2] + bb_max[2]) / 2.0,
        )
        ref_sk = _w2s(bb_center_w)
        along_w = (
            bb_center_w[0] + channel_axis[0],
            bb_center_w[1] + channel_axis[1],
            bb_center_w[2] + channel_axis[2],
        )
        along_sk = _w2s(along_w)
        ch_dir_x = along_sk.x - ref_sk.x
        ch_dir_y = along_sk.y - ref_sk.y

        length_along_x = abs(ch_dir_x) >= abs(ch_dir_y)

        if length_along_x:
            if ch_dir_x >= 0:
                len_min, len_max = sx_min, sx_max
            else:
                len_min, len_max = sx_max, sx_min
            cross_min, cross_max = sy_min, sy_max
            ch_span = abs(sx_max - sx_min)
            cr_span = abs(sy_max - sy_min)
        else:
            if ch_dir_y >= 0:
                len_min, len_max = sy_min, sy_max
            else:
                len_min, len_max = sy_max, sy_min
            cross_min, cross_max = sx_min, sx_max
            ch_span = abs(sy_max - sy_min)
            cr_span = abs(sx_max - sx_min)

        # --- Scale text sizes proportional to CHANNEL LENGTH ---
        # Much larger than before so they're clearly visible in screenshots
        text_h = max(0.5, min(ch_span * 0.08, 2.0))     # proportional to length
        text_h_side = max(0.4, min(cr_span * 0.3, 1.5))  # cross-section labels
        label_gap = max(0.4, ch_span * 0.04)              # gap from channel edge
        dim_gap = max(0.3, cr_span * 0.18)                # gap for dimension lines

        # -----------------------------------------------------------
        # Helper: add a large text label centred at (cx, cy)
        # -----------------------------------------------------------
        def _add_label(text_str, cx, cy, height):
            try:
                ti = texts.createInput2(text_str, height)
                half_h = height * 0.6
                ti.setAsMultiLine(
                    adsk.core.Point3D.create(cx - height * 0.5, cy - half_h, 0),
                    adsk.core.Point3D.create(cx + height * 0.5, cy + half_h, 0),
                    adsk.core.HorizontalAlignments.CenterHorizontalAlignment,
                    adsk.core.VerticalAlignments.MiddleVerticalAlignment,
                    0,
                )
                texts.add(ti)
            except Exception:
                pass

        # --- Draw hole circles and crosshairs (regular lines, NOT construction) ---
        CM_TO_MM = 10.0
        hole_sk_points = []  # SketchPoint objects for dimensions
        for hole in face_holes:
            center_cm = hole.get('center_cm', (0, 0, 0))
            radius_cm = hole.get('radius_mm', 2.0) / CM_TO_MM

            c = _w2s(center_cm)
            c2 = adsk.core.Point3D.create(c.x, c.y, 0)

            circles.addByCenterRadius(c2, radius_cm)

            # Crosshair (regular lines — visible in screenshots)
            cx_sz = max(radius_cm * 1.8, 0.15)
            try:
                for d in [(cx_sz, 0), (0, cx_sz)]:
                    lines.addByTwoPoints(
                        adsk.core.Point3D.create(c2.x - d[0], c2.y - d[1], 0),
                        adsk.core.Point3D.create(c2.x + d[0], c2.y + d[1], 0),
                    )
            except Exception:
                pass

            # Create a SketchPoint at the hole centre for dimension references
            try:
                hole_pt = sk_points.add(c2)
                hole_sk_points.append(hole_pt)
            except Exception:
                hole_sk_points.append(None)

        # -----------------------------------------------------------
        # Create End A reference point for dimensions
        # -----------------------------------------------------------
        try:
            if length_along_x:
                mid_cross_val = (sy_min + sy_max) / 2.0
                end_a_pt = sk_points.add(
                    adsk.core.Point3D.create(len_min, mid_cross_val, 0)
                )
            else:
                mid_cross_val = (sx_min + sx_max) / 2.0
                end_a_pt = sk_points.add(
                    adsk.core.Point3D.create(mid_cross_val, len_min, 0)
                )
        except Exception:
            end_a_pt = None

        # -----------------------------------------------------------
        # Add Fusion sketch dimensions from End A to each hole
        # -----------------------------------------------------------
        if end_a_pt is not None:
            for hi, hole_pt in enumerate(hole_sk_points):
                if hole_pt is None:
                    continue
                try:
                    if length_along_x:
                        orientation = adsk.fusion.DimensionOrientations.HorizontalDimensionOrientation
                        text_y = sy_max + dim_gap * (hi + 1)
                        text_pos = adsk.core.Point3D.create(
                            (len_min + hole_pt.geometry.x) / 2.0, text_y, 0
                        )
                    else:
                        orientation = adsk.fusion.DimensionOrientations.VerticalDimensionOrientation
                        text_x = sx_max + dim_gap * (hi + 1)
                        text_pos = adsk.core.Point3D.create(
                            text_x, (len_min + hole_pt.geometry.y) / 2.0, 0
                        )
                    sk_dims.addDistanceDimension(
                        end_a_pt, hole_pt, orientation, text_pos
                    )
                except Exception:
                    logger.debug('Distance dimension failed for hole %d', hi)

        # -----------------------------------------------------------
        # Side labels — placed well away from channel edges
        # A/B at lengthwise ends, C/D on visible cross-section edges,
        # E/F for the camera-facing / camera-away faces (noted in caption)
        # -----------------------------------------------------------
        if length_along_x:
            mid_cross = (sy_min + sy_max) / 2.0

            # A (left/start) and B (right/end)
            _add_label('A', len_min - label_gap - text_h * 0.6, mid_cross, text_h)
            _add_label('B', len_max + label_gap + text_h * 0.6, mid_cross, text_h)

            # C and D on the cross-section edges (top/bottom of visible rectangle)
            mid_len = (sx_min + sx_max) / 2.0
            _add_label('C', mid_len, cross_max + label_gap + text_h_side * 0.5, text_h_side)
            _add_label('D', mid_len, cross_min - label_gap - text_h_side * 0.5, text_h_side)

            # E label for camera-facing side (the face we're photographing)
            # Placed in the top-right corner area as a small indicator
            _add_label(
                face_label + ' (this side)',
                len_max - ch_span * 0.15, cross_max + label_gap * 2 + text_h_side,
                text_h_side * 0.6
            )
        else:
            mid_cross = (sx_min + sx_max) / 2.0

            # A (bottom/start) and B (top/end)
            _add_label('A', mid_cross, len_min - label_gap - text_h * 0.6, text_h)
            _add_label('B', mid_cross, len_max + label_gap + text_h * 0.6, text_h)

            # C and D on visible edges
            mid_len = (sy_min + sy_max) / 2.0
            _add_label('C', cross_min - label_gap - text_h_side * 0.5, mid_len, text_h_side)
            _add_label('D', cross_max + label_gap + text_h_side * 0.5, mid_len, text_h_side)

            # Face indicator
            _add_label(
                face_label + ' (this side)',
                cross_max + label_gap * 2 + text_h_side, len_max - ch_span * 0.15,
                text_h_side * 0.6
            )

        # Hide the construction plane so only the sketch lines show
        try:
            construction_plane.isLightBulbOn = False
        except Exception:
            pass

        # Hide construction points so only circles, crosshairs, text, dims show
        try:
            sketch.arePointsShown = False
        except Exception:
            pass

        return sketch

    except Exception:
        logger.debug('Annotation sketch creation failed: %s', traceback.format_exc())
        return None


def _delete_annotation_sketch(sketch: Optional[adsk.fusion.Sketch]) -> None:
    """Delete a temporary annotation sketch and its construction plane."""
    if sketch is None:
        return
    try:
        # Get the reference plane before deleting the sketch
        ref_plane = None
        try:
            ref_plane = sketch.referencePlane
        except Exception:
            pass

        sketch.deleteMe()

        # Also delete the construction plane if we created one
        if ref_plane is not None:
            try:
                cp = adsk.fusion.ConstructionPlane.cast(ref_plane)
                if cp and cp.name.startswith('_SmartCutList') or True:
                    cp.deleteMe()
            except Exception:
                pass
    except Exception:
        logger.debug('Annotation sketch cleanup failed')


def capture_hole_images(
    hole_summaries: list,
    output_dir: str,
    image_width: int = _IMAGE_WIDTH,
    image_height: int = _IMAGE_HEIGHT,
) -> dict:
    """Capture perpendicular photos of holes on channel bodies.

    For each body, groups holes by face direction and captures one image per
    unique face with the body visible, the creating sketch(es) overlaid, and
    a temporary annotation sketch showing channel outline, hole positions,
    End A/B labels, and dimension reference lines.

    Parameters
    ----------
    hole_summaries : list
        Output of ``hole_detection.detect_holes_on_channels()``.
    output_dir : str
        Directory for intermediate image files.

    Returns
    -------
    dict
        Mapping of ``{body_token: {face_label: base64_str}}``.
    """
    if not hole_summaries:
        return {}

    app = adsk.core.Application.get()
    design = adsk.fusion.Design.cast(app.activeProduct)
    viewport = app.activeViewport

    os.makedirs(output_dir, exist_ok=True)
    result: dict = {}
    vis_state = _save_visibility_state(design)

    try:
        for summary in hole_summaries:
            token = summary.get('body_token', '')
            body = export_dxf.resolve_body_token(token)
            if body is None:
                continue

            body_captures: dict = {}

            # Group holes by face label (axis direction)
            face_groups: dict = {}  # {face_label: [hole_dicts]}
            for group in summary.get('concentric_groups', []):
                for hole in group.get('holes', []):
                    fl = hole.get('face_label', 'Unknown')
                    face_groups.setdefault(fl, []).append(hole)

            for face_label, face_holes in face_groups.items():
                _hide_all_bodies(design)

                # Show only the target body
                try:
                    body.isVisible = True
                except Exception:
                    pass

                # Collect unique sketch tokens and show those sketches
                shown_sketches: list = []
                sketch_tokens_seen: set = set()
                for hole in face_holes:
                    st = hole.get('creating_sketch_token', '')
                    if st and st not in sketch_tokens_seen:
                        sketch_tokens_seen.add(st)
                        try:
                            entities = design.findEntityByToken(st)
                            if entities:
                                sketch = adsk.fusion.Sketch.cast(entities[0])
                                if sketch:
                                    sketch.isVisible = True
                                    shown_sketches.append(sketch)
                        except Exception:
                            pass

                # Create temporary annotation sketch with dimension overlays
                annotation_sketch = _create_annotation_sketch(
                    body, face_holes, summary, face_label
                )

                # Position camera perpendicular to the hole axis.
                # Orient so End A (bb_min along channel axis) is on the LEFT
                # and End B (bb_max) is on the RIGHT of the image.
                if face_holes:
                    axis = face_holes[0].get('axis', (0, 0, 1))
                    center = face_holes[0].get('center_cm', (0, 0, 0))

                    try:
                        camera = viewport.camera
                        camera.isFitView = False
                        camera.isSmoothTransition = False

                        pull = 30.0  # cm pullback
                        camera.target = adsk.core.Point3D.create(
                            center[0], center[1], center[2]
                        )
                        camera.eye = adsk.core.Point3D.create(
                            center[0] + axis[0] * pull,
                            center[1] + axis[1] * pull,
                            center[2] + axis[2] * pull,
                        )

                        # Compute up-vector so channel axis runs left→right
                        # (End A on left, End B on right).
                        # The "up" in the image must be perpendicular to both
                        # the viewing direction (hole axis) and the channel axis.
                        ch_axis = summary.get('channel_axis')
                        if ch_axis:
                            # up = hole_axis × channel_axis
                            # This makes channel_axis point right in the image
                            up = (
                                axis[1] * ch_axis[2] - axis[2] * ch_axis[1],
                                axis[2] * ch_axis[0] - axis[0] * ch_axis[2],
                                axis[0] * ch_axis[1] - axis[1] * ch_axis[0],
                            )
                            mag = (up[0]**2 + up[1]**2 + up[2]**2) ** 0.5
                            if mag > 1e-6:
                                up = (up[0]/mag, up[1]/mag, up[2]/mag)
                            else:
                                up = _compute_up_vector(axis)
                        else:
                            up = _compute_up_vector(axis)

                        camera.upVector = adsk.core.Vector3D.create(up[0], up[1], up[2])
                        viewport.camera = camera
                        viewport.fit()
                        viewport.refresh()
                    except Exception:
                        logger.debug('Camera positioning failed for %s/%s', token, face_label)

                # Capture
                image_name = 'hole_{}_{}.png'.format(
                    token[:16].replace('/', '_'), face_label.replace(' ', '_')
                )
                image_path = os.path.join(output_dir, image_name)

                captured = False
                try:
                    captured = viewport.saveAsImageFile(
                        image_path, image_width, image_height
                    )
                except Exception:
                    logger.exception('saveAsImageFile failed for hole image')

                if captured:
                    body_captures[face_label] = _image_to_base64(image_path)
                    # Clean up temp file
                    try:
                        os.remove(image_path)
                    except Exception:
                        pass

                # Clean up annotation sketch
                _delete_annotation_sketch(annotation_sketch)

                # Hide creating sketches
                for sketch in shown_sketches:
                    try:
                        sketch.isVisible = False
                    except Exception:
                        pass

            if body_captures:
                result[token] = body_captures
                summary['face_images'] = body_captures

    finally:
        _restore_visibility_state(vis_state)
        try:
            viewport.refresh()
        except Exception:
            pass

    return result


# ---------------------------------------------------------------------------
# Hole section HTML rendering (for weld plan document)
# ---------------------------------------------------------------------------

def _render_hole_section_html(
    weld: dict,
    hole_summaries: list,
    unit: str,
) -> str:
    """Render the 'Drilling Requirements' HTML section for a weld assembly.

    Returns empty string if no channel bodies in this assembly have holes.
    """
    # Find which body tokens are in this weld assembly
    weld_body_tokens = set()
    for step in weld.get('steps', []):
        token = step.get('body_token', '')
        if token:
            weld_body_tokens.add(token)

    # Filter hole summaries to this assembly's bodies
    relevant = [s for s in hole_summaries if s.get('body_token', '') in weld_body_tokens]
    if not relevant:
        return ''

    factor = 1.0 / 25.4 if unit == 'inches' else 1.0
    suffix = 'in' if unit == 'inches' else 'mm'

    def _fmt(val):
        if val is None:
            return '-'
        return '{:.1f}'.format(val * factor)

    body_sections = ''
    for summary in relevant:
        body_name = _escape_html(summary.get('body_name', ''))
        material = _escape_html(summary.get('material_name', ''))
        total = summary.get('total_holes', 0)

        # Perpendicular images with dimension callouts
        images_html = ''
        for face_label, b64 in summary.get('face_images', {}).items():
            if b64:
                # Build dimension callout legend for holes on this face
                callout_items = []
                for cg in summary.get('concentric_groups', []):
                    for h in cg.get('holes', []):
                        if h.get('face_label', '') == face_label:
                            callout_items.append(
                                '<span style="margin-right:12px;">'
                                '<strong>Grp {gi}:</strong> '
                                '&empty;{dia} {u}, {ea} {u} from End A'
                                '</span>'.format(
                                    gi=cg['group_index'] + 1,
                                    dia=_fmt(h.get('diameter_mm')),
                                    ea=_fmt(h.get('distance_from_end_a_mm')),
                                    u=suffix,
                                )
                            )
                            break  # one callout per group

                callout_html = ''
                if callout_items:
                    callout_html = (
                        '<div style="font-size:10px;color:#555;background:#f5f7fa;'
                        'border:1px solid #dde3ec;border-top:none;border-radius:0 0 4px 4px;'
                        'padding:4px 8px;">'
                        + ''.join(callout_items) +
                        '</div>'
                    )

                images_html += '''
                <div style="display:inline-block;width:48%;vertical-align:top;margin:4px 1%;">
                  <div style="font-size:10px;color:#666;margin-bottom:3px;">
                    Drilling into Side {face_label} &mdash; End A (left) &rarr; End B (right)
                  </div>
                  <img src="{b64}" alt="{face_label} view"
                       style="width:100%;border:1px solid #dde3ec;border-radius:4px 4px 0 0;cursor:zoom-in;"
                       onclick="window._openZoom(this.src)" />
                  {callout_html}
                </div>
                '''.format(
                    face_label=_escape_html(face_label),
                    b64=b64,
                    callout_html=callout_html,
                )

        # Hole table
        rows = ''
        for group in summary.get('concentric_groups', []):
            gi = group['group_index']
            group_type = _escape_html(group.get('hole_type_label', ''))
            is_through = group.get('is_through', False)

            for hi, hole in enumerate(group.get('holes', [])):
                through_badge = (
                    '<span style="color:#2e7d32;font-weight:600;">Yes</span>'
                    if hole.get('is_through') else
                    '<span style="color:#e65100;">No</span>'
                )
                # Cross-section position
                offset_mm = hole.get('cross_offset_mm')
                face_w = hole.get('face_width_mm')
                if offset_mm is not None and face_w and face_w > 0:
                    if abs(offset_mm) < 0.5:
                        pos_str = '<span style="color:#2e7d32;">Centered</span>'
                    else:
                        half_w = face_w / 2.0
                        sa = half_w - offset_mm
                        sb = half_w + offset_mm
                        pos_str = '{}/{} {}'.format(_fmt(min(sa, sb)), _fmt(max(sa, sb)), suffix)
                else:
                    pos_str = '-'

                rows += '''
                <tr style="background:{bg};">
                  <td>{grp}</td>
                  <td>{dia} {suffix}</td>
                  <td>{depth}</td>
                  <td>{through}</td>
                  <td>{thread}</td>
                  <td>{face}</td>
                  <td>{end_a} {suffix}</td>
                  <td>{end_b} {suffix}</td>
                  <td>{position}</td>
                </tr>
                '''.format(
                    bg='#f8fafd' if gi % 2 == 0 else '#ffffff',
                    grp=gi + 1,
                    dia=_fmt(hole.get('diameter_mm')),
                    suffix=suffix,
                    depth='{} {}'.format(_fmt(hole.get('depth_mm')), suffix)
                        if hole.get('depth_mm') else '-',
                    through=through_badge,
                    thread=_escape_html(hole.get('thread_hint', '') or '-'),
                    face=_escape_html(hole.get('face_label', '')),
                    end_a=_fmt(hole.get('distance_from_end_a_mm')),
                    end_b=_fmt(hole.get('distance_from_end_b_mm')),
                    position=pos_str,
                )

        body_sections += '''
        <div style="margin-bottom:16px;padding:10px;border:1px solid #dde3ec;border-radius:5px;">
          <h4 style="margin:0 0 4px;font-size:13px;color:#333;">
            {body_name}
            <span style="font-weight:normal;font-size:11px;color:#888;">&mdash; {material}, {total} hole(s)</span>
          </h4>
          {images_html}
          <p style="font-size:10px;color:#888;margin:4px 0;">
            Reference sides: A/B (ends), C/D/E/F (faces). Each photo shows drilling into the labelled side.
          </p>
          <table class="bom" style="margin-top:8px;">
            <thead>
              <tr>
                <th>Grp</th>
                <th>Diameter</th>
                <th>Depth</th>
                <th>Through</th>
                <th>Tap / Clearance</th>
                <th>Side</th>
                <th>End A</th>
                <th>End B</th>
                <th>Position</th>
              </tr>
            </thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
        '''.format(
            body_name=body_name,
            material=material,
            total=total,
            images_html=images_html,
            rows=rows,
        )

    return '''
    <div style="margin-top:20px;padding-top:12px;border-top:2px solid #e65100;">
      <h3 style="color:#e65100;margin-bottom:8px;">
        &#128736; Drilling Requirements &mdash; Complete before assembly
      </h3>
      <p style="font-size:11px;color:#666;margin-bottom:12px;">
        The following holes must be drilled into channel stock before welding.
        Measure from the indicated end of each piece.
      </p>
      {body_sections}
    </div>
    '''.format(body_sections=body_sections)


# ---------------------------------------------------------------------------
# HTML document generation
# ---------------------------------------------------------------------------

def _image_to_base64(image_path: str) -> str:
    """Read a PNG file and return a data URI string."""
    if not image_path or not os.path.isfile(image_path):
        return ''
    try:
        with open(image_path, 'rb') as f:
            data = f.read()
        return 'data:image/png;base64,' + base64.b64encode(data).decode('ascii')
    except Exception:
        logger.exception('Failed to encode image %s', image_path)
        return ''


def _escape_html(text: str) -> str:
    """Minimal HTML escaping."""
    return (
        str(text)
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
        .replace('"', '&quot;')
    )


def _format_dims(dims, unit: str = 'mm') -> str:
    """Format a dimensions list for display."""
    if not dims:
        return '-'
    factor = 1.0 / 25.4 if unit == 'inches' else 1.0
    suffix = 'in' if unit == 'inches' else 'mm'
    return ' x '.join('{:.1f}'.format(d * factor) for d in dims) + ' ' + suffix


def _render_step_html(
    capture: dict,
    step_number: int,
    unit: str,
) -> str:
    """Render one step card using a two-column layout (image | info)."""
    body_name = _escape_html(capture.get('body_name', ''))
    material = _escape_html(capture.get('material_name', '') or '')
    dims = _format_dims(capture.get('dimensions_mm'), unit)
    desc = _escape_html(capture.get('description', ''))
    is_sub = capture.get('is_sub_assembly', False)

    if step_number == 1:
        title = 'Base piece &mdash; {}'.format(body_name)
    elif is_sub:
        title = 'Add sub-assembly &mdash; {}'.format(body_name)
    else:
        title = 'Weld {} to assembly'.format(body_name)

    b64 = _image_to_base64(capture.get('image_path', ''))
    if b64:
        image_html = (
            '<img src="{}" alt="Step {}" style="cursor:zoom-in;" '
            'onclick="window._openZoom(this.src)" />'
        ).format(b64, step_number)
    else:
        image_html = '<div class="no-image">No image captured</div>'

    material_row = (
        '<div class="step-meta-row"><strong>Material:</strong> {}</div>'.format(material)
        if material else ''
    )
    desc_html = (
        '<div class="step-notes"><strong>Notes:</strong> {}</div>'.format(desc)
        if desc else ''
    )

    return '''
    <div class="step-card">
      <div class="step-card-header">
        <span class="step-badge">{num}</span>
        <h3>{title}</h3>
      </div>
      <div class="step-card-image">{image_html}</div>
      <div class="step-card-info">
        <div class="step-meta-row"><strong>Part:</strong> {body_name}</div>
        {material_row}
        <div class="step-meta-row"><strong>Dims:</strong> {dims}</div>
        {desc_html}
        <p class="step-bodies">{bodies_shown} bodies visible</p>
      </div>
    </div>
    '''.format(
        num=step_number,
        title=title,
        body_name=body_name,
        material_row=material_row,
        dims=dims,
        desc_html=desc_html,
        image_html=image_html,
        bodies_shown=capture.get('bodies_shown', 0),
    )


def _render_bom_table(weld: dict, unit: str) -> str:
    """Render a bill-of-materials table for one weld assembly."""
    rows = ''
    for i, step in enumerate(weld.get('steps', [])):
        rows += '''
        <tr>
          <td>{num}</td>
          <td>{name}</td>
          <td>{material}</td>
          <td>{dims}</td>
          <td>{desc}</td>
        </tr>
        '''.format(
            num=i + 1,
            name=_escape_html(step.get('body_name', '')),
            material=_escape_html(step.get('material_name', '') or '-'),
            dims=_format_dims(step.get('dimensions_mm'), unit),
            desc=_escape_html(step.get('description', '') or '-'),
        )

    return '''
    <table class="bom">
      <thead>
        <tr>
          <th>#</th>
          <th>Part Name</th>
          <th>Material</th>
          <th>Dimensions</th>
          <th>Notes</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
    '''.format(rows=rows)


def _generate_html_document(
    weld_plans: list,
    all_captures: Dict[str, List[dict]],
    output_dir: str,
    settings: dict,
    hole_summaries: Optional[list] = None,
) -> str:
    """Generate the self-contained HTML weld plan document.

    Returns the path to the generated file.
    """
    project = _escape_html(settings.get('project_name', 'SmartCutList'))
    unit = settings.get('unit', 'mm')
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    assembly_sections = ''
    for weld in weld_plans:
        weld_id = weld['weld_id']
        captures = all_captures.get(weld_id, [])
        comp_name = _escape_html(weld.get('component_name', weld_id))
        step_count = len(weld.get('steps', []))

        # Assembly notes
        notes_html = ''
        notes = weld.get('notes', '').strip()
        if notes:
            notes_html = '<div class="assembly-notes"><strong>Assembly Notes:</strong> {}</div>'.format(
                _escape_html(notes)
            )

        # Nested indicator
        nested_html = ''
        if weld.get('parent_weld_id'):
            nested_html = '<p class="nested-indicator">Nested inside another assembly</p>'
        if weld.get('children_weld_ids'):
            nested_html += '<p class="nested-indicator">Contains {} sub-assembly(ies)</p>'.format(
                len(weld['children_weld_ids'])
            )

        # BOM table
        bom_html = _render_bom_table(weld, unit)

        # Step cards
        steps_html = ''
        for capture in captures:
            steps_html += _render_step_html(
                capture, capture['step_index'] + 1, unit
            )

        # Drilling requirements section (holes in channels)
        hole_html = _render_hole_section_html(
            weld, hole_summaries or [], unit
        )

        assembly_sections += '''
        <div class="assembly" id="assembly-{weld_id}">
          <h2>{comp_name}</h2>
          <p class="assembly-meta">{step_count} weld step(s)</p>
          {nested_html}
          {notes_html}
          {hole_html}
          <h3>Bill of Materials</h3>
          {bom_html}
          <h3>Weld Sequence</h3>
          <div class="steps-grid">
            {steps_html}
          </div>
        </div>
        <div class="page-break"></div>
        '''.format(
            weld_id=weld_id,
            comp_name=comp_name,
            step_count=step_count,
            nested_html=nested_html,
            notes_html=notes_html,
            bom_html=bom_html,
            steps_html=steps_html,
            hole_html=hole_html,
        )

    # Build table-of-contents links
    toc_links = ' &nbsp;&middot;&nbsp; '.join(
        '<a href="#assembly-{weld_id}">{name}</a>'.format(
            weld_id=w['weld_id'],
            name=_escape_html(w.get('component_name', w['weld_id'])),
        )
        for w in weld_plans
    )
    toc_html = (
        '<div class="toc"><strong>Assemblies:</strong> ' + toc_links + '</div>'
    ) if toc_links else ''

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Weld Plan - {project}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 13px;
    line-height: 1.55;
    color: #222;
    max-width: 860px;
    margin: 0 auto;
    padding: 28px 24px;
    background: #fff;
  }}
  h1 {{ font-size: 24px; margin-bottom: 4px; }}
  h2 {{
    font-size: 18px;
    margin-top: 36px;
    margin-bottom: 8px;
    padding-bottom: 6px;
    border-bottom: 2px solid #1565c0;
    color: #1a1a2e;
  }}
  h3 {{ font-size: 14px; margin: 16px 0 6px; color: #333; }}
  .header-meta {{
    color: #777;
    font-size: 12px;
    margin-bottom: 18px;
  }}
  .toc {{
    background: #f5f7fa;
    border: 1px solid #dde3ec;
    padding: 9px 14px;
    border-radius: 5px;
    font-size: 12px;
    margin-bottom: 28px;
    line-height: 1.8;
  }}
  .toc strong {{ color: #333; margin-right: 6px; }}
  .toc a {{ color: #1565c0; text-decoration: none; margin: 0 4px; }}
  .toc a:hover {{ text-decoration: underline; }}
  .assembly {{ margin-bottom: 36px; }}
  .assembly-meta {{ color: #777; font-size: 12px; margin-bottom: 10px; }}
  .assembly-notes {{
    background: #fffde7;
    border-left: 4px solid #f9a825;
    padding: 8px 12px;
    margin: 8px 0 14px;
    font-size: 12px;
    border-radius: 0 4px 4px 0;
  }}
  .nested-indicator {{
    color: #1565c0;
    font-size: 11px;
    margin-bottom: 4px;
  }}
  /* BOM table */
  .bom {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 18px;
    font-size: 12px;
  }}
  .bom th, .bom td {{
    border: 1px solid #d0d7e3;
    padding: 5px 9px;
    text-align: left;
  }}
  .bom th {{
    background: #eef2f8;
    font-weight: 600;
    color: #333;
  }}
  .bom tr:nth-child(even) {{ background: #f8fafd; }}
  /* Step grid — 2 cards per row */
  .steps-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
  }}
  /* Individual step card — stacked layout (image on top, info below) */
  .step-card {{
    border: 1px solid #dde3ec;
    border-radius: 6px;
    break-inside: avoid;
    overflow: hidden;
    display: flex;
    flex-direction: column;
  }}
  .step-card-header {{
    background: #eef2f8;
    padding: 5px 10px;
    border-bottom: 1px solid #dde3ec;
    display: flex;
    align-items: center;
    gap: 8px;
  }}
  .step-badge {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 20px;
    height: 20px;
    background: #1565c0;
    color: #fff;
    border-radius: 50%;
    font-size: 10px;
    font-weight: 700;
    flex-shrink: 0;
  }}
  .step-card-header h3 {{
    margin: 0;
    font-size: 12px;
    font-weight: 600;
    color: #1a3a6a;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .step-card-image {{
    background: #f8f8f8;
    border-bottom: 1px solid #eaecf0;
    line-height: 0;
  }}
  .step-card-image img {{
    width: 100%;
    height: auto;
    display: block;
  }}
  .no-image {{
    padding: 28px 12px;
    text-align: center;
    color: #bbb;
    font-size: 11px;
    background: #f5f5f5;
    line-height: 1.4;
  }}
  .step-card-info {{
    padding: 8px 10px 10px;
    font-size: 11px;
    flex: 1;
  }}
  .step-meta-row {{
    margin-bottom: 3px;
    color: #444;
  }}
  .step-meta-row strong {{ color: #111; }}
  .step-notes {{
    margin-top: 6px;
    background: #f0f4f8;
    border-left: 3px solid #4a90d9;
    padding: 4px 8px;
    font-size: 11px;
    border-radius: 0 3px 3px 0;
    color: #333;
  }}
  .step-bodies {{
    margin-top: 6px;
    font-size: 9px;
    color: #bbb;
  }}
  .page-break {{ page-break-after: always; }}
  @media print {{
    body {{ padding: 0; max-width: none; font-size: 10px; }}
    .toc {{ display: none; }}
    h2 {{ margin-top: 16px; }}
    .steps-grid {{ grid-template-columns: 1fr 1fr; gap: 10px; }}
    .step-card {{ break-inside: avoid; }}
    .step-card-image img {{ max-height: 180px; width: auto; margin: 0 auto; }}
    .page-break {{ page-break-after: always; }}
  }}
</style>
</head>
<body>
  <h1>Weld Assembly Plan</h1>
  <div class="header-meta">
    Project: {project} &nbsp;|&nbsp; Generated: {date} &nbsp;|&nbsp; Units: {unit}
  </div>
  {toc}
  {sections}
  <footer style="text-align:center;color:#aaa;font-size:11px;margin-top:32px;padding-top:12px;border-top:1px solid #eee;">
    Generated by SmartCutList for Fusion 360
  </footer>
  <!-- Zoom modal with pan & zoom for drill/weld images -->
  <div id="_zoom_modal" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;
       background:rgba(0,0,0,0.92);z-index:9999;overflow:hidden;">
    <div style="position:absolute;top:10px;left:50%;transform:translateX(-50%);color:#aaa;
         font-size:11px;z-index:10001;pointer-events:none;background:rgba(0,0,0,0.6);
         padding:4px 12px;border-radius:4px;">Scroll to zoom &bull; Drag to pan &bull; Esc to close</div>
    <div style="position:absolute;top:10px;right:16px;color:#ccc;font-size:28px;cursor:pointer;
         z-index:10001;line-height:1;" onclick="document.getElementById('_zoom_modal').style.display='none'">&times;</div>
    <img id="_zoom_img" src="" alt="Zoomed view"
         style="position:absolute;cursor:grab;border-radius:6px;
                box-shadow:0 4px 24px rgba(0,0,0,0.5);user-select:none;
                transform-origin:0 0;" draggable="false" />
  </div>
  <script>
  (function(){{
    var m=document.getElementById('_zoom_modal'),im=document.getElementById('_zoom_img');
    var sc=1,px=0,py=0,dr=false,dsx=0,dsy=0,spx=0,spy=0;
    function upd(){{im.style.transform='translate('+px+'px,'+py+'px) scale('+sc+')';}}
    function rst(){{var w=m.clientWidth,h=m.clientHeight,iw=im.naturalWidth,ih=im.naturalHeight;
      sc=Math.min(w*0.92/iw,h*0.92/ih,1);px=(w-iw*sc)/2;py=(h-ih*sc)/2;
      im.style.width=iw+'px';im.style.height=ih+'px';upd();}}
    window._openZoom=function(src){{im.src=src;m.style.display='block';
      if(im.naturalWidth>0)rst();else im.onload=function(){{rst();im.onload=null;}};
    }};
    m.addEventListener('mousedown',function(e){{if(e.target===m)m.style.display='none';
      else if(e.target===im){{e.preventDefault();dr=true;dsx=e.clientX;dsy=e.clientY;spx=px;spy=py;im.style.cursor='grabbing';}}}});
    document.addEventListener('mousemove',function(e){{if(!dr)return;px=spx+(e.clientX-dsx);py=spy+(e.clientY-dsy);upd();}});
    document.addEventListener('mouseup',function(){{if(dr){{dr=false;im.style.cursor='grab';}}}});
    m.addEventListener('wheel',function(e){{e.preventDefault();var r=m.getBoundingClientRect();
      var mx=e.clientX-r.left,my=e.clientY-r.top,os=sc;
      sc=Math.max(0.1,Math.min(sc*(e.deltaY<0?1.15:1/1.15),20));
      px=mx-(mx-px)*(sc/os);py=my-(my-py)*(sc/os);upd();}},{{passive:false}});
    document.addEventListener('keydown',function(e){{if(e.key==='Escape'&&m.style.display!=='none')m.style.display='none';}});
  }})();
  </script>
</body>
</html>'''.format(
        project=project,
        date=now,
        unit=unit,
        toc=toc_html,
        sections=assembly_sections,
    )

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = 'WeldPlan_{}.html'.format(timestamp)
    filepath = os.path.join(output_dir, filename)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html)

    logger.info('Weld plan HTML written to %s', filepath)
    return filepath


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_weld_plan(
    weld_plans: list,
    output_dir: str,
    settings: dict,
    groups: Optional[list] = None,
) -> str:
    """Generate a complete weld planning document with progressive screenshots.

    Parameters
    ----------
    weld_plans : list
        Enhanced weld assembly dicts with ordered steps.
    output_dir : str
        Directory to write the output HTML and intermediate images.
    settings : dict
        Keys: ``project_name``, ``unit``.
    groups : list, optional
        Grouped classified body list.  When provided, hole detection runs
        automatically and the drilling requirements section is included.

    Returns
    -------
    str
        Path to the generated HTML file.
    """
    # Create images subdirectory
    images_dir = os.path.join(output_dir, 'weld_plan_images')
    os.makedirs(images_dir, exist_ok=True)

    all_captures: Dict[str, List[dict]] = {}

    for weld in weld_plans:
        weld_id = weld['weld_id']
        logger.info(
            'Capturing weld steps for %s (%s, %d steps)',
            weld_id, weld.get('component_name', ''), len(weld.get('steps', [])),
        )
        captures = _capture_weld_steps(weld, weld_plans, images_dir)
        all_captures[weld_id] = captures

    # Hole detection and photography
    hole_summaries: list = []
    if groups:
        try:
            hole_summaries = hole_detection.detect_holes_on_channels(groups)
            if hole_summaries:
                logger.info(
                    'Detected holes on %d channel bodies; capturing perpendicular images',
                    len(hole_summaries),
                )
                capture_hole_images(hole_summaries, images_dir)
        except Exception:
            logger.exception('Hole detection/capture failed; continuing without hole data')

    # Generate the HTML document (images are embedded as Base64)
    html_path = _generate_html_document(
        weld_plans, all_captures, output_dir, settings,
        hole_summaries=hole_summaries,
    )

    # Clean up intermediate image files now that they're embedded
    for captures in all_captures.values():
        for cap in captures:
            img = cap.get('image_path', '')
            if img and os.path.isfile(img):
                try:
                    os.remove(img)
                except Exception:
                    pass
    # Remove images dir if empty
    try:
        os.rmdir(images_dir)
    except Exception:
        pass

    return html_path
