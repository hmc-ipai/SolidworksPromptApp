"""
SolidWorks Shape Creator - COM Automation Module
=================================================
Creates 3D and 2D shapes in SolidWorks via pywin32 COM automation.
Supports stacking shapes on top of each other in a single part.

3D Shapes: sphere, cylinder, cube, box, hexagon, triangle, pentagon,
           octagon, ellipse, slot, washer, L-shape, cross, star
2D Shapes: circle, square, rectangle

Requirements:
    - SolidWorks must be running
    - pip install pywin32

Usage:
    from SolidworksCreate import SolidWorksCreator

    sw = SolidWorksCreator()
    sw.connect()

    # Standalone shape
    sw.create_cube(size=0.02)

    # Stacking shapes
    sw.create_cube(size=0.02)        # base shape
    sw.begin_stack()
    sw.create_sphere(radius=0.01)    # stacked on top
    sw.reset()
"""
import win32com.client
import pythoncom
import os
import math
import time


# =============================================================================
# SolidWorks Connection & Base Class
# =============================================================================

class SolidWorksCreator:
    """Manages the SolidWorks COM connection and creates shapes."""

    def __init__(self):
        self.swApp = None
        self.model = None
        self.template_path = None
        self._nothing = None
        # Stacking state
        self._stacking = False
        self._stack_height = 0.0
        self._last_shape_height = 0.0
        self._plane_height = None  # None = no active plane; float = Z height in meters

    def connect(self):
        """Connect to the running SolidWorks instance."""
        try:
            self.swApp = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            self.swApp = win32com.client.Dispatch("SldWorks.Application")

        self.swApp.Visible = True
        self._nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        self._find_template()
        return True

    def _find_template(self):
        """Locate the SolidWorks part template."""
        for year in ['2025', '2024', '2023', '2022', '2021', '2020']:
            path = rf"C:\ProgramData\SolidWorks\SOLIDWORKS {year}\templates\Part.prtdot"
            if os.path.exists(path):
                self.template_path = path
                return
        raise FileNotFoundError("Could not find SolidWorks part template")

    # -------------------------------------------------------------------------
    # Stacking API
    # -------------------------------------------------------------------------

    def begin_stack(self):
        """Enter stacking mode - next shapes reuse the current model."""
        self._stacking = True

    def reset(self):
        """Exit stacking mode - next shape gets a new document."""
        self._stacking = False
        self._stack_height = 0.0
        self._last_shape_height = 0.0
        self._plane_height = None

    def create_plane(self, height=None):
        """Create a visible reference plane at the given height.

        Args:
            height: Z offset in meters. If None, uses current stack_height.

        Returns:
            str: Description of the created plane.
        """
        if self.model is None:
            raise RuntimeError("Create a shape first, then add a plane")

        if height is None:
            height = self._stack_height

        ext = self.model.Extension
        ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0,
                        self._nothing, 0)
        refPlane = self.model.FeatureManager.InsertRefPlane(
            4, height, 0, 0, 0, 0)
        if refPlane is None:
            raise RuntimeError(f"InsertRefPlane failed at height={height}")
        self.model.ClearSelection2(True)

        self._plane_height = height
        self._stacking = True
        self._zoom_to_fit()
        return f"Plane at {height*1000:.1f}mm"

    def set_height_to_plane(self):
        """Set working height to the active plane so the next shape starts there."""
        if self._plane_height is None:
            raise RuntimeError("No active plane - create a plane first")
        self._stack_height = self._plane_height
        self._stacking = True

    # -------------------------------------------------------------------------
    # Part & Sketch Management
    # -------------------------------------------------------------------------

    def _new_part(self):
        """Create a new part document."""
        self.model = self.swApp.NewDocument(self.template_path, 0, 0, 0)
        if not self.model:
            raise RuntimeError("Failed to create new part document")
        return self.model

    def _get_or_create_part(self):
        """Reuse existing model when stacking, otherwise create new."""
        if self._stacking and self.model is not None:
            return self.model
        self._stack_height = 0.0
        self._last_shape_height = 0.0
        return self._new_part()

    def _next_sketch_name(self, model):
        """Predict the name of the next sketch by counting existing ones.

        SolidWorks names sketches sequentially: Sketch1, Sketch2, etc.
        Late-bound COM can't use GetActiveSketch2(), so we count
        ProfileFeature entries in the feature tree instead.
        """
        count = 0
        feat = model.FirstFeature
        while feat:
            if feat.GetTypeName2 == "ProfileFeature":
                count += 1
            feat = feat.GetNextFeature
        return f"Sketch{count + 1}"

    def _start_sketch_at_height(self, model, z_height):
        """Open a sketch on a plane at the given Z offset from Front Plane.

        If z_height is 0, uses Front Plane directly. Otherwise creates an
        offset reference plane.
        """
        ext = model.Extension
        skMgr = model.SketchManager

        if z_height == 0.0:
            ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0,
                            self._nothing, 0)
        else:
            # Select Front Plane as the base for the offset
            ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0,
                            self._nothing, 0)
            # Create offset reference plane
            # 4 = swRefPlaneReferenceConstraint_Distance
            refPlane = model.FeatureManager.InsertRefPlane(4, z_height,
                                                           0, 0, 0, 0)
            if refPlane is None:
                raise RuntimeError(
                    f"InsertRefPlane failed at z_height={z_height}")
            # Select the new plane for sketching
            model.ClearSelection2(True)
            ext.SelectByID2(refPlane.Name, "PLANE", 0, 0, 0, False, 0,
                            self._nothing, 0)

        skMgr.InsertSketch(True)

    def _advance_stack(self, shape_height):
        """Update stacking state after creating a shape."""
        self._stack_height += shape_height
        self._last_shape_height = shape_height

    def _zoom_to_fit(self):
        """Rebuild and zoom to fit the current model."""
        self.model.ViewZoomtofit2()
        self.model.ForceRebuild3(True)

    # -------------------------------------------------------------------------
    # Sketch Helpers
    # -------------------------------------------------------------------------

    def _draw_polygon(self, skMgr, cx, cy, radius, sides):
        """Draw a regular polygon in the active sketch."""
        points = []
        for i in range(sides):
            angle = 2 * math.pi * i / sides - math.pi / 2
            x = cx + radius * math.cos(angle)
            y = cy + radius * math.sin(angle)
            points.append((x, y))
        for i in range(sides):
            x1, y1 = points[i]
            x2, y2 = points[(i + 1) % sides]
            skMgr.CreateLine(x1, y1, 0, x2, y2, 0)

    def _draw_star(self, skMgr, cx, cy, outer_r, inner_r, num_points):
        """Draw a star shape in the active sketch."""
        vertices = []
        for i in range(num_points * 2):
            angle = math.pi * i / num_points - math.pi / 2
            r = outer_r if i % 2 == 0 else inner_r
            x = cx + r * math.cos(angle)
            y = cy + r * math.sin(angle)
            vertices.append((x, y))
        for i in range(len(vertices)):
            x1, y1 = vertices[i]
            x2, y2 = vertices[(i + 1) % len(vertices)]
            skMgr.CreateLine(x1, y1, 0, x2, y2, 0)

    def _extrude(self, featMgr, height):
        """Standard blind extrusion of the current sketch."""
        featMgr.FeatureExtrusion2(
            True, False, False, 0, 0, height, 0.00254,
            False, False, False, False,
            1.74532925199433E-02, 1.74532925199433E-02,
            False, False, False, False,
            True, True, True, 0, 0, False
        )

    def _revolve(self, featMgr):
        """Solid revolve 360 degrees around the pre-selected axis (mark=4).

        IMPORTANT: FeatureRevolve2 parameter order is:
            Dir1Type, Dir2Type, Dir1Angle, Dir2Angle
        NOT Dir1Type, Dir1Angle, Dir2Type, Dir2Angle.
        """
        feat = featMgr.FeatureRevolve2(
            True,               # SingleDir
            True,               # IsSolid
            False,              # IsThin
            False,              # IsCut
            False,              # ReverseDir
            False,              # BothDirectionUpToSameEntity
            0,                  # Dir1Type  (swEndCondBlind)
            0,                  # Dir2Type  (swEndCondBlind)
            2 * math.pi,       # Dir1Angle (360 degrees)
            0.0,                # Dir2Angle
            False,              # OffsetReverse1
            False,              # OffsetReverse2
            0.0,                # OffsetDistance1
            0.0,                # OffsetDistance2
            0,                  # ThinType
            0.0,                # ThinThickness1
            0.0,                # ThinThickness2
            True,               # Merge
            True,               # UseFeatScope
            True                # UseAutoSelect
        )
        if feat is None:
            raise RuntimeError("FeatureRevolve2 returned None")
        return feat

    # =========================================================================
    # 3D Shapes
    # =========================================================================

    def create_sphere(self, radius=0.01):
        """Create a solid sphere by revolving a semicircle 360 degrees.

        Args:
            radius: Sphere radius in meters (default 0.01 = 10 mm).

        Returns:
            str: Description of the created shape.
        """
        model = self._get_or_create_part()
        ext = model.Extension
        skMgr = model.SketchManager
        featMgr = model.FeatureManager

        # Predict the sketch name before opening it (Sketch1, Sketch2, etc.)
        sketch_name = self._next_sketch_name(model)

        # Sphere center must be offset so the bottom sits at _stack_height
        sphere_center_z = self._stack_height + radius
        self._start_sketch_at_height(model, sphere_center_z)

        # Semicircle arc: start (top), end (bottom), midpoint (right)
        skMgr.Create3PointArc(0, radius, 0, 0, -radius, 0, radius, 0, 0)

        # Centerline along Y axis as revolve axis
        skMgr.CreateCenterLine(0, -radius, 0, 0, radius, 0)

        # Close sketch
        skMgr.InsertSketch(True)
        time.sleep(0.3)

        # Select the centerline as revolve axis (mark=4)
        model.ClearSelection2(True)
        ext.SelectByID2(f"Line1@{sketch_name}", "EXTSKETCHSEGMENT", 0, 0, 0,
                         False, 4, self._nothing, 0)

        # Revolve 360 degrees
        self._revolve(featMgr)
        self._advance_stack(2 * radius)
        self._zoom_to_fit()
        return f"Sphere (r={radius*1000:.1f}mm)"

    def create_cylinder(self, radius=0.01, height=0.02):
        """Create a cylinder (extruded circle).

        Args:
            radius: Cylinder radius in meters.
            height: Cylinder height in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        skMgr.CreateCircle(0, 0, 0, radius, 0, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"Cylinder (r={radius*1000:.1f}mm, h={height*1000:.1f}mm)"

    def create_cube(self, size=0.02):
        """Create a cube (equal-sided box).

        Args:
            size: Side length in meters.
        """
        return self.create_box(size, size, size, name="Cube")

    def create_box(self, width=0.02, height=0.02, depth=0.02, name="Box"):
        """Create a rectangular box (extruded rectangle).

        Args:
            width:  X dimension in meters.
            height: Extrusion height in meters.
            depth:  Y dimension in meters.
            name:   Display name for the result string.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        hx, hy = width / 2, depth / 2
        skMgr.CreateLine(-hx, -hy, 0,  hx, -hy, 0)
        skMgr.CreateLine( hx, -hy, 0,  hx,  hy, 0)
        skMgr.CreateLine( hx,  hy, 0, -hx,  hy, 0)
        skMgr.CreateLine(-hx,  hy, 0, -hx, -hy, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"{name} ({width*1000:.1f}x{height*1000:.1f}x{depth*1000:.1f}mm)"

    def create_polygon_3d(self, radius=0.01, height=0.01, sides=6, name=None):
        """Create an extruded regular polygon.

        Args:
            radius: Circumscribed radius in meters.
            height: Extrusion height in meters.
            sides:  Number of polygon sides.
            name:   Display name (defaults based on side count).
        """
        if name is None:
            names = {3: "Triangle", 5: "Pentagon", 6: "Hexagon",
                     8: "Octagon"}
            name = names.get(sides, f"{sides}-gon")
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        self._draw_polygon(skMgr, 0, 0, radius, sides)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"{name} (r={radius*1000:.1f}mm, h={height*1000:.1f}mm)"

    def create_triangle_3d(self, base=0.02, tri_height=0.02, depth=0.01):
        """Create a triangular prism.

        Args:
            base:       Triangle base width in meters.
            tri_height: Triangle height in meters.
            depth:      Extrusion depth in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        hb = base / 2
        skMgr.CreateLine(-hb, 0, 0,  hb, 0, 0)
        skMgr.CreateLine( hb, 0, 0,  0, tri_height, 0)
        skMgr.CreateLine( 0, tri_height, 0, -hb, 0, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, depth)
        self._advance_stack(depth)
        self._zoom_to_fit()
        return f"Triangle Prism (base={base*1000:.1f}mm)"

    def create_ellipse_3d(self, major=0.02, minor=0.01, height=0.01):
        """Create an extruded ellipse.

        Args:
            major:  Major axis diameter in meters.
            minor:  Minor axis diameter in meters.
            height: Extrusion height in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        skMgr.CreateEllipse(0, 0, 0, major / 2, 0, 0, 0, minor / 2, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"Ellipse ({major*1000:.1f}x{minor*1000:.1f}mm, h={height*1000:.1f}mm)"

    def create_slot_3d(self, length=0.03, width=0.01, height=0.005):
        """Create an extruded slot (stadium / oblong).

        Args:
            length: Overall slot length in meters.
            width:  Slot width in meters.
            height: Extrusion height in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        r = width / 2
        half_len = (length - width) / 2
        skMgr.CreateLine(-half_len,  r, 0,  half_len,  r, 0)
        skMgr.CreateArc(half_len, 0, 0, half_len, r, 0, half_len, -r, 0, -1)
        skMgr.CreateLine( half_len, -r, 0, -half_len, -r, 0)
        skMgr.CreateArc(-half_len, 0, 0, -half_len, -r, 0, -half_len, r, 0, -1)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"Slot ({length*1000:.1f}x{width*1000:.1f}mm)"

    def create_washer(self, outer=0.02, inner=0.01, height=0.005):
        """Create a washer / ring (extruded annulus).

        Args:
            outer:  Outer diameter in meters.
            inner:  Inner diameter (hole) in meters.
            height: Extrusion height in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        skMgr.CreateCircle(0, 0, 0, outer / 2, 0, 0)
        skMgr.CreateCircle(0, 0, 0, inner / 2, 0, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"Washer (OD={outer*1000:.1f}mm, ID={inner*1000:.1f}mm)"

    def create_lshape(self, width=0.02, length=0.02, thickness=0.003,
                      depth=0.01):
        """Create an L-shaped extrusion.

        Args:
            width:     Horizontal leg length in meters.
            length:    Vertical leg length in meters.
            thickness: Wall thickness in meters.
            depth:     Extrusion depth in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        skMgr.CreateLine(0, 0, 0, width, 0, 0)
        skMgr.CreateLine(width, 0, 0, width, thickness, 0)
        skMgr.CreateLine(width, thickness, 0, thickness, thickness, 0)
        skMgr.CreateLine(thickness, thickness, 0, thickness, length, 0)
        skMgr.CreateLine(thickness, length, 0, 0, length, 0)
        skMgr.CreateLine(0, length, 0, 0, 0, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, depth)
        self._advance_stack(depth)
        self._zoom_to_fit()
        return f"L-Shape ({width*1000:.1f}x{length*1000:.1f}mm)"

    def create_cross(self, size=0.02, thickness=0.005, depth=0.005):
        """Create a cross / plus-shaped extrusion.

        Args:
            size:      Overall width and height in meters.
            thickness: Arm thickness in meters.
            depth:     Extrusion depth in meters.
        """
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        hs, ht = size / 2, thickness / 2
        skMgr.CreateLine(-ht,  hs, 0,  ht,  hs, 0)
        skMgr.CreateLine( ht,  hs, 0,  ht,  ht, 0)
        skMgr.CreateLine( ht,  ht, 0,  hs,  ht, 0)
        skMgr.CreateLine( hs,  ht, 0,  hs, -ht, 0)
        skMgr.CreateLine( hs, -ht, 0,  ht, -ht, 0)
        skMgr.CreateLine( ht, -ht, 0,  ht, -hs, 0)
        skMgr.CreateLine( ht, -hs, 0, -ht, -hs, 0)
        skMgr.CreateLine(-ht, -hs, 0, -ht, -ht, 0)
        skMgr.CreateLine(-ht, -ht, 0, -hs, -ht, 0)
        skMgr.CreateLine(-hs, -ht, 0, -hs,  ht, 0)
        skMgr.CreateLine(-hs,  ht, 0, -ht,  ht, 0)
        skMgr.CreateLine(-ht,  ht, 0, -ht,  hs, 0)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, depth)
        self._advance_stack(depth)
        self._zoom_to_fit()
        return f"Cross ({size*1000:.1f}mm)"

    def create_star_3d(self, outer=0.02, inner=None, points=5, height=0.005):
        """Create an extruded star.

        Args:
            outer:  Outer radius in meters.
            inner:  Inner radius in meters (default: 40% of outer).
            points: Number of star points.
            height: Extrusion height in meters.
        """
        if inner is None:
            inner = outer * 0.4
        model = self._get_or_create_part()
        skMgr = model.SketchManager
        self._start_sketch_at_height(model, self._stack_height)
        self._draw_star(skMgr, 0, 0, outer, inner, points)
        skMgr.InsertSketch(True)
        self._extrude(model.FeatureManager, height)
        self._advance_stack(height)
        self._zoom_to_fit()
        return f"{points}-Point Star (r={outer*1000:.1f}mm)"

    # =========================================================================
    # 2D Shapes (sketch only, no extrusion)
    # =========================================================================

    def create_circle_2d(self, radius=0.01):
        """Create a 2D circle sketch."""
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        skMgr.CreateCircle(0, 0, 0, radius, 0, 0)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"Circle 2D (r={radius*1000:.1f}mm)"

    def create_square_2d(self, size=0.02):
        """Create a 2D square sketch."""
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        hs = size / 2
        skMgr.CreateLine(-hs, -hs, 0,  hs, -hs, 0)
        skMgr.CreateLine( hs, -hs, 0,  hs,  hs, 0)
        skMgr.CreateLine( hs,  hs, 0, -hs,  hs, 0)
        skMgr.CreateLine(-hs,  hs, 0, -hs, -hs, 0)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"Square 2D ({size*1000:.1f}mm)"

    def create_rectangle_2d(self, width=0.02, length=0.01):
        """Create a 2D rectangle sketch."""
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        hw, hl = width / 2, length / 2
        skMgr.CreateLine(-hw, -hl, 0,  hw, -hl, 0)
        skMgr.CreateLine( hw, -hl, 0,  hw,  hl, 0)
        skMgr.CreateLine( hw,  hl, 0, -hw,  hl, 0)
        skMgr.CreateLine(-hw,  hl, 0, -hw, -hl, 0)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"Rectangle 2D ({width*1000:.1f}x{length*1000:.1f}mm)"

    def create_polygon_2d(self, radius=0.01, sides=6, name=None):
        """Create a 2D regular polygon sketch."""
        if name is None:
            names = {3: "Triangle", 5: "Pentagon", 6: "Hexagon",
                     8: "Octagon"}
            name = names.get(sides, f"{sides}-gon")
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        self._draw_polygon(skMgr, 0, 0, radius, sides)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"{name} 2D (r={radius*1000:.1f}mm)"

    def create_ellipse_2d(self, major=0.02, minor=0.01):
        """Create a 2D ellipse sketch."""
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        skMgr.CreateEllipse(0, 0, 0, major / 2, 0, 0, 0, minor / 2, 0)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"Ellipse 2D ({major*1000:.1f}x{minor*1000:.1f}mm)"

    def create_star_2d(self, outer=0.02, inner=None, points=5):
        """Create a 2D star sketch."""
        if inner is None:
            inner = outer * 0.4
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        self._draw_star(skMgr, 0, 0, outer, inner, points)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"{points}-Point Star 2D (r={outer*1000:.1f}mm)"

    def create_triangle_2d(self, base=0.02, tri_height=0.02):
        """Create a 2D triangle sketch."""
        model = self._new_part()
        skMgr = model.SketchManager
        skMgr.InsertSketch(True)
        hb = base / 2
        skMgr.CreateLine(-hb, 0, 0,  hb, 0, 0)
        skMgr.CreateLine( hb, 0, 0,  0, tri_height, 0)
        skMgr.CreateLine( 0, tri_height, 0, -hb, 0, 0)
        skMgr.InsertSketch(True)
        self._zoom_to_fit()
        return f"Triangle 2D (base={base*1000:.1f}mm)"


# =============================================================================
# Standalone demo
# =============================================================================

if __name__ == "__main__":
    sw = SolidWorksCreator()
    sw.connect()

    print("Creating a cube with a sphere on top...")
    sw.create_cube(size=0.02)
    sw.begin_stack()
    result = sw.create_sphere(radius=0.01)
    print(f"  {result}")
    sw.reset()
