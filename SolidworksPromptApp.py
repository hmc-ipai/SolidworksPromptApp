"""
SolidWorks AI Prompt Tool - GUI Version
=======================================
A floating button + text box that stays on top of SolidWorks.
Click the button, type a description, shape appears!

Supported 3D Shapes: cube, box, cylinder, hexagon, triangle prism, pentagon, octagon, ellipse/oval
Supported 2D Shapes: circle, square, rectangle, triangle, hexagon, pentagon, ellipse, line, arc

Requirements:
    pip install pywin32

Usage:
    python solidworks_ai_gui.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import os
import re
import math


# =============================================================================
# Unit Conversion
# =============================================================================

UNIT_TO_METERS = {
    "mm": 0.001, "millimeter": 0.001, "millimeters": 0.001,
    "cm": 0.01, "centimeter": 0.01, "centimeters": 0.01,
    "m": 1.0, "meter": 1.0, "meters": 1.0,
    "in": 0.0254, "inch": 0.0254, "inches": 0.0254, '"': 0.0254,
    "ft": 0.3048, "foot": 0.3048, "feet": 0.3048,
}

def convert_to_meters(value, unit="mm"):
    unit = unit.lower().strip().rstrip('.')
    return value * UNIT_TO_METERS.get(unit, 0.001)


# =============================================================================
# Prompt Parser
# =============================================================================

class ParsedShape:
    def __init__(self, shape_type, params, units="mm", is_2d=False):
        self.shape_type = shape_type
        self.params = params
        self.units = units
        self.is_2d = is_2d


def detect_units(prompt):
    prompt_lower = prompt.lower()
    if any(u in prompt_lower for u in ['inch', 'inches', '"', ' in ']):
        return "in"
    elif 'cm' in prompt_lower or 'centimeter' in prompt_lower:
        return "cm"
    elif 'ft' in prompt_lower or 'feet' in prompt_lower:
        return "ft"
    return "mm"


def extract_dimension(prompt, *keywords):
    prompt_lower = prompt.lower()
    unit_pattern = r'(mm|cm|m|in|inch|inches|ft|feet|foot|")?'
    
    for keyword in keywords:
        patterns = [
            rf'{keyword}\s*(?:of|=|:)?\s*(\d+\.?\d*)\s*{unit_pattern}',
            rf'(\d+\.?\d*)\s*{unit_pattern}\s*{keyword}',
        ]
        for pattern in patterns:
            match = re.search(pattern, prompt_lower, re.IGNORECASE)
            if match:
                value = float(match.group(1))
                unit = match.group(2) if match.group(2) else detect_units(prompt)
                return convert_to_meters(value, unit)
    return None


def extract_all_numbers(prompt):
    results = []
    default_unit = detect_units(prompt)
    pattern = r'(\d+\.?\d*)\s*(mm|cm|m|in|inch|inches|ft|feet|foot|")?'
    matches = re.findall(pattern, prompt.lower())
    for value_str, unit in matches:
        value = float(value_str)
        unit = unit if unit else default_unit
        results.append(convert_to_meters(value, unit))
    return results


def parse_prompt(prompt):
    prompt_lower = prompt.lower().strip()
    units = detect_units(prompt)
    
    # Check if 2D is requested
    is_2d = any(word in prompt_lower for word in ['2d', 'sketch', 'draw', 'flat'])
    
    # --- CYLINDER (3D) ---
    if any(word in prompt_lower for word in ['cylinder', 'cylindrical', 'tube', 'pipe']):
        radius = extract_dimension(prompt, 'radius', 'r')
        diameter = extract_dimension(prompt, 'diameter', 'dia', 'd')
        height = extract_dimension(prompt, 'height', 'tall', 'long', 'h')
        
        if diameter and not radius:
            radius = diameter / 2
        
        if not radius or not height:
            numbers = extract_all_numbers(prompt)
            if len(numbers) >= 2:
                radius = radius or numbers[0]
                height = height or numbers[1]
        
        radius = radius or 0.01
        height = height or 0.02
        return ParsedShape('cylinder', {'radius': radius, 'height': height}, units)
    
    # --- CUBE (3D) ---
    elif 'cube' in prompt_lower:
        size = extract_dimension(prompt, 'side', 'size', 'length')
        if not size:
            numbers = extract_all_numbers(prompt)
            size = numbers[0] if numbers else 0.02
        return ParsedShape('cube', {'size': size}, units)
    
    # --- BOX / RECTANGULAR PRISM (3D) ---
    elif any(word in prompt_lower for word in ['box', 'rectangular', 'prism', 'block']):
        width = extract_dimension(prompt, 'width', 'wide', 'w')
        height = extract_dimension(prompt, 'height', 'tall', 'h')
        depth = extract_dimension(prompt, 'depth', 'deep', 'long', 'length', 'd', 'l')
        
        numbers = extract_all_numbers(prompt)
        axb_match = re.search(r'(\d+\.?\d*)\s*x\s*(\d+\.?\d*)\s*x\s*(\d+\.?\d*)', prompt_lower)
        
        if axb_match:
            unit = detect_units(prompt)
            width = convert_to_meters(float(axb_match.group(1)), unit)
            height = convert_to_meters(float(axb_match.group(2)), unit)
            depth = convert_to_meters(float(axb_match.group(3)), unit)
        elif len(numbers) >= 3:
            width = width or numbers[0]
            height = height or numbers[1]
            depth = depth or numbers[2]
        elif len(numbers) == 1:
            width = height = depth = numbers[0]
        
        width = width or 0.02
        height = height or 0.02
        depth = depth or 0.02
        return ParsedShape('box', {'width': width, 'height': height, 'depth': depth}, units)
    
    # --- HEXAGON ---
    elif 'hexagon' in prompt_lower or 'hex' in prompt_lower:
        radius = extract_dimension(prompt, 'radius', 'r', 'size')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not radius and numbers:
            radius = numbers[0]
        if not height and len(numbers) >= 2:
            height = numbers[1]
        
        radius = radius or 0.01
        height = height or 0.01
        return ParsedShape('hexagon', {'radius': radius, 'height': height}, units, is_2d)
    
    # --- TRIANGLE ---
    elif 'triangle' in prompt_lower:
        base = extract_dimension(prompt, 'base', 'width', 'b', 'w')
        tri_height = extract_dimension(prompt, 'height', 'tall', 'h')
        depth = extract_dimension(prompt, 'depth', 'thick', 'extrude', 'd')
        numbers = extract_all_numbers(prompt)
        
        if not base and numbers:
            base = numbers[0]
        if not tri_height and len(numbers) >= 2:
            tri_height = numbers[1]
        if not depth and len(numbers) >= 3:
            depth = numbers[2]
        
        base = base or 0.02
        tri_height = tri_height or 0.02
        depth = depth or 0.01
        return ParsedShape('triangle', {'base': base, 'tri_height': tri_height, 'depth': depth}, units, is_2d)
    
    # --- PENTAGON ---
    elif 'pentagon' in prompt_lower:
        radius = extract_dimension(prompt, 'radius', 'r', 'size')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not radius and numbers:
            radius = numbers[0]
        if not height and len(numbers) >= 2:
            height = numbers[1]
        
        radius = radius or 0.01
        height = height or 0.01
        return ParsedShape('pentagon', {'radius': radius, 'height': height}, units, is_2d)
    
    # --- OCTAGON ---
    elif 'octagon' in prompt_lower:
        radius = extract_dimension(prompt, 'radius', 'r', 'size')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not radius and numbers:
            radius = numbers[0]
        if not height and len(numbers) >= 2:
            height = numbers[1]
        
        radius = radius or 0.01
        height = height or 0.01
        return ParsedShape('octagon', {'radius': radius, 'height': height}, units, is_2d)
    
    # --- ELLIPSE / OVAL ---
    elif any(word in prompt_lower for word in ['ellipse', 'oval']):
        major = extract_dimension(prompt, 'major', 'length', 'long', 'a')
        minor = extract_dimension(prompt, 'minor', 'width', 'short', 'b')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not major and numbers:
            major = numbers[0]
        if not minor and len(numbers) >= 2:
            minor = numbers[1]
        if not height and len(numbers) >= 3:
            height = numbers[2]
        
        major = major or 0.02
        minor = minor or 0.01
        height = height or 0.01
        return ParsedShape('ellipse', {'major': major, 'minor': minor, 'height': height}, units, is_2d)
    
    # --- CIRCLE (2D default, or 3D if extruded) ---
    elif 'circle' in prompt_lower:
        radius = extract_dimension(prompt, 'radius', 'r')
        diameter = extract_dimension(prompt, 'diameter', 'dia', 'd')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick', 'extrude')
        
        if diameter and not radius:
            radius = diameter / 2
        if not radius:
            numbers = extract_all_numbers(prompt)
            radius = numbers[0] if numbers else 0.01
        
        radius = radius or 0.01
        if height:
            return ParsedShape('cylinder', {'radius': radius, 'height': height}, units)
        return ParsedShape('circle', {'radius': radius}, units, is_2d=True)
    
    # --- SQUARE (2D) ---
    elif 'square' in prompt_lower:
        size = extract_dimension(prompt, 'side', 'size', 'length', 's')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick', 'extrude')
        numbers = extract_all_numbers(prompt)
        
        if not size and numbers:
            size = numbers[0]
        
        size = size or 0.02
        if height:
            return ParsedShape('cube', {'size': size}, units)
        return ParsedShape('square', {'size': size}, units, is_2d=True)
    
    # --- RECTANGLE (2D) ---
    elif 'rectangle' in prompt_lower or 'rect' in prompt_lower:
        width = extract_dimension(prompt, 'width', 'wide', 'w')
        length = extract_dimension(prompt, 'length', 'long', 'l', 'height', 'h')
        depth = extract_dimension(prompt, 'depth', 'thick', 'extrude', 'd')
        numbers = extract_all_numbers(prompt)
        
        axb_match = re.search(r'(\d+\.?\d*)\s*x\s*(\d+\.?\d*)', prompt_lower)
        if axb_match:
            unit = detect_units(prompt)
            width = convert_to_meters(float(axb_match.group(1)), unit)
            length = convert_to_meters(float(axb_match.group(2)), unit)
        elif len(numbers) >= 2:
            width = width or numbers[0]
            length = length or numbers[1]
        
        width = width or 0.02
        length = length or 0.01
        if depth:
            return ParsedShape('box', {'width': width, 'height': depth, 'depth': length}, units)
        return ParsedShape('rectangle', {'width': width, 'length': length}, units, is_2d=True)
    
    # --- SLOT ---
    elif 'slot' in prompt_lower:
        length = extract_dimension(prompt, 'length', 'long', 'l')
        width = extract_dimension(prompt, 'width', 'wide', 'w')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not length and numbers:
            length = numbers[0]
        if not width and len(numbers) >= 2:
            width = numbers[1]
        if not height and len(numbers) >= 3:
            height = numbers[2]
        
        length = length or 0.03
        width = width or 0.01
        height = height or 0.005
        return ParsedShape('slot', {'length': length, 'width': width, 'height': height}, units, is_2d)
    
    # --- WASHER / RING ---
    elif any(word in prompt_lower for word in ['washer', 'ring', 'donut', 'annulus']):
        outer = extract_dimension(prompt, 'outer', 'outside', 'od', 'diameter')
        inner = extract_dimension(prompt, 'inner', 'inside', 'id', 'hole')
        height = extract_dimension(prompt, 'height', 'tall', 'h', 'thick')
        numbers = extract_all_numbers(prompt)
        
        if not outer and numbers:
            outer = numbers[0]
        if not inner and len(numbers) >= 2:
            inner = numbers[1]
        if not height and len(numbers) >= 3:
            height = numbers[2]
        
        outer = outer or 0.02
        inner = inner or 0.01
        height = height or 0.005
        return ParsedShape('washer', {'outer': outer, 'inner': inner, 'height': height}, units, is_2d)
    
    # --- L-SHAPE ---
    elif any(word in prompt_lower for word in ['l-shape', 'lshape', 'l shape']):
        width = extract_dimension(prompt, 'width', 'w')
        length = extract_dimension(prompt, 'length', 'l', 'height', 'h')
        thickness = extract_dimension(prompt, 'thick', 't')
        depth = extract_dimension(prompt, 'depth', 'd', 'extrude')
        numbers = extract_all_numbers(prompt)
        
        if len(numbers) >= 1:
            width = width or numbers[0]
        if len(numbers) >= 2:
            length = length or numbers[1]
        if len(numbers) >= 3:
            thickness = thickness or numbers[2]
        if len(numbers) >= 4:
            depth = depth or numbers[3]
        
        width = width or 0.02
        length = length or 0.02
        thickness = thickness or 0.003
        depth = depth or 0.01
        return ParsedShape('lshape', {'width': width, 'length': length, 'thickness': thickness, 'depth': depth}, units, is_2d)
    
    # --- CROSS / PLUS ---
    elif any(word in prompt_lower for word in ['cross', 'plus']):
        size = extract_dimension(prompt, 'size', 's', 'width', 'w')
        thickness = extract_dimension(prompt, 'thick', 't', 'arm')
        depth = extract_dimension(prompt, 'depth', 'd', 'height', 'h')
        numbers = extract_all_numbers(prompt)
        
        if len(numbers) >= 1:
            size = size or numbers[0]
        if len(numbers) >= 2:
            thickness = thickness or numbers[1]
        if len(numbers) >= 3:
            depth = depth or numbers[2]
        
        size = size or 0.02
        thickness = thickness or 0.005
        depth = depth or 0.005
        return ParsedShape('cross', {'size': size, 'thickness': thickness, 'depth': depth}, units, is_2d)
    
    # --- STAR ---
    elif 'star' in prompt_lower:
        outer = extract_dimension(prompt, 'outer', 'radius', 'r', 'size')
        inner = extract_dimension(prompt, 'inner')
        points = 5
        height = extract_dimension(prompt, 'height', 'h', 'thick', 'depth')
        numbers = extract_all_numbers(prompt)
        
        points_match = re.search(r'(\d+)\s*point', prompt_lower)
        if points_match:
            points = int(points_match.group(1))
        
        if not outer and numbers:
            outer = numbers[0]
        if not height and len(numbers) >= 2:
            height = numbers[1]
        
        outer = outer or 0.02
        inner = inner or outer * 0.4
        height = height or 0.005
        return ParsedShape('star', {'outer': outer, 'inner': inner, 'points': points, 'height': height}, units, is_2d)
    
    return None


# =============================================================================
# SolidWorks Connection
# =============================================================================

class SolidWorksApp:
    def __init__(self):
        self.swApp = None
        self.model = None
        self.template_path = None
    
    def connect(self):
        try:
            self.swApp = win32com.client.GetActiveObject("SldWorks.Application")
        except:
            self.swApp = win32com.client.Dispatch("SldWorks.Application")
        
        self.swApp.Visible = True
        self._find_template()
        return True
    
    def _find_template(self):
        try:
            path = self.swApp.GetUserPreferenceStringValue(8)
            if path and os.path.exists(path):
                self.template_path = path
                return
        except:
            pass
        
        common_paths = [
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2024\templates\Part.prtdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Part.prtdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2022\templates\Part.prtdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2021\templates\Part.prtdot",
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                self.template_path = path
                return
        
        self.template_path = filedialog.askopenfilename(
            title="Select SolidWorks Part Template",
            filetypes=[("Part Template", "*.prtdot")],
        )
    
    def new_part(self):
        self.model = self.swApp.NewDocument(self.template_path, 0, 0, 0)
        if not self.model:
            raise Exception("Failed to create new part")
        return self.model
    
    def zoom_to_fit(self):
        self.model.ViewZoomtofit2()
        self.model.ForceRebuild3(True)


# =============================================================================
# Helper Functions
# =============================================================================

def draw_polygon(skMgr, cx, cy, radius, sides):
    """Draw a regular polygon."""
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


def draw_star(skMgr, cx, cy, outer_r, inner_r, points):
    """Draw a star shape."""
    vertices = []
    for i in range(points * 2):
        angle = math.pi * i / points - math.pi / 2
        r = outer_r if i % 2 == 0 else inner_r
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        vertices.append((x, y))
    
    for i in range(len(vertices)):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % len(vertices)]
        skMgr.CreateLine(x1, y1, 0, x2, y2, 0)


def extrude(featMgr, height):
    """Standard extrusion."""
    featMgr.FeatureExtrusion2(
        True, False, False, 0, 0, height, 0.00254,
        False, False, False, False,
        1.74532925199433E-02, 1.74532925199433E-02,
        False, False, False, False,
        True, True, True, 0, 0, False
    )


# =============================================================================
# Shape Creators
# =============================================================================

def create_cylinder(sw, radius, height):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    skMgr.CreateCircle(0, 0, 0, radius, 0, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"Cylinder (r={radius*1000:.1f}mm, h={height*1000:.1f}mm)"


def create_cube(sw, size):
    return create_box(sw, size, size, size, "Cube")


def create_box(sw, width, height, depth, name="Box"):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    hx, hy = width / 2, depth / 2
    skMgr.CreateLine(-hx, -hy, 0, hx, -hy, 0)
    skMgr.CreateLine(hx, -hy, 0, hx, hy, 0)
    skMgr.CreateLine(hx, hy, 0, -hx, hy, 0)
    skMgr.CreateLine(-hx, hy, 0, -hx, -hy, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"{name} ({width*1000:.1f}x{height*1000:.1f}x{depth*1000:.1f}mm)"


def create_polygon_3d(sw, radius, height, sides, name):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    draw_polygon(skMgr, 0, 0, radius, sides)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"{name} (r={radius*1000:.1f}mm, h={height*1000:.1f}mm)"


def create_triangle_3d(sw, base, tri_height, depth):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    hb = base / 2
    skMgr.CreateLine(-hb, 0, 0, hb, 0, 0)
    skMgr.CreateLine(hb, 0, 0, 0, tri_height, 0)
    skMgr.CreateLine(0, tri_height, 0, -hb, 0, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, depth)
    sw.zoom_to_fit()
    return f"Triangle Prism (base={base*1000:.1f}mm)"


def create_ellipse_3d(sw, major, minor, height):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    skMgr.CreateEllipse(0, 0, 0, major/2, 0, 0, 0, minor/2, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"Ellipse ({major*1000:.1f}x{minor*1000:.1f}mm, h={height*1000:.1f}mm)"


def create_slot_3d(sw, length, width, height):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    r = width / 2
    half_len = (length - width) / 2
    skMgr.CreateLine(-half_len, r, 0, half_len, r, 0)
    skMgr.CreateArc(half_len, 0, 0, half_len, r, 0, half_len, -r, 0, -1)
    skMgr.CreateLine(half_len, -r, 0, -half_len, -r, 0)
    skMgr.CreateArc(-half_len, 0, 0, -half_len, -r, 0, -half_len, r, 0, -1)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"Slot ({length*1000:.1f}x{width*1000:.1f}mm)"


def create_washer(sw, outer, inner, height):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    skMgr.CreateCircle(0, 0, 0, outer/2, 0, 0)
    skMgr.CreateCircle(0, 0, 0, inner/2, 0, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"Washer (OD={outer*1000:.1f}mm, ID={inner*1000:.1f}mm)"


def create_lshape(sw, width, length, thickness, depth):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    skMgr.CreateLine(0, 0, 0, width, 0, 0)
    skMgr.CreateLine(width, 0, 0, width, thickness, 0)
    skMgr.CreateLine(width, thickness, 0, thickness, thickness, 0)
    skMgr.CreateLine(thickness, thickness, 0, thickness, length, 0)
    skMgr.CreateLine(thickness, length, 0, 0, length, 0)
    skMgr.CreateLine(0, length, 0, 0, 0, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, depth)
    sw.zoom_to_fit()
    return f"L-Shape ({width*1000:.1f}x{length*1000:.1f}mm)"


def create_cross(sw, size, thickness, depth):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    hs, ht = size / 2, thickness / 2
    skMgr.CreateLine(-ht, hs, 0, ht, hs, 0)
    skMgr.CreateLine(ht, hs, 0, ht, ht, 0)
    skMgr.CreateLine(ht, ht, 0, hs, ht, 0)
    skMgr.CreateLine(hs, ht, 0, hs, -ht, 0)
    skMgr.CreateLine(hs, -ht, 0, ht, -ht, 0)
    skMgr.CreateLine(ht, -ht, 0, ht, -hs, 0)
    skMgr.CreateLine(ht, -hs, 0, -ht, -hs, 0)
    skMgr.CreateLine(-ht, -hs, 0, -ht, -ht, 0)
    skMgr.CreateLine(-ht, -ht, 0, -hs, -ht, 0)
    skMgr.CreateLine(-hs, -ht, 0, -hs, ht, 0)
    skMgr.CreateLine(-hs, ht, 0, -ht, ht, 0)
    skMgr.CreateLine(-ht, ht, 0, -ht, hs, 0)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, depth)
    sw.zoom_to_fit()
    return f"Cross ({size*1000:.1f}mm)"


def create_star_3d(sw, outer, inner, points, height):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    draw_star(skMgr, 0, 0, outer, inner, points)
    skMgr.InsertSketch(True)
    extrude(model.FeatureManager, height)
    sw.zoom_to_fit()
    return f"{points}-Point Star (r={outer*1000:.1f}mm)"


# 2D Shapes
def create_circle_2d(sw, radius):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    skMgr.CreateCircle(0, 0, 0, radius, 0, 0)
    skMgr.InsertSketch(True)
    sw.zoom_to_fit()
    return f"Circle 2D (r={radius*1000:.1f}mm)"


def create_square_2d(sw, size):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    hs = size / 2
    skMgr.CreateLine(-hs, -hs, 0, hs, -hs, 0)
    skMgr.CreateLine(hs, -hs, 0, hs, hs, 0)
    skMgr.CreateLine(hs, hs, 0, -hs, hs, 0)
    skMgr.CreateLine(-hs, hs, 0, -hs, -hs, 0)
    skMgr.InsertSketch(True)
    sw.zoom_to_fit()
    return f"Square 2D ({size*1000:.1f}mm)"


def create_rectangle_2d(sw, width, length):
    model = sw.new_part()
    skMgr = model.SketchManager
    skMgr.InsertSketch(True)
    hw, hl = width / 2, length / 2
    skMgr.CreateLine(-hw, -hl, 0, hw, -hl, 0)
    skMgr.CreateLine(hw, -hl, 0, hw, hl, 0)
    skMgr.CreateLine(hw, hl, 0, -hw, hl, 0)
    skMgr.CreateLine(-hw, hl, 0, -hw, -hl, 0)
    skMgr.InsertSketch(True)
    sw.zoom_to_fit()
    return f"Rectangle 2D ({width*1000:.1f}x{length*1000:.1f}mm)"


# =============================================================================
# Process Prompt
# =============================================================================

def process_prompt(sw, prompt):
    parsed = parse_prompt(prompt)
    
    if not parsed:
        return False, "Unknown shape. Try: cube, box, cylinder, hexagon, triangle, pentagon, octagon, ellipse, star, cross, slot, washer, L-shape, circle, square, rectangle"
    
    try:
        p = parsed.params
        t = parsed.shape_type
        
        if t == 'cylinder':
            name = create_cylinder(sw, p['radius'], p['height'])
        elif t == 'cube':
            name = create_cube(sw, p['size'])
        elif t == 'box':
            name = create_box(sw, p['width'], p['height'], p['depth'])
        elif t == 'hexagon':
            if parsed.is_2d:
                model = sw.new_part()
                model.SketchManager.InsertSketch(True)
                draw_polygon(model.SketchManager, 0, 0, p['radius'], 6)
                model.SketchManager.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"Hexagon 2D (r={p['radius']*1000:.1f}mm)"
            else:
                name = create_polygon_3d(sw, p['radius'], p['height'], 6, "Hexagon")
        elif t == 'triangle':
            if parsed.is_2d:
                model = sw.new_part()
                skMgr = model.SketchManager
                skMgr.InsertSketch(True)
                hb = p['base'] / 2
                skMgr.CreateLine(-hb, 0, 0, hb, 0, 0)
                skMgr.CreateLine(hb, 0, 0, 0, p['tri_height'], 0)
                skMgr.CreateLine(0, p['tri_height'], 0, -hb, 0, 0)
                skMgr.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"Triangle 2D (base={p['base']*1000:.1f}mm)"
            else:
                name = create_triangle_3d(sw, p['base'], p['tri_height'], p['depth'])
        elif t == 'pentagon':
            if parsed.is_2d:
                model = sw.new_part()
                model.SketchManager.InsertSketch(True)
                draw_polygon(model.SketchManager, 0, 0, p['radius'], 5)
                model.SketchManager.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"Pentagon 2D (r={p['radius']*1000:.1f}mm)"
            else:
                name = create_polygon_3d(sw, p['radius'], p['height'], 5, "Pentagon")
        elif t == 'octagon':
            if parsed.is_2d:
                model = sw.new_part()
                model.SketchManager.InsertSketch(True)
                draw_polygon(model.SketchManager, 0, 0, p['radius'], 8)
                model.SketchManager.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"Octagon 2D (r={p['radius']*1000:.1f}mm)"
            else:
                name = create_polygon_3d(sw, p['radius'], p['height'], 8, "Octagon")
        elif t == 'ellipse':
            if parsed.is_2d:
                model = sw.new_part()
                model.SketchManager.InsertSketch(True)
                model.SketchManager.CreateEllipse(0, 0, 0, p['major']/2, 0, 0, 0, p['minor']/2, 0)
                model.SketchManager.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"Ellipse 2D ({p['major']*1000:.1f}x{p['minor']*1000:.1f}mm)"
            else:
                name = create_ellipse_3d(sw, p['major'], p['minor'], p['height'])
        elif t == 'slot':
            name = create_slot_3d(sw, p['length'], p['width'], p['height'])
        elif t == 'washer':
            name = create_washer(sw, p['outer'], p['inner'], p['height'])
        elif t == 'lshape':
            name = create_lshape(sw, p['width'], p['length'], p['thickness'], p['depth'])
        elif t == 'cross':
            name = create_cross(sw, p['size'], p['thickness'], p['depth'])
        elif t == 'star':
            if parsed.is_2d:
                model = sw.new_part()
                model.SketchManager.InsertSketch(True)
                draw_star(model.SketchManager, 0, 0, p['outer'], p['inner'], p['points'])
                model.SketchManager.InsertSketch(True)
                sw.zoom_to_fit()
                name = f"{p['points']}-Point Star 2D"
            else:
                name = create_star_3d(sw, p['outer'], p['inner'], p['points'], p['height'])
        elif t == 'circle':
            name = create_circle_2d(sw, p['radius'])
        elif t == 'square':
            name = create_square_2d(sw, p['size'])
        elif t == 'rectangle':
            name = create_rectangle_2d(sw, p['width'], p['length'])
        else:
            return False, f"Shape '{t}' not implemented"
        
        return True, name
    except Exception as e:
        return False, str(e)


# =============================================================================
# GUI
# =============================================================================

class FloatingButton(tk.Toplevel):
    def __init__(self, parent, on_click):
        super().__init__(parent)
        self.on_click = on_click
        self.bg_color = '#2d2d2d'
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg=self.bg_color)
        self.button_size = 56
        self.geometry(f"{self.button_size}x{self.button_size}")
        self._position_bottom_right()
        self.canvas = tk.Canvas(self, width=self.button_size, height=self.button_size, highlightthickness=0, bg=self.bg_color)
        self.canvas.pack()
        self._draw_button()
        self.canvas.bind('<Button-1>', self._on_press)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.canvas.bind('<Enter>', lambda e: self._draw_button('#2563EB'))
        self.canvas.bind('<Leave>', lambda e: self._draw_button('#3B82F6'))
        self._drag_data = {'x': 0, 'y': 0, 'dragging': False}
    
    def _position_bottom_right(self):
        self.geometry(f"+{self.winfo_screenwidth() - self.button_size - 30}+{self.winfo_screenheight() - self.button_size - 80}")
    
    def _draw_button(self, color='#3B82F6'):
        self.canvas.delete('all')
        self.canvas.create_rectangle(0, 0, self.button_size, self.button_size, fill=self.bg_color, outline='')
        self.canvas.create_oval(4, 6, self.button_size - 2, self.button_size, fill='#1a1a1a', outline='')
        self.canvas.create_oval(2, 2, self.button_size - 4, self.button_size - 6, fill=color, outline='#2563EB', width=2)
        cx, cy = self.button_size // 2, self.button_size // 2 - 2
        pts = []
        for i in range(8):
            a = i * 45 * (3.14159 / 180)
            r = 12 if i % 2 == 0 else 5
            pts.extend([cx + r * math.sin(a), cy - r * math.cos(a)])
        self.canvas.create_polygon(pts, fill='white', outline='')
    
    def _on_press(self, e):
        self._drag_data = {'x': e.x, 'y': e.y, 'dragging': False}
        self._draw_button('#1D4ED8')
    
    def _on_drag(self, e):
        if abs(e.x - self._drag_data['x']) > 5 or abs(e.y - self._drag_data['y']) > 5:
            self._drag_data['dragging'] = True
        if self._drag_data['dragging']:
            self.geometry(f"+{self.winfo_x() + e.x - self._drag_data['x']}+{self.winfo_y() + e.y - self._drag_data['y']}")
    
    def _on_release(self, e):
        self._draw_button('#3B82F6')
        if not self._drag_data['dragging']:
            self.on_click()


class PromptDialog(tk.Toplevel):
    def __init__(self, parent, on_submit):
        super().__init__(parent)
        self.on_submit = on_submit
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.geometry("420x200")
        self.configure(bg='#2D2D2D')
        self._create_widgets()
        self.geometry(f"+{self.winfo_screenwidth() - 450}+{self.winfo_screenheight() - 300}")
        self.bind('<Escape>', lambda e: self.withdraw())
    
    def _create_widgets(self):
        main = tk.Frame(self, bg='#2D2D2D', padx=16, pady=12)
        main.pack(fill='both', expand=True)
        
        title_frame = tk.Frame(main, bg='#2D2D2D')
        title_frame.pack(fill='x', pady=(0, 8))
        title = tk.Label(title_frame, text="✨ Describe what you want to create", font=('Segoe UI', 11, 'bold'), fg='white', bg='#2D2D2D')
        title.pack(side='left')
        close = tk.Label(title_frame, text="✕", font=('Segoe UI', 12), fg='#888', bg='#2D2D2D', cursor='hand2')
        close.pack(side='right')
        close.bind('<Button-1>', lambda e: self.withdraw())
        
        for w in [title_frame, title]:
            w.bind('<Button-1>', lambda e: setattr(self, '_drag_start', (e.x, e.y)))
            w.bind('<B1-Motion>', lambda e: self.geometry(f"+{self.winfo_x() + e.x - self._drag_start[0]}+{self.winfo_y() + e.y - self._drag_start[1]}"))
        
        self.entry = tk.Entry(main, font=('Segoe UI', 11), bg='#3C3C3C', fg='white', insertbackground='white', relief='flat')
        self.entry.pack(fill='x', ipady=8)
        self.entry.bind('<Return>', lambda e: self._submit())
        
        tk.Label(main, text='3D: cube, box, cylinder, hexagon, triangle, pentagon, star, cross, slot, washer', font=('Segoe UI', 8), fg='#888', bg='#2D2D2D').pack(anchor='w', pady=(5,0))
        tk.Label(main, text='2D: circle, square, rectangle  •  Add "2d" for sketch only', font=('Segoe UI', 8), fg='#888', bg='#2D2D2D').pack(anchor='w', pady=(0,8))
        
        btn_frame = tk.Frame(main, bg='#2D2D2D')
        btn_frame.pack(fill='x')
        self.status = tk.Label(btn_frame, text="", font=('Segoe UI', 9), fg='#4ADE80', bg='#2D2D2D', anchor='w')
        self.status.pack(side='left', fill='x', expand=True)
        tk.Button(btn_frame, text="Cancel", font=('Segoe UI', 10), bg='#3C3C3C', fg='white', relief='flat', padx=16, pady=6, command=self.withdraw).pack(side='right', padx=(10,0))
        tk.Button(btn_frame, text="Create", font=('Segoe UI', 10, 'bold'), bg='#3B82F6', fg='white', relief='flat', padx=20, pady=6, command=self._submit).pack(side='right')
    
    def _submit(self):
        prompt = self.entry.get().strip()
        if not prompt:
            self.status.configure(text="Please enter a description", fg='#F87171')
            return
        self.status.configure(text="Creating...", fg='#4ADE80')
        self.update()
        self.on_submit(prompt)
    
    def show_status(self, msg, error=False):
        self.status.configure(text=msg, fg='#F87171' if error else '#4ADE80')
    
    def show(self):
        self.deiconify()
        self.entry.focus_set()
        self.entry.select_range(0, 'end')
    
    def clear(self):
        self.entry.delete(0, 'end')
        self.status.configure(text="")


class SolidWorksAIApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.sw = SolidWorksApp()
        self.sw.connect()
        self.button = FloatingButton(self.root, self._toggle_dialog)
        self.dialog = PromptDialog(self.root, self._on_submit)
        self.dialog.withdraw()
    
    def _toggle_dialog(self):
        if self.dialog.state() == 'withdrawn':
            self.dialog.show()
        else:
            self.dialog.withdraw()
    
    def _on_submit(self, prompt):
        success, result = process_prompt(self.sw, prompt)
        if success:
            self.dialog.show_status(f"✓ {result}")
            self.dialog.clear()
        else:
            self.dialog.show_status(f"✗ {result}", error=True)
    
    def run(self):
        print("=" * 50)
        print("  SolidWorks AI Prompt Tool")
        print("=" * 50)
        print("\n3D: cube, box, cylinder, hexagon, triangle,")
        print("    pentagon, octagon, ellipse, star, cross,")
        print("    slot, washer, L-shape")
        print("\n2D: circle, square, rectangle")
        print("    (add '2d' or 'sketch' for 2D versions)")
        print("\nLook for the BLUE BUTTON!")
        self.root.mainloop()


if __name__ == "__main__":
    SolidWorksAIApp().run()