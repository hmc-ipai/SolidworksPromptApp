"""
Create a sphere in SolidWorks using pywin32 COM automation.

Creates a solid sphere by revolving a semicircle arc 360 degrees
around a centerline axis.

Requirements:
    - SolidWorks must be running
    - pywin32 must be installed (pip install pywin32)

Key insight: FeatureRevolve2 parameter order is:
    Dir1Type, Dir2Type, Dir1Angle, Dir2Angle
    (NOT Dir1Type, Dir1Angle, Dir2Type, Dir2Angle)
"""
import win32com.client
import pythoncom
import os
import time
import math


def get_solidworks():
    """Connect to running SolidWorks instance."""
    swApp = win32com.client.GetActiveObject("SldWorks.Application")
    swApp.Visible = True
    return swApp


def find_template():
    """Find the SolidWorks part template."""
    for year in ['2025', '2024', '2023', '2022', '2021', '2020']:
        path = rf"C:\ProgramData\SolidWorks\SOLIDWORKS {year}\templates\Part.prtdot"
        if os.path.exists(path):
            return path
    raise FileNotFoundError("Could not find SolidWorks part template")


def create_sphere(radius=0.01):
    """
    Create a solid sphere in SolidWorks.

    Args:
        radius: Sphere radius in meters (default 0.01 = 10mm)

    Returns:
        tuple: (model, feature) - the SolidWorks model and revolve feature
    """
    swApp = get_solidworks()
    template = find_template()
    nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    # Create new part document
    model = swApp.NewDocument(template, 0, 0, 0)
    ext = model.Extension
    skMgr = model.SketchManager
    featMgr = model.FeatureManager

    # Open sketch on Front Plane
    ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, nothing, 0)
    skMgr.InsertSketch(True)

    # Draw semicircle arc on the right side of the Y axis
    # 3-point arc: start (top), end (bottom), midpoint (right)
    skMgr.Create3PointArc(0, radius, 0, 0, -radius, 0, radius, 0, 0)

    # Draw centerline along Y axis as the revolve axis
    skMgr.CreateCenterLine(0, -radius, 0, 0, radius, 0)

    # Close sketch
    skMgr.InsertSketch(True)
    time.sleep(0.3)

    # Select the centerline as revolve axis (mark=4)
    model.ClearSelection2(True)
    ext.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", 0, 0, 0, False, 4, nothing, 0)

    # Create solid revolve feature (360 degrees)
    # IMPORTANT: Parameter order is Dir1Type, Dir2Type, Dir1Angle, Dir2Angle
    feat = featMgr.FeatureRevolve2(
        True,               # SingleDir
        True,               # IsSolid
        False,              # IsThin
        False,              # IsCut
        False,              # ReverseDir
        False,              # BothDirectionUpToSameEntity
        0,                  # Dir1Type (swEndCondBlind)
        0,                  # Dir2Type (swEndCondBlind)
        2 * math.pi,        # Dir1Angle (360 degrees in radians)
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
        raise RuntimeError("FeatureRevolve2 returned None - sphere creation failed")

    model.ForceRebuild3(True)
    model.ViewZoomtofit2()

    return model, feat


if __name__ == "__main__":
    radius_mm = 10
    radius_m = radius_mm / 1000.0

    print(f"Creating sphere with radius {radius_mm}mm...")
    model, feat = create_sphere(radius_m)
    print(f"Feature: {feat.Name} ({feat.GetTypeName2})")
    print("Sphere created successfully!")
