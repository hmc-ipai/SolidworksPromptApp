"""
Debug: Check what the thin-wall revolve actually creates.
Also try GetBodies2 with all body type values.
Also try creating a VBA macro and running it.
"""
import win32com.client
import pythoncom
import os
import math
import time

swApp = win32com.client.GetActiveObject("SldWorks.Application")
swApp.Visible = True

template = None
for year in ['2025', '2024', '2023', '2022', '2021']:
    path = rf"C:\ProgramData\SolidWorks\SOLIDWORKS {year}\templates\Part.prtdot"
    if os.path.exists(path):
        template = path
        break

radius = 0.01
nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

# Create model with thin wall revolve (the approach that works)
model = swApp.NewDocument(template, 0, 0, 0)
ext = model.Extension
skMgr = model.SketchManager
featMgr = model.FeatureManager

ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, nothing, 0)
skMgr.InsertSketch(True)
skMgr.Create3PointArc(0, radius, 0, 0, -radius, 0, radius, 0, 0)
skMgr.CreateCenterLine(0, -radius, 0, 0, radius, 0)
skMgr.InsertSketch(True)
time.sleep(0.3)

model.ClearSelection2(True)
ext.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", 0, 0, 0, False, 4, nothing, 0)

feat = featMgr.FeatureRevolve2(
    True, False, True, False, False, False,
    0, 6.2831853071796, 0, 0.0,
    False, False, 0.0, 0.0,
    0, radius, radius,
    True, True, True
)
print(f"Feature created: {feat}")
model.ForceRebuild3(True)

# Check feature info
if feat:
    print(f"Feature name: {feat.Name}")
    print(f"Feature type: {feat.GetTypeName2}")

# Check bodies with ALL type values
print("\n--- GetBodies2 with different type values ---")
for btype in range(-1, 6):
    try:
        bodies = model.GetBodies2(btype, False)
        if bodies:
            print(f"  Type {btype}: {len(bodies)} bodies")
            for b in bodies:
                print(f"    - {b.Name}")
        else:
            print(f"  Type {btype}: None/empty")
    except Exception as e:
        print(f"  Type {btype}: Error - {e}")

# Check feature tree for body folders
print("\n--- Feature tree ---")
feat_iter = model.FirstFeature
while feat_iter:
    name = feat_iter.Name
    typename = feat_iter.GetTypeName2
    print(f"  {name} ({typename})")
    # Check sub-features for body folders
    if "Body" in name:
        subfeat = feat_iter.GetFirstSubFeature
        while subfeat:
            print(f"    -> {subfeat.Name} ({subfeat.GetTypeName2})")
            subfeat = subfeat.GetNextSubFeature
    feat_iter = feat_iter.GetNextFeature

# Try to get mass properties differently
print("\n--- Mass properties attempts ---")
try:
    # Create a MassProperty object
    mp = ext.CreateMassProperty
    if mp:
        print(f"MassProperty object: {mp}")
        # Try to get volume
        try:
            vol = mp.Volume
            print(f"Volume: {vol}")
        except Exception as e:
            print(f"Volume error: {e}")
except Exception as e:
    print(f"CreateMassProperty error: {e}")

model.ViewZoomtofit2()
print("\nDone!")
