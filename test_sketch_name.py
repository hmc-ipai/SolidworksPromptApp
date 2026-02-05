"""Quick test to find how to get the active sketch name."""
import win32com.client
import pythoncom
import os

swApp = win32com.client.GetActiveObject("SldWorks.Application")
nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

template = None
for year in ['2025', '2024', '2023', '2022', '2021']:
    path = rf"C:\ProgramData\SolidWorks\SOLIDWORKS {year}\templates\Part.prtdot"
    if os.path.exists(path):
        template = path
        break

model = swApp.NewDocument(template, 0, 0, 0)
ext = model.Extension
skMgr = model.SketchManager

# Create first sketch
ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, nothing, 0)
skMgr.InsertSketch(True)
skMgr.CreateCircle(0, 0, 0, 0.01, 0, 0)

# Try to get active sketch name while sketch is open
for attr in ['ActiveSketch', 'GetActiveSketch', 'GetActiveSketch2']:
    try:
        obj = getattr(skMgr, attr, None) or getattr(model, attr, None)
        if obj is not None:
            if callable(obj):
                obj = obj()
            print(f"{attr}: {obj}")
            if hasattr(obj, 'Name'):
                print(f"  .Name = {obj.Name}")
    except Exception as e:
        print(f"{attr}: Error - {e}")

# Close first sketch
skMgr.InsertSketch(True)

# Create second sketch (to see the naming pattern)
ext.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, nothing, 0)
skMgr.InsertSketch(True)
skMgr.CreateCircle(0, 0, 0, 0.005, 0, 0)

# Try again while second sketch is open
for attr in ['ActiveSketch']:
    try:
        obj = getattr(skMgr, attr, None) or getattr(model, attr, None)
        if obj is not None:
            if callable(obj):
                obj = obj()
            print(f"Second sketch - {attr}: {obj}")
            if hasattr(obj, 'Name'):
                print(f"  .Name = {obj.Name}")
    except Exception as e:
        print(f"Second sketch - {attr}: Error - {e}")

skMgr.InsertSketch(True)

# List all features to see sketch names
print("\nFeature tree:")
feat = model.FirstFeature
while feat:
    print(f"  {feat.Name} ({feat.GetTypeName2})")
    feat = feat.GetNextFeature
