import win32com.client
import os

swApp = win32com.client.GetActiveObject("SldWorks.Application")
swApp.Visible = True

# Find template
template = None
for year in ['2025', '2024', '2023', '2022', '2021']:
    path = rf"C:\ProgramData\SolidWorks\SOLIDWORKS {year}\templates\Part.prtdot"
    if os.path.exists(path):
        template = path
        break

print(f"Template: {template}")

model = swApp.NewDocument(template, 0, 0, 0)
skMgr = model.SketchManager
featMgr = model.FeatureManager

skMgr.InsertSketch(True)
skMgr.CreateCircle(0, 0, 0, 0.01, 0, 0)
skMgr.InsertSketch(True)

featMgr.FeatureExtrusion2(
    True, False, False, 0, 0, 0.02, 0.00254,
    False, False, False, False,
    1.74532925199433E-02, 1.74532925199433E-02,
    False, False, False, False,
    True, True, True, 0, 0, False
)

model.ViewZoomtofit2()
model.ForceRebuild3(True)
print("Cylinder created successfully!")
