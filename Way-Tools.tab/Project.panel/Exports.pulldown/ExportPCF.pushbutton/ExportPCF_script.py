__title__ = 'Place\nHangers'
__doc__ = """Populates Hangers on both ends of selected pipe.
1. Run Script.
2. Choose Service, Group, and Button by counting each (Starting from Zero)
3. Hangers are poplulated spaced from both ends of pipe.
"""


#Imports
import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

selected_elements = uidoc.Selection.PickObjects(ObjectType.Element)
element_ids = [elem.ElementId for elem in selected_elements]
Autodesk.Revit.DB.Fabrication.FabricationUtils.ExportToPCF(doc,element_ids,"C:\\temp\\My_PCF_File.pcf")

forms.toast(
    "PCF file saved to 'c:\Temp\' folder named 'My_PCF_File'",
    title="PCFExport",
    appid="Murray Tools",
    icon="C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\Murray.ico",
    click="https://murraycompany.com",)