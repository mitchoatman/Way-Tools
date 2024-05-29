# import Autodesk
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import ISet

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument

# #this is start of select element(s)
# sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element, 'Select Elements')
# selected_elements = [doc.GetElement( elId ) for elId in sel]
# #this is end of select element(s)

# # Create an ISet[ElementId] of the selected elements
# element_set = ISet[ElementId](list[ElementId](selected_elements))

# print element_set

# FabricationPart.SaveAsFabricationJob(doc, element_set, "C:/output.maj", True)

import Autodesk
import clr
import System
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB.FabricationPart import SaveAsFabricationJob
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
clr.AddReference("System.Core")
from System.Collections.Generic import HashSet

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

selected_elements = uidoc.Selection.PickObjects(ObjectType.Element)
element_ids = [elem.ElementId for elem in selected_elements]

# Create a new HashSet[ElementId] object
id_set = HashSet[ElementId]()
# Add the selected IDs to the HashSet
for id in element_ids:
    id_set.Add(id)

options = FabricationSaveJobOptions()

SaveAsFabricationJob(doc, id_set, "C:\\Temp\\Output.maj", options)