# coding: utf8
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode
from System.Windows.Controls import StackPanel, ComboBox, Button, Label
from System.Windows import HorizontalAlignment

# Revit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#-------------------------------------------------------
# SELECT ELEMENT
#-------------------------------------------------------
ref = uidoc.Selection.PickObject(ObjectType.Element, "Select fabrication part")
element = doc.GetElement(ref.ElementId)

param = element.LookupParameter("Fabrication Service")
if not param or not param.HasValue:
    TaskDialog.Show("Error", "No Fabrication Service found")
    raise Exception("Missing service")

service_name = param.AsValueString()

#-------------------------------------------------------
# GET SERVICE + BUTTONS
#-------------------------------------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()

target_service = None
for s in services:
    if s.Name == service_name:
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Service not found")
    raise Exception("Service not found")

buttons = []
button_map = {}

palette_count = target_service.PaletteCount

for p in range(palette_count):
    count = target_service.GetButtonCount(p)
    for i in range(count):
        btn = target_service.GetButton(p, i)
        name = btn.Name
        buttons.append(name)
        button_map[name] = btn

#-------------------------------------------------------
# WPF DIALOG
#-------------------------------------------------------
class PartPicker(Window):
    def __init__(self, items):
        self.Title = "Select Fabrication Part"
        self.Width = 300
        self.Height = 150
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        stack = StackPanel()
        stack.Margin = Thickness(10)

        label = Label()
        label.Content = "Choose Part:"
        stack.Children.Add(label)

        self.combo = ComboBox()
        self.combo.ItemsSource = items
        self.combo.SelectedIndex = 0
        stack.Children.Add(self.combo)

        btn = Button()
        btn.Content = "Place"
        btn.Margin = Thickness(0,10,0,0)
        btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.Click += self.ok
        stack.Children.Add(btn)

        self.Content = stack

    def ok(self, sender, args):
        self.DialogResult = True
        self.Close()

#-------------------------------------------------------
# SHOW DIALOG
#-------------------------------------------------------
dlg = PartPicker(buttons)

if not dlg.ShowDialog():
    raise Exception("Cancelled")

selected_name = dlg.combo.SelectedItem
fab_btn = button_map[selected_name]

#-------------------------------------------------------
# PICK POINT
#-------------------------------------------------------
ref = uidoc.Selection.PickObject(ObjectType.Element, "Select Fabrication Hanger")
hanger = doc.GetElement(ref.ElementId)

level_id = element.LevelId

rod_info = hanger.GetRodInfo()
rod_count = rod_info.RodCount

rod_points = []

for i in range(rod_count):
    pt = rod_info.GetRodEndPosition(i)
    rod_points.append(pt)

#-------------------------------------------------------
# CREATE PART
#-------------------------------------------------------
t = Transaction(doc, "Place Fabrication Part at Rods")
t.Start()

for pt in rod_points:
    try:
        new_part = FabricationPart.Create(doc, fab_btn, 0, hanger.LevelId)

        if not new_part:
            continue

        # Move to rod location
        bbox = new_part.get_BoundingBox(None)
        if bbox:
            center = (bbox.Min + bbox.Max) / 2
            move = pt - center
            ElementTransformUtils.MoveElement(doc, new_part.Id, move)

    except Exception as e:
        print("Error placing part: {}".format(e))

t.Commit()