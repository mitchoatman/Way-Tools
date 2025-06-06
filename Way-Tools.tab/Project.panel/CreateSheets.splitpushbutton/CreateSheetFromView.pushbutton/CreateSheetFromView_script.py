import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, BuiltInParameter
from Autodesk.Revit.UI.Selection import PickBoxStyle, ISelectionFilter, ObjectType
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)
from pyrevit import forms
from Autodesk.Revit.UI import TaskDialog
import sys

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

#>>>>>>>>>> GET SELECTED VIEWS
selected_views = forms.select_views(use_selection=True)
if not selected_views:
    forms.alert("No views selected. Please try again.", exitscript=True)

#>>>>>> SELECT TBLOCK
selected_titleblocks = forms.select_titleblocks(title='Select Titleblock', button_name='Select', no_tb_option='No Title Block', width=500, multiple=False, filterfunc=None)
if not selected_titleblocks:
    TaskDialog.Show("Error", "No title block selected.")
    sys.exit()

#>>>>>> DEFINE SHEET NUMBER AND NAME
# Set default sheet number text based on number of selected views
default_sheet_number = 'Varies - Don\'t Change' if len(selected_views) > 1 else 'Sheet Number'
components = [Label('Sheet Number:'),
              TextBox('sheetnumber', default_sheet_number),
              Label('Sheet Name:'),
              TextBox('sheetname', 'Sheet Name'),
              Button('Ok')]
form = FlexForm('Title', components)
form.show()

if not form.values:
    sys.exit()  # Exit if form is canceled

snumber = form.values['sheetnumber']
sname = form.values['sheetname']

# Check if sheet number already exists (only if single view and user provided a custom sheet number)
if len(selected_views) == 1 and snumber != 'Varies - Don\'t Change':
    sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(ViewSheet).ToElements())
    if sheet_exists:
        TaskDialog.Show("Error", "Sheet with number {} already exists.".format(snumber))
        sys.exit()

# Check if selected views are already placed on any sheet
for view in selected_views:
    viewports = fec(doc).OfClass(Viewport).ToElements()
    for viewport in viewports:
        if viewport.ViewId == view.Id:
            TaskDialog.Show("Error", "View '{}' is already placed on another sheet.".format(view.Name))
            sys.exit()

borders = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()

titleblockid = borders.FirstElementId()

#define a transaction variable and describe the transaction
t = Transaction(doc, 'Sheet From View')

# Begin new transaction
t.Start()
for i, view in enumerate(selected_views):
    # Use view name as sheet number, or user-provided sheet number for single view
    current_sheet_number = snumber if len(selected_views) == 1 and snumber != 'Varies - Don\'t Change' else view.Name
    
    # Check if the generated sheet number already exists
    sheet_exists = any(sheet.SheetNumber == current_sheet_number for sheet in fec(doc).OfClass(ViewSheet).ToElements())
    if sheet_exists:
        TaskDialog.Show("Error", "Sheet with number {} already exists.".format(current_sheet_number))
        t.RollBack()
        sys.exit()
    
    # Create new sheet
    SHEET = ViewSheet.Create(doc, selected_titleblocks)
    SHEET.Name = sname
    SHEET.SheetNumber = current_sheet_number
    x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
    y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
    ViewLocation = XYZ(x, y, 0.0)
    NEWSHEET = Viewport.Create(doc, SHEET.Id, view.Id, ViewLocation)
    
t.Commit()

# Set the active view to the last created sheet
if i == len(selected_views) - 1:
    uidoc.RequestViewChange(SHEET)