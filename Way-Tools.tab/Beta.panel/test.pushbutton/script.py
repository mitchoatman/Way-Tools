# coding: utf8
import clr
import math
import sys
import re
from fractions import Fraction

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode
from System.Windows.Controls import StackPanel, TextBox, ListBox, Label, ComboBox
from System.Windows.Input import Keyboard

# Revit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# -----------------------------
# DIAMETER MAP
# -----------------------------
DIAMETER_MAP = {
    (0.0, 1.0): 2.0, (1.0, 1.25): 2.5, (1.25, 1.5): 3.0,
    (1.5, 2.5): 4.0, (2.5, 3.5): 5.0, (3.5, 4.5): 6.0,
    (4.5, 7.5): 8.0, (7.5, 8.5): 10.0, (8.5, 10.5): 12.0,
    (10.5, 14.5): 16.0, (14.5, 16.5): 18.0, (16.5, 18.5): 20.0,
    (18.5, 20.5): 22.0, (20.5, 22.5): 24.0, (22.5, 24.5): 26.0,
    (24.5, 26.5): 28.0, (26.5, 28.5): 30.0, (28.5, 30.5): 32.0,
    (30.5, 32.5): 34.0, (32.5, 34.5): 36.0
}

# -----------------------------
# SIZE HELPERS
# -----------------------------
def clean_size_string(size_str):
    if not size_str:
        return ""
    return re.sub(r'["\']|ø', '', size_str.strip())

def parse_size_string_to_inches(size_str):
    cleaned = clean_size_string(size_str)

    try:
        return float(cleaned)
    except ValueError:
        pass

    # Supports:
    # 1 1/2
    # 1-1/2
    # 1/2
    m = re.match(r'^\s*(?:(\d+)[-\s])?(\d+/\d+)\s*$', cleaned)
    if m:
        int_part, frac_part = m.groups()
        val = float(Fraction(frac_part))
        if int_part:
            val += float(int_part)
        return val

    raise Exception("Could not parse Overall Size: {}".format(size_str))

def get_mapped_sleeve_diameter_feet(host_part):
    p = host_part.get_Parameter(BuiltInParameter.RBS_REFERENCE_OVERALLSIZE)
    if p is None:
        p = host_part.LookupParameter("Overall Size")
    if p is None:
        raise Exception("Could not find 'Overall Size' parameter.")

    raw = p.AsString()
    if not raw:
        raise Exception("Overall Size is empty.")

    pipe_dia_in = parse_size_string_to_inches(raw)

    for (lo, hi), sleeve_in in DIAMETER_MAP.items():
        if lo < pipe_dia_in <= hi:
            return sleeve_in / 12.0  # internal units = feet

    return 2.0 / 12.0

# -----------------------------
# PIPE POINTS & DIRECTION
# -----------------------------
def get_pipe_direction(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p1 - p0).Normalize()

def is_3d_view(view):
    return view.ViewType == ViewType.ThreeD

# -----------------------------
# ROTATION UTILITIES
# -----------------------------
def rotate_to_vector(doc, element, origin, from_vec, to_vec):
    from_vec = from_vec.Normalize()
    to_vec = to_vec.Normalize()
    axis = from_vec.CrossProduct(to_vec)

    if axis.GetLength() < 1e-8:
        dot = from_vec.DotProduct(to_vec)
        if dot < 0:
            axis = XYZ.BasisZ
            angle = math.pi
        else:
            return
    else:
        axis = axis.Normalize()
        angle = math.acos(max(min(from_vec.DotProduct(to_vec), 1.0), -1.0))

    rot_line = Line.CreateBound(origin, origin + axis)
    ElementTransformUtils.RotateElement(doc, element.Id, rot_line, angle)

# -----------------------------
# POINT CONVERTER
# -----------------------------
class PointConverter:
    """Convert coordinates between internal / project / survey systems."""
    def __init__(self, x, y, z, coord_sys='internal', doc=None):
        if doc is None:
            doc = __revit__.ActiveUIDocument.Document
        self.doc = doc
        pt = XYZ(x, y, z)

        srv_trans = self._get_survey_transform()
        proj_trans = self._get_project_transform()

        if coord_sys.lower() == 'internal':
            self.internal = pt
            self.survey   = srv_trans.Inverse.OfPoint(pt)
            self.project  = proj_trans.Inverse.OfPoint(pt)
        elif coord_sys.lower() == 'project':
            self.project  = pt
            self.internal = proj_trans.OfPoint(pt)
            self.survey   = srv_trans.Inverse.OfPoint(self.internal)
        elif coord_sys.lower() == 'survey':
            self.survey   = pt
            self.internal = srv_trans.OfPoint(pt)
            self.project  = proj_trans.Inverse.OfPoint(self.internal)
        else:
            raise ValueError("coord_sys must be 'internal', 'project' or 'survey'")

    def _get_survey_transform(self):
        return self.doc.ActiveProjectLocation.GetTotalTransform()

    def _get_project_transform(self):
        collector = FilteredElementCollector(self.doc).OfClass(ProjectLocation).WhereElementIsNotElementType()
        for loc in collector:
            if loc.Name == "Project":
                return loc.GetTotalTransform()
        return Transform.Identity

def is_vertical_pipe(pipe):
    if pipe.ItemCustomId != 2041:
        return False
    conns = list(pipe.ConnectorManager.Connectors)
    if len(conns) < 2:
        return False
    direction = (conns[1].Origin - conns[0].Origin).Normalize()
    return abs(direction.Z) > 0.99

# -----------------------------
# INTERSECTION CALC
# -----------------------------
def get_pipe_intersections(pipe, levels):
    curve = pipe.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    intersection_data = []

    dz = (p1.Z - p0.Z)
    if abs(dz) < 1e-6:
        return []

    for level in levels:
        z = level.Elevation
        t = (z - p0.Z) / dz

        if 0.0 <= t <= 1.0:
            pt = p0 + (p1 - p0) * t
            intersection_data.append((pt, level))

    return intersection_data

def align_top_to_point(doc, part, target_point):
    bbox = part.get_BoundingBox(None)
    if bbox is None:
        return

    top_z = bbox.Max.Z
    delta_z = target_point.Z - top_z
    move_vec = XYZ(0, 0, delta_z)
    ElementTransformUtils.MoveElement(doc, part.Id, move_vec)

# -----------------------------
# GET PRE-SELECTED OR VIEW PIPES
# -----------------------------
selected_ids = uidoc.Selection.GetElementIds()
host_parts = []

if selected_ids.Count > 0:
    # Pre-selection mode
    for eid in selected_ids:
        element = doc.GetElement(eid)
        if isinstance(element, FabricationPart) and isinstance(element.Location, LocationCurve):
            host_parts.append(element)
    
    if not host_parts:
        TaskDialog.Show("Error", "No valid fabrication pipes in selection.")
        sys.exit()
else:
    # Default mode - all visible in current view
    curview = doc.ActiveView
    is_3d = curview.ViewType == ViewType.ThreeD
    is_plan = curview.ViewType == ViewType.FloorPlan
    
    if not is_3d and not is_plan:
        TaskDialog.Show("Error", "This script supports 3D and Floor Plan views only.")
        sys.exit()
    
    visible_pipes = FilteredElementCollector(doc, curview.Id)\
                    .OfCategory(BuiltInCategory.OST_FabricationPipework)\
                    .WhereElementIsNotElementType().ToElements()
    
    for pipe in visible_pipes:
        if isinstance(pipe, FabricationPart) and isinstance(pipe.Location, LocationCurve):
            host_parts.append(pipe)
    
    if not host_parts:
        TaskDialog.Show("Info", "No fabrication pipes found in current view.")
        sys.exit()

# -----------------------------
# GET SERVICE & BUTTONS
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None

for s in services:
    if s.Name == "Plumbing: Sleeves":
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Could not find a fabrication service named 'Plumbing: Sleeves'.")
    sys.exit()

palette_names = []
button_records = []
for p in range(target_service.PaletteCount):
    palette_name = target_service.GetPaletteName(p)
    palette_names.append(palette_name)
    for i in range(target_service.GetButtonCount(p)):
        btn = target_service.GetButton(p, i)
        if btn.ConditionCount > 1:
            for c in range(btn.ConditionCount):
                display = u"{1}".format(btn.Name, btn.GetConditionName(c))
                button_records.append({
                    "palette_index": p,
                    "palette_name": palette_name,
                    "display": display,
                    "button": btn,
                    "condition_index": c
                })
        else:
            display = u"{0}".format(btn.Name)
            button_records.append({
                "palette_index": p,
                "palette_name": palette_name,
                "display": display,
                "button": btn,
                "condition_index": 0
            })

if not button_records:
    TaskDialog.Show("Error", "No fabrication buttons found for the 'Plumbing: Sleeves' service.")
    sys.exit()

# -----------------------------
# WPF PART PICKER
# -----------------------------
class PartPicker(Window):
    def __init__(self, records, palettes):
        self.all_records = list(records)
        self.filtered_records = list(records)
        self.selected_record = None
        self.Title = "Select Fabrication Part"
        self.Width = 400
        self.Height = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        stack = StackPanel()
        stack.Margin = Thickness(10)

        lbl_palette = Label()
        lbl_palette.Content = "Palette:"
        stack.Children.Add(lbl_palette)

        self.palette_combo = ComboBox()
        self.palette_combo.Margin = Thickness(0,0,0,10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes:
            self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        stack.Children.Add(self.palette_combo)

        lbl_search = Label()
        lbl_search.Content = "Search Part:"
        stack.Children.Add(lbl_search)

        self.search_box = TextBox()
        self.search_box.Margin = Thickness(0,0,0,10)
        self.search_box.TextChanged += self.apply_filters
        stack.Children.Add(self.search_box)

        lbl_instr = Label()
        lbl_instr.Content = "Double Click Item to Insert"
        lbl_instr.Margin = Thickness(0,0,0,5)
        stack.Children.Add(lbl_instr)

        self.list_box = ListBox()
        self.list_box.Height = 430
        self.list_box.Margin = Thickness(0,0,0,10)
        self.list_box.MouseDoubleClick += self.on_double_click
        stack.Children.Add(self.list_box)

        self.Content = stack
        self.refresh_list()
        self.search_box.Focus()
        Keyboard.Focus(self.search_box)

    def refresh_list(self):
        self.list_box.ItemsSource = [r["display"] for r in self.filtered_records]

    def apply_filters(self, sender, args):
        sel_palette = self.palette_combo.SelectedItem
        search_text = self.search_box.Text.lower().strip()
        records = self.all_records

        if sel_palette and sel_palette != "All Palettes":
            records = [r for r in records if r["palette_name"] == sel_palette]
        if search_text:
            records = [r for r in records if search_text in r["display"].lower()]

        self.filtered_records = records
        self.refresh_list()

    def on_double_click(self, sender, args):
        idx = self.list_box.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_records):
            TaskDialog.Show("Error", "Please select a part.")
            return
        self.selected_record = self.filtered_records[idx]
        self.DialogResult = True
        self.Close()

# -----------------------------
# SHOW DIALOG
# -----------------------------
dlg = PartPicker(button_records, palette_names)
if not dlg.ShowDialog():
    sys.exit()

selected_record = dlg.selected_record
fab_btn = selected_record["button"]
condition_index = selected_record["condition_index"]

# -----------------------------
# CREATE PART + MOVE + ROTATE
# -----------------------------
t = None
try:
    t = Transaction(doc, "Place Fabrication Sleeves at Level Intersections")
    t.Start()

    all_levels = list(FilteredElementCollector(doc).OfClass(Level))
    placed_count = 0

    for host_part in host_parts:
        try:
            new_diameter = get_mapped_sleeve_diameter_feet(host_part)
        except:
            continue

        intersections = get_pipe_intersections(host_part, all_levels)
        if not intersections:
            continue

        pipe_dir = get_pipe_direction(host_part)

        for pt, level in intersections:
            new_part = FabricationPart.Create(
                doc,
                fab_btn,
                condition_index,
                level.Id
            )

            doc.Regenerate()

            size_param = new_part.LookupParameter("Main Primary Diameter")
            if size_param and not size_param.IsReadOnly:
                size_param.Set(new_diameter)

            move_vec = pt - new_part.Origin
            ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)

            rotate_to_vector(doc, new_part, pt, XYZ.BasisX, pipe_dir.Multiply(-1))

            doc.Regenerate()

            align_top_to_point(doc, new_part, pt)

            placed_count += 1

    if placed_count == 0:
        TaskDialog.Show("Info", "No sleeves were placed.")
        t.RollBack()
        sys.exit()

    t.Commit()

except Exception as ex:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    TaskDialog.Show("Error", str(ex))