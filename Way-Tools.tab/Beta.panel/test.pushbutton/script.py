# coding: utf8
import clr
import math
import re
import sys

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
from fractions import Fraction

# Revit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# -----------------------------
# SELECTION FILTER
# -----------------------------
class FabricationStraightSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, FabricationPart) and isinstance(elem.Location, LocationCurve)
        except:
            return False
    def AllowReference(self, reference, point):
        return False

# -----------------------------
# PIPE POINTS & DIRECTION
# -----------------------------
def get_projected_insert_point(host_part):
    picked_point = uidoc.Selection.PickPoint("Pick a point along the pipe centerline")
    curve = host_part.Location.Curve
    result = curve.Project(picked_point)
    if not result:
        raise Exception("Could not project picked point onto pipe centerline.")
    return result.XYZPoint

def get_pipe_midpoint(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p0 + p1) * 0.5

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
        # Parallel or anti-parallel
        dot = from_vec.DotProduct(to_vec)
        if dot < 0:
            # 180 deg flip
            axis = XYZ.BasisZ
            angle = math.pi
        else:
            return
    else:
        axis = axis.Normalize()
        angle = math.acos(max(min(from_vec.DotProduct(to_vec), 1.0), -1.0))

    rot_line = Line.CreateBound(origin, origin + axis)
    ElementTransformUtils.RotateElement(doc, element.Id, rot_line, angle)

def ensure_facing_user(doc, element, origin):
    view = doc.ActiveView
    # Make part "face the user" in section/plan views
    view_dir = view.ViewDirection.Normalize()
    # Part's forward axis (assume after previous alignment it's along X)
    part_dir = XYZ.BasisX
    # Rotate about view_dir to align part's projected X to screen right
    # Project part_dir onto view plane
    part_proj = part_dir - view_dir.Multiply(part_dir.DotProduct(view_dir))
    part_proj_length = part_proj.GetLength()
    if part_proj_length < 1e-6:
        return
    part_proj = part_proj.Normalize()
    # Screen right in view
    screen_right = view.RightDirection.Normalize()
    axis = view_dir
    angle = math.acos(max(min(part_proj.DotProduct(screen_right), 1.0), -1.0))
    # Determine correct rotation direction
    if part_proj.CrossProduct(screen_right).DotProduct(axis) < 0:
        angle = -angle
    rot_line = Line.CreateBound(origin, origin + axis)
    ElementTransformUtils.RotateElement(doc, element.Id, rot_line, angle)

# -----------------------------
# SELECTION
# -----------------------------
straight_filter = FabricationStraightSelectionFilter()
try:
    host_parts = list(uidoc.Selection.PickElementsByRectangle(
        straight_filter,
        "Window-select fabrication pipes/straights"
    ))
except OperationCanceledException:
    sys.exit()

if not host_parts:
    TaskDialog.Show("Info", "No fabrication pipes selected.")
    sys.exit()


# -----------------------------
# GET SERVICE & BUTTONS
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None

# Search for the service named "Sleeves"
for s in services:
    if s.Name == "Plumbing: Sleeves":
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Could not find a fabrication service named 'Sleeves'.")
    sys.exit()

# Proceed to get the buttons for the 'Sleeves' service
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
    TaskDialog.Show("Error", "No fabrication buttons found for the 'Sleeves' service.")
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

        lbl_palette = Label(); lbl_palette.Content = "Palette:"; stack.Children.Add(lbl_palette)
        self.palette_combo = ComboBox(); self.palette_combo.Margin = Thickness(0,0,0,10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes: self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        stack.Children.Add(self.palette_combo)

        lbl_search = Label(); lbl_search.Content = "Search Part:"; stack.Children.Add(lbl_search)
        self.search_box = TextBox(); self.search_box.Margin = Thickness(0,0,0,10)
        self.search_box.TextChanged += self.apply_filters; stack.Children.Add(self.search_box)

        lbl_instr = Label(); lbl_instr.Content = "Double Click Item to Insert"; lbl_instr.Margin = Thickness(0,0,0,5); stack.Children.Add(lbl_instr)
        self.list_box = ListBox(); self.list_box.Height = 430; self.list_box.Margin = Thickness(0,0,0,10)
        self.list_box.MouseDoubleClick += self.on_double_click
        stack.Children.Add(self.list_box)

        self.Content = stack
        self.refresh_list()
        self.search_box.Focus(); Keyboard.Focus(self.search_box)

    def refresh_list(self):
        self.list_box.ItemsSource = [r["display"] for r in self.filtered_records]
    def apply_filters(self, sender, args):
        sel_palette = self.palette_combo.SelectedItem
        search_text = self.search_box.Text.lower().strip()
        records = self.all_records
        if sel_palette and sel_palette != "All Palettes": records = [r for r in records if r["palette_name"] == sel_palette]
        if search_text: records = [r for r in records if search_text in r["display"].lower()]
        self.filtered_records = records
        self.refresh_list()
    def on_double_click(self, sender, args):
        idx = self.list_box.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_records):
            TaskDialog.Show("Error","Please select a part."); return
        self.selected_record = self.filtered_records[idx]; self.DialogResult = True; self.Close()

# -----------------------------
# SHOW DIALOG
# -----------------------------
dlg = PartPicker(button_records, palette_names)
if not dlg.ShowDialog(): sys.exit()
selected_record = dlg.selected_record
fab_btn = selected_record["button"]
condition_index = selected_record["condition_index"]



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
        return []  # horizontal pipe → no level crossings

    for level in levels:
        z = level.Elevation

        # parametric intersection along pipe
        t = (z - p0.Z) / dz

        if 0.0 <= t <= 1.0:
            pt = p0 + (p1 - p0) * t
            intersection_data.append((pt, level))

    return intersection_data


def align_top_to_point(doc, part, target_point):
    bbox = part.get_BoundingBox(None)
    if bbox is None:
        return

    # Top Z of the part
    top_z = bbox.Max.Z

    # How far off we are
    delta_z = target_point.Z - top_z

    # Move only in Z
    move_vec = XYZ(0, 0, delta_z)
    ElementTransformUtils.MoveElement(doc, part.Id, move_vec)

# -----------------------------
# FORMATTING UNITS OF OVERALLSIZE
# -----------------------------
def clean_size_string(size_str):
    if not size_str:
        return ""
    return re.sub(r'["\']|ø', '', size_str.strip())

def parse_overall_size_to_feet(host_part):
    # Try built-in first
    p = host_part.get_Parameter(BuiltInParameter.RBS_REFERENCE_OVERALLSIZE)

    # Fallback to named parameter if needed
    if p is None:
        p = host_part.LookupParameter("Overall Size")

    if p is None:
        raise Exception("Could not find 'Overall Size' parameter.")

    raw = clean_size_string(p.AsString() or "")
    if not raw:
        raise Exception("Overall Size is empty.")

    try:
        dia_in = float(raw)
    except ValueError:
        m = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)$', raw)
        if m:
            int_part, frac_part = m.groups()
            dia_in = float(Fraction(frac_part))
            if int_part:
                dia_in += float(int_part)
        else:
            raise Exception("Could not parse Overall Size: {}".format(raw))

    return dia_in / 12.0  # Revit internal units = feet

# -----------------------------
# CREATE PART + MOVE + ROTATE
# -----------------------------
t = None
try:
    t = Transaction(doc, "Place Fabrication Sleeves at Level Intersections")
    t.Start()

    # Get all levels once
    all_levels = list(FilteredElementCollector(doc).OfClass(Level))

    placed_count = 0

    for host_part in host_parts:
        # Get pipe size from Overall Size
        try:
            current_diameter = parse_overall_size_to_feet(host_part)
        except:
            continue

        new_diameter = current_diameter + (2.0 / 12.0)  # keep your +2" clearance

        # Get intersections for this pipe
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

            # Size
            size_param = new_part.LookupParameter("Main Primary Diameter")
            if size_param and not size_param.IsReadOnly:
                size_param.Set(new_diameter)

            # Move to intersection
            move_vec = pt - new_part.Origin
            ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)

            # Rotate (flip-flopped version)
            rotate_to_vector(doc, new_part, pt, XYZ.BasisX, pipe_dir.Multiply(-1))

            doc.Regenerate()

            # Align TOP to intersection
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