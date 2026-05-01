# coding: utf8
import clr
import math
import sys
import re
from fractions import Fraction
import os

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

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
# SLEEVE LENGTH FROM FILE
# -----------------------------
temp_folder = r"c:\Temp"
sleeve_length_file = os.path.join(temp_folder, 'Ribbon_Sleeve.txt')

if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)

if not os.path.exists(sleeve_length_file):
    with open(sleeve_length_file, 'w') as f:
        f.write('6')

with open(sleeve_length_file, 'r') as f:
    sleeve_length = float(f.read().strip())

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
            return sleeve_in / 12.0

    return 2.0 / 12.0

def set_part_size_and_length(new_part, host_part):
    try:
        new_diameter = get_mapped_sleeve_diameter_feet(host_part)
        size_param = new_part.LookupParameter("Main Primary Diameter")
        if size_param and not size_param.IsReadOnly:
            size_param.Set(new_diameter)
    except:
        pass

    try:
        length_param = new_part.LookupParameter("Length")
        if length_param and not length_param.IsReadOnly:
            length_param.Set(sleeve_length)
    except:
        pass

# -----------------------------
# PIPE HELPERS
# -----------------------------
def get_pipe_direction(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p1 - p0).Normalize()

def get_horizontal_pipe_direction(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    v = (p1 - p0)
    horiz = XYZ(v.X, v.Y, 0.0)

    if horiz.GetLength() < 1e-8:
        raise Exception("Host pipe is vertical or too close to vertical for wall sleeve placement.")

    return horiz.Normalize()

def get_pipe_midpoint(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p0 + p1) * 0.5

def get_projected_insert_point(host_part):
    picked_point = uidoc.Selection.PickPoint("Pick a point along the pipe centerline")

    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    # Plan-based projection only to prevent drift on sloped pipes
    v2d = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
    w2d = XYZ(picked_point.X - p0.X, picked_point.Y - p0.Y, 0.0)

    v2d_len_sq = v2d.DotProduct(v2d)
    if v2d_len_sq < 1e-8:
        raise Exception("Pipe is vertical or too close to vertical for plan-based point projection.")

    t = w2d.DotProduct(v2d) / v2d_len_sq
    t = max(0.0, min(1.0, t))

    # Rebuild true 3D point on actual sloped pipe
    return p0 + (p1 - p0).Multiply(t)

def is_3d_view(view):
    return view.ViewType == ViewType.ThreeD

# -----------------------------
# ROTATION
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
# WALL SLEEVE END-POINT HELPERS
# -----------------------------
def get_part_connectors(part):
    try:
        return list(part.ConnectorManager.Connectors)
    except:
        return []

def get_end_connector_point(part, direction_vec):
    conns = get_part_connectors(part)
    if len(conns) < 2:
        return part.Origin

    best_conn = None
    best_val = None

    for c in conns:
        pt = c.Origin
        val = pt.X * direction_vec.X + pt.Y * direction_vec.Y + pt.Z * direction_vec.Z
        if best_conn is None or val < best_val:
            best_conn = c
            best_val = val

    return best_conn.Origin

# -----------------------------
# LEVEL INTERSECTIONS
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
# HOST COLLECTION FOR FLOOR MODE
# -----------------------------
def collect_vertical_hosts():
    selected_ids = uidoc.Selection.GetElementIds()
    host_parts = []

    if selected_ids.Count > 0:
        for eid in selected_ids:
            element = doc.GetElement(eid)
            if isinstance(element, FabricationPart) and isinstance(element.Location, LocationCurve):
                host_parts.append(element)

        if not host_parts:
            TaskDialog.Show("Error", "No valid fabrication pipes in selection.")
            sys.exit()
    else:
        curview = doc.ActiveView
        is_3d = curview.ViewType == ViewType.ThreeD
        is_plan = curview.ViewType == ViewType.FloorPlan

        if not is_3d and not is_plan:
            TaskDialog.Show("Error", "This script supports 3D and Floor Plan views only.")
            sys.exit()

        visible_pipes = FilteredElementCollector(doc, curview.Id) \
            .OfCategory(BuiltInCategory.OST_FabricationPipework) \
            .WhereElementIsNotElementType() \
            .ToElements()

        for pipe in visible_pipes:
            if isinstance(pipe, FabricationPart) and isinstance(pipe.Location, LocationCurve):
                host_parts.append(pipe)

        if not host_parts:
            TaskDialog.Show("Info", "No fabrication pipes found in current view.")
            sys.exit()

    return host_parts

# -----------------------------
# GET SERVICE & BUTTONS
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None

for s in services:
    if s.Name and 'sleeves' in s.Name.lower():
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Could not find a Fabrication Service name containing 'Sleeve'.")
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
                display = u"{} - {}".format(btn.Name, btn.GetConditionName(c))
                button_records.append({
                    "palette_index": p,
                    "palette_name": palette_name,
                    "display": display,
                    "button": btn,
                    "condition_index": c
                })
        else:
            display = u"{}".format(btn.Name)
            button_records.append({
                "palette_index": p,
                "palette_name": palette_name,
                "display": display,
                "button": btn,
                "condition_index": 0
            })

if not button_records:
    TaskDialog.Show("Error", "No fabrication buttons found for the sleeve service.")
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
is_wall_mode = "wall" in selected_record["display"].lower()

# -----------------------------
# MAIN
# -----------------------------
t = None

try:
    t = Transaction(doc, "Place Fabrication Sleeves")
    t.Start()

    if is_wall_mode:
        # ---------------------------------
        # WALL SLEEVE MODE
        # picked point = sleeve END
        # sleeve stays flat
        # ---------------------------------
        straight_filter = FabricationStraightSelectionFilter()

        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                straight_filter,
                "Select fabrication pipe/straight for wall sleeve"
            )
        except OperationCanceledException:
            t.RollBack()
            sys.exit()

        host_part = doc.GetElement(ref.ElementId)

        try:
            if is_3d_view(doc.ActiveView):
                insert_point = get_pipe_midpoint(host_part)
            else:
                insert_point = get_projected_insert_point(host_part)
        except OperationCanceledException:
            t.RollBack()
            sys.exit()
        except Exception as ex:
            t.RollBack()
            TaskDialog.Show("Error", str(ex))
            sys.exit()

        flat_pipe_dir = get_horizontal_pipe_direction(host_part)

        new_part = FabricationPart.Create(doc, fab_btn, condition_index, host_part.LevelId)
        doc.Regenerate()

        set_part_size_and_length(new_part, host_part)
        doc.Regenerate()

        # Rotate first so connector/end logic is based on final direction
        rotate_to_vector(doc, new_part, new_part.Origin, XYZ.BasisX, flat_pipe_dir)
        doc.Regenerate()

        # Move so picked point aligns to sleeve END, not center
        end_point = get_end_connector_point(new_part, flat_pipe_dir)
        move_vec = insert_point - end_point
        ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)
        doc.Regenerate()

    else:
        # ---------------------------------
        # FLOOR / VERTICAL SLEEVE MODE
        # ---------------------------------
        host_parts = collect_vertical_hosts()
        all_levels = list(FilteredElementCollector(doc).OfClass(Level))
        placed_count = 0

        for host_part in host_parts:
            try:
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

                    set_part_size_and_length(new_part, host_part)
                    doc.Regenerate()

                    move_vec = pt - new_part.Origin
                    ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)

                    rotate_to_vector(doc, new_part, pt, XYZ.BasisX, pipe_dir.Multiply(-1))
                    doc.Regenerate()

                    align_top_to_point(doc, new_part, pt)
                    placed_count += 1

            except:
                continue

        if placed_count == 0:
            TaskDialog.Show("Info", "No sleeves were placed.")
            t.RollBack()
            sys.exit()

    t.Commit()

except Exception as ex:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    TaskDialog.Show("Error", str(ex))