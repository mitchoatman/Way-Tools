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
# DEFAULT SLEEVE LENGTH FILE
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

def set_part_size_and_length(new_part, host_part, override_length=None):
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
            if override_length is not None:
                length_param.Set(override_length)
            else:
                length_param.Set(sleeve_length)
    except:
        pass

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
# CONNECTOR HELPERS
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
# SELECTION FILTERS
# -----------------------------
class FabricationStraightSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, FabricationPart) and isinstance(elem.Location, LocationCurve)
        except:
            return False

    def AllowReference(self, reference, point):
        return False

class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return isinstance(element, RevitLinkInstance)
        except:
            return False

    def AllowReference(self, reference, point):
        try:
            link_instance = doc.GetElement(reference.ElementId)
            if not isinstance(link_instance, RevitLinkInstance):
                return False

            link_doc = link_instance.GetLinkDocument()
            if not link_doc:
                return False

            linked_elem = link_doc.GetElement(reference.LinkedElementId)
            return isinstance(linked_elem, Wall)
        except:
            return False

# -----------------------------
# LINKED WALL HELPERS
# -----------------------------
def select_linked_wall():
    return uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        LinkedWallSelectionFilter(),
        "Select a wall in a Revit link"
    )

def get_linked_wall_thickness_and_curve(link_ref):
    link_instance = doc.GetElement(link_ref.ElementId)
    if link_instance is None:
        raise Exception("Could not get Revit link instance.")

    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        raise Exception("Could not access linked document.")

    wall = link_doc.GetElement(link_ref.LinkedElementId)
    if wall is None or not isinstance(wall, Wall):
        raise Exception("Selected linked element is not a valid wall.")

    wall_loc = wall.Location
    if not isinstance(wall_loc, LocationCurve):
        raise Exception("Selected linked wall does not have a valid location curve.")

    wall_curve = wall_loc.Curve
    link_transform = link_instance.GetTotalTransform()
    wall_curve_host = wall_curve.CreateTransformed(link_transform)

    thickness = wall.WallType.Width
    return thickness, wall_curve_host

def project_wall_curve_to_pipe_plane(wall_curve, pipe_curve):
    pipe_start = pipe_curve.GetEndPoint(0)
    pipe_end = pipe_curve.GetEndPoint(1)
    pipe_direction = (pipe_end - pipe_start).Normalize()

    plane_normal = XYZ(0, 0, 1)
    if abs(pipe_direction.Z) > 0.99:
        plane_normal = XYZ(1, 0, 0)

    plane = Plane.CreateByNormalAndOrigin(plane_normal, pipe_start)

    wall_start = wall_curve.GetEndPoint(0)
    wall_end = wall_curve.GetEndPoint(1)

    uv_start, dist_start = plane.Project(wall_start)
    uv_end, dist_end = plane.Project(wall_end)

    origin = plane.Origin
    x_axis = plane.XVec
    y_axis = plane.YVec

    projected_start = origin + x_axis.Multiply(uv_start.U) + y_axis.Multiply(uv_start.V)
    projected_end = origin + x_axis.Multiply(uv_end.U) + y_axis.Multiply(uv_end.V)

    return Line.CreateBound(projected_start, projected_end)

def get_wall_sleeve_data_from_link(host_part, link_ref):
    pipe_curve = host_part.Location.Curve
    wall_thickness, wall_curve = get_linked_wall_thickness_and_curve(link_ref)

    projected_wall_curve = project_wall_curve_to_pipe_plane(wall_curve, pipe_curve)

    result_array = clr.Reference[IntersectionResultArray]()
    result = pipe_curve.Intersect(projected_wall_curve, result_array)

    if result != SetComparisonResult.Overlap or not result_array.Value or result_array.Value.Size == 0:
        raise Exception("Pipe does not intersect the selected linked wall.")

    intersection_point = result_array.Value[0].XYZPoint

    p0 = pipe_curve.GetEndPoint(0)
    p1 = pipe_curve.GetEndPoint(1)

    if p0.DistanceTo(intersection_point) <= p1.DistanceTo(intersection_point):
        near_pt = p0
        far_pt = p1
    else:
        near_pt = p1
        far_pt = p0

    flat_dir = XYZ(far_pt.X - near_pt.X, far_pt.Y - near_pt.Y, 0.0)
    if flat_dir.GetLength() < 1e-8:
        raise Exception("Host pipe is vertical or too close to vertical for wall sleeve placement.")

    flat_dir = flat_dir.Normalize()

    start_point = intersection_point - flat_dir.Multiply(wall_thickness * 0.5)
    return start_point, wall_thickness, flat_dir

# -----------------------------
# GET SERVICE & WALL BUTTONS ONLY
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None

for s in services:
    if s.Name:
        n = s.Name.lower()
        if 'sleeve' in n or 'sleeves' in n:
            target_service = s
            break

if not target_service:
    TaskDialog.Show("Error", "Could not find a Fabrication Service name containing 'Sleeve'.")
    sys.exit()

palette_names = []
button_records = []

for p in range(target_service.PaletteCount):
    palette_name = target_service.GetPaletteName(p)

    local_records = []

    for i in range(target_service.GetButtonCount(p)):
        btn = target_service.GetButton(p, i)

        if btn.ConditionCount > 1:
            for c in range(btn.ConditionCount):
                display = u"{} - {}".format(btn.Name, btn.GetConditionName(c))
                is_wall = ("wall" in display.lower()) or ("wall" in palette_name.lower())
                if is_wall:
                    local_records.append({
                        "palette_index": p,
                        "palette_name": palette_name,
                        "display": display,
                        "button": btn,
                        "condition_index": c
                    })
        else:
            display = u"{}".format(btn.Name)
            is_wall = ("wall" in display.lower()) or ("wall" in palette_name.lower())
            if is_wall:
                local_records.append({
                    "palette_index": p,
                    "palette_name": palette_name,
                    "display": display,
                    "button": btn,
                    "condition_index": 0
                })

    if local_records:
        palette_names.append(palette_name)
        button_records.extend(local_records)

if not button_records:
    TaskDialog.Show("Error", "No wall sleeve fabrication buttons were found in the sleeve service.")
    sys.exit()

# -----------------------------
# WPF PART PICKER
# -----------------------------
class PartPicker(Window):
    def __init__(self, records, palettes):
        self.all_records = list(records)
        self.filtered_records = list(records)
        self.selected_record = None
        self.Title = "Select Wall Sleeve Fabrication Part"
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
        self.palette_combo.Margin = Thickness(0, 0, 0, 10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes:
            self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        stack.Children.Add(self.palette_combo)

        lbl_search = Label()
        lbl_search.Content = "Search Wall Sleeve:"
        stack.Children.Add(lbl_search)

        self.search_box = TextBox()
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self.apply_filters
        stack.Children.Add(self.search_box)

        lbl_instr = Label()
        lbl_instr.Content = "Double Click Item to Insert"
        lbl_instr.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(lbl_instr)

        self.list_box = ListBox()
        self.list_box.Height = 430
        self.list_box.Margin = Thickness(0, 0, 0, 10)
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
# MAIN
# -----------------------------
t = None

try:
    straight_filter = FabricationStraightSelectionFilter()

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            straight_filter,
            "Select fabrication pipe/straight for wall sleeve"
        )
    except OperationCanceledException:
        sys.exit()

    host_part = doc.GetElement(ref.ElementId)

    try:
        wall_ref = select_linked_wall()
    except OperationCanceledException:
        sys.exit()

    t = Transaction(doc, "Place Fabrication Wall Sleeve")
    t.Start()

    insert_point, wall_length, flat_pipe_dir = get_wall_sleeve_data_from_link(host_part, wall_ref)

    new_part = FabricationPart.Create(doc, fab_btn, condition_index, host_part.LevelId)
    doc.Regenerate()

    set_part_size_and_length(new_part, host_part, wall_length)
    doc.Regenerate()

    rotate_to_vector(doc, new_part, new_part.Origin, XYZ.BasisX, flat_pipe_dir)
    doc.Regenerate()

    end_point = get_end_connector_point(new_part, flat_pipe_dir)
    move_vec = insert_point - end_point
    ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)
    doc.Regenerate()

    t.Commit()

except Exception as ex:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    TaskDialog.Show("Error", str(ex))