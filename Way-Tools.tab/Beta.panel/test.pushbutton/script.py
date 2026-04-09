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
# DUPLICATE DETECTION FUNCTION
# -----------------------------
def is_duplicate_sleeve(point, existing_parts, tol=0.02):
    for el in existing_parts:
        try:
            loc = el.Origin
            if loc and all(abs(a - b) < tol for a, b in zip(
                (loc.X, loc.Y, loc.Z),
                (point.X, point.Y, point.Z)
            )):
                return True
        except:
            continue
    return False

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
    for eid in selected_ids:
        element = doc.GetElement(eid)
        if isinstance(element, FabricationPart) and isinstance(element.Location, LocationCurve):
            host_parts.append(element)
else:
    curview = doc.ActiveView
    visible_pipes = FilteredElementCollector(doc, curview.Id)\
                    .OfCategory(BuiltInCategory.OST_FabricationPipework)\
                    .WhereElementIsNotElementType().ToElements()
    for pipe in visible_pipes:
        if isinstance(pipe, FabricationPart) and isinstance(pipe.Location, LocationCurve):
            host_parts.append(pipe)

# -----------------------------
# SERVICE
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None

for s in services:
    if s.Name == "Plumbing: Sleeves":
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Could not find service.")
    sys.exit()

button = target_service.GetButton(0, 0)
condition_index = 0

# -----------------------------
# CREATE PARTS
# -----------------------------
t = Transaction(doc, "Place Sleeves")
t.Start()

all_levels = list(FilteredElementCollector(doc).OfClass(Level))

# --- DUPLICATE DETECTION (collect existing)
existing_sleeves = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_FabricationPipework)
    .WhereElementIsNotElementType()
    .ToElements()
)

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

        # --- DUPLICATE DETECTION (skip if exists)
        if is_duplicate_sleeve(pt, existing_sleeves):
            continue

        new_part = FabricationPart.Create(doc, button, condition_index, level.Id)

        doc.Regenerate()

        size_param = new_part.LookupParameter("Main Primary Diameter")
        if size_param and not size_param.IsReadOnly:
            size_param.Set(new_diameter)

        length_param = new_part.LookupParameter("Length")
        if length_param and not length_param.IsReadOnly:
            length_param.Set(sleeve_length)

        move_vec = pt - new_part.Origin
        ElementTransformUtils.MoveElement(doc, new_part.Id, move_vec)

        rotate_to_vector(doc, new_part, pt, XYZ.BasisX, pipe_dir.Multiply(-1))

        doc.Regenerate()

        align_top_to_point(doc, new_part, pt)

        placed_count += 1

        # --- DUPLICATE DETECTION (track newly placed)
        existing_sleeves.append(new_part)

if placed_count == 0:
    TaskDialog.Show("Info", "No sleeves placed.")
    t.RollBack()
else:
    t.Commit()