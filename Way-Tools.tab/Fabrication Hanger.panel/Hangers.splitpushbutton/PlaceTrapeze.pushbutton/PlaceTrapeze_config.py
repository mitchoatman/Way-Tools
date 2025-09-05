#------------------------------------------------------------------------------------IMPORTS
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, \
                                FabricationService, XYZ, ElementTransformUtils, BoundingBoxXYZ, Transform, Line
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")

from System.Windows.Forms import *
from System.Drawing import Point, Size, Font
from System import Array
#------------------------------------------------------------------------------------DEFINE SOME VARIABLES EASY USE
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

#------------------------------------------------------------------------------------SELECTING ELEMENTS
try:
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    if doc.GetElement(selected_element.ElementId).ItemCustomId != 916:
        selected_element1 = uidoc.Selection.PickObject(ObjectType.Element, 'Select OPPOSITE OUTSIDE Pipe')
        element = doc.GetElement(selected_element.ElementId)
        element1 = doc.GetElement(selected_element1.ElementId)
        selected_elements = [element, element1]

        level_id = element.LevelId
        print("Selected elements: {} and {}".format(element.Id, element1.Id))
        print("Level ID: {}".format(level_id))

        # FUNCTION TO GET PARAMETER VALUE
        def get_parameter_value(element, parameterName):
            param = element.LookupParameter(parameterName)
            if param and param.HasValue:
                return param.AsDouble()
            else:
                print("Parameter '{}' not found or has no value for element {}".format(parameterName, element.Id))
                return 0.0

        # Gets bottom elevation of selected pipe
        if RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
        else:
            PRTElevation = get_parameter_value(element, 'Bottom')
        print("Bottom Elevation: {}".format(PRTElevation))

        # Gets servicename of selection
        parameters = element.LookupParameter('Fabrication Service')
        if parameters and parameters.HasValue:
            service_name = parameters.AsValueString()
        else:
            print("Fabrication Service parameter not found for element {}".format(element.Id))
            raise Exception("Fabrication Service parameter missing.")
        print("Selected Service: {}".format(service_name))

        servicenamelist = []
        Config = FabricationConfiguration.GetFabricationConfiguration(doc)
        LoadedServices = Config.GetAllLoadedServices()

        for Item1 in LoadedServices:
            try:
                servicenamelist.append(Item1.Name)
            except:
                servicenamelist.append([])
        print("Available Services: {}".format(servicenamelist))

        # Gets matching index of selected element service from the servicenamelist
        try:
            Servicenum = servicenamelist.index(service_name)
            print("Service Index: {}".format(Servicenum))
        except ValueError:
            print("Service '{}' not found in loaded services.".format(service_name))
            raise Exception("Selected service not found.")

        # Find all hanger buttons across all palettes/groups
        buttonnames = []
        unique_hangers = set()

        for service_idx, service in enumerate(LoadedServices):
            palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
            for palette_idx in range(palette_count):
                buttoncount = service.GetButtonCount(palette_idx)
                for btn_idx in range(buttoncount):
                    bt = service.GetButton(palette_idx, btn_idx)
                    if bt.IsAHanger and bt.Name not in unique_hangers:
                        unique_hangers.add(bt.Name)
                        buttonnames.append(bt.Name)
        print("Available Hanger Buttons: {}".format(buttonnames))

        folder_name = "c:\\Temp"
        filepath = os.path.join(folder_name, 'Ribbon_PlaceTrapeze.txt')
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        if not os.path.exists(filepath):
            with open(filepath, 'w') as the_file:
                lines = ['1.625 Single Strut Trapeze\n', '1.0\n', '8.0\n', 'PLUMBING: DOMESTIC COLD WATER\n', 'True\n', 'True']
                the_file.writelines(lines)

        # Read text file for stored values
        with open(filepath, 'r') as file:
            lines = [line.rstrip() for line in file.readlines()]

        if len(lines) < 6:
            with open(filepath, 'w') as the_file:
                lines = ['1.625 Single Strut Trapeze\n', '1.0\n', '8.0\n', 'PLUMBING: DOMESTIC COLD WATER\n', 'True\n', 'True']
                the_file.writelines(lines)

        with open(filepath, 'r') as file:
            lines = [line.rstrip() for line in file.readlines()]
        print("Settings from file: {}".format(lines))

        checkboxdef = lines[4] != 'False'
        checkboxdefBOI = lines[5] != 'False'

        #-----------------------------------------------------------

        class HangerSpacingDialog(Form):
            def __init__(self, buttonnames, lines, checkboxdefBOI, checkboxdef):
                self.Text = "Hanger and Spacing"
                self.Size = Size(350, 360)
                self.StartPosition = FormStartPosition.CenterScreen
                self.FormBorderStyle = FormBorderStyle.FixedDialog

                # Choose Hanger Label
                label_hanger = Label()
                label_hanger.Text = "Choose Hanger:"
                label_hanger.Location = Point(10, 10)
                label_hanger.Size = Size(300, 20)
                label_hanger.Font = Font("Arial", 10)
                self.Controls.Add(label_hanger)

                # Hanger ComboBox
                self.combobox_hanger = ComboBox()
                self.combobox_hanger.Location = Point(10, 31)
                self.combobox_hanger.Size = Size(300, 20)
                self.combobox_hanger.DropDownStyle = ComboBoxStyle.DropDownList
                self.combobox_hanger.Items.AddRange(Array[object](buttonnames))
                if lines[0] in buttonnames:
                    self.combobox_hanger.SelectedItem = lines[0]
                self.Controls.Add(self.combobox_hanger)

                # Distance from End Label
                label_end_dist = Label()
                label_end_dist.Text = "Distance from End (Ft):"
                label_end_dist.Font = Font("Arial", 10)
                label_end_dist.Location = Point(10, 60)
                label_end_dist.Size = Size(300, 20)
                self.Controls.Add(label_end_dist)

                # Distance from End TextBox
                self.textbox_end_dist = TextBox()
                self.textbox_end_dist.Location = Point(10, 80)
                self.textbox_end_dist.Text = lines[1]
                self.Controls.Add(self.textbox_end_dist)

                # Hanger Spacing Label
                label_spacing = Label()
                label_spacing.Text = "Hanger Spacing (Ft):"
                label_spacing.Font = Font("Arial", 10)
                label_spacing.Location = Point(10, 110)
                label_spacing.Size = Size(300, 20)
                self.Controls.Add(label_spacing)

                # Hanger Spacing TextBox
                self.textbox_spacing = TextBox()
                self.textbox_spacing.Location = Point(10, 130)
                self.textbox_spacing.Text = lines[2]
                self.Controls.Add(self.textbox_spacing)

                # Align Trapeze CheckBox
                self.checkbox_boi = CheckBox()
                self.checkbox_boi.Text = "Align Trapeze to Bottom of Insulation"
                self.checkbox_boi.Font = Font("Arial", 10)
                self.checkbox_boi.Location = Point(10, 160)
                self.checkbox_boi.Size = Size(300, 20)
                self.checkbox_boi.Checked = checkboxdefBOI
                self.Controls.Add(self.checkbox_boi)

                # Attach to Structure CheckBox
                self.checkbox_attach = CheckBox()
                self.checkbox_attach.Text = "Attach to Structure"
                self.checkbox_attach.Font = Font("Arial", 10)
                self.checkbox_attach.Location = Point(10, 190)
                self.checkbox_attach.Size = Size(300, 20)
                self.checkbox_attach.Checked = checkboxdef
                self.Controls.Add(self.checkbox_attach)

                # Choose Service Label
                label_service = Label()
                label_service.Text = "Choose Service to Draw Hanger on:"
                label_service.Location = Point(10, 220)
                label_service.Size = Size(300, 20)
                label_service.Font = Font("Arial", 10)
                self.Controls.Add(label_service)

                # Service ComboBox
                self.combobox_service = ComboBox()
                self.combobox_service.Location = Point(10, 240)
                self.combobox_service.Size = Size(300, 20)
                self.combobox_service.DropDownStyle = ComboBoxStyle.DropDownList
                self.combobox_service.Items.AddRange(Array[object](servicenamelist))
                if lines[3] in servicenamelist:
                    self.combobox_service.SelectedItem = lines[3]
                self.Controls.Add(self.combobox_service)

                # OK Button
                self.button_ok = Button()
                self.button_ok.Text = "OK"
                self.button_ok.Font = Font("Arial", 10)
                self.button_ok.Location = Point((self.Width // 2) - 50, 280)
                self.button_ok.Click += self.ok_button_clicked
                self.Controls.Add(self.button_ok)

            def ok_button_clicked(self, sender, event):
                self.DialogResult = DialogResult.OK
                self.Close()

        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef)
        if form.ShowDialog() == DialogResult.OK:
            Selectedbutton = form.combobox_hanger.Text
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.Checked
            AtoS = form.checkbox_attach.Checked
            SelectedServiceName = form.combobox_service.Text
            print("Dialog Selections: Hanger='{}', Distance='{}', Spacing='{}', BOI={}, Attach={}, Service='{}'".format(
                Selectedbutton, distancefromend, Spacing, BOITrap, AtoS, SelectedServiceName))

            # Validate numeric inputs
            try:
                distancefromend = float(distancefromend)
                Spacing = float(Spacing)
            except ValueError:
                print("Invalid input: Distance from End or Spacing must be numeric.")
                raise Exception("Invalid numeric input.")

            # Gets matching index of selected service
            try:
                Servicenum = servicenamelist.index(SelectedServiceName)
                print("Selected Service Index: {}".format(Servicenum))
            except ValueError:
                print("Selected service '{}' not found.".format(SelectedServiceName))
                raise Exception("Selected service not found.")

            # Find the selected button and its service
            button_found = False
            fab_btn = None
            fab_service = None
            for servicenum, service in enumerate(LoadedServices):
                if service.Name == SelectedServiceName:
                    fab_service = service
                    palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
                    for palette_idx in range(palette_count):
                        button_count = service.GetButtonCount(palette_idx)
                        for btn_idx in range(button_count):
                            bt = service.GetButton(palette_idx, btn_idx)
                            if bt.Name == Selectedbutton:
                                fab_btn = bt
                                button_found = True
                                print("Found button '{}' in service '{}' at palette {}, button {}".format(
                                    Selectedbutton, SelectedServiceName, palette_idx, btn_idx))
                                # Check condition count
                                condition_count = bt.ConditionCount
                                print("Button '{}' has {} conditions".format(Selectedbutton, condition_count))
                                break
                        if button_found:
                            break
                    if button_found:
                        break

            if not button_found:
                print("Button '{}' not found in service '{}'".format(Selectedbutton, SelectedServiceName))
                raise Exception("Hanger button not found.")

            # Determine valid condition index
            condition_index = 0  # Default to 0
            if fab_btn.ConditionCount > 0:
                condition_index = 0  # Use the first condition
                print("Using condition index {} for button '{}'".format(condition_index, Selectedbutton))
            else:
                print("No valid conditions available for button '{}'. Attempting with index 0.".format(Selectedbutton))

            # Write values to text file
            with open(filepath, 'w') as the_file:
                lines = [Selectedbutton + '\n', str(distancefromend) + '\n', str(Spacing) + '\n',
                         SelectedServiceName + '\n', str(AtoS) + '\n', str(BOITrap)]
                the_file.writelines(lines)
            print("Settings saved to file.")

            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    center = (bBox.Max + bBox.Min) / 2
                    print("Center point for element {}: {}".format(ele_id, center))
                    return center
                else:
                    print("Bounding box not found for element {}".format(ele_id))
                    return XYZ(0, 0, 0)

            def myround(x, multiple):
                return multiple * math.ceil(x / multiple)

            # Get diameter from dimensions
            def get_diameter(pipe):
                for dim in pipe.GetDimensions():
                    dim_name = dim.Name.lower()
                    if dim_name in ["diameter", "od", "outside diameter"]:
                        value = pipe.GetDimensionValue(dim)
                        print("Diameter for pipe {}: {}".format(pipe.Id, value))
                        return value
                print("No diameter dimension found for pipe {}. Assuming 0.".format(pipe.Id))
                return 0.0

            # Compute direction and perpendicular vectors
            curve = element.Location.Curve
            if not curve or not curve.IsBound:
                print("Invalid location curve for element {}".format(element.Id))
                raise Exception("Pipe must have a valid location curve.")
            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            print("Direction vector: {}".format(dir_vec))

            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize()
            print("Perpendicular vector: {}".format(perp_vec))

            # Collect endpoints
            endpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.append(c.GetEndPoint(0))
                    endpoints.append(c.GetEndPoint(1))
                else:
                    print("Invalid location curve for pipe {}".format(pipe.Id))
                    raise Exception("Pipe must have a valid location curve.")
            print("Endpoints: {}".format(endpoints))

            # Compute projections along direction
            projs_along = [p.DotProduct(dir_vec) for p in endpoints]
            min_along = min(projs_along)
            max_along = max(projs_along)
            length_along = max_along - min_along
            print("Along projections: min={}, max={}, length={}".format(min_along, max_along, length_along))

            # Compute effective width and center in perpendicular direction
            min_perp = float('inf')
            max_perp = float('-inf')
            for pipe in selected_elements:
                mid = (pipe.Location.Curve.GetEndPoint(0) + pipe.Location.Curve.GetEndPoint(1)) / 2
                center_perp = mid.DotProduct(perp_vec)
                thick = pipe.InsulationThickness if pipe.HasInsulation else 0
                print("Insulation thickness for pipe {}: {}".format(pipe.Id, thick))
                half_size = get_diameter(pipe) / 2 + thick
                pipe_min_perp = center_perp - half_size
                pipe_max_perp = center_perp + half_size
                min_perp = min(min_perp, pipe_min_perp)
                max_perp = max(max_perp, pipe_max_perp)
                print("Pipe {}: centerline projection={}, half-size={}, extents=[{}, {}]".format(
                    pipe.Id, center_perp, half_size, pipe_min_perp, pipe_max_perp))
            width = max_perp - min_perp
            center_perp = (min_perp + max_perp) / 2
            print("Width: {}, Center perpendicular: {}".format(width, center_perp))

            # Compute combined bounding box for Z references
            combined_min = None
            combined_max = None
            for pipe in selected_elements:
                pipe_bounding_box = pipe.get_BoundingBox(curview)
                if not pipe_bounding_box:
                    print("Bounding box not found for pipe {}".format(pipe.Id))
                    continue
                if combined_min is None:
                    combined_min = pipe_bounding_box.Min
                    combined_max = pipe_bounding_box.Max
                else:
                    combined_min = XYZ(min(combined_min.X, pipe_bounding_box.Min.X),
                                       min(combined_min.Y, pipe_bounding_box.Min.Y),
                                       min(combined_min.Z, pipe_bounding_box.Min.Z))
                    combined_max = XYZ(max(combined_max.X, pipe_bounding_box.Max.X),
                                       max(combined_max.Y, pipe_bounding_box.Max.Y),
                                       max(combined_max.Z, pipe_bounding_box.Max.Z))
            if combined_min is None or combined_max is None:
                print("No valid bounding boxes found for selected pipes.")
                raise Exception("Invalid bounding boxes.")
            combined_bounding_box = BoundingBoxXYZ()
            combined_bounding_box.Min = combined_min
            combined_bounding_box.Max = combined_max
            combined_bounding_box_Center = (combined_max + combined_min) / 2
            print("Combined BB Center: {}".format(combined_bounding_box_Center))

            # Get reference level
            def get_reference_level(hanger):
                level_id = hanger.LevelId
                level = doc.GetElement(level_id)
                return level

            def get_level_elevation(level):
                if level:
                    return level.Elevation
                return 0.0

            # Setup spacing
            qtyofhgrs = int(math.ceil(length_along / Spacing))
            print("Number of hangers: {}".format(qtyofhgrs))

            # Place hangers at default location (0,0,0)
            hangers = []
            t = Transaction(doc, 'Place Trapeze Hanger')
            t.Start()
            for hgr in range(qtyofhgrs):
                try:
                    hanger = FabricationPart.CreateHanger(doc, fab_btn, condition_index, level_id)
                    if hanger:
                        hangers.append(hanger)
                        print("Created hanger {} with ID {} at default location (0,0,0)".format(hgr + 1, hanger.Id))
                    else:
                        print("Failed to create hanger {}".format(hgr + 1))
                except Exception as e:
                    print("Error creating hanger {}: {}".format(hgr + 1, str(e)))
            t.Commit()
            print("Hanger creation transaction committed. Hangers created: {}".format(len(hangers)))

            if not hangers:
                print("No hangers were created. Check fabrication service and button compatibility.")
                raise Exception("Hanger creation failed.")

            # Move and modify hangers
            t = Transaction(doc, 'Modify Trapeze Hanger')
            t.Start()
            IncrementSpacing = distancefromend
            for idx, hanger in enumerate(hangers):
                # Compute target position
                along = min_along + IncrementSpacing
                pos = dir_vec * along + perp_vec * center_perp + XYZ(0, 0, PRTElevation if BOITrap else combined_bounding_box_Center.Z)
                print("Hanger {} target position: {}".format(idx + 1, pos))
                IncrementSpacing += Spacing

                # Move hanger from default location
                center = GetCenterPoint(hanger.Id)
                translation = pos - center
                print("Translation vector for hanger {}: {}".format(hanger.Id, translation))
                try:
                    ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                    print("Moved hanger {} to {}".format(hanger.Id, pos))
                except Exception as e:
                    print("Error moving hanger {}: {}".format(hanger.Id, str(e)))

                # Rotate hanger
                z_axis = Line.CreateBound(pos, pos + XYZ(0, 0, 1))
                angle_rad = math.atan2(dir_vec.Y, dir_vec.X)
                try:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
                    print("Rotated hanger {} by {} radians".format(hanger.Id, angle_rad))
                except Exception as e:
                    print("Error rotating hanger {}: {}".format(hanger.Id, str(e)))

                # Set dimensions
                newwidth = myround(width * 12, 2) / 12
                for dim in hanger.GetDimensions():
                    dim_name = dim.Name
                    try:
                        if dim_name == "Width":
                            hanger.SetDimensionValue(dim, newwidth)
                            print("Set Width for hanger {} to {}".format(hanger.Id, newwidth))
                        if dim_name == "Bearer Extn":
                            hanger.SetDimensionValue(dim, 0.25)
                            print("Set Bearer Extn for hanger {} to 0.25".format(hanger.Id))
                    except Exception as e:
                        print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(e)))

                # Set offset
                reference_level = get_reference_level(hanger)
                elevation = get_level_elevation(reference_level)
                try:
                    offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                    if BOITrap:
                        offset_param.Set(PRTElevation)
                        print("Set offset for hanger {} to PRTElevation {}".format(hanger.Id, PRTElevation))
                    else:
                        offset_value = combined_bounding_box.Min.Z - elevation
                        offset_param.Set(offset_value)
                        print("Set offset for hanger {} to {}".format(hanger.Id, offset_value))
                except Exception as e:
                    print("Error setting offset for hanger {}: {}".format(hanger.Id, str(e)))

                if AtoS:
                    try:
                        hanger.GetRodInfo().AttachToStructure()
                        print("Attached hanger {} to structure".format(hanger.Id))
                    except Exception as e:
                        print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))

            t.Commit()
            print("Hanger modification transaction committed.")
        else:
            print("Dialog cancelled.")
    else:
        print("Coming Soon... \nYou will be able to place a trapeze on a ptrap")
except Exception as e:
    print("Script error: {}".format(str(e)))