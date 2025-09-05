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

        # FUNCTION TO GET PARAMETER VALUE
        def get_parameter_value(element, parameterName):
            param = element.LookupParameter(parameterName)
            if param and param.HasValue:
                return param.AsDouble()
            else:
                return 0.0

        # Gets bottom elevation of selected pipe
        if RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
        else:
            PRTElevation = get_parameter_value(element, 'Bottom')

        # Gets servicename of selection
        parameters = element.LookupParameter('Fabrication Service')
        if parameters and parameters.HasValue:
            service_name = parameters.AsValueString()
        else:
            raise Exception("Fabrication Service parameter missing.")

        servicenamelist = []
        Config = FabricationConfiguration.GetFabricationConfiguration(doc)
        LoadedServices = Config.GetAllLoadedServices()

        for Item1 in LoadedServices:
            try:
                servicenamelist.append(Item1.Name)
            except:
                servicenamelist.append([])

        # Gets matching index of selected element service from the servicenamelist
        try:
            Servicenum = servicenamelist.index(service_name)
        except ValueError:
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

        # Instantiate and show the dialog
        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef)
        if form.ShowDialog() == DialogResult.OK:
            Selectedbutton = form.combobox_hanger.Text
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.Checked
            AtoS = form.checkbox_attach.Checked
            SelectedServiceName = form.combobox_service.Text
            
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
            except ValueError:
                print("Selected service '{}' not found.".format(SelectedServiceName))
                raise Exception("Selected service not found.")
            
            # Find the selected button
            button_found = False
            fab_btn = None
            for servicenum, service in enumerate(LoadedServices):
                if service.Name == SelectedServiceName:
                    palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
                    for palette_idx in range(palette_count):
                        button_count = service.GetButtonCount(palette_idx)
                        for btn_idx in range(button_count):
                            bt = service.GetButton(palette_idx, btn_idx)
                            if bt.Name == Selectedbutton:
                                fab_btn = bt
                                button_found = True
                                break
                        if button_found:
                            break
                    if button_found:
                        break
            
            if not button_found:
                print("'{}' not found in '{}'".format(Selectedbutton, SelectedServiceName))
                raise Exception("Hanger button not found.")
            
            # Write values to text file
            with open(filepath, 'w') as the_file:
                lines = [str(Selectedbutton) + '\n', str(distancefromend) + '\n', str(Spacing) + '\n',
                         SelectedServiceName + '\n', str(AtoS) + '\n', str(BOITrap)]
                the_file.writelines(lines)
            
            # Helper functions
            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    center = (bBox.Max + bBox.Min) / 2
                    return center
                else:
                    print("No bounding box found for element {}".format(ele_id))
                    return XYZ(0, 0, 0)
            
            def myround(x, multiple):
                return multiple * math.ceil(x / multiple)
            
            def get_diameter(pipe):
                param = pipe.LookupParameter("Outside Diameter")
                if param and param.HasValue:
                    return param.AsDouble()
                print("Outside Diameter parameter not found for pipe {}. Assuming 0.".format(pipe.Id))
                return 0.0
            
            def get_reference_level(hanger):
                level_id = hanger.LevelId
                level = doc.GetElement(level_id)
                return level
            
            def get_level_elevation(level):
                if level:
                    return level.Elevation
                return 0.0
            
            # Determine pipe direction from the first pipe's location curve
            curve = element.Location.Curve
            if not curve or not curve.IsBound:
                print("Invalid location curve for element {}".format(element.Id))
                raise Exception("Pipe must have a valid location curve.")
            
            # Get direction vector in XY plane (ignore Z for rotation on Z-axis)
            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize()  # Ensure Z is 0
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize()  # Perpendicular vector in XY plane
            
            # Compute midpoints of both pipes
            midpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    mid = (c.GetEndPoint(0) + c.GetEndPoint(1)) / 2
                    midpoints.append(mid)
                else:
                    print("Invalid location curve for pipe {}".format(pipe.Id))
                    raise Exception("Pipe must have a valid location curve.")
            
            # Compute center point between the two pipes in the perpendicular direction
            perp_projs = [mid.DotProduct(perp_vec) for mid in midpoints]
            min_perp = min(perp_projs)
            max_perp = max(perp_projs)
            center_perp = (min_perp + max_perp) / 2
            
            # Adjust for pipe diameter and insulation
            widths = []
            for pipe in selected_elements:
                thick = pipe.InsulationThickness if pipe.HasInsulation else 0
                half_size = get_diameter(pipe) / 2 + thick
                widths.append(half_size)
            
            # Compute effective width (distance between outermost edges)
            width = abs(max_perp - min_perp) + widths[0] + widths[1]
            
            # Compute projections along direction to find length
            endpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.append(c.GetEndPoint(0))
                    endpoints.append(c.GetEndPoint(1))
            
            projs_along = [p.DotProduct(dir_vec) for p in endpoints]
            min_along = min(projs_along)
            max_along = max(projs_along)
            length_along = max_along - min_along
            
            # Compute Z bounds for elevation
            combined_min_z = float('inf')
            combined_max_z = float('-inf')
            for pipe in selected_elements:
                pipe_bb = pipe.get_BoundingBox(curview)
                if pipe_bb:
                    thick = pipe.InsulationThickness if pipe.HasInsulation else 0
                    if pipe.HasInsulation and BOITrap:
                        combined_min_z = min(combined_min_z, pipe_bb.Min.Z - thick)
                        combined_max_z = max(combined_max_z, pipe_bb.Max.Z + thick)
                    else:
                        combined_min_z = min(combined_min_z, pipe_bb.Min.Z)
                        combined_max_z = max(combined_max_z, pipe_bb.Max.Z)
            
            center_z = combined_min_z if BOITrap else (combined_min_z + combined_max_z) / 2
            
            # Calculate number of hangers
            qtyofhgrs = int(math.ceil(length_along / Spacing))
            
            # Place hangers at default location (0,0,0)
            hangers = []
            t = Transaction(doc, 'Place Trapeze Hanger')
            t.Start()
            for hgr in range(qtyofhgrs):
                try:
                    hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                    if hanger:
                        hangers.append(hanger)
                    else:
                        print("Failed to create hanger {}".format(hgr + 1))
                except Exception as e:
                    print("Error creating hanger {}: {}".format(hgr + 1, str(e)))
            t.Commit()
            
            if not hangers:
                print("No hangers were created. Check fabrication service and button compatibility.")
                raise Exception("Hanger creation failed.")
            
# Move and modify hangers
            t = Transaction(doc, 'Modify Trapeze Hanger')
            t.Start()
            IncrementSpacing = distancefromend
            
            # Find the closest endpoint of the first selected pipe
            first_pipe = selected_elements[0]
            curve = first_pipe.Location.Curve
            if not curve or not curve.IsBound:
                print("Invalid location curve for first pipe {}".format(first_pipe.Id))
                raise Exception("First pipe must have a valid location curve.")
            first_endpoints = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
            ref_point = min(first_endpoints, key=lambda p: p.DistanceTo(XYZ(0, 0, 0)))
            # print("Reference point (nearest endpoint of first pipe): X={:.2f}, Y={:.2f}, Z={:.2f}".format(ref_point.X, ref_point.Y, ref_point.Z))
            
            # Compute midpoints, diameters, and insulation for both pipes
            endpoints = []
            midpoints = []
            thicknesses = []
            diameters = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.extend([c.GetEndPoint(0), c.GetEndPoint(1)])
                    mid = (c.GetEndPoint(0) + c.GetEndPoint(1)) / 2
                    midpoints.append(mid)
                    thick = pipe.InsulationThickness if hasattr(pipe, 'InsulationThickness') and pipe.HasInsulation else 0.0
                    thicknesses.append(thick)
                    diam = get_diameter(pipe)
                    diameters.append(diam)
                else:
                    print("Invalid location curve for pipe {}".format(pipe.Id))
                    raise Exception("Pipe must have a valid location curve.")
            
            # Debug: Log pipe endpoints, midpoints, diameters, and insulation
            # for i, ep in enumerate(endpoints):
                # print("Pipe endpoint {}: X={:.2f}, Y={:.2f}, Z={:.2f}".format(i, ep.X, ep.Y, ep.Z))
            # for i, mid in enumerate(midpoints):
                # print("Pipe midpoint {}: X={:.2f}, Y={:.2f}, Z={:.2f}".format(i, mid.X, mid.Y, mid.Z))
            # for i, (diam, thick) in enumerate(zip(diameters, thicknesses)):
                # print("Pipe {} diameter: {:.2f}, insulation thickness: {:.2f}".format(i, diam, thick))
            # print("Distance from end: {:.2f}, Spacing: {:.2f}".format(distancefromend, Spacing))
            
            # Recalculate direction and perpendicular vectors
            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize()  # Ensure Z=0
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize()  # Perpendicular in XY plane
            # print("Direction vector: X={:.2f}, Y={:.2f}, Z={:.2f}".format(dir_vec.X, dir_vec.Y, dir_vec.Z))
            # print("Perpendicular vector: X={:.2f}, Y={:.2f}, Z={:.2f}".format(perp_vec.X, perp_vec.Y, perp_vec.Z))
            
            # Recalculate center_perp using projections onto perp_vec
            perp_projs = []
            for i, mid in enumerate(midpoints):
                proj = (mid - ref_point).DotProduct(perp_vec)
                half_size = diameters[i] / 2 + thicknesses[i]
                perp_projs.append(proj - half_size)  # Min edge
                perp_projs.append(proj + half_size)  # Max edge
            center_perp = (min(perp_projs) + max(perp_projs)) / 2
            # print("Perpendicular projections (adjusted for diameter/insulation): {}".format(["{:.2f}".format(float(p)) for p in perp_projs]))
            # print("Center perpendicular: {:.2f}".format(float(center_perp)))
            
            # Recalculate projections along the pipe direction for first pipe
            projs_along = [(p - ref_point).DotProduct(dir_vec) for p in first_endpoints]
            min_along = min(projs_along)
            # print("Min along projection (first pipe): {:.2f}".format(float(min_along)))
            
            # Set center_z to pipe elevation
            center_z = midpoints[0].Z  # Use pipe elevation
            if BOITrap:
                center_z = PRTElevation
            # print("Center Z: {:.2f}, BOITrap: {}, PRTElevation: {:.2f}".format(center_z, BOITrap, PRTElevation))
            
            for idx, hanger in enumerate(hangers):
                # Set dimensions first
                newwidth = myround(width * 12, 2) / 12
                for dim in hanger.GetDimensions():
                    dim_name = dim.Name
                    try:
                        if dim_name == "Width":
                            hanger.SetDimensionValue(dim, newwidth)
                        if dim_name == "Bearer Extn":
                            hanger.SetDimensionValue(dim, 0.25)
                    except Exception as e:
                        print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(e)))
                
                # Rotate hanger to align with pipe direction
                center = GetCenterPoint(hanger.Id)
                z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
                angle_rad = math.atan2(dir_vec.Y, dir_vec.X)
                try:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
                except Exception as e:
                    print("Error rotating hanger {}: {}".format(hanger.Id, str(e)))
                
                # Compute target position starting from nearest endpoint
                along = min_along + IncrementSpacing
                pos = ref_point + dir_vec * along + perp_vec * center_perp + XYZ(0, 0, center_z)
                IncrementSpacing += Spacing
                # print("Hanger {} target position: X={:.2f}, Y={:.2f}, Z={:.2f}".format(hanger.Id, pos.X, pos.Y, pos.Z))
                
                # Move hanger
                center = GetCenterPoint(hanger.Id)
                # print("Hanger {} current center: X={:.2f}, Y={:.2f}, Z={:.2f}".format(hanger.Id, center.X, center.Y, center.Z))
                translation = pos - center
                try:
                    ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                except Exception as e:
                    print("Error moving hanger {}: {}".format(hanger.Id, str(e)))
                
                # Set offset
                reference_level = get_reference_level(hanger)
                elevation = get_level_elevation(reference_level)
                try:
                    offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                    if BOITrap:
                        offset_param.Set(PRTElevation)
                    else:
                        offset_value = center_z - elevation
                        offset_param.Set(offset_value)
                except Exception as e:
                    print("Error setting offset for hanger {}: {}".format(hanger.Id, str(e)))
                
                if AtoS:
                    try:
                        hanger.GetRodInfo().AttachToStructure()
                    except Exception as e:
                        print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))
            
            t.Commit()
        
        else:
            print("Dialog cancelled.")
    
    else:
        print("Coming Soon... \nYou will be able to place a trapeze on a ptrap")

except Exception as e:
    print("Script error: {}".format(str(e)))