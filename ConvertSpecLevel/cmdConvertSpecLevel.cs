using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdConvertSpecLevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            #region Form

            // launch the form
            frmConvertSpecLevel curForm = new frmConvertSpecLevel();
            curForm.Topmost = true;

            curForm.ShowDialog();

            // check if user clicked Cancel
            if (curForm.DialogResult != true)
            {
                return Result.Cancelled;
            }

            // get user input from form
            // get user input from the form
            string selectedClient = curForm.GetSelectedClient();
            string selectedSpecLevel = curForm.GetSelectedSpecLevel();
            string selectedMWCabHeight = curForm.GetSelectedMWCabHeight();

            #endregion

            #region Transaction Group

            // create a transaction group
            using (TransactionGroup transgroup = new TransactionGroup(curDoc, "Convert Spec Level"))
            {
                // start the transaction group
                transgroup.Start();

                #region Floor Finish Updates

                // get a first floor annotation view & set it as the active view
                View curView = Utils.GetViewByNameContainsAndAssociatedLevel(curDoc, "Annotation", "First Floor");

                // check for null
                if (curView != null)
                {
                    uidoc.ActiveView = curView;
                }
                else
                {
                    // if null notify the user
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No view found with name containing 'Annotation' and associated level 'First Floor'");
                }

                // create a transaction for the flooring update
                using (Transaction t = new Transaction(curDoc, "Update Floor Finish"))
                {
                    // start the first transaction
                    t.Start();

                    // change the flooring for the specified rooms per the selected spec level
                    List<string> listUpdatedRooms = UpdateFloorFinishInActiveView(curDoc, selectedSpecLevel);

                    // manage the floor material breaks
                    ManageFloorMaterialBreaksInActiveView(curDoc, selectedSpecLevel);

                    // create a list of the rooms updated
                    string listRooms;
                    if (listUpdatedRooms.Count == 1)
                    {
                        listRooms = listUpdatedRooms[0];
                    }
                    else if (listUpdatedRooms.Count == 2)
                    {
                        listRooms = $"{listUpdatedRooms[0]} and {listUpdatedRooms[1]}";
                    }
                    else
                    {
                        listRooms = string.Join(", ", listUpdatedRooms.Take(listUpdatedRooms.Count - 1)) + $", and {listUpdatedRooms.Last()}";
                    }

                    // notify the user
                    Utils.TaskDialogInformation("Complete", "Spec Conversion", $"Flooring was changed at {listRooms} per the specified spec level.");

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Door Updates

                // get the door schedule & set it as the active view
                View curSched = Utils.GetScheduleByNameContains(curDoc, "Door Schedule");

                if (curSched != null)
                {
                    uidoc.ActiveView = curSched;
                }
                else
                {
                    // if not found alert the user
                    Utils.TaskDialogError("Error", "Spec Conversion", "No Door Schedule found.");
                }

                // create transaction for door updates
                using (Transaction t = new Transaction(curDoc, "Update Doors"))
                {
                    // start the second transaction
                    t.Start();

                    // update front door type
                    UpdateFrontDoor(curDoc, selectedSpecLevel);

                    // update rear door type
                    UpdateRearDoor(curDoc, selectedSpecLevel);

                    // notify the user
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "The front and rear doors were replaced per the specified spec level.");
                  
                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Cabinet Updates

                // get the Interior Elevations sheet & set it as the active view
                ViewSheet sheetIntr = Utils.GetSheetByName(curDoc, "Interior Elevations");

                if (sheetIntr != null)
                {
                    uidoc.ActiveView = sheetIntr;
                }
                else
                {
                    // if not found alert the user
                    Utils.TaskDialogError("Error", "Spec Conversion", "No Interior Elevations sheet found.");
                }

                // create transaction for cabinet updates
                using (Transaction t = new Transaction(curDoc, "Update Cabinets"))
                {
                    // start the third transaction
                    t.Start();

                    // revise the upper cabinets
                    if (selectedSpecLevel == "Complete Home")
                    {
                        // lower the upper cabinets to 36"
                        ReplaceWallCabinets(curDoc, "36");
                    }
                    else
                    {
                        // raise the upper cabinets to 42"
                        ReplaceWallCabinets(curDoc, "42");
                    }

                    // revise the MW cabinet
                    ReplaceMWCabinet(curDoc, selectedMWCabHeight);

                    //// add/remove the Ref Sp cabinet
                    //if (selectedSpecLevel == "Complete Home" && selectedCabinet != null)
                    //{
                    //    curDoc.Delete(((Element)selectedCabinet).Id);
                    //}
                    //else
                    //{
                    //    AddRefSpCabinet(curDoc, uidoc, selectedRefSpWall, selectedRefSp);
                    //}

                    // raise/lower the backsplash height
                    UpdateBacksplashHeight(curDoc, selectedSpecLevel);

                    UpdateBacksplashNote(curDoc, uidoc, selectedSpecLevel);

                    // notify the user
                    // upper cabinets were revised per the selected spec level
                    // Ref Sp cabinet was added/removed per the selected spec level
                    // backsplash height was raised/lowered per the selected spec level

                    // commit the transaction
                    t.Commit();
                }



                #endregion

                #region General Electrical Setup


                #endregion

                #region First Floor Electrical Updates


                #endregion

                #region Second Floor Electrical Updates


                #endregion


                // commit the transaction group
                transgroup.Assimilate();
            }

            #endregion

            return Result.Succeeded;
        }    

        #region Finish Floor Methods

        /// <summary>
        /// Updates the Floor Finish parameter for specified room types in the active view based on the selected specification level.
        /// </summary>
        /// <remarks>
        /// This method updates the following room types: Master Bedroom.
        /// Complete Home sets floor finish to "Carpet", Complete Home Plus sets it to "HS".
        /// Only processes rooms that are visible in the current active view.
        /// </remarks>
        internal static List<string> UpdateFloorFinishInActiveView(Document curDoc, string selectedSpecLevel)
        {
            // get the active view from the document
            View activeView = curDoc.ActiveView;

            // create a list of rooms to update
            List<string> m_RoomsToUpdateFloorFinish = new List<string>
            {
                "Master Bedroom"
            };

            // get the room element of the rooms to update
            List<Room> m_RoomstoUpdate = GetRoomsByNameContainsInActiveView(curDoc, m_RoomsToUpdateFloorFinish);

            // create the switch statement to determine the floor finish based on the spec level
            string floorFinish = selectedSpecLevel switch
            {
                "Complete Home" => "Carpet",
                "Complete Home Plus" => "HS",
                _ => null
            };

            // check if the floor finish is null
            if (floorFinish == null)
            {
                TaskDialog.Show("Error", "Invalid Spec Level selected.");
                return new List<string>();
            }

            // create an empty list to hold the room names fuond in the active view
            List<string> m_updatedRoomNames = new List<string>();

            // loop through the rooms to update
            foreach (Room curRoom in m_RoomstoUpdate)
            {
                // get the floor finish parameter
                Parameter paramFloorFinish = curRoom.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR);
                // check if the parameter is not null and has a value
                if (paramFloorFinish != null && !paramFloorFinish.IsReadOnly)
                {
                    // set the value of the floor finish parameter to the new value
                    paramFloorFinish.Set(floorFinish);

                    // add the room name to the ist
                    Parameter paramRoomName = curRoom.get_Parameter(BuiltInParameter.ROOM_NAME);
                    m_updatedRoomNames.Add(paramRoomName.AsString());
                }
            }

            // return the updated room names list
            return m_updatedRoomNames;
        }

        /// <summary>
        /// Gets all rooms in the active view whose names contain any of the specified strings.
        /// </summary>
        /// <returns>
        /// A list of Room elements in the active view whose names contain any of the specified strings.
        /// </returns>
        private static List<Room> GetRoomsByNameContainsInActiveView(Document curDoc, List<string> RoomsToUpdate)
        {
            // get the active view from the document
            View activeView = curDoc.ActiveView;

            // create a lsit to hold the matching rooms
            List<Room> m_matchingRooms = new List<Room>();

            // get all the rooms in the active view
            FilteredElementCollector m_colRooms = new FilteredElementCollector(curDoc, activeView.Id)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            // loop through the rooms and check if the room name contains the string in the list
            foreach (Room curRoom in m_colRooms)
            {
                // check if the room name contains any of the strings in the list
                if (RoomsToUpdate.Any(roomName => curRoom.Name.IndexOf(roomName, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    // if so, add the room to the matching rooms list
                    m_matchingRooms.Add(curRoom);
                }
            }

            // return the matching rooms
            return m_matchingRooms;
        }

        private void ManageFloorMaterialBreaksInActiveView(Document curDoc, string selectedSpecLevel)
        {
            if (selectedSpecLevel == "Complete Home Plus")
            {
                // remove the floor material breaks
                RemoveFloorMaterialBreaks(curDoc);
            }
            else if (selectedSpecLevel == "Complete Home")
            {
                // add floor material breaks
                AddFloorMaterialBreaks(curDoc);
            }
        }

        private void RemoveFloorMaterialBreaks(Document curDoc)
        {
            // get all the doors in the active view
            List<FamilyInstance> allDoorsInView = Utils.GetAllDoorsInActiveView(curDoc);

            // loop through each door
            foreach (FamilyInstance curDoor in allDoorsInView)
            {
                // get ToRoom & FromRoom values
                Room toRoom = curDoor.ToRoom;
                Room fromRoom = curDoor.FromRoom;

                // check for null
                if (toRoom == null || fromRoom == null)
                {
                    continue;
                }

                // check for match
                if (toRoom.LookupParameter("Floor Finish").AsString() == (fromRoom.LookupParameter("Floor Finish").AsString()))
                {
                    // if material matches, get the door's location point
                    LocationPoint doorLoc = curDoor.Location as LocationPoint;
                    XYZ doorPoint = doorLoc.Point;

                    // check for null
                    if (doorLoc == null || doorPoint == null)
                    {
                        continue;
                    }

                    // get all the floor material breaks in the active view
                    List<FamilyInstance> m_allMaterialBreaks = new FilteredElementCollector(curDoc, curDoc.ActiveView.Id)
                        .OfClass(typeof(FamilyInstance))
                        .OfCategory(BuiltInCategory.OST_FurnitureSystems)
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol.Family.Name == "Floor Material")
                        .ToList();

                    foreach (FamilyInstance curBreak in m_allMaterialBreaks)
                    {
                        // get the material break location point
                        LocationCurve breakLoc = curBreak.Location as LocationCurve;
                        Curve breakCurve = breakLoc.Curve;
                        XYZ breakPoint = breakCurve.Evaluate(0.5, true);

                        // check for null
                        if (breakLoc == null || breakPoint == null)
                        {
                            continue;
                        }

                        // calculate the distance between doorPoint & breakPoint
                        double distance = doorPoint.DistanceTo(breakPoint);

                        // check if distance is 18" or less
                        if (distance <= 1.5)
                        {
                            // if so, delete the material break
                            curDoc.Delete(curBreak.Id);
                        }
                    }
                }
            }
        }

        private void AddFloorMaterialBreaks(Document curDoc)
        {
            // get all the doors in the active view
            List<FamilyInstance> allDoorsInView = Utils.GetAllDoorsInActiveView(curDoc);

            // loop through each door
            foreach (FamilyInstance curDoor in allDoorsInView)
            {
                // get ToRoom & FromRoom values
                Room toRoom = curDoor.ToRoom;
                Room fromRoom = curDoor.FromRoom;

                // check for null
                if (toRoom == null || fromRoom == null)
                {
                    continue;
                }

                // check for non-match
                if (toRoom.LookupParameter("Floor Finish").AsString() != (fromRoom.LookupParameter("Floor Finish").AsString()))
                {
                    // if materials do not match, get the door's location point
                    LocationPoint doorLoc = curDoor.Location as LocationPoint;
                    XYZ doorPoint = doorLoc.Point;

                    // check for null
                    if (doorLoc == null || doorPoint == null)
                    {
                        continue;
                    }

                    // get all the floor material breaks in the active view
                    List<FamilyInstance> m_allMaterialBreaks = new FilteredElementCollector(curDoc, curDoc.ActiveView.Id)
                        .OfClass(typeof(FamilyInstance))
                        .OfCategory(BuiltInCategory.OST_FurnitureSystems)
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol.Family.Name == "Floor Material")
                        .ToList();

                    // declare a boolean flag
                    bool breakExists = false;

                    foreach (FamilyInstance curBreak in m_allMaterialBreaks)
                    {
                        // get the material break location point
                        LocationCurve breakLoc = curBreak.Location as LocationCurve;
                        Curve breakCurve = breakLoc.Curve;
                        XYZ breakPoint = breakCurve.Evaluate(0.5, true);

                        // check for null
                        if (breakLoc == null || breakPoint == null)
                        {
                            continue;
                        }

                        // calculate the distance between doorPoint & breakPoint
                        double distance = doorPoint.DistanceTo(breakPoint);

                        // check if distance is 18" or less
                        if (distance <= 1.5)
                        {
                            // skip this door & go to the next
                            breakExists = true;
                            break;
                        }
                    }

                    if (!breakExists)
                    {
                        // load the current family into the project
                        Family materialFamily = Utils.LoadFamilyFromLibrary(curDoc,
                            @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Annotation", "LD_AN_Floor_Material");

                        // get the type from the family
                        FamilySymbol materialSymbol = Utils.GetFamilySymbolByName(curDoc, "LD_AN_Floor_Material", "Type 1");

                        // activate the family & symbol
                        if (!materialSymbol.IsActive)
                        {
                            materialSymbol.Activate();
                        }

                        // get the door width
                        double drWidthParam = curDoor.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble();

                        // get the wall that hosts the door
                        Wall drWall = curDoor.Host as Wall;

                        // null check for host
                        if (drWall == null)
                        {
                            continue;
                        }

                        // get wall location curve
                        LocationCurve wallLoc = drWall.Location as LocationCurve;

                        // check for null curve
                        if (wallLoc == null)
                        {
                            continue;
                        }

                        // get the curve property from the wall
                        Curve wallCurve = wallLoc.Curve;

                        // get the work plane
                        Level workPlane = curDoc.GetElement(curDoc.ActiveView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsElementId()) as Level;

                        // cast the curve to a line
                        Line wallLine = wallCurve as Line;
                        if (wallLine != null)
                        {
                            // get the wall direction
                            XYZ wallDirection = wallLine.Direction;

                            //// get the center of the door opening
                            //XYZ doorCenter = wallLine.Evaluate(0.5, true);

                            //// create a perpendicular vector
                            //XYZ perpendicular = new XYZ(-wallDirection.Y, wallDirection.X, 0);

                            // create start point
                            XYZ startPoint = doorPoint - (wallDirection * (drWidthParam / 2));

                            // create end point  
                            XYZ endPoint = doorPoint + (wallDirection * (drWidthParam / 2));

                            // create a line to place the break
                            Line breakLine = Line.CreateBound(startPoint, endPoint);

                            // create a new Floor Material instance
                            FamilyInstance newBreak = curDoc.Create.NewFamilyInstance(breakLine, materialSymbol, workPlane, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                            // set the length of the break
                            //newBreak.LookupParameter("Length").Set(drWidthParam);

                            // set the floor finish for the FromRoom
                            string fromRmFinish = fromRoom.LookupParameter("Floor Finish").AsString();
                            newBreak.LookupParameter("Floor 1").Set(fromRmFinish);

                            // set the floor finish for the FromRoom
                            string toRmFinish = toRoom.LookupParameter("Floor Finish").AsString();
                            newBreak.LookupParameter("Floor 2").Set(toRmFinish);
                        }
                    }
                }
            }
        }

        #endregion

        #region Door Methods

        /// <summary>
        /// Gets all door instances in the document
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <returns>List of all door family instances</returns>
        private static List<FamilyInstance> GetAllDoors(Document curDoc)
        {
            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();
        }

        /// <summary>
        /// Loads a door family from the library, but only if it doesn't already exist in the project
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="typeDoor">The type of door ("Front" or "Rear")</param>
        /// <param name="specLevel">The spec level to determine which family to load</param>
        /// <returns>True if family exists or was loaded successfully, false if loading failed</returns>
        private static bool LoadDoorFamilyFromLibrary(Document curDoc, string typeDoor, string specLevel)
        {
            string familyName;

            if (typeDoor == "Front")
            {
                familyName = "LD_DR_Ext_Single 3_4 Lite_1 Panel";
            }
            else if (typeDoor == "Rear")
            {
                familyName = GetRearDoorFamilyName(specLevel);
            }
            else
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Invalid door type: {typeDoor}");
                return false;
            }

            // FIRST: Check if family already exists in the project
            var existingFamily = new FilteredElementCollector(curDoc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));

            if (existingFamily != null)
            {
                // Family already exists in project - no need to load from file
                return true;
            }

            // SECOND: If family doesn't exist, load it from the library
            Family loadedFamily = Utils.LoadFamilyFromLibrary(curDoc,
                @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Doors",
                familyName);

            // Return true if family was loaded successfully (not null)
            return loadedFamily != null;
        }

        private static FamilySymbol FindDoorSymbol(Document curDoc, string typeName, string familyName)
        {
            var comparer = StringComparer.OrdinalIgnoreCase;
            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(ds => comparer.Equals(ds.Family.Name, familyName) &&
                                      comparer.Equals(ds.Name, typeName));
        }

        #region Front Door Methods

        /// <summary>
        /// Updates the front door type based on spec level
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="specLevel">The spec level selection</param>
        public static void UpdateFrontDoor(Document curDoc, string specLevel)
        {
            var frontDoor = GetFrontDoor(curDoc);
            if (frontDoor == null)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "Front door not found.");
                return;
            }

            if (!LoadDoorFamilyFromLibrary(curDoc, "Front", specLevel))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to load door family from library.");
                return;
            }

            //string newDoorTypeName = GetFrontDoorType(specLevel);
            //if (string.IsNullOrEmpty(newDoorTypeName))
            //{
            //    Utils.TaskDialogError("Error", "Spec Conversion", "Unable to determine front door type for spec level: " + specLevel);
            //    return;
            //}

            string newDoorTypeName = GetFrontDoorType(specLevel);

            // DEBUG: List all available types in the family
            var allSymbols = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(ds => ds.Family.Name.Equals("LD_DR_Ext_Single 3_4 Lite_1 Panel", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var symbol in allSymbols)
            {
                System.Diagnostics.Debug.WriteLine($"Available type: {symbol.Name}, IsActive: {symbol.IsActive}");

                // Try activating all types
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    System.Diagnostics.Debug.WriteLine($"Activated: {symbol.Name}");
                }
            }

            var newDoorSymbol = FindDoorSymbol(curDoc, newDoorTypeName, "LD_DR_Ext_Single 3_4 Lite_1 Panel");
            if (newDoorSymbol == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Door type '{newDoorTypeName}' not found in the project after loading family.");
                return;
            }

            if (!newDoorSymbol.IsActive)
                newDoorSymbol.Activate();

            frontDoor.Symbol = newDoorSymbol;
        }

        /// <summary>
        /// Finds the front door based on room relationships
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <returns>The front door instance or null if not found</returns>
        public static FamilyInstance GetFrontDoor(Document curDoc)
        {
            return GetAllDoors(curDoc)
                .FirstOrDefault(door => IsFrontDoorMatch(door.FromRoom?.Name, door.ToRoom?.Name));
        }

        /// <summary>
        /// Checks if the room names match front door criteria
        /// </summary>
        /// <param name="fromRoomName">The "From Room: Name" value</param>
        /// <param name="toRoomName">The "To Room: Name" value</param>
        /// <returns>True if this appears to be the front door</returns>
        private static bool IsFrontDoorMatch(string fromRoomName, string toRoomName)
        {
            if (string.IsNullOrWhiteSpace(fromRoomName) || string.IsNullOrWhiteSpace(toRoomName))
                return false;

            bool toRoomMatch = toRoomName.Contains("Entry", StringComparison.OrdinalIgnoreCase) ||
                                 toRoomName.Contains("Foyer", StringComparison.OrdinalIgnoreCase);
            bool fromRoomMatch = fromRoomName.Contains("Covered Porch", StringComparison.OrdinalIgnoreCase);

            return fromRoomMatch && toRoomMatch;
        }

        /// <summary>
        /// Gets the front door type name based on spec level
        /// </summary>
        /// <param name="specLevel">The spec level</param>
        /// <returns>The door type name</returns>
        private static string GetFrontDoorType(string specLevel)
        {
            return specLevel switch
            {
                "Complete Home" => "36\"x80\"",
                "Complete Home Plus" => "36\"x96\"",
                _ => null
            };
        }

        #endregion

        #region Rear Door Methods

        /// <summary>
        /// Updates the rear door type based on spec level
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="specLevel">The spec level selection</param>
        public static void UpdateRearDoor(Document curDoc, string specLevel)
        {
            var rearDoor = GetRearDoor(curDoc);
            if (rearDoor == null)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "Rear door not found.");
                return;
            }

            if (!LoadDoorFamilyFromLibrary(curDoc, "Rear", specLevel))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to load rear door family from library.");
                return;
            }

            string newDoorTypeName = GetRearDoorType();
            var newDoorSymbol = FindDoorSymbol(curDoc, newDoorTypeName, GetRearDoorFamilyName(specLevel));
            if (newDoorSymbol == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Door type '{newDoorTypeName}' not found in the project after loading family.");
                return;
            }

            if (!newDoorSymbol.IsActive)
                newDoorSymbol.Activate();

            rearDoor.Symbol = newDoorSymbol;
        }

        /// <summary>
        /// Gets the rear door family name based on spec level
        /// </summary>
        /// <param name="specLevel">The spec level</param>
        /// <returns>The family name</returns>
        private static string GetRearDoorFamilyName(string specLevel)
        {
            return specLevel switch
            {
                "Complete Home" => "LD_DR_Ext_Single_Half Lite_2 Panel",
                "Complete Home Plus" => "LD_DR_Ext_Single_Full Lite",
                _ => null
            };
        }

        /// <summary>
        /// Finds the rear door based on width and description criteria
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <returns>The rear door instance or null if not found</returns>
        public static FamilyInstance GetRearDoor(Document curDoc)
        {
            return GetAllDoors(curDoc).FirstOrDefault(IsRearDoorMatch);
        }

        /// <summary>
        /// Checks if the door matches rear door criteria
        /// </summary>
        /// <param name="curDoor">The door to check</param>
        /// <returns>True if this appears to be the rear door</returns>
        private static bool IsRearDoorMatch(FamilyInstance curDoor)
        {
            // Check if door width is 32"
            var widthParam = curDoor.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            bool widthMatch = widthParam != null &&
                              Math.Abs((widthParam.AsDouble() * 12.0) - 32.0) < 0.1;

            // Check if description contains "Exterior Entry"
            var descParam = curDoor.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            bool descriptionMatch = descParam?.AsString()?.Contains("Exterior Entry", StringComparison.OrdinalIgnoreCase) == true;

            return widthMatch && descriptionMatch;
        }

        /// <summary>
        /// Gets the rear door type name (always 32"x80" for both spec levels)
        /// </summary>
        /// <returns>The door type name</returns>
        private static string GetRearDoorType()
        {
            return "32\"x80\"";
        }

        #endregion

        #endregion

        #region Cabinet Methods

        private void ReplaceWallCabinets(Document curDoc, string cabHeight)
        {
            // load the new cabinet families
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Casework\Kitchen", "LD_CW_Wall_1-Dr_Flush");
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Casework\Kitchen", "LD_CW_Wall_2-Dr_Flush");

            // get all wall cabinets in the document
            List<FamilyInstance> m_allWallCabs = GetAllStandardWallCabinets(curDoc);

            // loop through the wall cabinet instances
            foreach (FamilyInstance curCabinet in m_allWallCabs)
            {
                if (curCabinet.Symbol.Family.Name.Contains("Single Door") || curCabinet.Symbol.Family.Name.Contains("1-Dr"))
                {
                    // create string variable for new cabinet family name
                    string newCabinetFamilyName = "LD_CW_Wall_1-Dr_Flush";

                    // get the current cabinet type name
                    string curCabinetTypeName = curCabinet.Symbol.Name;

                    // get the new cabinet type based on the spec level height
                    string[] curDimensions = curCabinetTypeName.Split('x');
                    string curWidth = curDimensions[0].Trim();
                    string newCabinetTypeName = curWidth + "x" + cabHeight + "\"";

                    // add single door cabinet replacement logic
                    FamilySymbol newCabinetType = Utils.GetFamilySymbolByName(curDoc, newCabinetFamilyName, newCabinetTypeName);

                    // null check for the new cabinet type
                    if (newCabinetType == null)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", $"Cabinet type '{newCabinetTypeName}' not found in the project after loading family.");
                        continue;
                    }

                    // check if the new cabinet type is active
                    if (!newCabinetType.IsActive)
                    {
                        newCabinetType.Activate();
                    }

                    // replace the cabinet type
                    curCabinet.Symbol = newCabinetType;
                }
                else if (curCabinet.Symbol.Family.Name.Contains("Double Door") || curCabinet.Symbol.Family.Name.Contains("2-Dr"))
                {
                    // create string variable for new cabinet family name
                    string newCabinetFamilyName = "LD_CW_Wall_2-Dr_Flush";

                    // get the current cabinet type name
                    string curCabinetTypeName = curCabinet.Symbol.Name;

                    // get the new cabinet type based on the spec level height
                    string[] curDimensions = curCabinetTypeName.Split('x');
                    string curWidth = curDimensions[0].Trim();
                    string newCabinetTypeName = curWidth + "x" + cabHeight + "\"";

                    // add double door cabinet replacement logic
                    FamilySymbol newCabinetType = Utils.GetFamilySymbolByName(curDoc, newCabinetFamilyName, newCabinetTypeName);

                    // null check for the new cabinet type
                    if (newCabinetType == null)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", $"Cabinet type '{newCabinetTypeName}' not found in the project after loading family.");
                        continue;
                    }

                    // check if the new cabinet type is active
                    if (!newCabinetType.IsActive)
                    {
                        newCabinetType.Activate();
                    }

                    // replace the cabinet type
                    curCabinet.Symbol = newCabinetType;
                }
            }
        }

        private List<FamilyInstance> GetAllStandardWallCabinets(Document curDoc)
        {
            // get all wall cabinets in the document
            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Casework)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(cab => (cab.Symbol.Family.Name.Contains("Upper") ||
                 cab.Symbol.Family.Name.Contains("Wall")) &&
                 cab.Symbol.Name.Split('x').Length == 2)
                .ToList();
        }

        private void ReplaceMWCabinet(Document curDoc, string selectedMWCabHeight)
        {
            // get all wall cabinets with non-standard depth in the document
            List<FamilyInstance> m_allMWCabs = GetAllNonStandardWallCabinets(curDoc);

            // declare variables
            FamilyInstance curMWCab = null;
            string curWidth = "";
            string curDepth = "";

            // loop through the list and find the one being used for the MW cabinet (30" wide x 15" deep)
            foreach (FamilyInstance curCab in m_allMWCabs)
            {
                // parse the type name
                curWidth = curCab.Symbol.Name.Split('x')[0];
                curDepth = curCab.Symbol.Name.Split("x")[2];

                if (curWidth == "30\"" && curDepth == "15\"")
                {
                    curMWCab = curCab;
                    break;
                }
                else
                    continue;
            }

            // null check for cabinet found
            if (curMWCab == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "No MW cabinet found in the project.");
                return;
            }

            // get the current cabinet type name
            string curMWCabName = curMWCab.Symbol.Name;

            // create a string variable for the new cabinet type name
            string newMWCabTypeName = $"{curWidth}x{selectedMWCabHeight}x{curDepth}\"";

            // create the new cabinet type name based on the selected height
            FamilySymbol newMWCab = Utils.GetFamilySymbolByName(curDoc, "LD_CW_Wall_2-Dr_Flush", newMWCabTypeName);

            // null check for the new cabinet type
            if (newMWCab == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Cabinet type '{newMWCabTypeName}' not found in the project after loading family.");
                return;
            }

            // check if the new cabinet type is active
            if (!newMWCab.IsActive)
            {
                newMWCab.Activate();
            }

            // replace the cabinet type
            curMWCab.Symbol = newMWCab;

            // create variable for the new MW cabinet family name
            string newMWCabFamilyName = curMWCab.Symbol.Family.Name;

            // notify the user that the MW cabinet was updated
            Utils.TaskDialogInformation("Complete", "Spec Conversion",
                $"The MW cabinet was updated to {newMWCabFamilyName}:{newMWCabTypeName} per the selected height.");
        }

        private List<FamilyInstance> GetAllNonStandardWallCabinets(Document curDoc)
        {
            // get all wall cabinets in the document
            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Casework)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(cab => cab.Symbol.Family.Name.Contains("Wall") &&
                              (cab.Symbol.Name.Contains("Single") || cab.Symbol.Name.Contains("Double")) &&
                              cab.Symbol.Name.Split('x').Length == 3)
                .ToList();
        }       

        private void UpdateBacksplashHeight(Document curDoc, string selectedSpecLevel)
        {
            // load the new counter & backsplash families
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Generic Model\Kitchen", "LD_GM_Kitchen_Counter_Top-Mount");
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Generic Model\Kitchen", "LD_GM_Kitchen_Backsplash");

            // get all generic model instances in the document
            List<FamilyInstance> m_allGenericModels = Utils.GetAllGenericFamilies(curDoc);

            // filter the list for counter tops and backsplashes
            List<FamilyInstance> listBacksplashGMs = m_allGenericModels
                .Where(gm => gm.Symbol.Family.Name.Contains("Kitchen Counter") || gm.Symbol.Name.Contains("Kitchen Backsplash"))
                .ToList();

            // null check for the list
            if (listBacksplashGMs == null || !listBacksplashGMs.Any())
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "No Kitchen Counter or Backsplash generic models found in the project.");
                return;
            }

            // loop through the list and update the height based on the spec level
            foreach (FamilyInstance curGM in listBacksplashGMs)
            {
                // get the current type name
                string curTypeName = curGM.Symbol.Name;

                // replace the famile instance based on the current name
                if (curTypeName.Contains("Kitchen Counter"))
                {
                    // get the new counter type
                    FamilySymbol newCounterType = Utils.GetFamilySymbolByName(curDoc, "LD_GM_Kitchen_Counter_Top-Mount", "Type 1");

                    // null check for the new counter type
                    if (newCounterType == null)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", $"Counter type not found in the project after loading family.");
                        continue;
                    }

                    // check if the new counter type is active
                    if (!newCounterType.IsActive)
                    {
                        newCounterType.Activate();
                    }

                    // replace the family instance
                    curGM.Symbol = newCounterType;

                    // set the height based on the spec level
                    if (selectedSpecLevel == "Complete Home")
                    {
                        // set the height to 4"
                        curGM.Symbol.LookupParameter("Backsplash Height").Set(4.0 / 12.0);
                    }
                    else
                    {
                        // set the height to 18"
                        curGM.Symbol.LookupParameter("Backsplash Height").Set(18.0 / 12.0);
                    }
                }
                else if (curTypeName.Contains("Kitchen Backsplash"))
                {
                    // get the new backsplash type
                    FamilySymbol newBacksplashType = Utils.GetFamilySymbolByName(curDoc, "LD_GM_Kitchen_Backsplash", "Type 1");

                    // null check for the new backsplash type
                    if (newBacksplashType == null)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", $"Backsplash type not found in the project after loading family.");
                        continue;
                    }

                    // check if the new backsplash type is active
                    if (!newBacksplashType.IsActive)
                    {
                        newBacksplashType.Activate();
                    }

                    // replace the family instance
                    curGM.Symbol = newBacksplashType;

                    // set the height based on the spec level
                    if (selectedSpecLevel == "Complete Home")
                    {
                        // set the height to 4"
                        curGM.Symbol.LookupParameter("Height").Set(4.0 / 12.0);
                    }
                    else
                    {
                        // set the height to 18"
                        curGM.Symbol.LookupParameter("Height").Set(18.0 / 12.0);
                    }
                }
            }
        }

        private void UpdateBacksplashNote(Document curDoc, UIDocument uidoc, string selectedSpecLevel)
        {
            // get the selected spec level
            if (selectedSpecLevel == "Complete Home")
            {
                // get all text notes & filter for backsplash note
                List<TextNote> backsplashNotes = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(TextNote))
                    .Cast<TextNote>()
                    .Where(note => note.Text.Contains("Full Tile Backsplash"))
                    .ToList();

                // check if any backsplash notes were found
                if (backsplashNotes == null || !backsplashNotes.Any())
                {
                    // continue if no backsplash notes found
                    return;
                }

                // loop through and delete each note
                foreach (TextNote curNote in backsplashNotes)
                {
                    curDoc.Delete(curNote.Id);
                }
            }
            else if (selectedSpecLevel == "Complete Home Plus")
            {
                // get the TextNoteType
                TextNoteType backsplashNoteType = Utils.GetTextNoteTypeByName(curDoc, "STANDARD");

                // null check for the TextNoteType
                if (backsplashNoteType == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "Text Note Type 'STANDARD' not found in the project.");
                    return;
                }

                // get all the interior elevation views
                List<ViewSection> allIntElevs = GetAllIntElevViews(curDoc);

                // check if any interior elevation views were found
                if (allIntElevs == null || !allIntElevs.Any())
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "No Interior Elevation views found in the project.");
                    return;
                }

                // loop through each interior elevation views & add the backsplash note
                foreach (ViewSection curIntElev in allIntElevs)
                {
                    //// set the active view
                    //uidoc.ActiveView = curIntElev;

                    // set the text note location

                    // get the view boundaries
                    BoundingBoxXYZ curViewBounds = curIntElev.get_BoundingBox(curIntElev);

                    // calculate center point, 1' below bottom of bounding box
                    XYZ centerPoint = new XYZ(
                        (curViewBounds.Min.X + curViewBounds.Max.X) / 2, // horizontal center
                        curViewBounds.Min.Y - 1.0, // 1' below the bottom of the bounding box
                        0); // Z = 0 for view plane

                    try
                    {
                        // create a new text note
                        TextNote backsplashNote = TextNote.Create(curDoc, curIntElev.Id, centerPoint, "Full Tile Backsplash", backsplashNoteType.Id);

                        // set text note properties
                        backsplashNote.HorizontalAlignment = HorizontalTextAlignment.Center;
                        backsplashNote.VerticalAlignment = VerticalTextAlignment.Top;

                        // add leader lines
                        Leader leaderRight = backsplashNote.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R);
                        Leader leaderLeft = backsplashNote.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L);
                    }
                    catch (Exception ex)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", $"Error creating backsplash note in view {curIntElev.Name}: {ex.Message}");
                        continue;
                    }
                }
            }
        }

        private List<ViewSection> GetAllIntElevViews(Document curDoc)
        {
            // get all the ViewSection views and filter for Interior Elevaitons named Kitchen
            List<ViewSection> m_allIntElevs = Utils.GetAllSectionViews(curDoc)
                .OfType<ViewSection>()
                .Where(view => view.Name.Contains("Kitchen") && view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)?.AsString() == "Kitchen")
                .ToList();

            return m_allIntElevs;
        }

        #endregion

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            clsButtonData myButtonData = new clsButtonData(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }
}
