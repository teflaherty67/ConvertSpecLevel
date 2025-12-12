using Autodesk.Revit.DB;
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

            #region Ref Sp cleanup

            // get a first floor electrical view
            View curElecView = Utils.GetViewByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor", "Floor");

            // check for null
            if (curElecView != null)
            {
                // set it as active view
                uidoc.ActiveView = curElecView;
            }
            else
            {
                // if null notify the user
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "No view found with name containing 'Electrical' and associated level 'First Floor'");
            }

            // check for instance of new Ref Sp family
            bool isNewRefSpPresent = Utils.IsFamilyInstancePresent(curDoc, "LD_GR_Kitchen_Ref-Sp");

            // if not found, delete any existing Ref Sp, CW connection, outlet, & wall cabinet (if present)
            if (!isNewRefSpPresent)
            {
                // get the exisitng Ref Sp instance & supporting elements to delete
                List<ElementId> elementsToDelete = GetElementsToDelete(curDoc);

                // delete the elements
                if (elementsToDelete.Count > 0)
                {
                    // create a transaction to delete the elements
                    using (Transaction t = new Transaction(curDoc, "Delete Existing Ref Sp and Supporting Elements"))
                    {
                        // start the transaction
                        t.Start();

                        // delete the elements
                        curDoc.Delete(elementsToDelete);

                        // load the new Ref Sp family
                        Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Groups", "LD_GR_Kitchen_Ref-Sp");

                        // commit the transaction
                        t.Commit();
                    }

                    // notify the user
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "No instance of 'LD_GR_Kitchen_Ref-Sp' was found in the project." +
                        "The existing Ref Sp family and it's supporting elements have been deleted, and the new Ref Sp family has been loaded." +
                        "Please place an instance of 'LD_GR_Kitchen_Ref-Sp' in the project and re-run the command.");
                    return Result.Succeeded;
                }
            }

            #endregion

            // if found, proceed with the rest of code
            else
            {                
                #region Form

                // launch the form
                frmConvertSpecLevel curForm = new frmConvertSpecLevel(curDoc, uidoc);
                curForm.Topmost = true;

                curForm.ShowDialog();

                // check if user clicked Cancel
                if (curForm.DialogResult != true)
                {
                    return Result.Cancelled;
                }

                // get user input from form
                string selectedClient = curForm.GetSelectedClient();
                string selectedSpecLevel = curForm.GetSelectedSpecLevel();


                Reference selectedOutlet = curForm.SelectedOutlet;
                Reference selectedSprinklerWall = curForm.SelectedOutletWall;
                Reference selectedGarageWall = curForm.SelectedGarageWall;
               
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
                        ManageFloorMaterialBreaksInActiveView(curDoc, listUpdatedRooms);

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

                        // commit the transaction
                        t.Commit();

                        // notify the user
                        Utils.TaskDialogInformation("Complete", "Spec Conversion", $"Flooring was changed at {listRooms} per the specified spec level.");
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

                        // commit the transaction
                        t.Commit();
                    }

                    // notify the user
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "The front and rear doors were replaced per the specified spec level.");

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
                            ReplaceCabinetFillers(curDoc, "36");
                            ManageCaseworkTags(curDoc, selectedSpecLevel);
                        }
                        else
                        {
                            // raise the upper cabinets to 42"
                            ReplaceWallCabinets(curDoc, "42");
                            ReplaceCabinetFillers(curDoc, "42");
                            ManageCaseworkTags(curDoc, selectedSpecLevel);
                        }

                        // get MW cabinet height from cabinet spec map
                        string mwHeight = clsCabSpecMap.GetMWCabHeight(selectedClient, selectedSpecLevel);
                        if (mwHeight == null)
                        {
                            Utils.TaskDialogError("Error", "Spec Conversion", $"No MW cabinet height found for {selectedClient} - {selectedSpecLevel}");
                            return Result.Failed;
                        }

                        // revise the MW cabinet
                        ReplaceMWCabinet(curDoc, mwHeight);

                        // get Ref Sp settings from the cabinet spec map
                        var refSpSettings = clsCabSpecMap.GetRefSpSettings(selectedClient, selectedSpecLevel);
                        if (refSpSettings == null)
                        {
                            Utils.TaskDialogError("Error", "Spec Conversion", $"No RefSp settings found for {selectedClient} - {selectedSpecLevel}");
                            return Result.Failed;
                        }

                        // apply the Ref Sp settings
                        ManageRefSpCabinet(curDoc, refSpSettings);

                        // raise/lower the backsplash height
                        UpdateBacksplash(curDoc, selectedSpecLevel);

                        // commit the transaction
                        t.Commit();

                        // notify the user
                        // upper cabinets were revised per the selected spec level
                        // Ref Sp cabinet was added/removed per the selected spec level
                        // backsplash height was raised/lowered per the selected spec level
                    }

                    #endregion

                    #region General Electrical Setup

                    // variables for sprinkler family
            FamilySymbol sprinklerSymbol = null;
            FamilyInstance outletInstance = null;
            string outletFamilyName = "LD_EF_Recep_None";
            string outletTypeName = "Sprinkler";
#if REVIT2025
            string outletFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical";
#endif

#if REVIT2026
            string outletFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\Electrical";
#endif

            // variables for tag family
            FamilySymbol tagSymbol = null;
            string tagFamilyName = "LD_AN_Tag_EF_Type-Comments";
            string tagTypeName = "Type 2";
#if REVIT2025
            string tagFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Annotation\Tags";
#endif

#if REVIT2026
            string tagFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\Annotation\Tags";
#endif

                    // get all electrical plan views
                    List<View> electricalViews = Utils.GetAllViewsByNameContains(curDoc, "Electrical");

                    // verify if project is two story
                    bool isPlanTwoStory = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .Any(level => level.Name.Contains("Second Floor"));

                    // load the new electrical families (outlet & light fixture)
                    Family lightSymbol = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Lighting", "LD_LF_None");
                    Family outletSymbol = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical", "LD_EF_Recep_Wall");

                    #endregion

                    #region First Floor Electrical Updates

                    // get all views with Electrical in the name & associated with the First Floor
                    List<View> firstFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor");

                    // get the first view in the list and set it as the active view
                    if (firstFloorElecViews.Any())
                    {
                        uidoc.ActiveView = firstFloorElecViews.First();

                        // create transaction for first floor electrical updates
                        using (Transaction t = new Transaction(curDoc, "Update First Floor Electrical"))
                        {
                            // start the fourth transaction
                            t.Start();

                            // replace the light fixtures in the specified rooms per the selected spec level
                            var (roomsUpdated, fixtureCount) = UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                            // add/remove the ceiling fan note in the views
                            var (added, deleted, viewCount) = ManageClgFanNotes(curDoc, uidoc, selectedSpecLevel, firstFloorElecViews);

                            // add/remove the sprinkler outlet in the Garage
                            ManageSprinklerOutlet(curDoc, uidoc, selectedSpecLevel, selectedSprinklerWall, selectedGarageWall, selectedOutlet);

                            // add/remove sprinkler outlet note
                            RemoveSprinklerOutletNote(curDoc, uidoc, selectedSpecLevel, firstFloorElecViews);

                            // commit the transaction
                            t.Commit();

                            // Show summary message with proper grammar
                            string roomList;
                            if (roomsUpdated.Count == 0)
                            {
                                roomList = "No rooms";
                            }
                            else if (roomsUpdated.Count == 1)
                            {
                                roomList = roomsUpdated[0];
                            }
                            else if (roomsUpdated.Count == 2)
                            {
                                roomList = $"{roomsUpdated[0]} and {roomsUpdated[1]}";
                            }
                            else
                            {
                                roomList = string.Join(", ", roomsUpdated.Take(roomsUpdated.Count - 1)) + $", and {roomsUpdated.Last()}";
                            }

                            // Determine action based on spec level
                            string action = selectedSpecLevel switch
                            {
                                "Complete Home" => "added",
                                "Complete Home Plus" => "deleted",
                                _ => "processed"
                            };

                            // Grammar for fixtures and views
                            string fixtureText = roomsUpdated.Count == 1 ? "Fixture updated" : "Fixtures updated";
                            string viewText = viewCount == 1 ? "view" : "views";

                            // Create the final message
                            string messageSummary = $"{fixtureText} in {roomList}. Ceiling fan notes {action} across {viewCount} {viewText}.";

                            // Show the summary dialog
                            Utils.TaskDialogInformation("Complete", "First Floor Electrical", messageSummary);
                        }
                    }
                    else
                    {
                        // if not found alert the user
                        Utils.TaskDialogError("Error", "Spec Conversion", "No Electrical views found for First Floor");
                    }

                    #endregion

                    #region Second Floor Electrical Updates

                    // get all views with Electrical in the name & associated with the Second Floor
                    List<View> secondFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "Second Floor");

                    // get the first view in the list and set it as the active view
                    if (secondFloorElecViews.Any())
                    {
                        uidoc.ActiveView = secondFloorElecViews.First();
                    }
                    else
                    {
                        // if not found alert the user
                        Utils.TaskDialogError("Error", "Spec Conversion", "No Electrical views found for Second Floor");
                    }

                    // create transaction for Second Floor Electrical updates
                    using (Transaction t = new Transaction(curDoc, "Update Second Floor Electrical"))
                    {

                        // start the transaction
                        t.Start();

                        // replace the light fixtures in the specified rooms per the selected spec level
                        var (roomsUpdated, fixtureCount) = UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                        // add/remove the ceiling fan note in the views
                        var (added, deleted, viewCount) = ManageClgFanNotes(curDoc, uidoc, selectedSpecLevel, secondFloorElecViews);

                        // commit the transaction
                        t.Commit();

                        // Show summary message with proper grammar
                        string roomList;
                        if (roomsUpdated.Count == 0)
                        {
                            roomList = "No rooms";
                        }
                        else if (roomsUpdated.Count == 1)
                        {
                            roomList = roomsUpdated[0];
                        }
                        else if (roomsUpdated.Count == 2)
                        {
                            roomList = $"{roomsUpdated[0]} and {roomsUpdated[1]}";
                        }
                        else
                        {
                            roomList = string.Join(", ", roomsUpdated.Take(roomsUpdated.Count - 1)) + $", and {roomsUpdated.Last()}";
                        }

                        // Determine action based on spec level
                        string action = selectedSpecLevel switch
                        {
                            "Complete Home" => "added",
                            "Complete Home Plus" => "deleted",
                            _ => "processed"
                        };

                        // Grammar for fixtures and views
                        string fixtureText = roomsUpdated.Count == 1 ? "Fixture updated" : "Fixtures updated";
                        string viewText = viewCount == 1 ? "view" : "views";

                        // Create the final message
                        string messageSummary = $"{fixtureText} in {roomList}. Ceiling fan notes {action} across {viewCount} {viewText}.";

                        // Show the summary dialog
                        Utils.TaskDialogInformation("Complete", "First Floor Electrical", messageSummary);
                    }

                    #endregion

                    // commit the transaction group
                    transgroup.Assimilate();
                }

                #endregion
            }

            return Result.Succeeded;
        }

        #region Ref Sp Cleanup Methods

        private List<ElementId> GetElementsToDelete(Document curDoc)
        {
            List<ElementId> elementsToDelete = new List<ElementId>();

            // Find existing Ref Sp instances directly
            var existingRefSp = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_SpecialityEquipment)
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("Refrigerator") || fi.Symbol.Family.Name.Contains("Refrigeration"))
                .ToList();

            // If none found, notify the user
            if (existingRefSp.Count == 0)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "No existing refrigerator found in the project.");
                return elementsToDelete;
            }

            foreach (FamilyInstance curRefSp in existingRefSp)
            {
                elementsToDelete.Add(curRefSp.Id);

                LocationPoint locPoint = curRefSp.Location as LocationPoint;
                if (locPoint == null) continue;

                XYZ fridgeOrigin = locPoint.Point;

                // Get the family's transform to determine its orientation
                Transform familyTransform = curRefSp.GetTransform();

                // Get the family's local coordinate system directions
                XYZ familyXDirection = familyTransform.BasisX; // Right/Left direction
                XYZ familyYDirection = familyTransform.BasisY; // Forward/Back direction

                // get the Center (L/R) direction (perpendicular to the fridge face)
                XYZ centerLRDirection = familyYDirection.Normalize();

                // Create line extending 18" in both directions along Center (L/R) reference
                double searchDistance = 1.5; // 18" in feet
                XYZ startPoint = fridgeOrigin - (centerLRDirection * searchDistance);
                XYZ endPoint = fridgeOrigin + (centerLRDirection * searchDistance);
                Line searchLine = Line.CreateBound(startPoint, endPoint);

                // Find perpendicular wall to this line
                Wall perpendicularWall = FindPerpendicularWall(curDoc, searchLine, fridgeOrigin);
                if (perpendicularWall == null)
                {
                    Utils.TaskDialogWarning("Warning", "Spec Conversion",
                        "Could not find perpendicular wall to refrigerator");
                    continue;
                }

                // Use the wall to establish search area for supporting elements
                XYZ wallSearchPoint = GetWallSearchPoint(perpendicularWall, fridgeOrigin);

                // Define search area around the wall point
                double searchRadius = 1.79167; // 21.5" search radius

                // -------- proximty search for nearby outlets --------
                var nearbyOutlets = new FilteredElementCollector(curDoc)
                    .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(outlet =>
                    {
                        // Filter for only "Outlet-Duplex" type fixtures
                        if (outlet.Symbol.Name != "Outlet-Duplex") return false;

                        LocationPoint outletLoc = outlet.Location as LocationPoint;
                        if (outletLoc == null) return false;

                        XYZ outletPoint = outletLoc.Point;
                        double horizontalDistance = Math.Sqrt(
                            Math.Pow(wallSearchPoint.X - outletPoint.X, 2) +
                            Math.Pow(wallSearchPoint.Y - outletPoint.Y, 2));
                        double verticalDistance = Math.Abs(fridgeOrigin.Z - outletPoint.Z);

                        return horizontalDistance <= searchRadius && verticalDistance <= 4.0;
                    })
                    .ToList();

                // -------- select outlet closest to fridge origin --------
                if (nearbyOutlets.Count > 0)
                {
                    var closestOutlet = nearbyOutlets
                        .OrderBy(outlet =>
                        {
                            LocationPoint outletLoc = outlet.Location as LocationPoint;
                            return fridgeOrigin.DistanceTo(outletLoc.Point);
                        })
                        .First();

                    // Add ONLY the closest outlet to delete list
                    elementsToDelete.Add(closestOutlet.Id);                    
                }

                // -------- proximity search for nearby CW connections --------
                var nearbyCWConnections = new FilteredElementCollector(curDoc)
                    .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(connection =>
                    {
                        LocationPoint connectionLoc = connection.Location as LocationPoint;
                        if (connectionLoc == null) return false;

                        XYZ connectionPoint = connectionLoc.Point;
                        double horizontalDistance = Math.Sqrt(
                            Math.Pow(wallSearchPoint.X - connectionPoint.X, 2) +
                            Math.Pow(wallSearchPoint.Y - connectionPoint.Y, 2));
                        double verticalDistance = Math.Abs(fridgeOrigin.Z - connectionPoint.Z);

                        return horizontalDistance <= searchRadius && verticalDistance <= 2.5;
                    })
                    .ToList();

                elementsToDelete.AddRange(nearbyCWConnections.Select(cw => cw.Id));

                // -------- search for wall cabinets above fridge --------
                var wallCabinetsAbove = new FilteredElementCollector(curDoc)
                    .OfCategory(BuiltInCategory.OST_Casework)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(cabinet =>
                    {
                        if (cabinet.SuperComponent != null) return false;

                        LocationPoint cabinetLoc = cabinet.Location as LocationPoint;
                        if (cabinetLoc == null) return false;

                        XYZ cabinetPoint = cabinetLoc.Point;

                        // Cabinet bottom height between 5'-9" and 6'-6"
                        if (cabinetPoint.Z < 5.75 || cabinetPoint.Z > 6.5)
                            return false;

                        // Use wallSearchPoint instead of fridgeOrigin for better positioning
                        double horizontalDistance = Math.Sqrt(
                            Math.Pow(wallSearchPoint.X - cabinetPoint.X, 2) +
                            Math.Pow(wallSearchPoint.Y - cabinetPoint.Y, 2));

                        return horizontalDistance <= searchRadius;
                    })
                    .ToList();

                elementsToDelete.AddRange(wallCabinetsAbove.Select(cab => cab.Id));

                // -------- FINAL DEBUG SUMMARY --------               
                int outletsToDelete = nearbyOutlets.Count > 0 ? 1 : 0;

                Utils.TaskDialogInformation("DEBUG", "Final Element Count",
                    $"Outlets: Adding {outletsToDelete} to delete\n" +
                    $"CW Connections: Adding {nearbyCWConnections.Count} to delete\n" +
                    $"Wall Cabinets: Adding {wallCabinetsAbove.Count} to delete\n" +
                    $"Total Elements to Delete: {elementsToDelete.Count}");
            }

            return elementsToDelete;
        }

        private Wall FindPerpendicularWall(Document curDoc, Line searchLine, XYZ fridgeOrigin)
        {
            // Get all walls in the document
            var m_allWalls = new FilteredElementCollector(curDoc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(wall => wall.Location is LocationCurve)
                .ToList();

            XYZ searchDirection = searchLine.Direction;
            double closestDistance = double.MaxValue;
            Wall closestPerpendicularWall = null;

            foreach (Wall curWall in m_allWalls)
            {
                LocationCurve wallLoc = curWall.Location as LocationCurve;
                if (wallLoc != null)
                {
                    Line wallLine = wallLoc.Curve as Line;
                    if (wallLine != null)
                    {
                        XYZ wallDirection = wallLine.Direction;

                        // Check if wall is perpendicular to search line (dot product near 0)
                        double dotProduct = Math.Abs(searchDirection.DotProduct(wallDirection));
                        if (dotProduct < 0.1) // Perpendicular tolerance
                        {
                            // Find closest point on wall to fridge origin
                            XYZ closestPointOnWall = wallLine.Project(fridgeOrigin).XYZPoint;
                            double distanceToWall = fridgeOrigin.DistanceTo(closestPointOnWall);

                            // Check if this wall is closer and within reasonable distance
                            if (distanceToWall < closestDistance && distanceToWall < 5.0) // 5' max distance
                            {
                                closestDistance = distanceToWall;
                                closestPerpendicularWall = curWall;
                            }
                        }
                    }
                }
            }

            return closestPerpendicularWall;
        }

        private XYZ GetWallSearchPoint(Wall wall, XYZ fridgeOrigin)
        {
            LocationCurve wallLoc = wall.Location as LocationCurve;
            if (wallLoc != null)
            {
                Line wallLine = wallLoc.Curve as Line;
                if (wallLine != null)
                {
                    // Project fridge origin onto the wall line to get search point
                    XYZ projectedPoint = wallLine.Project(fridgeOrigin).XYZPoint;
                    return projectedPoint;
                }
            }

            return fridgeOrigin; // Fallback to fridge origin if wall projection fails
        }

        #endregion

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

        private void ManageFloorMaterialBreaksInActiveView(Document curDoc, List<string> listUpdatedRooms)
        {
            // get all the doors in the active view
            List<FamilyInstance> allDoorsInView = Utils.GetAllDoorsInActiveView(curDoc);

            // filter the list for doors with ToRoom or FromRoom matching any of the updated rooms
            var filteredDoors = allDoorsInView.Where(door => listUpdatedRooms.Any(roomName =>
            (door.ToRoom != null && door.ToRoom.Name.IndexOf(roomName, StringComparison.OrdinalIgnoreCase) >= 0) ||
            (door.FromRoom != null && door.FromRoom.Name.IndexOf(roomName, StringComparison.OrdinalIgnoreCase) >= 0)))
                .ToList();

            // loop through filtereedDoors
            foreach (FamilyInstance curDoor in filteredDoors)
            {
                // get ToRoom & FromRoom values
                Room toRoom = curDoor.ToRoom;
                Room fromRoom = curDoor.FromRoom;

                // check for null
                if (toRoom == null || fromRoom == null)
                {
                    continue;
                }

                // if materials match, remove any material break
                if (toRoom.LookupParameter("Floor Finish").AsString() == (fromRoom.LookupParameter("Floor Finish").AsString()))
                {
                    RemoveFloorMaterialBreak(curDoc, curDoor);
                }
                else // if materials do not match, add a material break if needed
                {
                    AddFloorMaterialBreak(curDoc, curDoor);
                }
            }
        }

       

        private void RemoveFloorMaterialBreak(Document curDoc,FamilyInstance curDoor)
        {
            // get the door's location point
            LocationPoint doorLoc = curDoor.Location as LocationPoint;
            XYZ doorPoint = doorLoc.Point;

            // check for null
            if (doorLoc == null || doorPoint == null)
            {
                return;
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

        private void AddFloorMaterialBreak(Document curDoc, FamilyInstance curDoor)
        {
            // get ToRoom & FromRoom values
            Room toRoom = curDoor.ToRoom;
            Room fromRoom = curDoor.FromRoom;

            // check for null
            if (toRoom == null || fromRoom == null)
            {
                return;
            }

            // if materials do not match, get the door's location point
            LocationPoint doorLoc = curDoor.Location as LocationPoint;
            XYZ doorPoint = doorLoc.Point;

            // check for null
            if (doorLoc == null || doorPoint == null)
            {
                return;
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
                    return;
                }

                // get wall location curve
                LocationCurve wallLoc = drWall.Location as LocationCurve;

                // check for null curve
                if (wallLoc == null)
                {
                    return;
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

                    // create start point
                    XYZ startPoint = doorPoint - (wallDirection * (drWidthParam / 2));

                    // create end point  
                    XYZ endPoint = doorPoint + (wallDirection * (drWidthParam / 2));

                    // create a line to place the break
                    Line breakLine = Line.CreateBound(startPoint, endPoint);

                    // create a new Floor Material instance
                    FamilyInstance newBreak = curDoc.Create.NewFamilyInstance(breakLine, materialSymbol, workPlane, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    // set the floor finish for the FromRoom
                    string fromRmFinish = fromRoom.LookupParameter("Floor Finish").AsString();
                    if (fromRmFinish == "Carpet") fromRmFinish = "C";
                    newBreak.LookupParameter("Floor 1").Set(fromRmFinish);

                    // set the floor finish for the ToRoom
                    string toRmFinish = toRoom.LookupParameter("Floor Finish").AsString();
                    if (toRmFinish == "Carpet") toRmFinish = "C";
                    newBreak.LookupParameter("Floor 2").Set(toRmFinish);
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

        private void ReplaceCabinetFillers(Document curDoc, string fillerHeight)
        {
            // load the new filler families
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Casework\Kitchen", "LD_CW_Wall_Filler");

            // get all wall fillers in the document
            List<FamilyInstance> m_allWallFillers = GetAllWallFillers(curDoc);

            // loop through all wall fillers
            foreach (FamilyInstance curFiller in m_allWallFillers)
            {
                // create string variable for new filler family name
                string newFillerFamilyName = "LD_CW_Wall_Filler";

                // get the current filler type name
                string curFillerTypeName = curFiller.Symbol.Name;

                // get the new filler type based on the spec level height
                string[] curDimensions = curFillerTypeName.Split('x');
                string curWidth = curDimensions[0].Trim();
                string newFillerTypeName = curWidth + "x" + fillerHeight + "\"";

                // add filler replacement logic
                FamilySymbol newFillerType = Utils.GetFamilySymbolByName(curDoc, newFillerFamilyName, newFillerTypeName);

                // null check for the new cabinet type
                if (newFillerType == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", $"Filler type '{newFillerTypeName}' not found in the project after loading family.");
                    continue;
                }

                // check if the new cabinet type is active
                if (!newFillerType.IsActive)
                {
                    newFillerType.Activate();
                }

                // replace the cabinet type
                curFiller.Symbol = newFillerType;
            }
        }

        private List<FamilyInstance> GetAllWallFillers(Document curDoc)
        {
            // get all wall cabinets in the document
            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Casework)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(cab => (cab.Symbol.Family.Name.Contains("Filler")) &&
                 cab.Symbol.Name.Split('x').Length == 2)
                .ToList();
        }

        private void ManageRefSpCabinet(Document curDoc, clsCabSpecMap.RefSpSettings refSpSettings)
        {
            // find existing Ref Sp instances
            var instanceRefSp = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("LD_GR_Kitchen_Ref-Sp"))
                .ToList();

            bool isVisible = refSpSettings.IsVisible;
            string typeName = refSpSettings.TypeName;
            double heightOffset = refSpSettings.HeightOffsetFromLevel;

            if (instanceRefSp.Count == 0)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "No Ref Sp instances found in the project.");
                return;
            }

            // if found apply settings
            foreach (FamilyInstance curRefSp in instanceRefSp)
            {
                // apply cabinet visibility parameter
                Parameter visibilityParam = curRefSp.LookupParameter("Ref Sp Cabinet");
                if (visibilityParam != null)
                {
                    visibilityParam.Set(isVisible ? 1 : 0);  // 1 = visible, 0 = hidden
                }

                // apply cabinet type if visible
                Parameter typeParam = curRefSp.LookupParameter("Cabinet");
                if (typeParam != null && !string.IsNullOrEmpty(typeName))
                {
                    // Construct the full format: "FamilyName : TypeName"
                    string fullCabinetValue = $"LD_CW_Wall_2-Dr_Flush : {typeName}";
                    typeParam.Set(fullCabinetValue);
                }

                // apply height offset from level
                Parameter heightParam = curRefSp.LookupParameter("Cabinet Offset AFF");
                if (heightParam != null)
                {
                    heightParam.Set(heightOffset);
                }
            }
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
            string newMWCabTypeName = $"{curWidth}x{selectedMWCabHeight}x{curDepth}";

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
                .Where(cab => cab.Symbol.Family.Name.Contains("Upper Cabinet-Double") || cab.Symbol.Family.Name.Contains("Wall_2-Dr") &&
                              cab.Symbol.Name.Split('x').Length == 3)
                .ToList();
        }

        private void ManageCaseworkTags(Document curDoc, string selectedSpecLevel)
        {
            // get all wall cabinet tags in the document
            List<IndependentTag> m_allWallCabTags = new FilteredElementCollector(curDoc)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .Where(tag => tag.TagText.StartsWith("W"))
                .ToList();

            // loop through the list & raise or lower the tag based on the spec level
            foreach (IndependentTag curTag in m_allWallCabTags)
            {
                // get the location of the tag
                XYZ curTagPoint = curTag.TagHeadPosition;

                // check for null
                if (curTagPoint == null)
                {
                    continue;
                }

                // adjust tag locan based on spec level
                if (selectedSpecLevel == "Complete Home")
                {
                    // lower the tag by 6"
                    XYZ newTagPoint = new XYZ(curTagPoint.X, curTagPoint.Y, curTagPoint.Z - 0.5);
                    curTag.TagHeadPosition = newTagPoint;
                }
                else
                {
                    // raise the tag by 6"
                    XYZ newTagPoint = new XYZ(curTagPoint.X, curTagPoint.Y, curTagPoint.Z + 0.5);
                    curTag.TagHeadPosition = newTagPoint;
                }
            }
        }

        private void UpdateBacksplash(Document curDoc, string selectedSpecLevel)
        {
            // load the new counter & backsplash families
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Generic Model\Kitchen", "LD_GM_Kitchen_Counter_Top-Mount");
            Utils.LoadFamilyFromLibrary(curDoc, $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Generic Model\Kitchen", "LD_GM_Kitchen_Backsplash");

            // get all generic model instances in the document
            List<FamilyInstance> m_allGenericModels = Utils.GetAllGenericFamilies(curDoc);

            // filter the list for counter tops and backsplashes
            List<FamilyInstance> listBacksplashGMs = m_allGenericModels
                .Where(gm => gm.Symbol.Family.Name.Contains("Kitchen Counter") || gm.Symbol.Family.Name.Contains("Kitchen_Counter"))
                .ToList();           

            // null check for the list
            if (listBacksplashGMs == null || !listBacksplashGMs.Any())
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "No Kitchen Counter or Backsplash generic models found in the project.");
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

            // loop through the list and update the height based on the spec level
            foreach (FamilyInstance curGM in listBacksplashGMs)
            {
                // get the current type name
                string curTypeName = curGM.Symbol.Name;

                // replace the family instance based on the current name
                if (curGM.Symbol.Family.Name.Contains("Kitchen Counter"))
                {
                    // store existing patrameter values
                    clsCountertopParams storedParams = new clsCountertopParams(curGM);

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

                    // restore parameter values
                    storedParams.RestoreToElement(curGM);

                    // check the value of the Backsplash Back parameter
                    Parameter paramBacksplashBack = curGM.LookupParameter("Backsplash Back");

                    // if Backsplash Back is not null and is equal to Yes
                    if (paramBacksplashBack != null && paramBacksplashBack.AsInteger() == 1) // 1 = yes/true
                    {
                        SetBacksplashHeight(curGM, selectedSpecLevel, "Backsplash Height");

                        ManageBacksplashNote(curDoc, curGM, allIntElevs, selectedSpecLevel);
                    }
                }
                else if (curGM.Symbol.Family.Name.Contains("Kitchen_Counter"))
                {
                    // check the value of the Backsplash Back parameter
                    Parameter paramBacksplashBack = curGM.LookupParameter("Backsplash Back");

                    // if Backsplash Back is not null and is equal to Yes
                    if (paramBacksplashBack != null && paramBacksplashBack.AsInteger() == 1) // 1 = yes/true
                    {
                        SetBacksplashHeight(curGM, selectedSpecLevel, "Backsplash Height");

                        ManageBacksplashNote(curDoc, curGM, allIntElevs, selectedSpecLevel);
                    }
                }
            }
        }

        private void SetBacksplashHeight(FamilyInstance curGM, string selectedSpecLevel, string paramName)
        {
            // then set the height based on the spec level
            if (selectedSpecLevel == "Complete Home")
            {
                // set the height to 4"
                curGM.LookupParameter(paramName).Set(4.0 / 12.0);
            }
            else
            {
                // set the height to 18"
                curGM.LookupParameter(paramName).Set(18.0 / 12.0);
            }
        }

        private void ManageBacksplashNote(Document curDoc, FamilyInstance curGM, List<ViewSection> allIntElevs, string selectedSpecLevel)
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

                // loop through interior elevation views & find the one with the current countertop
                foreach (ViewSection curIntElev in allIntElevs)
                {
                    // check if the current countertop is in the view
                    var counterInView = new FilteredElementCollector(curDoc, curIntElev.Id)
                        .OfCategory(BuiltInCategory.OST_GenericModel)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .FirstOrDefault(gm => gm.Id == curGM.Id);

                    // null check for the counter in view
                    if (counterInView != null)
                    {
                        // get all text notes & filter for backsplash note
                        List<TextNote> backsplashNotes = new FilteredElementCollector(curDoc, curIntElev.Id)
                            .OfClass(typeof(TextNote))
                            .Cast<TextNote>()
                            .Where(note => note.Text.Contains("Full Tile Backsplash"))
                            .ToList();

                        // check if any backsplash notes were found
                        if (backsplashNotes == null || !backsplashNotes.Any())
                        {
                            // get countertop location for note positioning
                            LocationCurve countertopLoc = curGM.Location as LocationCurve;

                            // null check for location
                            if (countertopLoc != null)
                            {
                                // get the midpoint of the countertop curve
                                Curve curve = countertopLoc.Curve;
                                XYZ countertopPoint = curve.Evaluate(0.5, true); // midpoint of countertop

                                // set note locaiton based on countertop point
                                XYZ notePosition = new XYZ(
                                    countertopPoint.X, // X = countertop X
                                    countertopPoint.Y, // y = countertop Y
                                    countertopPoint.Z + 3.75); // countertop Z + 9"

                                try
                                {
                                    // create a new text note
                                    TextNote backsplashNote = TextNote.Create(curDoc, curIntElev.Id, notePosition, "Full Tile Backsplash", backsplashNoteType.Id);

                                    // set text note properties
                                    backsplashNote.HorizontalAlignment = HorizontalTextAlignment.Center;
                                    backsplashNote.VerticalAlignment = VerticalTextAlignment.Top;

                                    // add leader lines
                                    Leader leaderRight = backsplashNote.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R);
                                    Leader leaderLeft = backsplashNote.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L);

                                    // set leader attachment to midpoint
                                    backsplashNote.LeaderLeftAttachment = LeaderAtachement.Midpoint;
                                    backsplashNote.LeaderRightAttachment = LeaderAtachement.Midpoint;

                                    break; // exit loop after creating note in correct view
                                }
                                catch (Exception ex)
                                {
                                    Utils.TaskDialogError("Error", "Spec Conversion", $"Error creating backsplash note in view {curIntElev.Name}: {ex.Message}");
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
        }    

        private List<FamilyInstance> GetAllCountersWithBacksplashBack(Document curDoc, ElementId viewId)
        {
            return new FilteredElementCollector(curDoc, viewId)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(gm => gm.Symbol.Family.Name.Contains("Kitchen_Counter") &&
                             gm.LookupParameter("Backsplash Back")?.AsInteger() == 1)
                .ToList();
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

        #region Electrical Methods

        /// <summary>
        /// Updates lighting fixtures in specified rooms based on the given specification level.
        /// Only processes fixtures in rooms visible in the active view.
        /// </summary>
        private static (List<string> updatedRooms, int fixtureCount) UpdateLightingFixturesInActiveView(Document curDoc, string selectedSpecLevel)
        {
            // Define rooms that need lighting fixture updates
            List<string> roomsToUpdate = new List<string>
        {
            "Master Bedroom",
            "Covered Patio",
            "Gameroom",
            "Loft"
        };

            // Determine target family type based on spec level
            string targetFamilyType = selectedSpecLevel switch
            {
                "Complete Home" => "LED",
                "Complete Home Plus" => "Ceiling Fan",
                _ => null
            };

            if (targetFamilyType == null)
            {
                TaskDialog.Show("Error", "Invalid Spec Level selected.");
                return (new List<string>(), 0);
            }

            // Get the active view
            View activeView = curDoc.ActiveView;
            if (activeView == null)
            {
                TaskDialog.Show("Error", "No active view found.");
                return (new List<string>(), 0);
            }

            // Find target family symbol
            FamilySymbol targetFamilySymbol = Utils.FindFamilySymbol(curDoc, "LT-No Base", targetFamilyType);
            if (targetFamilySymbol == null)
            {
                TaskDialog.Show("Error", $"Family symbol '{targetFamilyType}' not found.");
                return (new List<string>(), 0);
            }

            // Activate the family symbol if not already active
            if (!targetFamilySymbol.IsActive)
            {
                targetFamilySymbol.Activate();
            }

            // counters and tracking
            int updatedCount = 0;
            int roomsNotFound = 0;

            List<string> updatedRooms = new List<string>();

            // Iterate through each room to update lighting fixtures
            foreach (string roomName in roomsToUpdate)
            {
                // Get the room by name in the active view only
                List<Room> rooms = GetRoomByNameContainsInActiveView(curDoc, activeView, roomName);

                // If no room found, show an error message and continue to the next room
                if (rooms.Count == 0)
                {
                    roomsNotFound++;
                    continue;
                }

                // Process each matched room
                foreach (Room room in rooms)
                {
                    // Find the lighting fixture of the specified family in the room (active view only)
                    List<FamilyInstance> lightingFixtures = GetLightFixtureInRoomInActiveView(curDoc, activeView, room, "LT-No Base");

                    // Update the lighting fixture type
                    foreach (FamilyInstance curFixture in lightingFixtures)
                    {
                        // Change the family type of the fixture
                        curFixture.Symbol = targetFamilySymbol;
                        updatedCount++;
                    }

                    // Add room name to updated rooms list
                    if (!updatedRooms.Contains(room.Name))
                    {
                        Parameter paramRoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                        updatedRooms.Add(paramRoomName.AsValueString());
                    }
                }
            }

            // return the counts & room list
            return (updatedRooms, updatedCount);
        }

        /// <summary>
        /// Gets rooms by name that contain the specified string and are visible in the active view
        /// </summary>        
        private static List<Room> GetRoomByNameContainsInActiveView(Document curDoc, View activeView, string roomNameContains)
        {
            List<Room> matchingRooms = new List<Room>();

            // Get all rooms visible in the active view
            FilteredElementCollector roomCollector = new FilteredElementCollector(curDoc, activeView.Id)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            foreach (Room room in roomCollector.Cast<Room>())
            {
                if (room.Name.IndexOf(roomNameContains, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    matchingRooms.Add(room);
                }
            }

            return matchingRooms;
        }

        /// <summary>
        /// Gets all light fixtures (family instances) in a specific room that are visible in the active view
        /// </summary>        
        private static List<FamilyInstance> GetLightFixtureInRoomInActiveView(Document curDoc, View activeView, Room room, string familyName = null)
        {
            List<FamilyInstance> m_lightFixtures = new List<FamilyInstance>();

            // Get all lighting fixtures visible in the active view
            var familyInstances = new FilteredElementCollector(curDoc, activeView.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_LightingFixtures)
                .Cast<FamilyInstance>();

            foreach (FamilyInstance curInstance in familyInstances)
            {
                // Check the Room parameter
                if (curInstance.Room != null && curInstance.Room.Id == room.Id)
                {
                    // Filter by family name if specified
                    if (string.IsNullOrEmpty(familyName) ||
                        curInstance.Symbol.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    {
                        m_lightFixtures.Add(curInstance);
                    }
                }
            }

            return m_lightFixtures;
        }

        /// <summary>
        /// Manages ceiling fan notes in specified rooms based on spec level conversion
        /// </summary>        
        public static (int totalAdded, int totalDeleted, int viewCount) ManageClgFanNotes(Document curDoc, UIDocument uidoc, string specLevel, List<View> viewsElectrical)
        {
            // define the variables
            int totalAdded = 0;
            int totalDeleted = 0;
            int viewCount = viewsElectrical.Count;

            // Define rooms that need note management
            List<string> roomsToUpdate = new List<string>
            {
                "Master Bedroom",
                "Covered Patio",
                "Gameroom",
                "Loft"
            };

            string noteText = "Block & pre-wire for clg fan";

            // loop through each view to ensure notes are added/removed in all relevant views
            foreach (View curView in viewsElectrical)
            {
                if (specLevel == "Complete Home Plus")
                {
                    // CHP to CH conversion - DELETE notes in all rooms
                    int deletedCount = DeleteCeilingFanNotes(curDoc, curView.Id, roomsToUpdate, noteText);

                    // increment the counter
                    totalDeleted += deletedCount;
                }
                else if (specLevel == "Complete Home")
                {
                    // CH to CHP conversion - ADD notes in all rooms EXCEPT Covered Patio
                    List<string> roomsForNotes = roomsToUpdate.Where(r => r != "Covered Patio").ToList();
                    int addedCount = AddCeilingFanNotes(curDoc, curView.Id, roomsForNotes, noteText);

                    // increment the counter
                    totalAdded += addedCount;
                }
            }

            // return the counts
            return (totalAdded, totalDeleted, viewCount);
        }

        /// <summary>
        /// Deletes ceiling fan notes from specified rooms
        /// </summary>       
        private static int DeleteCeilingFanNotes(Document curDoc, ElementId curViewId, List<string> roomNames, string noteText)
        {
            int deletedCount = 0;

            foreach (string roomName in roomNames)
            {
                // Get rooms containing this name
                List<Room> rooms = Utils.GetRoomByNameContains(curDoc, roomName);

                foreach (Room room in rooms)
                {
                    // Find text notes in this room
                    List<TextNote> notesToDelete = GetTextNotesInRoom(curDoc, curViewId, room, noteText);

                    // Delete each matching note
                    foreach (TextNote note in notesToDelete)
                    {
                        curDoc.Delete(note.Id);
                        deletedCount++;
                    }
                }
            }

            // return the count
            return deletedCount;
        }

        /// <summary>
        /// Adds ceiling fan notes to specified rooms
        /// </summary>       
        private static int AddCeilingFanNotes(Document curDoc, ElementId curViewId, List<string> roomNames, string noteText)
        {
            int addedCount = 0;

            // get the TextNoteType
            TextNoteType textNoteType = Utils.GetTextNoteTypeByName(curDoc, "STANDARD");

            // null check for the TextNoteType
            if (textNoteType == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Text Note Type 'STANDARD' not found in the project.");
                return 0;
            }

            foreach (string roomName in roomNames)
            {
                // Skip Covered Patio
                if (roomName == "Covered Patio")
                    continue;

                // Get rooms containing this name
                List<Room> rooms = Utils.GetRoomByNameContains(curDoc, roomName);

                foreach (Room curRoom in rooms)
                {
                    // Check if note already exists
                    List<TextNote> existingNotes = GetTextNotesInRoom(curDoc, curViewId, curRoom, noteText);
                    if (existingNotes.Count > 0)
                        continue; // Note already exists, skip

                    // insertion point for note placement
                    XYZ roomCenter = Utils.GetRoomCenterPoint(curRoom);
                    if (roomCenter != null)
                    {
                        // Get the view's directions
                        View currentView = curDoc.GetElement(curViewId) as View;
                        XYZ viewUp = currentView.UpDirection;

                        // Create position 2' down from room center relative to the view
                        XYZ notePosition = roomCenter - (viewUp * .5); // move 2' in the opposite of "up" direction

                        // Create the text note
                        TextNote newNote = TextNote.Create(curDoc, curViewId, notePosition, noteText, textNoteType.Id);

                        // set the justification
                        newNote.HorizontalAlignment = HorizontalTextAlignment.Center;

                        // set the width of the note
                        newNote.Width = .078250;

                        // increment the counter
                        addedCount++;
                    }
                }
            }

            // return the count
            return addedCount;
        }

        /// <summary>
        /// Gets text notes in a specific room that contain the specified text
        /// </summary>        
        /// <returns>List of matching text notes</returns>
        private static List<TextNote> GetTextNotesInRoom(Document curDoc, ElementId curViewId, Room room, string searchText)
        {
            List<TextNote> m_textNotes = new List<TextNote>();

            // Get all text notes in the document
            var textNotes = new FilteredElementCollector(curDoc, curViewId)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>();

            foreach (TextNote curNote in textNotes)
            {
                // Check if note text contains the search text
                if (curNote.Text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Check if note is in the room (using bounding box intersection)
                    if (IsTextNoteInRoom(curNote, room))
                    {
                        m_textNotes.Add(curNote);
                    }
                }
            }

            return m_textNotes;
        }

        /// <summary>
        /// Checks if a text note is within a room's boundaries
        /// </summary>
        /// <returns>True if text note is in room</returns>
        private static bool IsTextNoteInRoom(TextNote curNote, Room room)
        {
            try
            {
                XYZ notePosition = curNote.Coord;
                var roomAtPoint = room.Document.GetRoomAtPoint(notePosition);
                return roomAtPoint != null && roomAtPoint.Id == room.Id;
            }
            catch
            {
                return false;
            }
        }

        private void ManageSprinklerOutlet(Document curDoc, UIDocument uidoc, string selectedSpecLevel, Reference selectedSprinklerWall, Reference selectedGarageWall, Reference selectedOutlet)
        {
            if (selectedSpecLevel == "Complete Home")
                RemoveSprinklerOutlet(curDoc, selectedOutlet);
            else
                AddSprinklerOutlet(curDoc, uidoc, selectedSprinklerWall, selectedGarageWall);
        }

        private void RemoveSprinklerOutlet(Document curDoc, Reference selectedOutlet)
        {
            // delete the selected outlet
            Element curOutlet = curDoc.GetElement(selectedOutlet);
            if (curOutlet != null)
            {
                curDoc.Delete(curOutlet.Id);
            }
        }

        private void AddSprinklerOutlet(Document curDoc, UIDocument uidoc, Reference selectedSprinklerWall, Reference selectedGarageWall)
        {
            // Get selected Garage wall
            Wall garageWall = curDoc.GetElement(selectedGarageWall) as Wall;

            // Create options for geometry extraction - controls how detailed the geometry will be
            Options geometryOptions = new Options();

            // Get the wall's 3D geometry as a collection of solids and surfaces
            GeometryElement wallGeometry = garageWall.get_Geometry(geometryOptions);

            // Variables to store our target point when found
            XYZ targetPoint = null;
            bool foundTargetPoint = false;

            // Loop through the geometry to find the solid objects (the actual wall layers)
            foreach (GeometryObject geomObj in wallGeometry)
            {
                // Check if this geometry object is a solid with volume and not null
                if (geomObj is Solid solid && solid != null && solid.Volume > 0)
                {
                    // Get the material of this solid to identify if it's the stud layer
                    ElementId materialId = solid.GraphicsStyleId;
                    if (materialId != null && materialId != ElementId.InvalidElementId)
                    {
                        // Get the material element to check its name
                        Material material = curDoc.GetElement(materialId) as Material;
                        if (material != null && material.Name.Contains("2x"))
                        {
                            // Loop through all faces of this solid to find the structural core face
                            foreach (Face face in solid.Faces)
                            {
                                // Get the face normal vector (direction perpendicular to the face)
                                XYZ faceNormal = face.ComputeNormal(new UV(0.5, 0.5));

                                // Check if this face is vertical (Z component near 0)
                                if (Math.Abs(faceNormal.Z) < 0.1)
                                {
                                    // Get a point on the face to use for calculations
                                    XYZ pointOnFace = face.Evaluate(new UV(0.5, 0.5));

                                    // Check if this is the outside face by comparing with wall location line
                                    LocationCurve wallLocationCurve = garageWall.Location as LocationCurve;

                                    // Get the centerline of the wall
                                    Line wallCenterLine = wallLocationCurve.Curve as Line;

                                    // Get the closest point on the wall centerline to our face point
                                    XYZ closestPointOnCenterline = wallCenterLine.Project(pointOnFace).XYZPoint;

                                    // Calculate vector from centerline to face point
                                    XYZ centerToFace = pointOnFace - closestPointOnCenterline;

                                    // Check if this face is pointing away from the building interior (outside face)
                                    if (centerToFace.DotProduct(faceNormal) > 0)
                                    {
                                        // This is the outside face - calculate target point 60" perpendicular from this face
                                        targetPoint = pointOnFace + (faceNormal * 5.0);
                                        foundTargetPoint = true;
                                        break; // Exit face loop - we found our outside face
                                    }

                                    if (foundTargetPoint)
                                        break; // Exit geometry loop - we found our target point
                                }
                            }
                        }
                    }

                    // Check if we successfully found our target point
                    if (!foundTargetPoint || targetPoint == null)
                    {
                        Utils.TaskDialogError("Error", "Spec Conversion", "Could not find the outside face of the garage stud wall.");
                        return;
                    }

                    // get the outlet wall to project the tartget point onto
                    Wall outletWall = curDoc.GetElement(selectedSprinklerWall) as Wall;

                    // Get the outlet wall's geometry to find where to place the outlet
                    GeometryElement outletWallGeometry = outletWall.get_Geometry(geometryOptions);

                    // Variables to store the outlet placement point
                    XYZ outletPlacementPoint = null;
                    bool foundOutletFace = false;

                    // Loop through the outlet wall geometry to find the GWB layer
                    foreach (GeometryObject outletGeomObj in outletWallGeometry)
                    {
                        // Check if this geometry object is a solid with volume and not null
                        if (outletGeomObj is Solid outletSolid && outletSolid != null && outletSolid.Volume > 0)
                        {
                            // Get the material of this solid to identify if it's the GWB layer
                            ElementId outletMaterialId = outletSolid.GraphicsStyleId;

                            if (outletMaterialId != null && outletMaterialId != ElementId.InvalidElementId)
                            {
                                // Get the material element to check its name for GWB
                                Material outletMaterial = curDoc.GetElement(outletMaterialId) as Material;
                                if (outletMaterial != null && outletMaterial.Name.Contains("GWB"))
                                {
                                    // Loop through all faces of the GWB solid to find the inside face
                                    foreach (Face outletFace in outletSolid.Faces)
                                    {
                                        // Get the face normal vector (direction perpendicular to the face)
                                        XYZ outletFaceNormal = outletFace.ComputeNormal(new UV(0.5, 0.5));

                                        // Check if this face is vertical (Z component near 0)
                                        if (Math.Abs(outletFaceNormal.Z) < 0.1)
                                        {
                                            // Project our target point onto this face to see if it intersects
                                            IntersectionResult intersection = outletFace.Project(targetPoint);
                                            if (intersection != null)
                                            {
                                                // Get the intersection point - this is where the outlet will be placed
                                                outletPlacementPoint = intersection.XYZPoint;
                                                foundOutletFace = true;
                                                break; // Exit face loop - we found our intersection
                                            }

                                        }
                                    }

                                    if (foundOutletFace)
                                    {
                                        break; // Exit geometry loop - we found our outlet face
                                    }
                                }
                            }

                            // Check if we successfully found our outlet placement point
                            if (!foundOutletFace || outletPlacementPoint == null)
                            {
                                Utils.TaskDialogError("Error", "Spec Conversion", "Could not find intersection point on the outlet wall.");
                                return;
                            }

                            // Load the sprinkler outlet family
                            Family outletFamily = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical", "LD_EF_Recep_Wall");

                            // Check if family loaded successfully
                            if (outletFamily == null)
                            {
                                Utils.TaskDialogError("Error", "Spec Conversion", "Could not load sprinkler outlet family.");
                                return;
                            }

                            // Get the family symbol (type) to place
                            FamilySymbol outletSymbol = Utils.GetFamilySymbolByName(curDoc, "LD_EF_Recep_Wall", "Sprinkler");

                            // Check if the symbol was found and activate it if needed
                            if (outletSymbol == null)
                            {
                                Utils.TaskDialogError("Error", "Spec Conversion", "Could not find 'Sprinkler' type in the receptacle family.");
                                return;
                            }

                            // Activate the family symbol if it's not already active
                            if (!outletSymbol.IsActive)
                            {
                                outletSymbol.Activate();
                            }

                            // Place the sprinkler outlet at the calculated intersection point
                            FamilyInstance sprinklerOutlet = curDoc.Create.NewFamilyInstance(outletPlacementPoint, outletSymbol, outletWall, StructuralType.NonStructural);

                            // Check if the outlet was placed successfully
                            if (sprinklerOutlet == null)
                            {
                                Utils.TaskDialogError("Error", "Spec Conversion", "Failed to place sprinkler outlet.");
                                return;
                            }

                            // Create dimension from garage wall face to outlet center
                            // First, get references for dimensioning - garage wall face and outlet center
                            Reference refGarageWall = new Reference(garageWall);

                            // Get reference to the outlet's 'Center (Left/Right)' reference point
                            Reference refOutlet = null;
                            foreach (Reference familyRef in sprinklerOutlet.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
                            {
                                refOutlet = familyRef;
                                break; // Take the first (and likely only) Center (Left/Right) reference
                            }

                            // Check if we found the outlet reference
                            if (refOutlet == null)
                            {
                                Utils.TaskDialogError("Error", "Spec Conversion", "Could not find outlet center reference for dimensioning.");
                                return;
                            }

                            // Create a reference array for the dimension (from garage wall to outlet)
                            ReferenceArray dimReferenceArray = new ReferenceArray();
                            dimReferenceArray.Append(refGarageWall);
                            dimReferenceArray.Append(refOutlet);

                            // Create a line to define where the dimension will be placed (2' from outlet wall exterior)
                            XYZ dimensionOffset = (outletPlacementPoint - targetPoint).Normalize() * 2.0; // 2 feet perpendicular to outlet wall
                            XYZ dimStartPoint = targetPoint + dimensionOffset;
                            XYZ dimEndPoint = outletPlacementPoint + dimensionOffset;
                            Line dimensionLine = Line.CreateBound(dimStartPoint, dimEndPoint);

                            // Create the dimension
                            Dimension sprinklerDimension = curDoc.Create.NewDimension(curDoc.ActiveView, dimensionLine, dimReferenceArray);

                            // Create a tag for the sprinkler outlet with leader
                            // Determine outlet wall orientation and calculate tag position accordingly
                            LocationCurve outletWallLocation = outletWall.Location as LocationCurve;
                            Line outletWallLine = outletWallLocation.Curve as Line;
                            XYZ wallDirection = outletWallLine.Direction;

                            // Check if wall is more horizontal or vertical by comparing X and Y components
                            bool isHorizontalWall = Math.Abs(wallDirection.X) > Math.Abs(wallDirection.Y);

                            XYZ tagPosition;
                            if (isHorizontalWall)
                            {
                                // Horizontal wall: 2' left and 2' up
                                tagPosition = outletPlacementPoint + new XYZ(-2.0, 0, 2.0);
                            }
                            else
                            {
                                // Vertical wall: 2' right and 2' up  
                                tagPosition = outletPlacementPoint + new XYZ(2.0, 0, 2.0);
                            }

                            // Create the tag for the sprinkler outlet
                            IndependentTag sprinklerTag = IndependentTag.Create(curDoc, curDoc.ActiveView.Id, new Reference(sprinklerOutlet), true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, tagPosition);

                            // Set the leader to free end
                            sprinklerTag.LeaderEndCondition = LeaderEndCondition.Free;

                            // Success - sprinkler outlet added successfully
                            Utils.TaskDialogInformation("Complete", "Spec Conversion", "Sprinkler outlet added successfully.");
                        }
                    }
                }
            }
        }

        private void RemoveSprinklerOutletNote(Document curDoc, UIDocument uidoc, string selectedSpecLevel, List<View> firstFloorElecViews)
        {
            // Remove any existing sprinkler text notes from previous conversions
            var sprinklerNotes = new FilteredElementCollector(curDoc)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(note => note.Text.Contains("Sprinkler", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (TextNote curNote in sprinklerNotes)
            {
                curDoc.Delete(curNote.Id);
            }
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