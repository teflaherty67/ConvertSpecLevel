using Autodesk.Revit.DB.Architecture;
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
            foreach(FamilyInstance curDoor in allDoorsInView)
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
                        FamilySymbol materialSymbol = Utils.GetFamilySymbolByName(curDoc, materialFamily.Name, "Type 1");

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

                            // get the center of the door opening
                            XYZ doorCenter = wallLine.Evaluate(0.5, true);

                            // create a perpendicular vector
                            XYZ perpendicular = new XYZ(-wallDirection.Y, wallDirection.X, 0);

                            // create start point
                            XYZ startPoint = doorCenter-(perpendicular * (drWidthParam / 2));

                            // create end point
                            XYZ endPoint = doorCenter + (perpendicular * (drWidthParam / 2));

                            // create a line to place the break
                            Line breakLine = Line.CreateBound(startPoint, endPoint);

                            // create a new Floor Material instance
                            FamilyInstance newBreak = curDoc.Create.NewFamilyInstance(doorPoint, materialSymbol, workPlane, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

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
