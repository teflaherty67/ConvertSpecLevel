using Autodesk.Revit.DB.Architecture;
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

            // launch the form
            frmConvertSpecLevel curForm = new frmConvertSpecLevel()
            {
                Topmost = true,
            };

            curForm.ShowDialog();

            // check if user clicked OK
            if (curForm.DialogResult != true)
            {
                return Result.Cancelled;
            }

            // get user input from the form
            string selectedClient = curForm.GetSelectedClient();
            string selectedSpecLevel = curForm.GetSelectedSpecLevel();
            string selectedMWCabHeight = curForm.GetSelectedMWCabHeight();

            // create a transaction group
            using (TransactionGroup transGroup = new TransactionGroup(curDoc, "Convert Spec Level"))
            {
                // start the transaction group
                transGroup.Start();

                #region Floor Finish Updates

                // get the first floor annotation view & set it as the active view
                View curView = Utils.GetViewByNameContainsAndAssociatedLevel(curDoc, "Annotation", "First Floor");

                if (curView != null)
                {
                    uidoc.ActiveView = curView;
                }
                else
                {
                    // if not found alert the user
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No view found with name containing 'Annotation' and associated level 'First Floor'");                                       
                }

                // create transaction for flooring update
                using (Transaction t = new Transaction(curDoc, "Update Floor Finish"))
                {
                    // start the first transaction
                    t.Start();

                    // change the flooring for the specified rooms per the selected spec level
                    List<string> listUpdatedRooms = UpdateFloorFinishInActiveView(curDoc, selectedSpecLevel);

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
                    TaskDialog tdDrUpdate = new TaskDialog("Complete");
                    tdDrUpdate.MainIcon = Icon.TaskDialogIconInformation;
                    tdDrUpdate.Title = "Spec Conversion";
                    tdDrUpdate.TitleAutoPrefix = false;
                    tdDrUpdate.MainContent = "The front and rear doors were replaced per the specified spec level.";
                    tdDrUpdate.CommonButtons = TaskDialogCommonButtons.Close;

                    TaskDialogResult tdDrUpdateSuccess = tdDrUpdate.Show();

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Cabinet Updates

                // get the Interior Elevations sheet & set it as the active view
                ViewSheet sheetIntr = Utils.GetViewSheetByName(curDoc, "Interior Elevations");

                if (sheetIntr != null)
                {
                    uidoc.ActiveView = sheetIntr;
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Interior Elevation sheet found");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for cabinet updates
                using (Transaction t = new Transaction(curDoc, "Update Cabinets"))
                {
                    // start the third transaction
                    t.Start();

                    // revise the upper cabinets

                    // revise the MW cabinet

                    // add/remove the Ref Sp cabinet

                    // raise/lower the backsplash height

                    // notify the user
                    // upper cabinets were revised per the selected spec level
                    // Ref Sp cabinet was added/removed per the selected spec level
                    // backsplash height was raised/lowered per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region First Floor Electrical Updates

                // get all views with Electrical in the name & associated with the First Floor
                List<View> firstFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor");

                // check for Second Floor Electrical Plan views

                // if found
                // string nameView = "First Floor Electrical Plan";

                // if not found
                // string nameView = "Electrical Plan";

                // get the first view in the list and set it as the active view
                if (firstFloorElecViews.Any())
                {
                    uidoc.ActiveView = firstFloorElecViews.First();
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Electrical views found for First Floor");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for first floor electrical updates
                using (Transaction t = new Transaction(curDoc, "Update First Floor Electrical"))
                {
                    // start the fourth transaction
                    t.Start();

                    // replace the light fixtures in the specified rooms per the selected spec level
                    Utils.UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                    // make a list of the rooms that were updated

                    // add/remove the sprinkler outlet in the Garage

                    // loop through all the views
                    foreach (View curElecView in firstFloorElecViews)
                    {
                        // add/remove ceiling fan note

                        // add/remove sprinkler outlet note                        
                    }

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {View Name} per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Second Floor Electrical Updates

                // get all views with Electrical in the name & associated with the Second Floor
                List<View> secondFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "Second Floor");

                // set the view name variable
                string nameView = "Second Floor Electrical Plan";

                // get the first view in the list and set it as the active view
                if (secondFloorElecViews.Any())
                {
                    uidoc.ActiveView = secondFloorElecViews.First();
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Electrical views found for Second Floor");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for Second Floor Electrical updates
                using (Transaction t = new Transaction(curDoc, "Update Second Floor Electrical"))
                {

                    // start the transaction
                    t.Start();

                    // replace the light fixtures in the specified rooms per the selected spec level
                    Utils.UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                    // loop through the views and add/remove the clg fan note
                    foreach (View curElecView in secondFloorElecViews)
                    {
                        // add/remove ceiling fan note                        
                    }

                    // make a list of the rooms that were updated            

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {nameView} per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                // commit the transaction group
                transGroup.Assimilate(); // this will commit all the transactions in the group
            }

            return Result.Succeeded;

            // notify user conversion successful
        }

        #region Front Door Update

        /// <summary>
        /// Updates the front door type based on spec level
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="specLevel">The spec level selection</param>
        public static void UpdateFrontDoor(Document curDoc, string specLevel)
        {
            // Find the front door
            FamilyInstance frontDoor = GetFrontDoor(curDoc);

            if (frontDoor == null)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "Front door not found.");
                return;
            }

            // Always load the door family from library to get the most recent version
            if (!LoadDoorFamilyFromLibrary(curDoc))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to load door family from library.");
                return;
            }

            // Determine the new door type based on spec level
            string newDoorTypeName = GetFrontDoorType(specLevel);
            if (string.IsNullOrEmpty(newDoorTypeName))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to determine front door type for spec level: " + specLevel);
                return;
            }

            // Find the new door family symbol
            FamilySymbol newDoorSymbol = FindDoorSymbol(curDoc, newDoorTypeName);
            if (newDoorSymbol == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Door type '{newDoorTypeName}' not found in the project after loading family.");
                return;
            }

            // Activate the symbol if needed
            if (!newDoorSymbol.IsActive)
            {
                newDoorSymbol.Activate();
            }           

            // Change the door type
            frontDoor.Symbol = newDoorSymbol;

            // notify the user
            Utils.TaskDialogInformation("Complete", "Spec Conversion", $"Front door updated to '{newDoorTypeName}' for {specLevel} spec level.");
        }

        /// <summary>
        /// Always loads the front door family from the library to ensure most recent version
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <returns>True if family loaded successfully, false if loading failed</returns>
        private static bool LoadDoorFamilyFromLibrary(Document curDoc)
        {
            string familyPath = @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Doors\LD_DR_Ext_Single 3_4 Lite_1 Panel.rfa";

            // Check if family file exists
            if (!System.IO.File.Exists(familyPath))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Family file not found at: {familyPath}");
                return false;
            }

            try
            {
                // Create family load options that overwrite existing family and all parameters
                FamilyLoadOptions loadOptions = new FamilyLoadOptions();

                // Always load the family from library with overwrite options
                bool familyLoaded = curDoc.LoadFamily(familyPath, loadOptions, out Family loadedFamily);

                return true; // consider success regardless of return value
            }
            catch (Exception ex)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error loading family: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Family load options class to handle overwrite behavior
        /// </summary>
        public class FamilyLoadOptions : IFamilyLoadOptions
        {
            /// <summary>
            /// Called when a family being loaded already exists in the project
            /// </summary>
            /// <param name="familyInUse">True if family is in use in the project</param>
            /// <param name="overwriteParameterValues">Set to true to overwrite parameter values</param>
            /// <returns>True to overwrite the existing family</returns>
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                // Always overwrite existing family and all its parameter values
                overwriteParameterValues = true;
                return true;
            }

            /// <summary>
            /// Called when shared parameters are found
            /// </summary>
            /// <param name="sharedParameters">The shared parameters</param>
            /// <param name="overwriteParameterValues">Set to true to overwrite parameter values</param>
            /// <returns>True to continue loading</returns>
            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                // Use the new family from the library and overwrite parameter values
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }

        /// <summary>
        /// Finds the front door based on room relationships
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <returns>The front door instance or null if not found</returns>
        public static FamilyInstance GetFrontDoor(Document curDoc)
        {
            // Get all doors in the document
            List<FamilyInstance> allDoors = GetAllDoors(curDoc);

            foreach (FamilyInstance curDoor in allDoors)
            {
                // Get the FromRoom and ToRoom properties
                Room fromRoom = curDoor.FromRoom;
                Room toRoom = curDoor.ToRoom;

                if (fromRoom != null && toRoom != null)
                {
                    string fromRoomName = fromRoom.Name;
                    string toRoomName = toRoom.Name;

                    // Check if this matches front door criteria
                    if (IsFrontDoorMatch(fromRoomName, toRoomName))
                    {
                        return curDoor;
                    }
                }
            }

            return null; // Front door not found
        }

        /// <summary>
        /// Checks if the room names match front door criteria
        /// </summary>
        /// <param name="fromRoomName">The "From Room: Name" value</param>
        /// <param name="toRoomName">The "To Room: Name" value</param>
        /// <returns>True if this appears to be the front door</returns>
        private static bool IsFrontDoorMatch(string fromRoomName, string toRoomName)
        {
            if (string.IsNullOrEmpty(fromRoomName) || string.IsNullOrEmpty(toRoomName))
                return false;

            // Check if From Room is Entry or Foyer
            bool fromRoomMatch = fromRoomName.IndexOf("Entry", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                fromRoomName.IndexOf("Foyer", StringComparison.OrdinalIgnoreCase) >= 0;

            // Check if To Room is Covered Porch
            bool toRoomMatch = toRoomName.IndexOf("Covered Porch", StringComparison.OrdinalIgnoreCase) >= 0;

            return fromRoomMatch && toRoomMatch;
        }

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
        /// Gets the front door type name based on spec level
        /// </summary>
        /// <param name="specLevel">The spec level</param>
        /// <returns>The door type name</returns>
        private static string GetFrontDoorType(string specLevel)
        {
            return specLevel switch
            {
                "Complete Home" => "36\"x80\"",     // CH uses 36"x80"
                "Complete Home Plus" => "36\"x96\"", // CHP uses 36"x96"
                _ => null
            };
        }

        /// <summary>
        /// Finds a door symbol by type name within the specific family
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="typeName">The door type name (e.g., "36\"x80\"")</param>
        /// <returns>The door symbol or null if not found</returns>
        private static FamilySymbol FindDoorSymbol(Document curDoc, string typeName)
        {
            string familyName = "LD_DR_Ext_Single 3_4 Lite_1 Panel";

            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(ds => ds.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                      ds.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
        }

        #endregion

        #region Rear Door Update

        /// <summary>
        /// Updates the rear door type based on spec level
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="specLevel">The spec level selection</param>
        public static void UpdateRearDoor(Document curDoc, string specLevel)
        {
            // Find the rear door
            FamilyInstance rearDoor = GetRearDoor(curDoc);

            if (rearDoor == null)
            {
                Utils.TaskDialogWarning("Warning", "Spec Conversion", "Rear door not found.");
                return;
            }

            // Always load the door family from library to get the most recent version
            if (!LoadRearDoorFamilyFromLibrary(curDoc, specLevel))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to load rear door family from library.");
                return;
            }

            // Get the door type name (always 32"x80" for both spec levels)
            string newDoorTypeName = GetRearDoorType();

            // Find the new door family symbol
            FamilySymbol newDoorSymbol = FindRearDoorSymbol(curDoc, newDoorTypeName, specLevel);
            if (newDoorSymbol == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Door type '{newDoorTypeName}' not found in the project after loading family.");
                return;
            }

            // Activate the symbol if needed
            if (!newDoorSymbol.IsActive)
            {
                newDoorSymbol.Activate();
            }

            // Store original swing parameter value
            Parameter swingParam = rearDoor.LookupParameter("Swing");
            string originalSwing = swingParam?.AsString();

            // Change the door type
            rearDoor.Symbol = newDoorSymbol;           

            string familyName = GetRearDoorFamilyName(specLevel);
            Utils.TaskDialogInformation("Complete", "Spec Conversion", $"Rear door updated to '{familyName} - {newDoorTypeName}' for {specLevel} spec level.");
        }

        /// <summary>
        /// Always loads the rear door family from the library to ensure most recent version
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="specLevel">The spec level to determine which family to load</param>
        /// <returns>True if family loaded successfully, false if loading failed</returns>
        private static bool LoadRearDoorFamilyFromLibrary(Document curDoc, string specLevel)
        {
            string familyName = GetRearDoorFamilyName(specLevel);
            string familyPath = $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Doors\{familyName}.rfa";

            // Check if family file exists
            if (!System.IO.File.Exists(familyPath))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Rear door family file not found at: {familyPath}");
                return false;
            }

            try
            {
                // Create family load options that overwrite existing family and all parameters
                FamilyLoadOptions loadOptions = new FamilyLoadOptions();

                // Always load the family from library with overwrite options
                bool familyLoaded = curDoc.LoadFamily(familyPath, loadOptions, out Family loadedFamily);

                return true; // Consider success regardless of return value
            }
            catch (Exception ex)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error loading rear door family: {ex.Message}");
                return false;
            }
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
                "Complete Home" => "LD_DR_Ext_Single_Half Lite_2 Panel",        // CH uses Half Lite 2 Panel
                "Complete Home Plus" => "LD_DR_Ext_Single_Full Lite",           // CHP uses Full Lite
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
            // Get all doors in the document
            List<FamilyInstance> allDoors = GetAllDoors(curDoc);

            foreach (FamilyInstance curDoor in allDoors)
            {
                // Check if this matches rear door criteria
                if (IsRearDoorMatch(curDoor))
                {
                    return curDoor;
                }
            }

            return null; // Rear door not found
        }

        /// <summary>
        /// Checks if the door matches rear door criteria
        /// </summary>
        /// <param name="curDoor">The door to check</param>
        /// <returns>True if this appears to be the rear door</returns>
        private static bool IsRearDoorMatch(FamilyInstance curDoor)
        {
            // Check if door width is 32"
            bool widthMatch = false;
            Parameter widthParam = curDoor.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            if (widthParam != null)
            {
                // Convert from internal units (feet) to inches and check if it's 32"
                double widthInFeet = widthParam.AsDouble();
                double widthInInches = widthInFeet * 12.0;

                // Check if width is approximately 32" (allowing for small tolerance)
                widthMatch = Math.Abs(widthInInches - 32.0) < 0.1;
            }

            // Check if description contains "Exterior Entry"
            bool descriptionMatch = false;
            Parameter descParam = curDoor.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (descParam != null)
            {
                string description = descParam.AsString() ?? "";
                descriptionMatch = description.IndexOf("Exterior Entry", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            // Door must be 32" wide AND have "Exterior Entry" in description
            return widthMatch && descriptionMatch;
        }

        /// <summary>
        /// Gets the rear door type name (always 32"x80" for both spec levels)
        /// </summary>
        /// <returns>The door type name</returns>
        private static string GetRearDoorType()
        {
            return "32\"x80\"";     // Both CH and CHP use 32"x80"
        }

        /// <summary>
        /// Finds a rear door symbol by type name within the specific family for the spec level
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="typeName">The door type name (32"x80")</param>
        /// <param name="specLevel">The spec level to determine which family to search</param>
        /// <returns>The door symbol or null if not found</returns>
        private static FamilySymbol FindRearDoorSymbol(Document curDoc, string typeName, string specLevel)
        {
            string familyName = GetRearDoorFamilyName(specLevel);

            return new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(ds => ds.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                      ds.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
        }

        #endregion

        #region Finish Floor Methods

        /// <summary>
        /// Updates the Floor Finish parameter for specified room types in the active view based on the selected specification level.
        /// </summary>
        /// <param name="curDoc">The current Revit document.</param>
        /// <param name="selectedSpecLevel">The specification level ("Complete Home" or "Complete Home Plus") that determines the floor finish value.</param>
        /// <remarks>
        /// This method updates the following room types: Master Bedroom, Family, and Hall.
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
        /// <param name="curDoc">The current Revit document.</param>
        /// <param name="RoomsToUpdate">List of room name strings to search for (case-insensitive matching).</param>
        /// <returns>A list of Room elements in the active view whose names contain any of the specified strings.</returns>
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

        #endregion

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
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
