using Autodesk.Revit.DB.Architecture;
using ConvertSpecLevel.Common;
using System.Linq;

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
            frmConvertSpecLevel curForm = new frmConvertSpecLevel(uidoc, curDoc)
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
            
            object selectedCabinet = curForm.SelectedCabinet; 
           

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

                // create variable to hold the wall cabinet family
                Family wallCabinetFamily = null;

                // load the wall cabinet family
                if (selectedSpecLevel == "Complete Home Plus")
                {
                    // load the wall cabinet family for Complete Home Plus spec level
                    wallCabinetFamily = Utils.LoadFamilyFromLibrary(curDoc, "LD_Cab_Wall_Complete Home Plus");
                }               

                // create transaction for cabinet updates
                using (Transaction t = new Transaction(curDoc, "Update Cabinets"))
                {
                    // start the third transaction
                    t.Start();

                    // revise the upper cabinets

                    // revise the MW cabinet

                    // add/remove the Ref Sp cabinet
                    if (selectedSpecLevel == "Complete Home")
                    {
                        RemoveRefSpCabinet(curDoc, selectedCabinet);
                    }
                    else
                    {
                        AddRefSpCabinet(curDoc, wallCabinetFamily);
                    }

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

        private void AddRefSpCabinet(Document curDoc, Family wallCabinetFamily)
        {
            
            throw new NotImplementedException();
        }

        private void RemoveRefSpCabinet(Document curDoc, object selectedCabinet)
        {
            throw new NotImplementedException();
        }

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
        /// Loads a door family from the library to ensure most recent version
        /// </summary>
        /// <param name="curDoc">The Revit document</param>
        /// <param name="typeDoor">The type of door ("Front" or "Rear")</param>
        /// <param name="specLevel">The spec level to determine which family to load</param>
        /// <returns>True if family loaded successfully, false if loading failed</returns>
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

            string familyPath = $@"S:\Shared Folders\Lifestyle USA Design\Library 2025\Doors\{familyName}.rfa";

            if (!System.IO.File.Exists(familyPath))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Door family file not found at: {familyPath}");
                return false;
            }

            try
            {
                var loadOptions = new FamilyLoadOptions();
                bool familyLoaded = curDoc.LoadFamily(familyPath, loadOptions, out Family loadedFamily);
                return familyLoaded;
            }
            catch (Exception ex)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error loading door family: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Family load options class to handle overwrite behavior
        /// </summary>
        public class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
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

        #region Front Door Update

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

            string newDoorTypeName = GetFrontDoorType(specLevel);
            if (string.IsNullOrEmpty(newDoorTypeName))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Unable to determine front door type for spec level: " + specLevel);
                return;
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

            bool fromRoomMatch = fromRoomName.Contains("Entry", StringComparison.OrdinalIgnoreCase) ||
                                 fromRoomName.Contains("Foyer", StringComparison.OrdinalIgnoreCase);
            bool toRoomMatch = toRoomName.Contains("Covered Porch", StringComparison.OrdinalIgnoreCase);

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

        #region Rear Door Update

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
            var widthParam = curDoor.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            bool widthMatch = widthParam != null &&
                              Math.Abs((widthParam.AsDouble() * 12.0) - 32.0) < 0.1;

            // Check if description contains "Exterior Entry"
            var descParam = curDoor.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
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
