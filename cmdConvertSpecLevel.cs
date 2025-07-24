using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Common;
using System.ComponentModel;
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

            #region Form

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

            Reference selectedSprinklerWall = curForm.SelectedOutletWall;
            Reference selectedGarageWall = curForm.SelectedGarageWall;
            Reference selectedRefSpWall = curForm.SelectedRefSpWall;
            FamilyInstance selectedRefSp = curForm.SelectedRefSp;
            Reference selectedOutlet = curForm.SelectedOutlet;

            #endregion

            #region Transaction Group

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

                    // add/remove the Ref Sp cabinet
                    if (selectedSpecLevel == "Complete Home" && curForm.SelectedCabinet != null)
                    {
                        curDoc.Delete(((Element)curForm.SelectedCabinet).Id);
                    }
                    else
                    {
                        AddRefSpCabinet(curDoc, uidoc, selectedRefSpWall, selectedRefSp);
                    }

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

                // get all electrical plan views
                List<View> electricalViews = Utils.GetAllViewsByNameContains(curDoc, "Electrical");

                // verify if project is two story
                bool isPlanTwoStory = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .Any(level => level.Name.Contains("Second Floor"));

                // load the new electrical families (outlet & light fixture)
                Family lightSymbol = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Lighting", "LD_LF_No Base");
                Family outletSymbol = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical", "LD_EF_Wall Base");

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
                        UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                        // add/remove the sprinkler outlet in the Garage
                        ManageSprinklerOutlet(curDoc, uidoc, selectedSpecLevel, selectedSprinklerWall, selectedGarageWall, selectedOutlet);

                        // add/remove the ceiling fan note in the views
                        ManageClgFanNotes(curDoc, uidoc, selectedSpecLevel, firstFloorElecViews);

                        // add/remove sprinkler outlet note
                        RemoveSprinklerOutletNote(curDoc, uidoc, selectedSpecLevel, firstFloorElecViews);

                        // commit the transaction
                        t.Commit();
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
                    UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                    // add/remove the ceiling fan note in the views
                    ManageClgFanNotes(curDoc, uidoc, selectedSpecLevel, secondFloorElecViews);

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                // commit the transaction group
                transGroup.Assimilate(); // this will commit all the transactions in the group
            }

            #endregion

            return Result.Succeeded;

            // notify user conversion successful
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
                var loadOptions = new Utils.FamilyLoadOptions();
                bool familyLoaded = curDoc.LoadFamily(familyPath, loadOptions, out Family loadedFamily);
                return familyLoaded;
            }
            catch (Exception ex)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error loading door family: {ex.Message}");
                return false;
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
                if (curCabinet.Symbol.Family.Name.Contains("Single"))
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
                else
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
                .Where(cab => cab.Symbol.Family.Name.Contains("Wall") &&
                              (cab.Symbol.Name.Contains("Single") || cab.Symbol.Name.Contains("Double")) &&
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

        private void AddRefSpCabinet(Document curDoc, UIDocument uidoc, Reference selectedRefSpWall, FamilyInstance selectedRefSp)
        {
            // get the wall where the Ref Sp cabinet will be added
            Wall wallRefSp = curDoc.GetElement(selectedRefSpWall) as Wall;

            // create a variable to hold the Ref Sp's Center (Left/Right) reference for placement calculations
            Reference refCenterLR = null;

            // loop through the selected Ref Sp cabinet references to find the Center (Left/Right) reference
            foreach (Reference familyRef in selectedRefSp.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
            {
                // Store the centerline reference from the refrigerator
                refCenterLR = familyRef;
                // Take the first (and likely only) Center (Left/Right) reference
                break;
            }
            // null check for the Center (Left/Right) reference
            if (refCenterLR == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Could not fins centerline reference fore the refrigerator.");
                return;
            }

            // Load the Ref Sp cabinet family from the library
            Family refSpFamily = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Casework\Kitchen", "LD_CW_Wall_2-Dr_Flush");

            // Check if the cabinet family loaded successfully
            if (refSpFamily == null)
            {
                // Show error message if cabinet family failed to load
                Utils.TaskDialogError("Error", "Spec Conversion", "Could not load Ref Sp cabinet family from library.");
                return;
            }

            // Get the cabinet family type for placement
            FamilySymbol refSpSymbol = Utils.GetFamilySymbolByName(curDoc, "LD_CW_Wall_2-Dr_Flush", "39\"x27\"x15\"");

            throw new NotImplementedException();
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
                foreach(ViewSection curIntElev in allIntElevs)
                {
                    // set the active view
                    uidoc.ActiveView = curIntElev;

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
            List<ViewSection> m_allIntElevs= Utils.GetAllSectionViews(curDoc)
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
        private static void UpdateLightingFixturesInActiveView(Document curDoc, string selectedSpecLevel)
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
                return;
            }

            // Get the active view
            View activeView = curDoc.ActiveView;
            if (activeView == null)
            {
                TaskDialog.Show("Error", "No active view found.");
                return;
            }

            // Find target family symbol
            FamilySymbol targetFamilySymbol = Utils.FindFamilySymbol(curDoc, "LT-No Base", targetFamilyType);
            if (targetFamilySymbol == null)
            {
                TaskDialog.Show("Error", $"Family symbol '{targetFamilyType}' not found.");
                return;
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
                        updatedRooms.Add(room.Name);
                    }
                }
            }

            // Show summary message with proper grammar
            string roomList;
            if (updatedRooms.Count == 0)
            {
                roomList = "No rooms";
            }
            else if (updatedRooms.Count == 1)
            {
                roomList = updatedRooms[0];
            }
            else if (updatedRooms.Count == 2)
            {
                roomList = $"{updatedRooms[0]} and {updatedRooms[1]}";
            }
            else
            {
                roomList = string.Join(", ", updatedRooms.Take(updatedRooms.Count - 1)) + $", and {updatedRooms.Last()}";
            }

            string message = $"Updated {updatedCount} light fixtures to '{targetFamilyType}' in: {roomList}";
            Utils.TaskDialogInformation("Complete", "Spec Conversion", message);
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
        public static void ManageClgFanNotes(Document curDoc, UIDocument uidoc, string specLevel, List<View> firstFloorElecViews)
        {
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
            foreach (View curView in firstFloorElecViews)
            {
                // Set the active view
                uidoc.ActiveView = curView;

                if (specLevel == "Complete Home Plus")
                {
                    // CHP to CH conversion - DELETE notes in all rooms
                    DeleteCeilingFanNotes(curDoc, roomsToUpdate, noteText);
                }
                else if (specLevel == "Complete Home")
                {
                    // CH to CHP conversion - ADD notes in all rooms EXCEPT Covered Patio
                    List<string> roomsForNotes = roomsToUpdate.Where(r => r != "Covered Patio").ToList();
                    AddCeilingFanNotes(curDoc, roomsForNotes, noteText);
                }                
            }
        }

        /// <summary>
        /// Deletes ceiling fan notes from specified rooms
        /// </summary>       
        private static void DeleteCeilingFanNotes(Document curDoc, List<string> roomNames, string noteText)
        {
            int deletedCount = 0;

            foreach (string roomName in roomNames)
            {
                // Get rooms containing this name
                List<Room> rooms = Utils.GetRoomByNameContains(curDoc, roomName);

                foreach (Room room in rooms)
                {
                    // Find text notes in this room
                    List<TextNote> notesToDelete = GetTextNotesInRoom(curDoc, room, noteText);

                    // Delete each matching note
                    foreach (TextNote note in notesToDelete)
                    {
                        curDoc.Delete(note.Id);
                        deletedCount++;
                    }
                }
            }

            if (deletedCount > 0)
            {
                TaskDialog.Show("Notes Deleted", $"Deleted {deletedCount} ceiling fan notes.");
            }
        }

        /// <summary>
        /// Adds ceiling fan notes to specified rooms
        /// </summary>       
        private static void AddCeilingFanNotes(Document curDoc, List<string> roomNames, string noteText)
        {
            int addedCount = 0;

            // get the TextNoteType
            TextNoteType textNoteType = Utils.GetTextNoteTypeByName(curDoc, "STANDARD");

            // null check for the TextNoteType
            if (textNoteType == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Text Note Type 'STANDARD' not found in the project.");
                return;
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
                    List<TextNote> existingNotes = GetTextNotesInRoom(curDoc, curRoom, noteText);
                    if (existingNotes.Count > 0)
                        continue; // Note already exists, skip

                    // insertion point for note placement
                    XYZ roomCenter = Utils.GetRoomCenterPoint(curRoom);
                    if (roomCenter != null)
                    {
                        // Create point 2' below room center
                        XYZ notePosition = new XYZ(roomCenter.X, roomCenter.Y - 2.0, roomCenter.Z);

                        // Create the text note
                        TextNote.Create(curDoc, curDoc.ActiveView.Id, notePosition, noteText, textNoteType.Id);
                        addedCount++;
                    }
                }
            }

            if (addedCount > 0)
            {
                TaskDialog.Show("Notes Added", $"Added {addedCount} ceiling fan notes.");
            }
        }

        /// <summary>
        /// Gets text notes in a specific room that contain the specified text
        /// </summary>        
        /// <returns>List of matching text notes</returns>
        private static List<TextNote> GetTextNotesInRoom(Document curDoc, Room room, string searchText)
        {
            List<TextNote> m_textNotes = new List<TextNote>();

            // Get all text notes in the document
            var textNotes = new FilteredElementCollector(curDoc)
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
