
using Autodesk.Revit.DB.Architecture;

namespace ConvertSpecLevel.Common
{
    internal static class Utils
    {
        #region Elements - Architectural

        //return list of all doors in the current model
        public static List<FamilyInstance> GetAllDoors(Document curDoc)
        {
            //get all doors
            var returnList = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>()
                .ToList();

            return returnList;
        }

        internal static List<FamilyInstance> GetAllDoorsInActiveView(Document curDoc)
        {
            // get all the doors in the current view
           var m_returnList = new FilteredElementCollector(curDoc, curDoc.ActiveView.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>()
                .ToList();

            // return the list
            return m_returnList;
        }



        //return list of all windows in the current model
        public static List<FamilyInstance> GetAllWindows(Document curDoc)
        {
            //get all windows
            var returnList = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            return returnList;
        }

        #endregion

        #region Families

        /// <summary>
        /// Finds a family symbol by family name and type name
        /// </summary>        
        /// <returns>The family symbol or null if not found</returns>
        internal static FamilySymbol FindFamilySymbol(Document curDoc, string familyName, string typeName)
        {
            return new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs => fs.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                     fs.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
        }

        internal static Family LoadFamilyFromLibrary(Document curDoc, String filePath, string familyName)
        {
            // create the full path to the family file
            string familyPath = Path.Combine(filePath, familyName + ".rfa");

            // Check if the family file exists at the specified path
            if (!System.IO.File.Exists(familyPath))
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Family file not found at: {familyPath}");
                return null;
            }

            try
            {
                var loadOptions = new FamilyLoadOptions();
                curDoc.LoadFamily(familyPath, loadOptions, out Family loadedFamily);
                return loadedFamily; // This will be null if loading failed
            }
            catch (Exception ex)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error loading family: {ex.Message}");
                return null; // Return null if an error occurs during loading
            }
        }

        internal static FamilySymbol GetFamilySymbolByName(Document curDoc, string familyName, string typeName)
        {
            List<Family> m_famList = GetAllFamilies(curDoc);

            // loop through families in current document and look for match
            foreach (Family curFam in m_famList)
            {
                if (curFam.Name == familyName)
                {
                    // get family symbol from family
                    ISet<ElementId> fsList = curFam.GetFamilySymbolIds();

                    // loop through family symbol ids and look for match
                    foreach (ElementId fsID in fsList)
                    {
                        FamilySymbol fs = curDoc.GetElement(fsID) as FamilySymbol;

                        if (fs.Name == typeName)
                        {
                            return fs;
                        }
                    }
                }
            }

            return null;
        }

        private static List<Family> GetAllFamilies(Document curDoc)
        {
            List<Family> m_returnList = new List<Family>();

            FilteredElementCollector m_colFamilies = new FilteredElementCollector(curDoc)
                .OfClass(typeof(Family));

            foreach (Family family in m_colFamilies)
            {
                m_returnList.Add(family);
            }

            return m_returnList;
        }

        public static List<FamilyInstance> GetAllGenericFamilies(Document curDoc)
        {
            ElementClassFilter m_famFilter = new ElementClassFilter(typeof(FamilyInstance));
            ElementCategoryFilter m_typeFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
            LogicalAndFilter andFilter = new LogicalAndFilter(m_famFilter, m_typeFilter);

            FilteredElementCollector m_colGM = new FilteredElementCollector(curDoc);
            m_colGM.WherePasses(andFilter);

            List<FamilyInstance> m_famList = new List<FamilyInstance>();

            foreach (FamilyInstance curFam in m_colGM)
            {
                m_famList.Add(curFam);
            }

            return m_famList;
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

        #endregion

        #region Ribbon Panel

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        #endregion

        #region Rooms

        /// <summary>
        /// Retrieves all rooms from the document whose names contain the specified string,
        /// and filters out rooms with zero or invalid area.
        /// </summary>        
        internal static List<Room> GetRoomByNameContains(Document curDoc, string nameRoom)
        {
            // Retrieve all rooms in the document
            List<Room> m_roomList = GetAllRooms(curDoc);

            // Initialize the list to hold the matching rooms
            List<Room> m_returnList = new List<Room>();

            // Iterate through all rooms
            foreach (Room curRoom in m_roomList)
            {
                // Check if the room name contains the specified substring
                if (curRoom != null &&
                curRoom.Name != null &&
                curRoom.Name.IndexOf(nameRoom, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Check if the room has a valid area (greater than 0)
                    if (curRoom.Area > 0)
                    {
                        // Add the room to the result list
                        m_returnList.Add(curRoom);
                    }
                }
            }

            // Return the filtered list of rooms
            return m_returnList;
        }


        /// <summary>
        /// Retrieves all <see cref="Room"/> elements from the specified Revit document.
        /// </summary>        
        public static List<Room> GetAllRooms(Document curDoc)
        {
            return new FilteredElementCollector(curDoc)         // Initialize a collector for the given document
                .OfCategory(BuiltInCategory.OST_Rooms)           // Filter elements to include only Rooms
                .Cast<Room>()                                    // Cast the elements to Room type
                .ToList();                                       // Convert the collection to a List<Room> and return
        }

        /// <summary>
        /// Gets the center point of a room for note placement
        /// </summary>
        /// <returns>The center point or null if not found</returns>
        internal static XYZ GetRoomCenterPoint(Room curRoom)
        {
            try
            {
                LocationPoint location = curRoom.Location as LocationPoint;
                return location?.Point;
            }
            catch
            {
                // Fallback: use bounding box center
                var bbox = curRoom.get_BoundingBox(null);
                if (bbox != null)
                {
                    return (bbox.Min + bbox.Max) / 2;
                }
                return null;
            }
        }

        #endregion

        #region Schedules

        /// <summary>
        /// Retrieves the first schedule that contains the specified string in its name
        /// </summary>
        /// <param name="curDoc">The current Revit document</param>
        /// <param name="scheduleString">The string to search for in schedule names</param>
        /// <returns>The first matching ViewSchedule, or null if no match is found</returns>
        internal static ViewSchedule GetScheduleByNameContains(Document curDoc, string scheduleString)
        {
            // Get all schedules in the document
            List<ViewSchedule> m_scheduleList = GetAllSchedules(curDoc);

            // Loop through each schedule to find one containing the specified string
            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                // Check if the schedule name contains the search string
                if (curSchedule.Name.Contains(scheduleString))
                    return curSchedule;
            }

            // Return null if no matching schedule is found
            return null;
        }

        /// <summary>
        /// Retrieves all schedules from the document, excluding templates and revision schedules
        /// </summary>
        /// <param name="curDoc">The current Revit document</param>
        /// <returns>A list of all ViewSchedule elements in the document</returns>
        internal static List<ViewSchedule> GetAllSchedules(Document curDoc)
        {
            // Initialize the list to hold the schedules
            List<ViewSchedule> m_schedList = new List<ViewSchedule>();

            // Create a collector to find all ViewSchedule elements
            FilteredElementCollector curCollector = new FilteredElementCollector(curDoc);
            curCollector.OfClass(typeof(ViewSchedule));
            curCollector.WhereElementIsNotElementType();

            // Loop through views and check if schedule - if so then put into schedule list
            foreach (ViewSchedule curView in curCollector)
            {
                // Check if the view is actually a schedule type
                if (curView.ViewType == ViewType.Schedule)
                {
                    // Skip schedule templates
                    if (curView.IsTemplate == false)
                    {
                        // Skip revision schedules (which have < > in their names)
                        if (curView.Name.Contains("<") && curView.Name.Contains(">"))
                            continue;
                        else
                            // Add the valid schedule to our list
                            m_schedList.Add((ViewSchedule)curView);
                    }
                }
            }

            // Return the filtered list of schedules
            return m_schedList;
        }

        #endregion

        #region Sheets

        internal static ViewSheet GetSheetByName(Document curDoc, string sheetName)
        {
            //get all sheets
            List<ViewSheet> curSheets = GetAllSheets(curDoc);

            //loop through sheets and check sheet name
            foreach (ViewSheet curSheet in curSheets)
            {
                if (curSheet.Name == sheetName)
                {
                    return curSheet;
                }
            }

            return null;
        }

        internal static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector m_colViews = new FilteredElementCollector(curDoc);
            m_colViews.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> m_sheets = new List<ViewSheet>();
            foreach (ViewSheet x in m_colViews.ToElements())
            {
                m_sheets.Add(x);
            }

            return m_sheets;
        }

        #endregion      

        #region Task Dialog

        /// <summary>
        /// Displays a warning dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogWarning(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconWarning;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        /// <summary>
        /// Displays an information dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogInformation(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconInformation;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        /// <summary>
        /// Displays an error dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogError(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconError;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        #endregion

        #region Text Notes

        public static List<TextNoteType> GetAllTextNoteTypes(Document curDoc)
        {
            List<TextNoteType> returnList = new FilteredElementCollector(curDoc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .ToList();

            return returnList;
        }

        public static TextNoteType GetTextNoteTypeByName(Document curDoc, string name)
        {
            List<TextNoteType> textNoteList = GetAllTextNoteTypes(curDoc).ToList();

            return textNoteList.FirstOrDefault(curStyle => curStyle.Name == name);
        }

        #endregion

        #region Views

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in m_colviews.ToElements())
            {
                m_views.Add(x);
            }

            return m_views;
        }

        public static List<View> GetAllSectionViews(Document m_doc)
        {
            //get all views
            FilteredElementCollector m_colViews = new FilteredElementCollector(m_doc);
            m_colViews.OfCategory(BuiltInCategory.OST_Views);
            m_colViews.OfClass(typeof(ViewSection));

            List<View> m_Views = new List<View>();
            foreach (View x in m_colViews)
            {
                if (x.IsTemplate == false)
                {
                    m_Views.Add(x);
                }
            }

            return m_Views;
        }

        internal static List<View> GetAllViewsByNameContains(Document curDoc, string viewName)
        {
            // create an empty list to hold the results
            List<View> m_returnList = new List<View>();

            // get all views in the document
            List<View> m_allViews = GetAllViews(curDoc);

            // loop through all views
            foreach (View curView in m_allViews)
            {
                // check if the view name contains the specified string
                if (curView.Name.IndexOf(viewName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // add the view to the return list
                    m_returnList.Add(curView);
                }
            }

            // return the list of views that match the criteria
            return m_returnList;
        }

        internal static List<View> GetAllViewsByNameContainsAndAssociatedLevel(Document curDoc, string viewName, string levelName)
        {
            // create an empty list to hold the results
            List<View> m_returnList = new List<View>();

            // get all views in the document
            List<View> m_allViews = GetAllViews(curDoc);

            // loop through all views
            foreach (View curView in m_allViews)
            {
                // check if the view name contains the specified string
                if (curView.Name.IndexOf(viewName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // get the associated level parameter
                    Parameter associatedLevelParam = curView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL);

                    // check if the parameter is not null and has a value
                    if (associatedLevelParam != null && associatedLevelParam.HasValue)
                    {
                        // get the level name from the parameter
                        string levelNameFromParam = associatedLevelParam.AsString();

                        // check if the level name matches the specified level name
                        if (levelNameFromParam.Equals(levelName, StringComparison.OrdinalIgnoreCase))
                        {
                            // filter out dependent views
                            if (curView.GetPrimaryViewId() == ElementId.InvalidElementId)
                            {
                                // add the view to the return list
                                m_returnList.Add(curView);
                            }
                        }
                    }
                }
            }

            // return the list of views that match the criteria
            return m_returnList;
        }

        internal static View GetViewByNameContainsAndAssociatedLevel(Document curDoc, string viewName, string levelName)
        {
            // create an empty list to hold the results
            List<View> m_returnList = new List<View>();

            // get all views in the document
            List<View> m_allViews = GetAllViews(curDoc);

            // loop through all views
            foreach (View curView in m_allViews)
            {
                // check if the view name contains the specified string
                if (curView.Name.IndexOf(viewName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // get the associated level parameter
                    Parameter paramAssociatedLevel = curView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL);

                    // check if the parameter is not null and has a value
                    if (paramAssociatedLevel != null && paramAssociatedLevel.HasValue)
                    {
                        // get the level name from the parameter
                        string levelNameFromParam = paramAssociatedLevel.AsString();

                        // check if the level name matches the specified level name
                        if (levelNameFromParam.Equals(levelName, StringComparison.OrdinalIgnoreCase))
                        {
                            // add the view to the return list
                            m_returnList.Add(curView);
                        }
                    }
                }
            }

            // return the list of views that match the criteria
            if (m_returnList.Count > 0)
            {
                return m_returnList.FirstOrDefault();
            }
            else
            {
                return null;
            }
        }

        internal static View GetViewByNameContainsAndAssociatedLevel(Document curDoc, string viewName, string levelName, string levelNameAlt)
        {
            // create an empty list to hold the results
            List<View> m_returnList = new List<View>();

            // get all views in the document
            List<View> m_allViews = GetAllViews(curDoc);

            // loop through all views
            foreach (View curView in m_allViews)
            {
                // check if the view name contains the specified string
                if (curView.Name.IndexOf(viewName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // get the associated level parameter
                    Parameter paramAssociatedLevel = curView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL);

                    // check if the parameter is not null and has a value
                    if (paramAssociatedLevel != null && paramAssociatedLevel.HasValue)
                    {
                        // get the level name from the parameter
                        string levelNameFromParam = paramAssociatedLevel.AsString();

                        // check if the level name matches the specified level name
                        if (levelNameFromParam.Equals(levelName, StringComparison.OrdinalIgnoreCase) || levelNameFromParam.Equals(levelNameAlt, StringComparison.OrdinalIgnoreCase))
                        {
                            // add the view to the return list
                            m_returnList.Add(curView);
                        }
                    }
                }
            }

            // return the list of views that match the criteria
            if (m_returnList.Count > 0)
            {
                return m_returnList.FirstOrDefault();
            }
            else
            {
                return null;
            }
        }

        internal static bool IsFamilyInstancePresent(Document curDoc, string familyName)
        {
            return new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Any(fi => fi.Symbol.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
        }

        #endregion

        internal static Wall SelectWall(UIDocument uidoc, string prompt)
        {
            try
            {
                Reference picked = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), prompt);
                return uidoc.Document.GetElement(picked.ElementId) as Wall;
            }
            catch
            {
                return null;
            }
        }

        internal class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Wall; // Allows only Wall elements to be selected
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
