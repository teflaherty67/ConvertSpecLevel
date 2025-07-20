

using static ConvertSpecLevel.cmdConvertSpecLevel;

namespace ConvertSpecLevel.Common
{
    internal static class Utils
    {
        #region Families

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

        internal static FamilySymbol GetFamilySymbolByName(Document curDoc, string newCabinetFamilyName, string newCabinetTypeName)
        {
            throw new NotImplementedException();
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

        internal static List<View> GetAllViewsByNameContainsAndAssociatedLevel(Document curDoc, string v1, string v2)
        {
            throw new NotImplementedException();
        }      

        internal static View GetViewByNameContainsAndAssociatedLevel(Document curDoc, string v1, string v2)
        {
            throw new NotImplementedException();
        }

        internal static ViewSheet GetViewSheetByName(Document curDoc, string v)
        {
            throw new NotImplementedException();
        }      

        internal static void UpdateLightingFixturesInActiveView(Document curDoc, string selectedSpecLevel)
        {
            throw new NotImplementedException();
        }


    }
}
