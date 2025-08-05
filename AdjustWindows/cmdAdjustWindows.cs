using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;
using System.Windows.Input;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdAdjustWindows : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // get all the window instances in the project
            List<FamilyInstance> allWindows = Utils.GetAllWindows(curDoc);

            // create a dictionary to hold the window data
            Dictionary<ElementId, clsWindowData> dictionaryWinData = new Dictionary<ElementId, clsWindowData>();

            // loop through the windows and get the data to store
            foreach (FamilyInstance curWindow in allWindows)
            {
                // store the data
                clsWindowData curData = new clsWindowData(curWindow);
                dictionaryWinData.Add(curWindow.Id, curData);
            }

            // get all the ViewSection views
            List<View> listViews = Utils.GetAllSectionViews(curDoc);

            // get the first view whose Title on Sheet is "Front Elevation"
            View elevFront = listViews
                .FirstOrDefault(v => v.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)?.AsString() == "Front Elevation");

            // set that view as the active view
            if (elevFront != null)
            {
                uidoc.ActiveView = elevFront;
            }
            else
            {
                Utils.TaskDialogInformation("Information", "Spec Conversion", "Front Elevation view not found. Proceeding with level adjustments in current view.");
            }

            // launch the form
            frmAdjustWindows curForm = new frmAdjustWindows();
            curForm.Topmost = true;

            // check if the user clicks OK
            if (curForm.ShowDialog() == true)
            {
                // get the selected spec level from the form
                string selectedSpecLevel = curForm.GetSelectedSpecLevel();

                // check if adjust head heights is checked
                bool adjustHeadHeights = curForm.IsAdjustWindowHeadHeightsChecked();

                // test for raising or lowering windows
                bool raiseWindows = (selectedSpecLevel == "Complete Home Plus");

                // create counter for windows changed
                int countWindows = 0;

                if (adjustHeadHeights)
                {
                    // create and start a transaction
                    using (Transaction t = new Transaction(curDoc, "Adjust Window Head Heights"))
                    {
                        t.Start();

                        foreach (var kvp in dictionaryWinData)
                        {
                            clsWindowData curData = kvp.Value;
                            double plateAdjustment = 1.0;
                            double newHeadHeight;

                            if (!raiseWindows)
                            {
                                // lower window head heights by 12"
                                newHeadHeight = curData.CurHeadHeight - plateAdjustment;
                            }
                            else
                            {
                                // raise window head height by by 12"
                                newHeadHeight = curData.CurHeadHeight + plateAdjustment;
                            }

                            if (curData.HeadHeightParam != null && !curData.HeadHeightParam.IsReadOnly)
                            {
                                // adjust the head heihgt
                                curData.HeadHeightParam.Set(newHeadHeight);

                                // increment the counter
                                countWindows++;
                            }
                        }

                        t.Commit();
                    }

                    // notify user of results
                    Utils.TaskDialogInformation("Information", "Spec Conversion",
                        $"Adjusted head heights for {countWindows} windows per the selected spec level.");
                }
                else
                {
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "No window head height adjustments selected.");
                }

                return Result.Succeeded;
            }

            else
            {
                return Result.Cancelled;
            }            
        }

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
