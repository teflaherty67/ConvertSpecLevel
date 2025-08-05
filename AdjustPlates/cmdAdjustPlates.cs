using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdAdjustPlates : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // get all the levels in the project
            List<Level> listLevels = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .OfType<Level>()
                .ToList();

            // check for two story plan
            foreach (Level curLevel in listLevels)
            {
                // look for a level named Second Floor or Upper Level
                if (curLevel.Name == "Second Floor" || curLevel.Name == "Upper Level")
                {
                    // if found notify user & end command
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "Multi-story plan detected. Plate change not applicable.");
                    return Result.Succeeded;
                }
            }

            // Filter out First Floor/Main Level
            listLevels = listLevels.Where(level => level.Name != "First Floor" && level.Name != "Main Level").ToList();

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
            frmAdjustPlates curForm = new frmAdjustPlates();
            curForm.Topmost = true;

            // check if the user clicks OK
            if (curForm.ShowDialog() == true)
            {
                // get the selected spec level from the form
                string selectedSpecLevel = curForm.GetSelectedSpecLevel();

                // test for raising or lowering plates
                bool raisePlates = (selectedSpecLevel == "Complete Home Plus");                

                // create counter for levels changed
                int countLevels = 0;

                // create and start a transaction
                using (Transaction t = new Transaction(curDoc, "Adjust Plate Heights"))
                {
                    t.Start();

                    if (!raisePlates)
                    {
                        // lower the plates by 12"
                        foreach (Level curLevel in listLevels)
                        {
                            curLevel.Elevation = curLevel.Elevation - 1.0;

                            // increment the counter
                            countLevels++;
                        }
                    }
                    else
                    {
                        // raise the plates by 12"
                        foreach(Level curLevel in listLevels)
                        {
                            curLevel.Elevation = curLevel.Elevation + 1.0;

                            // increment the counter
                            countLevels++;
                        }
                    }

                    t.Commit();
                }

                // notify the user
                Utils.TaskDialogInformation("Information", "Spec Conversion", $"{countLevels} Level{(countLevels == 1 ? "" : "s")}" +
                    $" {(countLevels == 1 ? "was" : "were")} adjusted per the specified spec level.");

                return Result.Succeeded;
            }

            // if user clicks on Cancel
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
