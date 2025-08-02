using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdPlateChange : IExternalCommand
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

            // launch the form
            frmPlateChange curForm = new frmPlateChange()
            {
                Topmost = true,
            };

            curForm.ShowDialog();

            // check if user clicked OK
            if (curForm.DialogResult != true)
            {
                return Result.Cancelled;
            }



            // increase current value of plate heights by 12"

            // notify user how many plates were raised



            return Result.Succeeded;
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
