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

            // get user input from the form
            string selectedSpecLevel = curForm.GetSelectedSpecLevel();

            // counters
            int updatedPlates = 0;

            // get all the levels in the document
            List<Level> allLevels = Utils.GetAllLevels(curDoc)
                // filter out the First Floor/Main Level levels
                .Where(x => x.Name != "First Floor" && x.Name != "Main Level")
                .ToList(); // converts the result to a list           

            // start a transaction to change the plate heights
            using(Transaction t = new Transaction(curDoc, "Change Plate Heights"))
            {
                t.Start();

                // loop through the levels and change the plate heights
                foreach (Level curLevel in allLevels)
                {
                    // raise/lower the plate height based on the selected spec level
                    ManagePlateHeights(curDoc, curLevel, selectedSpecLevel);

                    // increment the counter
                    updatedPlates++;
                }

                // commit the transaction
                t.Commit();
            }           

            // show a message box to confirm the changes
            Utils.TaskDialogInformation("Information", "Spec Conversion", 
                $"Successfully updated {updatedPlates} plates per the selected spec level.");

            return Result.Succeeded;
        }

        private void ManagePlateHeights(Document curDoc, Level curLevel, string selectedSpecLevel)
        {
            throw new NotImplementedException();
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData.Data;
        }
    }
}
