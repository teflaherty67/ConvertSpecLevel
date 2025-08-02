using ConvertSpecLevel.Classes;

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

            // check for two story plan
            // look for a level named Second Floor or Upper Level
            // if found notify user & exit command
            // if not found continue executing the code

            // filter the list to remove First Floor (or Main Level) level

            // increase current value of plate heights by 12"

            // notify user how many plates were raised

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
