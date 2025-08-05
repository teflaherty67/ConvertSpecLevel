using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

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
                clsWindowData curData = new clsWindowData(curWindow);
                dictionaryWinData.Add(curWindow.Id, curData);
            }



            // store the current head heights

            // store the current window height

            // launch the form
            frmAdjustWindows curForm = new frmAdjustWindows()
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
