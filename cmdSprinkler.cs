using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;
using System.Windows.Controls;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSprinkler : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // set a first floor electrical plan view as the active view

            // prompt the user to select a wall to place the sprinkler outlet on

            // prompt the user to select a the garage front wall

            View planView = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor").FirstOrDefault();
            if (planView == null)
            {
                Utils.TaskDialogError("Error", "Sprinkler Outlet", "No First Floor electrical plan found.");
                return Result.Failed;
            }

            uidoc.ActiveView = planView;

            Wall outletWall = SelectWall(uidoc, "Select wall to host sprinkler outlet.");
            if (outletWall == null) return Result.Cancelled;

            Wall garageWall = SelectWall(uidoc, "Select garage (perpendicular) wall.");
            if (garageWall == null) return Result.Cancelled;










            return Result.Succeeded;

        }


    }
   
}