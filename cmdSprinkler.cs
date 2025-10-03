using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;
using System.Windows.Controls;
using static ConvertSpecLevel.cmdRefSp;

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
                       
            // get the first electrical plan associated with First Floor level
            View planView = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor").FirstOrDefault();

            // null check
            if (planView == null)
            {
                // notify the user and exit
                Utils.TaskDialogError("Error", "Sprinkler Outlet", "No First Floor electrical plan found.");
                return Result.Failed;
            }

            // set it as the active view
            uidoc.ActiveView = planView;

            // prompt the user to select the wall to host the sprinkler outlet
            Wall outletWall = SelectWall(uidoc, "Select wall to host sprinkler outlet.");
            if (outletWall == null) return Result.Cancelled;

            // prompt the user to select the front garage wall
            Wall garageWall = SelectWall(uidoc, "Select front garage wall.");
            if (garageWall == null) return Result.Cancelled;

            // get garage wall exterior face point
            XYZ facePoint = GetWallExteriorStructuralFace(garageWall);

            // null check
            if (facePoint == null)
            {
                Utils.TaskDialogError("Error", "Sprinkler Outlet", "Could not determine garage wall exterior structural face.");
                return Result.Failed;
            }

            // Compute offset in direction of garage wall's orientation (which points OUT)
            double offsetFeet = UnitUtils.ConvertToInternalUnits(60, UnitTypeId.Inches);

            // reverse direction to go IN from garage face
            XYZ garageToOutletDir = -garageWall.Orientation; // Going inward from garage face
            XYZ offsetPoint = facePoint + (garageToOutletDir * offsetFeet);

            // create variable for outlet insertion point
            XYZ outletPoint = new XYZ(offsetPoint.X, offsetPoint.Y, 0);










            return Result.Succeeded;

        }

        private XYZ GetWallExteriorStructuralFace(Wall garageWall)
        {
            throw new NotImplementedException();
        }

        private Wall SelectWall(UIDocument uidoc, string prompt)
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
    }
   
}