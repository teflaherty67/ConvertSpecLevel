using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdRefSp : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // automatically switch to a First Floor view if not already in one
            View firstFloorPlan = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Annotation", "First Floor")
                .FirstOrDefault();

            // null check
            if (firstFloorPlan != null)
            {
                // set the active view to the first floor plan
                uidoc.ActiveView = firstFloorPlan;
            }
            else
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "No First Floor plan views found in the project.");
                return Result.Failed;
            }            

            // prompt the user to select the wall behind the Ref Sp
            Reference selectedRefSpWall = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select wall to place Ref Sp cabinet");

            // cast the selected element to a Wall
            Wall selectedWall = curDoc.GetElement(selectedRefSpWall.ElementId) as Wall;

            // verify the selected element is a wall
            if (selectedWall == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                return Result.Failed;
            }

            // prompt the user to select the Ref Sp
            FamilyInstance selectedRefSp = curDoc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, new ApplianceSelectionFilter(), "Select Ref Sp")) as FamilyInstance;

            // verify the selected element is a family instance
            if (selectedRefSp == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a refrigerator. Please try again.");
                return Result.Failed;
            }          

            // create a transaction
            using (Transaction t = new Transaction(curDoc, "Place Ref Sp"))
            {
                t.Start();                

                // check if family is loaded
                FamilySymbol cabRefSp = Utils.GetFamilySymbolByName(curDoc, "LD_CW_Wall_2-Dr_Flush", "39\"x27\"x15\"");

                // null check
                if (cabRefSp == null)
                {
                    // load the family
                    Family cabFam = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Casework\Kitchen", "LD_CW_Wall_2-Dr_Flush");

                    // Check if the cabinet family loaded successfully
                    if (cabFam == null)
                    {
                        // Show error message if cabinet family failed to load
                        Utils.TaskDialogError("Error", "Spec Conversion", "Could not load Ref Sp cabinet family from library.");
                        return Result.Failed;
                    }

                    // try to get the family symbol again
                    cabRefSp = Utils.GetFamilySymbolByName(curDoc, "LD_CW_Wall_2-Dr_Flush", "39\"x27\"x15\"");

                    // Check if the cabinet type was found
                    if (cabRefSp == null)
                    {
                        // Show error message if cabinet type not found
                        Utils.TaskDialogError("Error", "Spec Conversion", "Could not find Ref Sp cabinet type in the project after loading family.");
                        return Result.Failed;
                    }
                }

                // activate family symbol if needed
                if (cabRefSp != null && !cabRefSp.IsActive)
                {
                    cabRefSp.Activate();
                    curDoc.Regenerate();
                }

                // Get refrigerator location point
                LocationPoint fridgeLocation = selectedRefSp.Location as LocationPoint;
                XYZ fridgePoint = fridgeLocation.Point;                            

                // check if the fridge location point is null
                if (fridgePoint == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "Refrigerator centerline is not a straight line. Cannot calculate cabinet placement.");
                    return Result.Failed;
                }
               
                // Get the wall's location curve to find intersection point
                LocationCurve wallLocation = selectedWall.Location as LocationCurve;

                // get the curve geometry from the wall location
                Line wallCenterLine = wallLocation.Curve as Line;

                // Calculate the wall direction vector to determine left offset direction
                XYZ wallDirection = wallCenterLine.Direction;

                // Create a perpendicular vector pointing left relative to the wall direction
                XYZ leftDirection = wallDirection;

                // Calculate the final cabinet placement point by offsetting 19.5" to the left and setting elevation to 75" AFF
                XYZ cabinetPlacementPoint = new XYZ(
                    fridgePoint.X + (leftDirection.X * (19.5 / 12.0)),
                    fridgePoint.Y + (leftDirection.Y * (19.5 / 12.0)),
                    6.25); // 75" AFF (75/12 = 6.25 feet)

                // Place the Ref Sp cabinet at the calculated placement point
                FamilyInstance refSpCabinet = curDoc.Create.NewFamilyInstance(cabinetPlacementPoint, cabRefSp, selectedWall, StructuralType.NonStructural);

                // Check if the cabinet was placed successfully
                if (refSpCabinet == null)
                {
                    // Show error message if cabinet placement failed
                    Utils.TaskDialogError("Error", "Spec Conversion", "Failed to place Ref Sp cabinet at calculated location.");
                    return Result.Failed;
                }

                // Success - Ref Sp cabinet placed successfully
                Utils.TaskDialogInformation("Complete", "Spec Conversion", "Ref Sp cabinet placed successfully at calculated location.");

                t.Commit();
            }

            return Result.Succeeded;
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

        internal class ApplianceSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                // allow any specialty equipment family
                return elem.Category != null && elem.Category.Name == "Specialty Equipment";
            }
            public bool AllowReference(Reference reference, XYZ position)
            {
                // Allow all references
                return true;
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
