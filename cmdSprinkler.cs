using Autodesk.Revit.DB;
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

            // variables for sprinkler family
            FamilySymbol sprinklerSymbol = null;
            string outletFamilyName = "LD_EF_Recep_None";
            string outletTypeName = "Sprinkler";
#if REVIT2025
            string outletFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical";
#endif

#if REVIT2026
            string outletFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\Electrical";
#endif

            // variables for tag family
            FamilySymbol tagSymbol = null;
            string tagFamilyName = "LD_AN_Tag_EF_Type-Comments";
            string tagTypeName = "Type 2";
#if REVIT2025
            string tagFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Annotation\Tags";
#endif

#if REVIT2026
            string tagFilePath = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\Annotation\Tags";
#endif

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

            // create a transaction group to place the outlet
            using (TransactionGroup tg = new TransactionGroup(curDoc, "Add Sprinkler Outlet"))
            {
                // start the transaction group
                tg.Start();

                // create a transaction to load the families
                using (Transaction t = new Transaction(curDoc, "Load Families"))
                {
                    // start the transaction
                    t.Start();

                    // check if the sprinkler outlet family & type is already loaded
                    sprinklerSymbol = Utils.GetFamilySymbolByName(curDoc, outletFamilyName, outletTypeName);
                    if (sprinklerSymbol == null)
                    {
                        // if not, load it
                        Family loadedFamily = Utils.LoadFamilyFromLibrary(curDoc, outletFilePath, outletFamilyName);

                        // check again if the family is loaded
                        sprinklerSymbol = Utils.GetFamilySymbolByName(curDoc, outletFamilyName, outletTypeName);
                    }

                    // check if symbol is found
                    if (sprinklerSymbol == null)
                    {
                        // if not, notify the user and exit
                        Utils.TaskDialogError("Error", "Sprinkler Outlet", $"Unable to find type '{outletTypeName}' in family '{outletFamilyName}'.");
                        return Result.Failed;
                    }

                    // activate the outlet family symbol if not already active
                    if (!sprinklerSymbol.IsActive) sprinklerSymbol.Activate();
                    curDoc.Regenerate();

                    // check if the tag family & type is already loaded
                    tagSymbol = Utils.GetFamilySymbolByName(curDoc, tagFamilyName, tagTypeName);
                    if (tagSymbol == null)
                    {
                        // if not, load it
                        Family loadedFamily = Utils.LoadFamilyFromLibrary(curDoc, tagFilePath, tagFamilyName);

                        // check again if the family is loaded
                        tagSymbol = Utils.GetFamilySymbolByName(curDoc, tagFamilyName, tagTypeName);
                    }

                    // check if symbol is found
                    if (tagSymbol == null)
                    {
                        // if not, notify the user and exit
                        Utils.TaskDialogError("Error", "Sprinkler Outlet", $"Unable to find type '{tagTypeName}' in family '{tagFamilyName}'.");
                        return Result.Failed;
                    }

                    // activate the outlet family symbol if not already active
                    if (!tagSymbol.IsActive) tagSymbol.Activate();
                    curDoc.Regenerate();

                    // commit the transaction
                    t.Commit();
                }

                // get the level to place the outlet on
                Level outletLevel = curDoc.GetElement(planView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsElementId()) as Level;
                
                // create a transaction to place the outlet
                using (Transaction  t1 = new Transaction(curDoc, "Place Sprinkler Outlet"))
                {
                    // start the transaction
                    t1.Start();

                    // create the outlet family instance
                    FamilyInstance outletInstance = curDoc.Create.NewFamilyInstance(outletPoint, sprinklerSymbol, outletLevel, StructuralType.NonStructural);

                    // get facing of outlet instance
                    XYZ outletFacing = outletInstance.FacingOrientation;

                    // get facing of outlet wall
                    XYZ wallFacing = outletWall.Orientation;

                    // normalize both to be safe
                    outletFacing = outletFacing.Normalize();
                    wallFacing = wallFacing.Normalize();

                    // set rotation point as the outlet insertion point
                    LocationPoint loc = outletInstance.Location as LocationPoint;
                    XYZ origin = loc.Point;

                    // check if rotation is needed

                    double dot = outletFacing.DotProduct(wallFacing);

                    const double angleTolerance = 1e-6;

                    if (Math.Abs(dot - 1) < angleTolerance)
                    {
                        // Already aligned – no rotation needed
                    }
                    else
                    {
                        // Compute the angle between them
                        double angle = outletFacing.AngleTo(wallFacing); // in radians

                        // Define the axis of rotation: vertical (Z-axis) at the outlet point
                        Line rotationAxis = Line.CreateBound(outletPoint, outletPoint + XYZ.BasisZ);

                        // Determine rotation direction
                        // Use cross product to check if angle should be negative (clockwise)
                        XYZ cross = outletFacing.CrossProduct(wallFacing);
                        if (cross.Z < 0)
                        {
                            angle = -angle;
                        }

                        // Apply rotation
                        ElementTransformUtils.RotateElement(curDoc, outletInstance.Id, rotationAxis, angle);
                    }

                    // commit the transaction
                    t1.Commit();
                }

                // create a transaction to place the annotations
                using (Transaction t2 = new Transaction(curDoc, "Add Sprinkler Tag and Dimension"))
                {
                    // start the transaction
                    t2.Start();

                    // commit the transaction
                    t2.Commit();
                }

                // assimilate the transaction group
                tg.Assimilate();
            }

            // notify the user of success


            return Result.Succeeded;
        }

        private XYZ GetWallExteriorStructuralFace(Wall wall)
        {
            WallType wallType = wall.Document.GetElement(wall.GetTypeId()) as WallType;
            CompoundStructure structure = wallType?.GetCompoundStructure();
            if (structure == null) return null;

            bool hasStructure = structure.GetLayers().Any(l => l.Function == MaterialFunctionAssignment.Structure);
            if (!hasStructure) return null;

            Options opts = new Options { ComputeReferences = false };
            GeometryElement geom = wall.get_Geometry(opts);
            XYZ orientation = wall.Orientation;

            foreach (GeometryObject obj in geom)
            {
                if (obj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                        if (Math.Abs(normal.Z) < 0.01 && normal.DotProduct(orientation) > 0.5)
                        {
                            return face.Evaluate(new UV(0.5, 0.5));
                        }
                    }
                }
            }

            return null;
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