using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSprinkler : IExternalCommand
    {
        const double OFFSET_INCHES = 60.0;       // Outlet offset from garage wall
        const double DIM_LINE_OFFSET_FEET = 2.0; // Offset of dimension line itself

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            View planView = Utils.GetAllViewsByNameContainsAndAssociatedLevel(doc, "Electrical", "First Floor").FirstOrDefault();
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

            using (Transaction t = new Transaction(doc, "Place Sprinkler Outlet"))
            {
                try
                {
                    t.Start();

                    // Load and activate the sprinkler family symbol
                    FamilySymbol sprinklerSymbol = Utils.GetFamilySymbolByName(doc, "LD_EF_Recep_None", "Sprinkler");
                    if (sprinklerSymbol == null)
                    {
                        Utils.LoadFamilyFromLibrary(doc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical", "LD_EF_Recep_None");
                        sprinklerSymbol = Utils.GetFamilySymbolByName(doc, "LD_EF_Recep_None", "Sprinkler");
                    }

                    if (sprinklerSymbol == null)
                    {
                        Utils.TaskDialogError("Error", "Sprinkler Outlet", "Unable to load sprinkler family.");
                        return Result.Failed;
                    }

                    if (!sprinklerSymbol.IsActive) sprinklerSymbol.Activate();
                    doc.Regenerate();

                    // 1. Get garage wall exterior face point
                    XYZ facePoint = GetWallExteriorStructuralFace(garageWall);
                    if (facePoint == null)
                    {
                        Utils.TaskDialogError("Error", "Sprinkler Outlet", "Could not determine garage wall exterior structural face.");
                        return Result.Failed;
                    }

                    // 2. Compute 60" offset in direction of garage wall's orientation (which points OUT)
                    double offsetFeet = UnitUtils.ConvertToInternalUnits(OFFSET_INCHES, UnitTypeId.Inches);
                    XYZ garageToOutletDir = -garageWall.Orientation; // Going inward from garage face
                    XYZ offsetPoint = facePoint + (garageToOutletDir * offsetFeet);

                    // 3. Final outlet point at 18" AFF
                    // 3. Place outlet at level elevation (Z=0)
                    XYZ outletPoint = new XYZ(offsetPoint.X, offsetPoint.Y, 0);

                    // 4. Place the outlet
                    Level level = doc.GetElement(planView.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsElementId()) as Level;
                    FamilyInstance outletInstance = doc.Create.NewFamilyInstance(outletPoint, sprinklerSymbol, level, StructuralType.NonStructural);

                    // 5. Get face references
                    Reference garageFaceRef = GetWallExteriorFaceReference(garageWall);
                    Reference outletCenterRef = outletInstance.GetReferences(FamilyInstanceReferenceType.CenterLeftRight).FirstOrDefault();

                    // 6. Create dimension if both references exist
                    if (garageFaceRef != null && outletCenterRef != null)
                    {
                        ReferenceArray dimRefs = new ReferenceArray();
                        dimRefs.Append(garageFaceRef);
                        dimRefs.Append(outletCenterRef);

                        // Direction perpendicular to the outlet wall (i.e., facing OUTSIDE)
                        XYZ outletWallNormal = outletWall.Orientation; // Revit defines this as exterior direction

                        // Offset dimension line in that direction
                        XYZ dimOffset = outletWallNormal * DIM_LINE_OFFSET_FEET;

                        // Flatten to plan view (Z = 0)
                        XYZ p1 = new XYZ((facePoint + dimOffset).X, (facePoint + dimOffset).Y, 0);
                        XYZ p2 = new XYZ((outletPoint + dimOffset).X, (outletPoint + dimOffset).Y, 0);
                        Line dimLine = Line.CreateBound(p1, p2);

                        doc.Create.NewDimension(planView, dimLine, dimRefs);
                    }

                    // 7. Place tag

                    // Tag expects a Reference to the instance itself when using TagMode.TM_ADDBY_CATEGORY
                    Reference tagRef = new Reference(outletInstance);

                    // Offset direction: move tag away from outlet wall face (using wall orientation)
                    XYZ tagDirection = outletWall.Orientation; // This points toward the exterior of the outlet wall
                    XYZ tagOffset = tagDirection * 2.0 + new XYZ(0, 0, 2.0); // 2 ft away, lifted slightly in Z
                    XYZ tagPosition = outletPoint + tagOffset;

                    IndependentTag tag = IndependentTag.Create(doc, planView.Id,
                        tagRef,
                        true,
                        TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal,
                        tagPosition);

                    if (tag != null)
                    {
                        tag.LeaderEndCondition = LeaderEndCondition.Free;
                    }

                    t.Commit();

                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error Details", ex.Message + "\n\n" + ex.StackTrace);
                    t.RollBack();
                    return Result.Failed;
                }
                
            }

            Utils.TaskDialogInformation("Complete", "Sprinkler Outlet", "Sprinkler outlet placed successfully.");
            return Result.Succeeded;
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

        private Reference GetWallExteriorFaceReference(Wall wall)
        {
            Options opts = new Options { ComputeReferences = true };
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
                            return face.Reference;
                        }
                    }
                }
            }

            return null;
        }

        private XYZ ProjectPointOntoWall(Wall wall, XYZ point)
        {
            Options opts = new Options { ComputeReferences = false };
            GeometryElement geom = wall.get_Geometry(opts);

            foreach (GeometryObject obj in geom)
            {
                if (obj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                        // Check if this face is vertical
                        if (Math.Abs(normal.Z) < 0.01)
                        {
                            // Project point onto this face
                            IntersectionResult result = face.Project(point);
                            if (result != null)
                            {
                                return result.XYZPoint;
                            }
                        }
                    }
                }
            }

            return null; // Fallback if projection fails
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnSprinkler";
            string buttonTitle = "Sprinkler Outlet";

            clsButtonData myButtonData = new clsButtonData(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Place sprinkler outlet with dimensioning and tagging");

            return myButtonData.Data;
        }
    }

    // Wall selection filter class
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
}