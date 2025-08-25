using Autodesk.Revit.DB.Structure;
using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;
using System.IO.Packaging;
using static ConvertSpecLevel.cmdRefSp;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSprinkler : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // automatically switch to a First Floor Electrical Plan view if not already in one
            View firstFloorElecPlan = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor")
                .FirstOrDefault();

            // null check
            if (firstFloorElecPlan != null)
            {
                // set the active view to the first floor electrical plan
                uidoc.ActiveView = firstFloorElecPlan;
            }
            else
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "No first floor electrical views found in the project.");
                return Result.Failed;
            }

            // prompt the user to select the wall for the sprinkler outlet
            Reference selectedOutletWall = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select wall for sprinkler outlet.");

            // cast the selected element to a Wall
            Wall outletWall = curDoc.GetElement(selectedOutletWall.ElementId) as Wall;

            // verify the selected element is a wall
            if (outletWall == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                return Result.Failed;
            }

            // prompt user to select the front wall of the garage
            Reference selectedGarageWall = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select front wall of Garage.");

            // cast the selected element to a Wall
            Wall garageWall = curDoc.GetElement(selectedGarageWall.ElementId) as Wall;

            // verify the selected element is a wall
            if (garageWall == null)
            {
                Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                return Result.Failed;
            }

            // start a transaction
            using (Transaction t = new Transaction(curDoc, "Sprinkler Outlet"))
            {
                t.Start();

                // check if family is loaded
                FamilySymbol outletSprinkler = Utils.GetFamilySymbolByName(curDoc, "LD_EF_Recep_Wall", "Sprinkler");

                // null check
                if (outletSprinkler == null)
                {
                    // load the family
                    Family outletFam = Utils.LoadFamilyFromLibrary(curDoc, @"S:\Shared Folders\Lifestyle USA Design\Library 2025\Electrical", "LD_EF_Recep_Wall");

                    // Check if the receptacle family loaded successfully
                    if (outletFam == null)
                    {
                        // Show error message if receptacle family failed to load
                        Utils.TaskDialogError("Error", "Spec Conversion", "Could not load receptacle family from library.");
                        return Result.Failed;
                    }

                    // try to get the family symbol again
                    outletSprinkler = Utils.GetFamilySymbolByName(curDoc, "LD_EF_Recep_Wall", "Sprinkler");

                    // Check if the receptacle type was found
                    if (outletSprinkler == null)
                    {
                        // Show error message if receptacle type not found
                        Utils.TaskDialogError("Error", "Spec Conversion", "Could not find receptacle type in the project after loading family.");
                        return Result.Failed;
                    }
                }

                // activate family symbol if needed
                if (outletSprinkler != null && !outletSprinkler.IsActive)
                {
                    outletSprinkler.Activate();
                    curDoc.Regenerate();
                }

                // Create options for geometry extraction - controls how detailed the geometry will be
                Options geometryOptions = new Options();

                // Get the wall's 3D geometry as a collection of solids and surfaces
                GeometryElement wallGeometry = garageWall.get_Geometry(geometryOptions);

                // Variables to store our target point when found
                XYZ targetPoint = null;
                bool foundTargetPoint = false;

                // Loop through the geometry to find the solid objects (the actual wall layers)
                foreach (GeometryObject geomObj in wallGeometry)
                {
                    // Check if this geometry object is a solid with volume and not null
                    if (geomObj is Solid solid && solid != null && solid.Volume > 0)
                    {
                        // Get the material of this solid to identify if it's the stud layer
                        ElementId materialId = solid.GraphicsStyleId;
                        if (materialId != null && materialId != ElementId.InvalidElementId)
                        {
                            // Get the material element to check its name
                            Material material = curDoc.GetElement(materialId) as Material;
                            if (material != null && material.Name.Contains("2x"))
                            {
                                // Loop through all faces of this solid to find the structural core face
                                foreach (Face face in solid.Faces)
                                {
                                    // Get the face normal vector (direction perpendicular to the face)
                                    XYZ faceNormal = face.ComputeNormal(new UV(0.5, 0.5));

                                    // Check if this face is vertical (Z component near 0)
                                    if (Math.Abs(faceNormal.Z) < 0.1)
                                    {
                                        // Get a point on the face to use for calculations
                                        XYZ pointOnFace = face.Evaluate(new UV(0.5, 0.5));

                                        // Check if this is the outside face by comparing with wall location line
                                        LocationCurve wallLocationCurve = garageWall.Location as LocationCurve;

                                        // Get the centerline of the wall
                                        Line wallCenterLine = wallLocationCurve.Curve as Line;

                                        // Get the closest point on the wall centerline to our face point
                                        XYZ closestPointOnCenterline = wallCenterLine.Project(pointOnFace).XYZPoint;

                                        // Calculate vector from centerline to face point
                                        XYZ centerToFace = pointOnFace - closestPointOnCenterline;

                                        // Check if this face is pointing away from the building interior (outside face)
                                        if (centerToFace.DotProduct(faceNormal) > 0)
                                        {
                                            // This is the outside face - calculate target point 60" perpendicular from this face
                                            targetPoint = pointOnFace + (faceNormal * 5.0);
                                            foundTargetPoint = true;
                                            break; // Exit face loop - we found our outside face
                                        }

                                        if (foundTargetPoint)
                                            break; // Exit geometry loop - we found our target point
                                    }
                                }
                            }
                        }

                        // Check if we successfully found our target point
                        if (!foundTargetPoint || targetPoint == null)
                        {
                            Utils.TaskDialogError("Error", "Spec Conversion", "Could not find the outside face of the garage stud wall.");
                            return Result.Failed;
                        }                       

                        // Get the outlet wall's geometry to find where to place the outlet
                        GeometryElement outletWallGeometry = outletWall.get_Geometry(geometryOptions);

                        // Variables to store the outlet placement point
                        XYZ outletPlacementPoint = null;
                        bool foundOutletFace = false;

                        // Loop through the outlet wall geometry to find the GWB layer
                        foreach (GeometryObject outletGeomObj in outletWallGeometry)
                        {
                            // Check if this geometry object is a solid with volume and not null
                            if (outletGeomObj is Solid outletSolid && outletSolid != null && outletSolid.Volume > 0)
                            {
                                // Get the material of this solid to identify if it's the GWB layer
                                ElementId outletMaterialId = outletSolid.GraphicsStyleId;

                                if (outletMaterialId != null && outletMaterialId != ElementId.InvalidElementId)
                                {
                                    // Get the material element to check its name for GWB
                                    Material outletMaterial = curDoc.GetElement(outletMaterialId) as Material;
                                    if (outletMaterial != null && outletMaterial.Name.Contains("GWB"))
                                    {
                                        // Loop through all faces of the GWB solid to find the inside face
                                        foreach (Face outletFace in outletSolid.Faces)
                                        {
                                            // Get the face normal vector (direction perpendicular to the face)
                                            XYZ outletFaceNormal = outletFace.ComputeNormal(new UV(0.5, 0.5));

                                            // Check if this face is vertical (Z component near 0)
                                            if (Math.Abs(outletFaceNormal.Z) < 0.1)
                                            {
                                                // Project our target point onto this face to see if it intersects
                                                IntersectionResult intersection = outletFace.Project(targetPoint);
                                                if (intersection != null)
                                                {
                                                    // Get the intersection point - this is where the outlet will be placed
                                                    outletPlacementPoint = intersection.XYZPoint;
                                                    foundOutletFace = true;
                                                    break; // Exit face loop - we found our intersection
                                                }

                                            }
                                        }

                                        if (foundOutletFace)
                                        {
                                            break; // Exit geometry loop - we found our outlet face
                                        }
                                    }
                                }

                                // Check if we successfully found our outlet placement point
                                if (!foundOutletFace || outletPlacementPoint == null)
                                {
                                    Utils.TaskDialogError("Error", "Spec Conversion", "Could not find intersection point on the outlet wall.");
                                    return Result.Failed;
                                }                               

                                // Place the sprinkler outlet at the calculated intersection point
                                FamilyInstance sprinklerOutlet = curDoc.Create.NewFamilyInstance(outletPlacementPoint, outletSprinkler, outletWall, StructuralType.NonStructural);

                                // Check if the outlet was placed successfully
                                if (sprinklerOutlet == null)
                                {
                                    Utils.TaskDialogError("Error", "Spec Conversion", "Failed to place sprinkler outlet.");
                                    return Result.Failed;
                                }

                                // Create dimension from garage wall face to outlet center
                                // First, get references for dimensioning - garage wall face and outlet center
                                Reference refGarageWall = new Reference(garageWall);

                                // Get reference to the outlet's 'Center (Left/Right)' reference point
                                Reference refOutlet = null;
                                foreach (Reference familyRef in sprinklerOutlet.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
                                {
                                    refOutlet = familyRef;
                                    break; // Take the first (and likely only) Center (Left/Right) reference
                                }

                                // Check if we found the outlet reference
                                if (refOutlet == null)
                                {
                                    Utils.TaskDialogError("Error", "Spec Conversion", "Could not find outlet center reference for dimensioning.");
                                    return Result.Failed;
                                }

                                // Create a reference array for the dimension (from garage wall to outlet)
                                ReferenceArray dimReferenceArray = new ReferenceArray();
                                dimReferenceArray.Append(refGarageWall);
                                dimReferenceArray.Append(refOutlet);

                                // Create a line to define where the dimension will be placed (2' from outlet wall exterior)
                                XYZ dimensionOffset = (outletPlacementPoint - targetPoint).Normalize() * 2.0; // 2 feet perpendicular to outlet wall
                                XYZ dimStartPoint = targetPoint + dimensionOffset;
                                XYZ dimEndPoint = outletPlacementPoint + dimensionOffset;
                                Line dimensionLine = Line.CreateBound(dimStartPoint, dimEndPoint);

                                // Create the dimension
                                Dimension sprinklerDimension = curDoc.Create.NewDimension(curDoc.ActiveView, dimensionLine, dimReferenceArray);

                                // Create a tag for the sprinkler outlet with leader
                                // Determine outlet wall orientation and calculate tag position accordingly
                                LocationCurve outletWallLocation = outletWall.Location as LocationCurve;
                                Line outletWallLine = outletWallLocation.Curve as Line;
                                XYZ wallDirection = outletWallLine.Direction;

                                // Check if wall is more horizontal or vertical by comparing X and Y components
                                bool isHorizontalWall = Math.Abs(wallDirection.X) > Math.Abs(wallDirection.Y);

                                XYZ tagPosition;
                                if (isHorizontalWall)
                                {
                                    // Horizontal wall: 2' left and 2' up
                                    tagPosition = outletPlacementPoint + new XYZ(-2.0, 0, 2.0);
                                }
                                else
                                {
                                    // Vertical wall: 2' right and 2' up  
                                    tagPosition = outletPlacementPoint + new XYZ(2.0, 0, 2.0);
                                }

                                // Create the tag for the sprinkler outlet
                                IndependentTag sprinklerTag = IndependentTag.Create(curDoc, curDoc.ActiveView.Id, new Reference(sprinklerOutlet), true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, tagPosition);

                                // Set the leader to free end
                                sprinklerTag.LeaderEndCondition = LeaderEndCondition.Free;

                                // Success - sprinkler outlet added successfully
                                Utils.TaskDialogInformation("Complete", "Spec Conversion", "Sprinkler outlet added successfully.");
                            }
                        }
                    }
                }

                t.Commit();
            }
            
            // notify user

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            clsButtonData myButtonData = new clsButtonData(
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
