using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdConvertSpecLevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // launch the form
            frmConvertSpecLevel curForm = new frmConvertSpecLevel()
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
            string selectedClient = curForm.GetSelectedClient();
            string selectedSpecLevel = curForm.GetSelectedSpecLevel();
            string selectedMWCabHeight = curForm.GetSelectedMWCabHeight();

            // create a transaction group
            using (TransactionGroup transGroup = new TransactionGroup(curDoc, "Convert Spec Level"))
            {
                // start the transaction group
                transGroup.Start();

                #region Floor Finish Updates

                // get the first floor annotation view & set it as the active view
                View curView = Utils.GetViewByNameContainsAndAssociatedLevel(curDoc, "Annotation", "First Floor");

                if (curView != null)
                {
                    uidoc.ActiveView = curView;
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No view found with name containing 'Annotation' and associated with 'First Floor'");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for flooring update
                using (Transaction t = new Transaction(curDoc, "Update Floor Finish"))
                {
                    // start the first transaction
                    t.Start();

                    // change the flooring for the specified rooms per the selected spec level
                    List<string> listUpdatedRooms = Utils.UpdateFloorFinishInActiveView(curDoc, selectedSpecLevel);

                    // create a list of the rooms updated
                    string listRooms;
                    if (listUpdatedRooms.Count == 1)
                    {
                        listRooms = listUpdatedRooms[0];
                    }
                    else if (listUpdatedRooms.Count == 2)
                    {
                        listRooms = $"{listUpdatedRooms[0]} and {listUpdatedRooms[1]}";
                    }
                    else
                    {
                        listRooms = string.Join(", ", listUpdatedRooms.Take(listUpdatedRooms.Count - 1)) + $", and {listUpdatedRooms.Last()}";
                    }

                    // notify the user
                    TaskDialog tdFloorUpdate = new TaskDialog("Complete");
                    tdFloorUpdate.MainIcon = Icon.TaskDialogIconInformation;
                    tdFloorUpdate.Title = "Spec Conversion";
                    tdFloorUpdate.TitleAutoPrefix = false;
                    tdFloorUpdate.MainContent = $"Flooring was changed at {listRooms} per the specified spec level.";
                    tdFloorUpdate.CommonButtons = TaskDialogCommonButtons.Close;

                    TaskDialogResult tdFloorUpdateSuccess = tdFloorUpdate.Show();

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Door Updates

                // get the door schedule & set it as the active view
                View curSched = Utils.GetScheduleByNameContains(curDoc, "Door Schedule");

                if (curSched != null)
                {
                    uidoc.ActiveView = curSched;
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Door Schedule found");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for door updates
                using (Transaction t = new Transaction(curDoc, "Update Doors"))
                {
                    // start the second transaction
                    t.Start();

                    // update front door type
                    Utils.UpdateFrontDoorType(curDoc, selectedSpecLevel);

                    // update rear door type
                    Utils.UpdateRearDoorType(curDoc, selectedSpecLevel);

                    // notify the user
                    TaskDialog tdDrUpdate = new TaskDialog("Complete");
                    tdDrUpdate.MainIcon = Icon.TaskDialogIconInformation;
                    tdDrUpdate.Title = "Spec Conversion";
                    tdDrUpdate.TitleAutoPrefix = false;
                    tdDrUpdate.MainContent = "The front and rear doors were replaced per the specified spec level.";
                    tdDrUpdate.CommonButtons = TaskDialogCommonButtons.Close;

                    TaskDialogResult tdDrUpdateSuccess = tdDrUpdate.Show();

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Cabinet Updates

                // get the Interior Elevations sheet & set it as the active view
                ViewSheet sheetIntr = Utils.GetViewSheetByName(curDoc, "Interior Elevations");

                if (sheetIntr != null)
                {
                    uidoc.ActiveView = sheetIntr;
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Interior Elevation sheet found");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for cabinet updates
                using (Transaction t = new Transaction(curDoc, "Update Cabinets"))
                {
                    // start the third transaction
                    t.Start();

                    // revise the upper cabinets

                    // revise the MW cabinet

                    // add/remove the Ref Sp cabinet

                    // raise/lower the backsplash height

                    // notify the user
                    // upper cabinets were revised per the selected spec level
                    // Ref Sp cabinet was added/removed per the selected spec level
                    // backsplash height was raised/lowered per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region First Floor Electrical Updates

                // get all views with Electrical in the name & associated with the First Floor
                List<View> firstFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "First Floor");

                // check for Second Floor Electrical Plan views

                // if found
                // string nameView = "First Floor Electrical Plan";

                // if not found
                // string nameView = "Electrical Plan";

                // get the first view in the list and set it as the active view
                if (firstFloorElecViews.Any())
                {
                    uidoc.ActiveView = firstFloorElecViews.First();
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Electrical views found for First Floor");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for first floor electrical updates
                using (Transaction t = new Transaction(curDoc, "Update First Floor Electrical"))
                {
                    // start the fourth transaction
                    t.Start();

                    // replace the light fixtures in the specified rooms per the selected spec level
                    Utils.UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                    // make a list of the rooms that were updated

                    // add/remove the sprinkler outlet in the Garage

                    // loop through all the views
                    foreach (View curElecView in firstFloorElecViews)
                    {
                        // add/remove ceiling fan note

                        // add/remove sprinkler outlet note                        
                    }

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {View Name} per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                #region Second Floor Electrical Updates

                // get all views with Electrical in the name & associated with the Second Floor
                List<View> secondFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(curDoc, "Electrical", "Second Floor");

                // set the view name variable
                string nameView = "Second Floor Electrical Plan";

                // get the first view in the list and set it as the active view
                if (secondFloorElecViews.Any())
                {
                    uidoc.ActiveView = secondFloorElecViews.First();
                }
                else
                {
                    // if not found alert the user
                    TaskDialog.Show("Error", "No Electrical views found for Second Floor");
                    transGroup.RollBack();
                    return Result.Failed;
                }

                // create transaction for Second Floor Electrical updates
                using (Transaction t = new Transaction(curDoc, "Update Second Floor Electrical"))
                {

                    // start the transaction
                    t.Start();

                    // replace the light fixtures in the specified rooms per the selected spec level
                    Utils.UpdateLightingFixturesInActiveView(curDoc, selectedSpecLevel);

                    // loop through the views and add/remove the clg fan note
                    foreach (View curElecView in secondFloorElecViews)
                    {
                        // add/remove ceiling fan note                        
                    }

                    // make a list of the rooms that were updated            

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {nameView} per the selected spec level

                    // commit the transaction
                    t.Commit();
                }

                #endregion

                // commit the transaction group
                transGroup.Assimilate(); // this will commit all the transactions in the group
            }

            return Result.Succeeded;

            // notify user conversion successful
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
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
