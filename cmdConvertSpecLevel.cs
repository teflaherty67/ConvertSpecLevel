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

                    // if not found alert the user

                    // create transaction for door updates

                    // start the transaction

                    // update the front door

                    // update the rear door

                    // notify user of results

                    // commit the transaction

                    #endregion

                    #region Cabinet Updates

                    // get the Interior Elevations & set it as the active view

                    // if not found alert the user

                    // create transaction for cabinet updates

                    // start the transaction

                    // revise the upper cabinets

                    // revise the MW cabinet

                    // add/remove the Ref Sp cabinet

                    // raise/lower the backsplash height

                    // notify the user
                    // upper cabinets were revised per the selected spec level
                    // Ref Sp cabinet was added/removed per the selected spec level
                    // backsplash height was raised/lowered per the selected spec level

                    // commit the transaction

                    #endregion

                    #region First Floor Electrical Updates

                    // get all the First Floor Electrical views

                    // get the first view in the list and set it as the active view

                    // if not found alert the user

                    // create transaction for first floor electrical updates

                    // start the transaction

                    // replace the light fixtures in the specified rooms per the selected spec level

                    // add/remove the sprinkler outlet in the Garage

                    // loop through the views and add/remove the clg fan note & sprinkler note

                    // make a list of the rooms that were updated            

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {curView.Name} per the selected spec level

                    // commit the transaction

                    #endregion

                    #region Second Floor Electrical Updates

                    // get all the Second Floor Electrical views

                    // get the first view in the list and set it as the active view

                    // if not found notify the user

                    // create transaction for Second Floor Electrical updates

                    // start the transaction

                    // replace the light fixtures in the specified rooms per the selected spec level          

                    // loop through the views and add/remove the clg fan note

                    // make a list of the rooms that were updated            

                    // notify the user
                    // Lighting fixtures were replaced at {listRooms} at {curView.Name} per the selected spec level

                    // commit the transaction

                    #endregion

                    // end the transaction group using Assimilate
            }

                


                return Result.Succeeded;
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
