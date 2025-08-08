using ConvertSpecLevel.Classes;
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

            #region Form

            // launch the form
            frmConvertSpecLevel curForm = new frmConvertSpecLevel();
            curForm.Topmost = true;

            curForm.ShowDialog();

            // check if user clicked Cancel
            if (curForm.DialogResult != true)
            {
                return Result.Cancelled;
            }

            // get user input from form
            // get user input from the form
            string selectedClient = curForm.GetSelectedClient();
            string selectedSpecLevel = curForm.GetSelectedSpecLevel();
            string selectedMWCabHeight = curForm.GetSelectedMWCabHeight();

            #endregion

            #region Transaction Group

            // create a transaction group
            using (TransactionGroup transgroup = new TransactionGroup(curDoc, "Convert Spec Level"))
            {
                // start the transaction group
                transgroup.Start();

                #region Floor Finish Updates

                // get a first floor annotation view & set it as the active view
                View curView = Utils.GetViewByNameContainsAndAssociatedLevel(curDoc, "Annotation", "First Floor");

                // check for null
                if (curView != null)
                {
                    uidoc.ActiveView = curView;
                }
                else
                {
                    // if null notify the user
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No view found with name containing 'Annotation' and associated level 'First Floor'");
                }

                // create a transaction for the flooring update
                using (Transaction t = new Transaction(curDoc, "Update Floor Finish"))
                {
                    // start the first transaction
                    t.Start();

                    // change the flooring for the specified rooms per the selected spec level
                    List<string> listUpdatedRooms = UpdateFloorFinishInActiveView(curDoc, selectedSpecLevel);

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
                    Utils.TaskDialogInformation("Complete", "Spec Conversion", $"Flooring was changed at {listRooms} per the specified spec level.");

                    // commit the transaction
                    t.Commit();
                }

                #endregion




                // commit the transaction group
                transgroup.Assimilate();
            }

            #endregion

            return Result.Succeeded;
        }

        private List<string> UpdateFloorFinishInActiveView(Document curDoc, string selectedSpecLevel)
        {
            throw new NotImplementedException();
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
