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

            // get user input from the form

            // create a transaction group

            // start the transaction group

            #region Flooring Updates

            // get the first floor annotation view & set it as the active view

            // if not found alert the user

            // create transaction for flooring update

            // start the transaction

            // change the flooring for the specified rooms per the slected spec level

            // create a list of the rooms updated

            // notify user of results

            // commit the transaction

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
