using ConvertSpecLevel.Classes;
using ConvertSpecLevel.Common;
using System.Windows.Input;

namespace ConvertSpecLevel
{
    [Transaction(TransactionMode.Manual)]
    public class cmdAdjustWindows : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // get all the window instances in the project
            List<FamilyInstance> allWindows = Utils.GetAllWindows(curDoc);

            // create a dictionary to hold the window data
            Dictionary<ElementId, clsWindowData> dictionaryWinData = new Dictionary<ElementId, clsWindowData>();

            // loop through the windows and get the data to store
            foreach (FamilyInstance curWindow in allWindows)
            {
                // store the data
                clsWindowData curData = new clsWindowData(curWindow);
                dictionaryWinData.Add(curWindow.Id, curData);
            }

            // get all the ViewSection views
            List<View> listViews = Utils.GetAllSectionViews(curDoc);

            // get the first view whose Title on Sheet is "Front Elevation"
            View elevFront = listViews
                .FirstOrDefault(v => v.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)?.AsString() == "Front Elevation");

            // set that view as the active view
            if (elevFront != null)
            {
                uidoc.ActiveView = elevFront;
            }
            else
            {
                Utils.TaskDialogInformation("Information", "Spec Conversion", "Front Elevation view not found. Proceeding with level adjustments in current view.");
            }

            // launch the form
            frmAdjustWindows curForm = new frmAdjustWindows();
            curForm.Topmost = true;

            // check if the user clicks OK
            if (curForm.ShowDialog() == true)
            {
                // get the selected spec level from the form
                string selectedSpecLevel = curForm.GetSelectedSpecLevel();

                // check if adjust head heights is checked
                bool adjustHeadHeights = curForm.IsAdjustWindowHeadHeightsChecked();

                // check is adjust window heights is checked
                bool adjustWindowHeights = curForm.IsAdjustWindowHeightsChecked();

                // test for raising or lowering windows
                bool raiseWindows = (selectedSpecLevel == "Complete Home Plus");

                // create counter for windows changed
                int countWindows = 0;

                // create a list for windows skipped
                List<string> skippedWindows = new List<string>();

                #region Adjust Head Heights

                // execute this code if adjust head heights is checked
                if (adjustHeadHeights)
                {
                    // create and start a transaction
                    using (Transaction t = new Transaction(curDoc, "Adjust Window Head Heights"))
                    {
                        t.Start();

                        foreach (var kvp in dictionaryWinData)
                        {
                            clsWindowData curData = kvp.Value;
                            double plateAdjustment = 1.0;
                            double newHeadHeight;

                            if (!raiseWindows)
                            {
                                // lower window head heights by 12"
                                newHeadHeight = curData.CurHeadHeight - plateAdjustment;
                            }
                            else
                            {
                                // raise window head height by by 12"
                                newHeadHeight = curData.CurHeadHeight + plateAdjustment;
                            }

                            if (curData.HeadHeightParam != null && !curData.HeadHeightParam.IsReadOnly)
                            {
                                // adjust the head heihgt
                                curData.HeadHeightParam.Set(newHeadHeight);

                                // increment the counter
                                countWindows++;
                            }
                        }

                        t.Commit();
                    }

                    // notify user of results
                    Utils.TaskDialogInformation("Information", "Spec Conversion",
                        $"Adjusted head heights for {countWindows} windows per the selected spec level.");
                }

                #endregion

                #region Adjust Head Height & Window Height

                // execute this code if both boxes are checked
                if (adjustHeadHeights && adjustWindowHeights)
                {
                    // create and start a transaction
                    using (Transaction t = new Transaction(curDoc, "Adjust Window Head Heights & Window Heights"))
                    {
                        t.Start();

                        foreach (var kvp in dictionaryWinData)
                        {
                            clsWindowData curData = kvp.Value;
                            double plateAdjustment = 1.0;
                            double newHeadHeight;

                            if (!raiseWindows)
                            {
                                // lower window head heights by 12"
                                newHeadHeight = curData.CurHeadHeight - plateAdjustment;
                            }
                            else
                            {
                                // raise window head height by by 12"
                                newHeadHeight = curData.CurHeadHeight + plateAdjustment;
                            }

                            if (curData.HeadHeightParam != null && !curData.HeadHeightParam.IsReadOnly)
                            {
                                // adjust the head heihgt
                                curData.HeadHeightParam.Set(newHeadHeight);

                                // increment the counter
                                countWindows++;

                                // adjust window heights
                                AdjustWindowHeights(curDoc, curData, plateAdjustment, raiseWindows, skippedWindows);
                            }
                        }

                        t.Commit();
                    }
                }

                #endregion

                if (!adjustHeadHeights)
                {
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "No window head height adjustments selected.");
                }

                return Result.Succeeded;
            }
            else
            {
                return Result.Cancelled;
            }
        }

        private void AdjustWindowHeights(Document curDoc, clsWindowData curData, double plateAdjustment, bool raiseWindows, List<string> skippedWindows)
        {
            // get the current family
            Family curFam = curData.WindowInstance.Symbol.Family;

            // get the current window instance type name
            string curTypeName = curData.WindowInstance.Symbol.Name;

            // split the Type Name into parts
            string[] stringParts = curTypeName.Split(' ');
            string sizePart = stringParts[0];

            // store the width & mull indicator if present
            string wndwPrefix = sizePart.Substring(0, sizePart.Length - 2);

            // get the current window height
            string wndwHeight = sizePart.Substring(sizePart.Length - 2);

            // change the string to an interger
            int curHeight = int.Parse(wndwHeight);

            // create variable for new height
            string newHeightPart;

            if (raiseWindows)
            {
                // set the new height number
                int newHeight = curHeight + 10;
                
                // convert to a string
                newHeightPart = newHeight.ToString();
            }
            else
            {
                // set the new height number
                int newHeight = curHeight - 10;

                // convert to a string
                newHeightPart = newHeight.ToString();
            }

            // set the new type name
            string newTypeName = wndwPrefix + newHeightPart + " " + string.Join(" ", stringParts.Skip(1));

            // get all the tpyes from the family
            foreach (ElementId curTypeId in curFam.GetFamilySymbolIds())
            {
                // find the correct type
                FamilySymbol curFamType = curDoc.GetElement(curTypeId) as FamilySymbol;
                string typeName = curFamType.Name;

                // compare type names
                if (typeName == newTypeName)
                {
                    // if match found, change the type
                    curData.WindowInstance.ChangeTypeId(curFamType.Id);
                }
                else
                {
                    // if not found, add to list of skipped windows
                }
            }

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