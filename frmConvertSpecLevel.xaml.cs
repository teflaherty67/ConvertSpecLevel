using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ConvertSpecLevel.Common;

namespace ConvertSpecLevel
{
    /// <summary>
    /// Interaction logic for frmConvertSpecLevel.xaml
    /// </summary>
    public partial class frmConvertSpecLevel : Window
    {
        // Properties to hold Revit references
        public UIDocument UIDoc { get; set; }
        public Document CurDoc { get; set; }

        // properties for selected elements
        public object SelectedCabinet { get; set; }
        public Reference SelectedOutlet { get; set; }
        public object SelectedRefSpWall { get; set; }
        public Line SelectedRefSp { get; set; }
        public Reference SelectedOutletWall { get; set; }
        public Reference SelectedGarageWall { get; set; }

        public frmConvertSpecLevel(UIDocument uidoc, Document curDoc)
        {
            InitializeComponent();

            // store the UIDocument and Document references
            UIDoc = uidoc;
            CurDoc = curDoc;

            // create a list of LGI division clients
            List<string> listClients = new List<string> { "Central Texas", "Dallas/Ft Worth",
                "Florida", "Houston", "Maryland", "Minnesota", "Oklahoma", "Pennsylvania",
                "Southeast", "Virginia", "West Virginia" };

            // add each client to the combobox
            foreach (string client in listClients)
            {
                cmbClient.Items.Add(client);
            }

            // set the default selection to the first client in the list (Central Texas)
            cmbClient.SelectedIndex = 0;

            // create a list of MW cabinet heights
            List<string> listMWCabinets = new List<string> { "18\"", "21\"", "24\"", "27\"", "30\"" };

            // add each client to the comboxbox
            foreach (string height in listMWCabinets)
            {
                cmbMWCabHeight.Items.Add(height);
            }

            // set the default selection to the first height in the list
            cmbMWCabHeight.SelectedIndex = 0;

            // intialize dynamic content based on selected spec level
            UpdateDynamicContent();
        }

        #region Form Controls

        public string GetSelectedClient()
        {
            return cmbClient.SelectedItem as string;
        }

        public string GetSelectedSpecLevel()
        {
            if (rbCompleteHome.IsChecked == true)
            {
                return rbCompleteHome.Content.ToString();
            }
            else
            {
                return rbCompleteHomePlus.Content.ToString();
            }
        }

        public string GetSelectedMWCabHeight()
        {
            return cmbMWCabHeight.SelectedItem as string;
        }

        #endregion

        #region Dynamic Controls

        private void SpecLevel_Changed(object sender, RoutedEventArgs e)
        {
            UpdateDynamicContent();
        }

        private void UpdateDynamicContent()
        {
            // Add null check to prevent errors during initialization
            if (spDynamicRow1 == null || spDynamicRow2 == null)
                return;

            if (rbCompleteHome.IsChecked == true)
            {
                // Complete Home content
                Label row1Label = spDynamicRow1.Children[0] as Label;
                row1Label.Content = "Select cabinet to remove:";
                btnDynamicRow1.Content = "Select";

                Label row2Label = spDynamicRow2.Children[0] as Label;
                row2Label.Content = "Select outlet to remove:";
                btnDynamicRow2.Content = "Select";
            }
            else
            {
                // Complete Home Plus content
                Label row1Label = spDynamicRow1.Children[0] as Label;
                row1Label.Content = "Select wall behind Ref Sp:";
                btnDynamicRow1.Content = "Select";

                Label row2Label = spDynamicRow2.Children[0] as Label;
                row2Label.Content = "Select wall for sprinkler outlet:";
                btnDynamicRow2.Content = "Select";
            }
        }

        private void btnDynamicRow1_Click(object sender, RoutedEventArgs e)
        {
            if (rbCompleteHome.IsChecked == true)
            {
                // Handle cabinet selection for Complete Home
                SelectCabinet();
            }
            else
            {
                // Handle wall selection for Complete Home Plus
                SelectRefSpWall();
            }
        }

        private void btnDynamicRow2_Click(object sender, RoutedEventArgs e)
        {
            if (rbCompleteHome.IsChecked == true)
            {
                // Handle outlet selection for Complete Home
                SelectOutlet();
            }
            else
            {
                // Handle sprinkler wall selection for Complete Home Plus
                SelectSprinklerWall();
            }
        }

        #endregion

        #region Selection Methods

        private void SelectCabinet()
        {
            try
            {
                // Hide the form so user can see Revit
                this.Hide();

                // get all views with Annotation in the name & associated with the First Floor
                List<View> firstFloorAnnoViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(CurDoc, "Annotation", "First Floor");

                // get the first view in the list
                if (firstFloorAnnoViews.Any())
                {
                    // set that view as the active view
                    UIDoc.ActiveView = firstFloorAnnoViews.First();

                    // prompt the user to select the outlet to delete
                    SelectedOutlet = UIDoc.Selection.PickObject(ObjectType.Element, new CabinetSelectionFilter(), "Select cabinet to remove");
                }
                else
                {
                    // notify the user that no Electrical views were found
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No Annotation views found associated with the First Floor.");
                }

                // Update button text and appearance
                btnDynamicRow1.Content = "Selected";
                btnDynamicRow1.IsEnabled = true; // Keep enabled so they can reselect if needed

                // Show the form again
                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Error selecting cabinet: {ex.Message}", "Error");
            }
        }

        private void SelectOutlet()
        {
            try
            {
                this.Hide();                

                // get all views with Electrical in the name & associated with the First Floor
                List<View> firstFloorElecViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(CurDoc, "Electrical", "First Floor");

                // get the first view in the list
                if (firstFloorElecViews.Any())
                {
                    // set that view as the active view
                    UIDoc.ActiveView = firstFloorElecViews.First();
                
                    // prompt the user to select the outlet to delete
                    SelectedOutlet = UIDoc.Selection.PickObject(ObjectType.Element, new OutlettSelectionFilter(), "Select sprinkler outlet to remove");
                }
                else
                {
                    // notify the user that no Electrical views were found
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No Electrical views found associated with the First Floor.");
                }

                btnDynamicRow2.Content = "Selected";
                btnDynamicRow2.IsEnabled = true;

                this.Show();
            }
            catch (Exception ex)
            {
                // notify the user of the error
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error selecting outlet: {ex.Message}");
            }
        }

        private void SelectRefSpWall()
        {
            try
            {
                this.Hide();
                // get all views with Annotation in the name & associated with the First Floor
                List<View> firstFloorAnnoViews = Utils.GetAllViewsByNameContainsAndAssociatedLevel(CurDoc, "Annotation", "First Floor");

                // get the first view in the list
                if (firstFloorAnnoViews.Any())
                {
                    // set that view as the active view
                    UIDoc.ActiveView = firstFloorAnnoViews.First();

                    // prompt the user to select the outlet to delete
                    SelectedOutlet = UIDoc.Selection.PickObject(ObjectType.Element, new CabinetSelectionFilter(), "Select cabinet to remove");
                }
                else
                {
                    // notify the user that no Electrical views were found
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No Annotation views found associated with the First Floor.");
                }

                btnDynamicRow1.Content = "Selected";
                btnDynamicRow1.IsEnabled = true;

                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Error selecting wall: {ex.Message}", "Error");
            }
        }

        private void SelectSprinklerWall()
        {
            try
            {
                this.Hide();

                // check if the active view is a view plan type
                if (CurDoc.ActiveView.ViewType != ViewType.FloorPlan)
                {
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "Please switch to a floor plan view to select the wall for the sprinkler outlet.");
                    return;
                }

                // prompt the user to select the wall for the sprinkler outlet
                SelectedOutletWall = UIDoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select wall for sprinkler outlet.");

                // cast the selected element to a Wall
                Wall selectedWall = CurDoc.GetElement(SelectedOutletWall.ElementId) as Wall;

                // verify the selected element is a wall
                if (selectedWall == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                    this.Show();
                    return;
                }                

                SelectGarageFrontWall();
            }
            catch (OperationCanceledException)
            {
                // User pressed Esc or cancelled - just show the form again, no error message needed
                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error selecting wall: {ex.Message}");
            }
        }

        private void SelectGarageFrontWall()
        {
            try
            {
                // check if the active view is a view plan type
                if (CurDoc.ActiveView.ViewType != ViewType.FloorPlan)
                {
                    Utils.TaskDialogInformation("Information", "Spec Conversion", "Please switch to a floor plan view to select the front wall of the Garage.");
                    return;
                }

                // prompt the user to select the wall for the sprinkler outlet
                SelectedGarageWall = UIDoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select front wall of Garage.");

                // cast the selected element to a Wall
                Wall selectedWall = CurDoc.GetElement(SelectedGarageWall.ElementId) as Wall;

                // verify the selected element is a wall
                if (selectedWall == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                    return;
                }

                btnDynamicRow2.Content = "Selected";
                btnDynamicRow2.IsEnabled = true;

                // show the form again
                this.Show();
            }
            catch (OperationCanceledException)
            {
                // User pressed Esc or cancelled - just show the form again, no error message needed
                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error selecting wall: {ex.Message}");
            }
        }

        #endregion

        #region Buttons Section

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // launch the help site with user's default browser
                string helpUrl = "https://lifestyle-usa-design.atlassian.net/wiki/spaces/MFS/pages/472711169/Spec+Level+Conversion?atlOrigin=eyJpIjoiMmU4MzM3NzFmY2NlNDdiNjk1MjY2M2MyYzZkMjY2YWQiLCJwIjoiYyJ9";
                Process.Start(new ProcessStartInfo
                {
                    FileName = helpUrl,
                    UseShellExecute = true
                });

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("An error occurred while trying to display help: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }       

        #endregion
    }

    #region Selection Filters

    internal class WallSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Wall; // Allows only Wall elements to be selected
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            // For selecting walls, we are primarily interested in the element itself.
            // If you need to allow selection based on faces or edges of a wall,
            // you might adjust this method and potentially the PickObject/PickObjects call.
            return false;
        }
    }

    internal class CabinetSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Allow only cabinets
            return elem.Category != null && elem.Category.Name == "Casework";
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            // Allow all references
            return true;
        }
    }

    internal class OutlettSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Check if element is an electrical fixture
            if (elem.Category != null && elem.Category.Name == "Electrical Fixtures")
            {
                // Cast to FamilyInstance to access the Symbol
                if (elem is FamilyInstance familyInstance)
                {
                    // Check if the type name contains "Outlet" or "Sprinkler"
                    string typeName = familyInstance.Symbol.Name;
                    return typeName.Contains("Outlet", StringComparison.OrdinalIgnoreCase) ||
                           typeName.Contains("Sprinkler", StringComparison.OrdinalIgnoreCase);
                }
            }
            return false;
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            // Allow all references
            return true;
        }
    }

    #endregion
}
