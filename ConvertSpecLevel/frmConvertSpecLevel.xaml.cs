using ConvertSpecLevel.Common;
using System;
using System;
using System.Collections.Generic;
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
using static ConvertSpecLevel.cmdRefSp;

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

        // properties for form return values
        public string SelectedClientName { get; private set; }
        public bool IsCompleteHomePlus { get; private set; }

        // properties for selected elements
        public Reference SelectedOutlet { get; set; }        
        public Reference SelectedOutletWall { get; set; }
        public Reference SelectedGarageWall { get; set; }



        #region Constructors

        // Constructor 1: initial form - no selections
        public frmConvertSpecLevel(Document curDoc, UIDocument uidoc)
        {
            InitializeComponent();

            UIDoc = uidoc;

            // intitilize form with populated combo boxes
            InitializeForm();
        }

        #endregion

        #region Constructor Helpers

        // call methods to populate comboboxes and restore form values
        private void InitializeForm()
        {
            PopulateComboBoxes();
            //RestoreFormValues(); // Restore any previously set values
            //UpdateDynamicContent();
        }

        private void PopulateComboBoxes()
        {
            // Create a list of LGI division clients
            List<string> listClients = new List<string> { "Central Texas", "Dallas/Ft Worth",
                "Florida", "Houston", "Maryland", "Minnesota", "Oklahoma", "Pennsylvania",
                "Southeast", "Virginia", "West Virginia" };

            // Add each client to the combobox
            foreach (string client in listClients)
            {
                cmbClient.Items.Add(client);
            }

            // Set the default selection to the first client in the list (Central Texas)
            if (cmbClient.Items.Count > 0)
                cmbClient.SelectedIndex = 0;
        }

        #endregion

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

        #endregion

        #region Dynamic Controls

        private void SpecLevel_Changed(object sender, RoutedEventArgs e)
        {
            // UpdateDynamicContent();
        }

        private void UpdateDynamicContent()
        {
            // Add null check to prevent errors during initialization
            if (spDynamicRow == null)
                return;

            if (rbCompleteHome.IsChecked == true)
            {
                Label row1Label = spDynamicRow.Children[0] as Label;
                row1Label.Content = "Select outlet to remove:";
                btnDynamicRow.Content = "Select";
            }
            else
            {
                Label row1Label = spDynamicRow.Children[0] as Label;
                row1Label.Content = "Select wall for sprinkler outlet:";
                btnDynamicRow.Content = "Select";
            }
        }

        private void btnDynamicRow_Click(object sender, RoutedEventArgs e)
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
                    SelectedOutlet = UIDoc.Selection.PickObject(ObjectType.Element, new OutletSelectionFilter(), "Select sprinkler outlet to remove");
                }
                else
                {
                    // notify the user that no Electrical views were found
                    Utils.TaskDialogWarning("Warning", "Spec Conversion", "No Electrical views found associated with the First Floor.");
                }

                btnDynamicRow.Content = "Selected";
                btnDynamicRow.IsEnabled = true;

                this.Show();
            }
            catch (Exception ex)
            {
                // notify the user of the error
                Utils.TaskDialogError("Error", "Spec Conversion", $"Error selecting outlet: {ex.Message}");
            }
        }

        private void SelectSprinklerWall()
        {
            try
            {
                this.Hide();

                // get the first electrical plan associated with First Floor level
                View planView = Utils.GetAllViewsByNameContainsAndAssociatedLevel(CurDoc, "Electrical", "First Floor").FirstOrDefault();

                // null check
                if (planView == null)
                {
                    // notify the user and exit
                    Utils.TaskDialogError("Error", "Sprinkler Outlet", "No First Floor electrical plan found.");
                    return;
                }

                // set it as the active view
                UIDoc.ActiveView = planView;

                // prompt the user to select the wall to host the sprinkler outlet
                Wall outletWall = Utils.SelectWall(UIDoc, "Select wall to host sprinkler outlet.");
                if (outletWall == null) return;

                // verify the selected element is a wall
                if (outletWall == null)
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
                // prompt the user to select the front garage wall
                Wall garageWall = Utils.SelectWall(UIDoc, "Select front garage wall.");
                if (garageWall == null) return;

                // verify the selected element is a wall
                if (garageWall == null)
                {
                    Utils.TaskDialogError("Error", "Spec Conversion", "Selected element is not a wall. Please try again.");
                    return;
                }

                btnDynamicRow.Content = "Selected";
                btnDynamicRow.IsEnabled = true;

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

        internal class OutletSelectionFilter : ISelectionFilter
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
}
