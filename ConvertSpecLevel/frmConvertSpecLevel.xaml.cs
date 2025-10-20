using System;
using System.Collections.Generic;
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
        // properties for form return values
        public string SelectedClientName { get; private set; }
        public bool IsCompleteHomePlus { get; private set; }
        public Reference SelectedOutlet { get; set; }
        public Reference SelectedOutletWall { get; set; }
        public Reference SelectedGarageWall { get; set; }


        #region Constructors

        // Constructor 1: initial form - no selections
        public frmConvertSpecLevel()
        {
            InitializeComponent();

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
    }
}
