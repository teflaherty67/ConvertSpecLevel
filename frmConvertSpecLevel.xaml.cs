using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
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
        public object SelectedOutlet { get; set; }
        public object SelectedRefSpWall { get; set; }
        public object SelectedRefSp { get; set; }
        public object SelectedOutletWall { get; set; }
        public object SelectedGarageWall { get; set; }

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

        private void SelectCabinet()
        {
            try
            {
                // Hide the form so user can see Revit
                this.Hide();

                // prompt the user to select the cabinet to delete
                SelectedCabinet = UIDoc.Selection.PickObject(ObjectType.Element, new CabinetSelectionFilter(), "Select cabinet to remove");

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

                // TODO: Implement actual Revit outlet selection logic here
                System.Threading.Thread.Sleep(1000);

                btnDynamicRow2.Content = "Selected";
                btnDynamicRow2.IsEnabled = true;

                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Error selecting outlet: {ex.Message}", "Error");
            }
        }

        private void SelectRefSpWall()
        {
            try
            {
                this.Hide();

                // TODO: Implement actual Revit wall selection logic here
                System.Threading.Thread.Sleep(1000);

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

                // TODO: Implement actual Revit wall selection logic here
                System.Threading.Thread.Sleep(1000);

                btnDynamicRow2.Content = "Selected";
                btnDynamicRow2.IsEnabled = true;

                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Error selecting wall: {ex.Message}", "Error");
            }
        }

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
}
