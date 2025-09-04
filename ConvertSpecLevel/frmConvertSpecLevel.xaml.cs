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
}
