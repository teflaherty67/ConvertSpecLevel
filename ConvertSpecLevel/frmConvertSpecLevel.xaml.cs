using ConvertSpecLevel.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using static ConvertSpecLevel.cmdRefSp;

namespace ConvertSpecLevel
{
    /// <summary>
    /// Interaction logic for frmConvertSpecLevel.xaml
    /// </summary>
    public partial class frmConvertSpecLevel : Window
    {
        // What selection the form is requesting when it closes without DialogResult
        public enum SelectionRequest { None, Outlet, Walls }

        // Properties to hold Revit references
        public UIDocument UIDoc { get; set; }
        public Document CurDoc { get; set; }

        // Set by the command before ShowDialog() to restore state on re-opens
        public int InitialClientIndex { get; set; } = 0;
        public bool InitialIsCompleteHomePlus { get; set; } = false;

        // Set by the command before ShowDialog() to show "Selected" on the button
        public bool ShowOutletAsSelected { get; set; } = false;
        public bool ShowWallsAsSelected { get; set; } = false;

        // Read by the command after ShowDialog() returns null
        public SelectionRequest PendingSelection { get; private set; } = SelectionRequest.None;

        #region Constructor

        public frmConvertSpecLevel(Document curDoc, UIDocument uiDoc)
        {
            InitializeComponent();

            CurDoc = curDoc;
            UIDoc = uiDoc;

            InitializeForm();
        }

        #endregion

        #region Constructor Helpers

        private void InitializeForm()
        {
            PopulateComboBoxes();

            // Restore client selection
            if (InitialClientIndex >= 0 && InitialClientIndex < cmbClient.Items.Count)
                cmbClient.SelectedIndex = InitialClientIndex;

            // Restore spec level selection
            if (InitialIsCompleteHomePlus)
                rbCompleteHomePlus.IsChecked = true;
            else
                rbCompleteHome.IsChecked = true;

            // Show button as "Selected" if the command has already captured a selection
            if (ShowOutletAsSelected || ShowWallsAsSelected)
                btnDynamicRow.Content = "Selected";
        }

        private void PopulateComboBoxes()
        {
            List<string> listClients = new List<string> { "Central Texas", "Dallas/Ft Worth",
                "Florida", "Houston", "Maryland", "Minnesota", "Oklahoma", "Pennsylvania",
                "Southeast", "Virginia", "West Virginia" };

            foreach (string client in listClients)
                cmbClient.Items.Add(client);

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
                return rbCompleteHome.Content.ToString();
            else
                return rbCompleteHomePlus.Content.ToString();
        }

        public int GetClientIndex()
        {
            return cmbClient.SelectedIndex;
        }

        public bool GetIsCompleteHomePlus()
        {
            return rbCompleteHomePlus.IsChecked == true;
        }

        #endregion

        #region Dynamic Controls

        private void SpecLevel_Changed(object sender, RoutedEventArgs e)
        {
            if (btnDynamicRow == null)
                return;

            // Reset the button when spec level changes so the user re-selects for the new mode
            btnDynamicRow.Content = "Select";
        }

        private void btnDynamicRow_Click(object sender, RoutedEventArgs e)
        {
            // Signal to the command what kind of selection is needed, then close so
            // Revit's modal loop is released and PickObject can receive user input.
            PendingSelection = rbCompleteHome.IsChecked == true
                ? SelectionRequest.Outlet
                : SelectionRequest.Walls;

            this.Close(); // no DialogResult — ShowDialog() returns null
        }

        #endregion

        #region Buttons Section

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            string specLevel = GetSelectedSpecLevel();

            if (specLevel == "Complete Home" && !ShowOutletAsSelected)
            {
                MessageBox.Show("Please select the sprinkler outlet to remove before continuing.",
                    "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (specLevel == "Complete Home Plus" && !ShowWallsAsSelected)
            {
                MessageBox.Show("Please select both the sprinkler outlet wall and the front garage wall before continuing.",
                    "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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
                string helpUrl = "https://lifestyle-usa-design.atlassian.net/wiki/spaces/MFS/pages/472711169/Spec+Level+Conversion?atlOrigin=eyJpIjoiMmU4MzM3NzFmY2NlNDdiNjk1MjY2M2MyYzZkMjY2YWQiLCJwIjoiYyJ9";
                Process.Start(new ProcessStartInfo
                {
                    FileName = helpUrl,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while trying to display help: " + ex.Message,
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Selection Filters

        internal class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Wall;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        internal class OutletSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category != null && elem.Category.Name == "Electrical Fixtures")
                {
                    if (elem is FamilyInstance familyInstance)
                    {
                        string typeName = familyInstance.Symbol.Name;
                        return typeName.Contains("Outlet", StringComparison.OrdinalIgnoreCase) ||
                               typeName.Contains("Sprinkler", StringComparison.OrdinalIgnoreCase);
                    }
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return true;
            }
        }

        #endregion
    }
}
